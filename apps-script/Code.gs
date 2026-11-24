/**
 * Robbie v3.10 backend (Google Apps Script web app). Zero cost: no billing account anywhere.
 * Deploy: Execute as "Me", Who has access "Anyone".
 * Secrets live only in Script Properties (set by the one-time setup page): GEMINI_API_KEY, PASSPHRASE.
 * This file contains no keys, passphrases or spreadsheet IDs, so it can sit in a public repository.
 */

var DEFAULT_MODEL = 'gemini-3.5-flash-lite';
var DAILY_TURN_LIMIT = 300;      // Gemini calls for Robbie per London day
var DAILY_CONSULT_LIMIT = 20;    // Claude-to-Gemini review calls per London day
var MAX_AUDIO_B64 = 1500000;     // about 30 s of 16 kHz mono WAV
var MAX_CONSULT_CHARS = 200000;
var SHORT_LIMIT = 20;
var LONG_LIMIT = 80;
var PAGES_URL = 'https://tisantosuk.github.io/robbie/';
var DEFAULT_TTS_MODEL = 'gemini-2.5-flash-preview-tts';
var FALLBACK_MODELS = ['gemini-flash-lite-latest', 'gemini-3.1-flash-lite']; // conversation model when the main one is overloaded
var DEFAULT_TTS_VOICE = 'Aoede';
var TTS_STYLE_VERSION = 's2';
var TTS_BACKUP_MODELS = []; // v3.7 (F-5): one voice model for normal turns, for a consistent voice
var TTS_DAILY_BACKUP = 'gemini-3.1-flash-tts-preview'; // v3.8 (H-1b): only after the main model's DAILY limit (same voice name)    // change when the TTS style text below changes
var DAILY_BENCH_LIMIT = 80;      // Claude's latency/voice measurements per London day
var TTS_VOICES = ['Zephyr','Puck','Charon','Kore','Fenrir','Leda','Orus','Aoede','Callirrhoe','Autonoe','Enceladus','Iapetus','Umbriel','Algieba',
  'Despina','Erinome','Algenib','Rasalgethi','Laomedeia','Achernar','Alnilam','Schedar','Gacrux','Pulcherrima','Achird','Zubenelgenubi',
  'Vindemiatrix','Sadachbia','Sadaltager','Sulafat'];

// ---------------------------------------------------------------- entry points

function doGet(e) {
  var props = PropertiesService.getScriptProperties();
  if (e && e.parameter && e.parameter.health === '1') {
    return json_({ ok: true, configured: Boolean(props.getProperty('GEMINI_API_KEY')) });
  }
  return HtmlService.createHtmlOutput(setupHtml_(Boolean(props.getProperty('GEMINI_API_KEY'))))
    .setTitle('Robbie setup').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ error: 'bad request' }); }
  var props = PropertiesService.getScriptProperties();
  if (!passOk_(req.passphrase, props.getProperty('PASSPHRASE'))) return json_({ error: 'unauthorised' });
  try {
    switch (req.action) {
      case 'start': return json_(start_(req));
      case 'turn': return json_(turn_(req));
      case 'end': return json_(end_(req));
      case 'consult': return json_(consult_(req));
      case 'models': return json_(listModels_());
      case 'bench': return json_(bench_(req));       // Claude: measure one model setting on a recording
      case 'tts_test': return json_(ttsTest_(req));  // Claude: measure one voice setting
      case 'set_config': return json_(setConfig_(req)); // Claude: apply the agreed model/voice settings
      case 'diag': return json_({ calendar: calendarDigest_(req).status }); // Claude: calendar status, no event data
      default: return json_({ error: 'unknown action' });
    }
  } catch (err) {
    console.error(err && err.stack || err);
    return json_({ error: String(err && err.message || err) });
  }
}

// ---------------------------------------------------------------- setup (called from Setup.html)

function saveSetup(key) {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('GEMINI_API_KEY')) throw new Error('Já configurado.');
  key = String(key || '').trim();
  if (!/^[A-Za-z0-9_.\-]{20,200}$/.test(key)) throw new Error('A chave não parece válida.'); // AIza... and newer AQ.... keys; Google's own check follows
  var test = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=1000',
    { headers: { 'x-goog-api-key': key }, muteHttpExceptions: true });
  if (test.getResponseCode() !== 200) throw new Error('O Google recusou a chave (' + test.getResponseCode() + ').');
  var models = modelNames_(JSON.parse(test.getContentText()));
  var model = pickModel_(models); // Gemini review 1445, Defect 1: use a model this key can actually call
  if (!model) throw new Error('Nenhum modelo Flash disponível para esta chave.');
  if (model !== DEFAULT_MODEL) console.warn('DEFAULT_MODEL ' + DEFAULT_MODEL + ' not offered; using ' + model);

  var pass = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  var ss = memorySheet_(true);
  props.setProperties({ GEMINI_API_KEY: key, PASSPHRASE: pass, GEMINI_MODEL: model, SHEET_ID: ss.getId() });

  var url = ScriptApp.getService().getUrl();
  var phoneLink = PAGES_URL + '#u=' + encodeURIComponent(url) + '&p=' + encodeURIComponent(pass);
  return {
    phoneLink: phoneLink,
    claudeLine: 'ROBBIE_LINK ' + url + ' ' + pass,
    model: model,
    models: models,
  };
}

// Default if offered; otherwise the first stable flash-lite, then the first stable flash.
function pickModel_(names) {
  if (names.indexOf(DEFAULT_MODEL) >= 0) return DEFAULT_MODEL;
  var stable = names.filter(function (n) { return !/(preview|exp|tts|image|live|audio|thinking|embedding)/.test(n); });
  var lite = stable.filter(function (n) { return /flash-lite/.test(n); });
  return lite[0] || stable[0] || null;
}

// ---------------------------------------------------------------- Robbie actions

function start_(req) {
  var cfg = config_();
  var memory = '';
  var cache = CacheService.getScriptCache();
  memory = cache.get('memory_check');
  if (memory === null) {
    memory = String(memorySheet_(false).getSheetByName('Config').getRange('B1').getValue() || '');
    cache.put('memory_check', memory, 600);
  }
  var greeting = greetingText_(req && req.hour);
  var out = { greeting: greeting, memoryValue: memory, model: cfg.model };
  var g = greetingAudio_(cfg, greeting);
  var audio = g && !g.failed ? g : null;
  if (audio) { out.audio = audio.data; out.audioRate = audio.rate; out.audioFormat = audio.format; out.ttsModel = audio.model; }
  var gw = (g && g.why) || {};
  try { logTurn_(req || {}, { transcript: '', reason: 'start', speak: true, text: greeting, ttsModel: out.ttsModel || '',
    ttsFallback: audio ? '' : (gw.reason || 'off'), ttsDetail: gw.detail || '' }, null, ''); } catch (e) {}
  return out;
}

// Ti, 23 Nov: no self-introduction each time; "Oi Ti" + time of day + a casual check-in, varied.
var CHECK_INS = ['E aí, beleza?', 'Tudo certo?']; // M-3: 2 per period (6 greeting texts a day at most), to save the daily voice quota
function greetingText_(hour) {
  var h = Number(hour);
  if (!(h >= 0 && h <= 23) || hour === null || hour === undefined || hour === '') h = Number(Utilities.formatDate(new Date(), 'Europe/London', 'H'));
  var part = h >= 5 && h < 12 ? 'bom dia' : h >= 12 && h < 18 ? 'boa tarde' : 'boa noite';
  return 'Oi Ti, ' + part + '. ' + CHECK_INS[Math.floor(Math.random() * CHECK_INS.length)];
}

// v3.4: every turn is written to the Turns tab of Ti's memory sheet (diagnostics; DEBUG_LOG=false turns it off).
function turn_(req) {
  var holder = {};
  var out = turnCore_(req, holder);
  // Gemini review 1053, Defect 2: analysis passed explicitly; silent segments (no speech) are not written
  if (out.transcript || out.busy || out.error || out.closed) {
    try { logTurn_(req, out, holder.analysis || null, holder.calendar); } catch (e) { console.warn('turn log failed: ' + e); }
  }
  return out;
}

function turnCore_(req, holder) {
  var cfg = config_();
  var pending = (req.pending || []).map(String).slice(-20);
  var history = (req.history || []).slice(-20);
  if (!req.audio || String(req.audio).length > MAX_AUDIO_B64) return { error: 'audio missing or longer than 30 seconds' };

  if (!bump_('turns', DAILY_TURN_LIMIT)) {
    return { transcript: '', speak: true, closed: true, reason: 'daily_limit',
      text: 'Chegamos ao limite de hoje. A gente continua amanhã.', pending: [], history: history };
  }

  var t0 = Date.now();
  var r;
  var cal = calendarDigest_(req); // Increment 1: today's agenda, read-only, cached 5 min
  holder.calendar = cal.status;
  var proposal = cleanProposal_(req.proposal);
  var partialText = req.partial && typeof req.partial === 'object' ? JSON.stringify(stripEmpty_({ title: req.partial.title, date: req.partial.date, start: req.partial.start, end: req.partial.end, said: req.partial.said })) : '';
  try { r = analyseSegment_(cfg, req.audio, pending.join(' '), history, cal.text, proposal ? proposalSentence_(proposal) : '', partialText); holder.analysis = r; }
  catch (err) { // Gemini review 1703, Defect 1: a second 500/503 is not a failure on Ti's screen; the phone re-sends once
    if (/Gemini (500|503)/.test(String(err.message))) return { busy: true, speak: false, transcript: '', pending: pending, history: history };
    throw err;
  }
  var geminiMs = Date.now() - t0;
  var transcript = String(r.transcript || '').trim();
  if (!transcript) return { transcript: '', speak: false, reason: 'no-speech', pending: pending, history: history, geminiMs: geminiMs };

  if (isClosing(transcript)) {
    return { transcript: transcript, speak: true, closed: true, reason: 'voice', text: 'Tchau, Ti. Guardei nossa conversa.',
      pending: [], history: history, geminiMs: geminiMs };
  }

  // Increment 1b: create one event after a spoken confirmation
  var action = String(r.calendar_action || 'none');
  var partialIn = req.partial && typeof req.partial === 'object' ? req.partial : null;
  if (partialIn && action === 'none' && transcript) action = 'propose'; // H-4b: the next segment completes the request
  // v3.9 (K-1, K-2): Ti answering Robbie's own question, right after it was asked (first segment, within 90 s)
  var lastRobbie = history.length ? String(history[history.length - 1].robbie || '') : '';
  var lastTi = history.length ? String(history[history.length - 1].ti || '') : '';
  var since = Number(req.secondsSinceRobbie);
  var answering = pending.length === 0 && /\?\s*$/.test(lastRobbie) && since >= 0 && since <= 90 && isAnswer_(transcript);
  if (answering && action === 'none' && !proposal && !partialIn && isYes_(transcript)) {
    var offered = offerFrom_(lastRobbie, lastTi, phoneDay_(req).date);
    if (offered) { action = 'propose'; r.event = offered; } // K-2: "Quer que eu anote um lembrete...?" -> "Quero sim" -> propose
  }
  if (proposal || action === 'propose') {
    var ev = calendarStep_(req, proposal, action, r, transcript);
    if (ev) {
      holder.calendar = ev.log;
      history.push({ ti: pending.concat([transcript]).join(' '), robbie: ev.text });
      return withVoice_(cfg, { transcript: transcript, speak: true, reason: 'calendar-' + ev.kind, text: ev.text, proposal: ev.proposal || null,
        partial: ev.partial && !partialIn ? ev.partial : null, pending: [], history: history.slice(-20), geminiMs: geminiMs });
    }
  }

  var firstSegment = pending.length === 0;
  pending.push(transcript);
  var decision = r.urgent ? { complete: true, reason: 'urgent' } : isTurnComplete(transcript, r.addressed_question, firstSegment);
  if (!decision.complete && answering) decision = { complete: true, reason: 'answer' };
  if (!decision.complete) return { transcript: transcript, speak: false, reason: decision.reason, pending: pending, history: history, geminiMs: geminiMs };

  var full = pending.join(' ');
  var reply;
  if (r.urgent) {
    reply = String(r.reply || '').trim() || 'Ti, estou aqui com você. Você está seguro agora?';
    if (!/116 ?123/.test(reply)) reply += ' O Samaritans atende 24 horas no 116 123.';
    if (!/\b188\b/.test(reply)) reply += ' No Brasil, o CVV atende no 188.';
  } else {
    reply = capReply(r.reply, wantsExplanation(full) ? LONG_LIMIT : SHORT_LIMIT) || 'Hmm, me conta mais?';
  }
  history.push({ ti: full, robbie: reply });
  return withVoice_(cfg, { transcript: transcript, speak: true, reason: decision.reason, text: reply, pending: [], history: history.slice(-20), geminiMs: geminiMs });
}

function withVoice_(cfg, out) {
  var t1 = Date.now();
  var why = {};
  var audio = ttsSafe_(cfg, out.text, why); // V-1: same round trip; on any failure the phone uses its own voice
  if (audio) { out.audio = audio.data; out.audioRate = audio.rate; out.audioFormat = audio.format; out.ttsMs = Date.now() - t1; out.ttsModel = audio.model; }
  else { out.ttsFallback = why.reason || 'unknown'; if (why.detail) out.ttsDetail = why.detail; }
  return out;
}

// ---------------------------------------------------------------- Increment 1b: create events (Calendar Advanced Service, calendar.events scope)

var DEFAULT_REMINDERS = [10080, 1440, 180, 60, 30]; // Ti's rule: 1 week, 1 day, 3 h, 1 h, 30 min before (popup)
var DAILY_EVENT_LIMIT = 10;
var YES_WORDS = /\b(sim|pode|confirma|confirmo|isso|claro|ok|beleza|fechado|pode gravar|grava|manda)\b/;
var WEEKDAYS = ['domingo', 'segunda', 'terça', 'quarta', 'quinta', 'sexta', 'sábado'];

function stripEmpty_(o) { var out = {}; Object.keys(o).forEach(function (k) { if (o[k] !== '' && o[k] !== null && o[k] !== undefined) out[k] = o[k]; }); return out; }
function partialFrom_(ev, transcript, earlier, today) { // what Ti has said so far, kept on the phone for ONE turn
  var p = {};
  ['title', 'date', 'start', 'end'].forEach(function (k) { if (ev[k]) p[k] = String(ev[k]).slice(0, 80); });
  if (p.date && today) p.date = normaliseDate_(p.date, today) || p.date;
  p.said = (String((earlier || {}).said || '') + ' ' + transcript).trim().slice(-300);
  return p;
}

var ANSWER_WORDS = /^(sim|nao|quero|pode|claro|isso|ok|okay|beleza|com certeza|tambem|talvez|acho que sim|acho que nao|prefiro|bora|vamos)\b/;
function isAnswer_(t) { var n = normalise(t).replace(/\?/g, ' ').trim(); return n.split(' ').filter(Boolean).length <= 6 || ANSWER_WORDS.test(n); }
function isYes_(t) { var n = normalise(t).replace(/\?/g, ' ').trim(); return /^(sim|quero|pode|claro|isso|ok|okay|beleza|com certeza|bora|vamos|por favor)\b/.test(n) && !/\bnao\b/.test(n); }
// K-2: Robbie offered a reminder or event in his last reply; build the draft from the offer (the server asks for what is missing).
var OFFER = /quer que eu (?:anote|crie|marque|agende|coloque|faca|faça)(?: um| uma)? (?:lembrete|evento|compromisso)(?: (?:para|pra) (?:voce|você))?(?: (?:para|pra|de))? ?([^?]*)\?/i;
function offerFrom_(lastRobbie, lastTi, today) {
  var m = OFFER.exec(lastRobbie);
  if (!m) return null;
  var what = String(m[1] || '').replace(/(^|\s)(depois de amanhã|depois de amanha|amanhã|amanha|hoje)(?=\s|$)/gi, ' ').replace(/\s+/g, ' ').trim();
  var title = what ? what.charAt(0).toUpperCase() + what.slice(1) : 'Lembrete';
  return { title: title.slice(0, 80), date: normaliseDate_(lastRobbie, today) || normaliseDate_(lastTi, today) || '' };
}

function cleanProposal_(p) {
  if (!p || typeof p !== 'object') return null;
  var e = { title: String(p.title || '').replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, 80), date: String(p.date || ''),
    start: String(p.start || ''), end: String(p.end || '') };
  if (!e.title || !/^\d{4}-\d{2}-\d{2}$/.test(e.date) || !/^\d{2}:\d{2}$/.test(e.start) || !/^\d{2}:\d{2}$/.test(e.end)) return null;
  e.hash = hash_(e.title + '|' + e.date + '|' + e.start + '|' + e.end);
  if (p.hash && String(p.hash) !== e.hash) return null; // must be exactly what Robbie spoke
  return e;
}

// G-5: deterministic backup for the end time when the model leaves event.end empty.
var NUM_WORDS = { uma: 1, um: 1, duas: 2, dois: 2, tres: 3, quatro: 4, cinco: 5, seis: 6, sete: 7, oito: 8, nove: 9, dez: 10, onze: 11, doze: 12,
  meio: 12, 'meio dia': 12, meia: 30, quinze: 15, trinta: 30, quarenta: 40, cinquenta: 50, vinte: 20 };
function hm_(h, m) { return pad2_(h) + ':' + pad2_(m || 0); }
function parseClock_(txt) { // "10", "10:30", "10h30", "dez", "dez e meia"
  var m = /^(\d{1,2})(?:[: h](\d{2})\b)?/.exec(txt); // normalise() turns "10:30" into "10 30"
  if (m) return { h: Number(m[1]), m: Number(m[2] || 0) };
  var w = /^(meio dia|[a-z]+)(?: e (meia|quinze|trinta|quarenta|vinte))?/.exec(txt);
  if (w && NUM_WORDS[w[1]] !== undefined && NUM_WORDS[w[1]] <= 12) return { h: NUM_WORDS[w[1]], m: w[2] ? NUM_WORDS[w[2]] : 0 };
  return null;
}
function endFromTranscript_(transcript, start) {
  var t = normalise(transcript).replace(/\?/g, ' ');
  var sMin = Number(start.slice(0, 2)) * 60 + Number(start.slice(3, 5));
  var dur = /\bpor (\d+|uma|um|duas|meia) (hora|horas|minuto|minutos)\b/.exec(t);
  if (dur) {
    var n = /^\d+$/.test(dur[1]) ? Number(dur[1]) : dur[1] === 'meia' ? 0.5 : NUM_WORDS[dur[1]];
    var mins = /hora/.test(dur[2]) ? Math.round(n * 60) : n;
    return { said: true, end: sMin + mins < 24 * 60 ? hm_(Math.floor((sMin + mins) / 60), (sMin + mins) % 60) : null };
  }
  var NUMW = '(?:\\d{1,2}(?:[: h]\\d{2})?|uma|duas|tres|quatro|cinco|seis|sete|oito|nove|dez|onze|doze|meio dia)';
  var rng = new RegExp('\\b(?:das?|de) ' + NUMW + '(?: e (?:meia|quinze|trinta|quarenta))?(?: (?:da|de) (?:manha|tarde|noite))? (?:as|a|ate as|ate) (.+)$').exec(t) ||
    /\bate (?:as )?(.+)$/.exec(t);
  if (!rng) return { said: false, end: null };
  var c = parseClock_(rng[1].trim());
  if (!c || c.h > 23 || c.m > 59) return { said: true, end: null };
  var e = c.h * 60 + c.m;
  if (e <= sMin && e + 720 > sMin && e + 720 < 24 * 60) e += 720; // "das 2 às 4 da tarde"
  return { said: true, end: e > sMin ? hm_(Math.floor(e / 60), e % 60) : null };
}

// v3.8 (H-4a): spoken dates from the phone's own date; the model's value first, then the transcript.
var WEEKDAY_NAMES = ['domingo', 'segunda', 'terca', 'quarta', 'quinta', 'sexta', 'sabado'];
function normaliseDate_(value, today) {
  var v = String(value || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(v)) return v;
  var t = normalise(v).replace(/\?/g, ' ');
  if (/\bdepois de amanha\b/.test(t)) return addDays_(today, 2);
  if (/\bamanha\b/.test(t)) return addDays_(today, 1);
  if (/\bhoje\b/.test(t)) return today;
  var dm = /\b(\d{1,2})\s*\/\s*(\d{1,2})(?:\s*\/\s*(\d{2,4}))?\b/.exec(v) || /\bdia (\d{1,2})\b/.exec(t);
  if (dm) {
    var p = ymd_(today), year = p[0], month = dm[2] ? Number(dm[2]) : p[1] + 1, d = Number(dm[1]);
    if (dm[3]) year = Number(dm[3].length === 2 ? '20' + dm[3] : dm[3]);
    var cand = year + '-' + pad2_(month) + '-' + pad2_(d);
    if (!dm[3] && cand < today) cand = (dm[2] ? year + 1 : (month === 12 ? year + 1 : year)) + '-' + pad2_(dm[2] ? month : (month % 12) + 1) + '-' + pad2_(d);
    return cand;
  }
  for (var i = 0; i < 7; i++) {
    if (new RegExp('\\b' + WEEKDAY_NAMES[i] + '\\b').test(t)) {
      var q = ymd_(today), wd = new Date(Date.UTC(q[0], q[1], q[2])).getUTCDay();
      return addDays_(today, ((i - wd + 7) % 7) || 7);
    }
  }
  return '';
}

function proposalFrom_(event, day, transcript) {
  if (!event) return { error: 'missing' };
  var title = String(event.title || '').replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, 80);
  var date = normaliseDate_(event.date, day.date) || (transcript ? normaliseDate_(transcript, day.date) : '');
  var start = String(event.start || '').slice(0, 5), end = String(event.end || '').slice(0, 5);
  if (/^\d:\d\d$/.test(String(event.start || ''))) start = '0' + String(event.start).slice(0, 4);
  if (!title) return { error: 'title' };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return { error: 'date' };
  if (!/^([01]\d|2[0-3]):[0-5]\d$/.test(start)) return { error: 'start' };
  var toMin = function (hm) { var a = hm.split(':').map(Number); return a[0] * 60 + a[1]; };
  if (!/^([01]\d|2[0-3]):[0-5]\d$/.test(end) && transcript) {
    var fromText = endFromTranscript_(transcript, start);
    if (fromText.end) end = fromText.end;
    else if (fromText.said) return { error: 'end' }; // Ti said an end but it could not be read: ask, do not guess
  }
  if (!/^([01]\d|2[0-3]):[0-5]\d$/.test(end)) { var m = Math.min(toMin(start) + 30, 23 * 60 + 59); end = pad2_(Math.floor(m / 60)) + ':' + pad2_(m % 60); }
  var dur = toMin(end) - toMin(start); // end must be after start on the same day
  if (!(dur >= 5 && dur <= 240)) return { error: 'duration' };
  var daysAhead = (Date.UTC.apply(null, ymd_(date)) - Date.UTC.apply(null, ymd_(day.date))) / 86400000;
  if (!(daysAhead >= 0 && daysAhead <= 365)) return { error: 'date range' };
  var e = { title: title, date: date, start: start, end: end };
  e.hash = hash_(e.title + '|' + e.date + '|' + e.start + '|' + e.end);
  return { proposal: e, daysAhead: daysAhead };
}

function addDays_(date, n) { var p = ymd_(date); var d = new Date(Date.UTC(p[0], p[1], p[2]) + n * 86400000); return d.getUTCFullYear() + '-' + pad2_(d.getUTCMonth() + 1) + '-' + pad2_(d.getUTCDate()); }
function pad2_(n) { return (n < 10 ? '0' : '') + n; }
function ymd_(d) { var p = d.split('-').map(Number); return [p[0], p[1] - 1, p[2]]; }

function proposalSentence_(e) { return '"' + e.title + '", ' + spokenDay_(e.date) + ', das ' + e.start + ' às ' + e.end; }
function spokenDay_(date) {
  var p = ymd_(date), d = new Date(Date.UTC(p[0], p[1], p[2]));
  return WEEKDAYS[d.getUTCDay()] + ', ' + p[2] + '/' + pad2_(p[1] + 1);
}

function isoLocal_(date, hm, offset) { // offset = Date.getTimezoneOffset minutes; returns RFC 3339 with the phone's offset
  var sign = offset <= 0 ? '+' : '-', a = Math.abs(offset);
  return date + 'T' + hm + ':00' + sign + pad2_(Math.floor(a / 60)) + ':' + pad2_(a % 60);
}

// Returns null when the turn is not about the calendar after all (normal conversation continues).
function calendarStep_(req, proposal, action, r, transcript) {
  var day = phoneDay_(req);
  var plain = normalise(transcript);
  if (proposal && action === 'confirm' && YES_WORDS.test(plain)) {
    var cache = CacheService.getScriptCache();
    if (cache.get('created_' + proposal.hash)) {
      return { kind: 'duplicate', text: 'Esse evento já está no calendário.', log: 'duplicate: ' + proposal.title };
    }
    var recheck = proposalFrom_(proposal, day); // Gemini review 2255, Defect 1: same limits again right before writing
    if (recheck.error || recheck.proposal.hash !== proposal.hash) {
      return { kind: 'unclear', text: 'Me diz o dia e o horário do evento?', log: 'refused at confirm: ' + (recheck.error || 'changed') };
    }
    if (!bump_('events', DAILY_EVENT_LIMIT)) return { kind: 'limit', text: 'Já criei o limite de 10 eventos hoje.', log: 'limit' };
    try {
      var created = Calendar.Events.insert({
        summary: proposal.title,
        start: { dateTime: isoLocal_(proposal.date, proposal.start, day.offset) },
        end: { dateTime: isoLocal_(proposal.date, proposal.end, day.offset) },
        reminders: { useDefault: false, overrides: DEFAULT_REMINDERS.map(function (m) { return { method: 'popup', minutes: m }; }) },
      }, 'primary');
      cache.put('created_' + proposal.hash, String(created && created.id || '1'), 21600);
      try { cache.remove('cal_' + day.date + '_' + day.offset); } catch (e) {} // today's digest must show it
      return { kind: 'created', text: 'Pronto, criei ' + proposalSentence_(proposal) + ', com os 5 avisos.', log: 'created: ' + proposal.date + ' ' + proposal.start };
    } catch (err) {
      console.warn('event insert failed: ' + err);
      return { kind: 'error', text: 'Não consegui gravar no calendário agora.', log: 'create error: ' + String(err && err.message || err).slice(0, 100) };
    }
  }
  if (action === 'propose') {
    var ev0 = Object.assign({}, req.partial || {}, stripEmpty_(r.event || {})); // H-4b: complete an earlier partial request
    var p = proposalFrom_(ev0, day, transcript + ' ' + String((req.partial || {}).said || ''));
    var partial = partialFrom_(ev0, transcript, req.partial, day.date);
    if (p.error === 'end') return { kind: 'unclear', partial: partial, text: 'Até que horas vai o evento?', log: 'proposal incomplete: end' };
    if (p.error === 'title') return { kind: 'unclear', partial: partial, text: 'Qual é o assunto do evento?', log: 'proposal incomplete: title' };
    if (p.error === 'start') return { kind: 'unclear', partial: partial, text: 'Que horas eu marco?', log: 'proposal incomplete: start' };
    if (p.error === 'date') return { kind: 'unclear', partial: partial, text: 'Pra que dia eu marco?', log: 'proposal incomplete: date' };
    if (p.error) return { kind: 'unclear', partial: partial, text: 'Me diz o dia e o horário do evento?', log: 'proposal incomplete: ' + p.error };
    return { kind: 'proposed', proposal: p.proposal, text: 'Vou criar ' + proposalSentence_(p.proposal) + ', com os 5 avisos padrão. Confirma?',
      log: 'proposed: ' + p.proposal.date + ' ' + p.proposal.start };
  }
  if (proposal) { // anything other than a clear yes or a new proposal discards it (Gemini 1831, D-2)
    if (action === 'cancel') return { kind: 'cancelled', text: 'Tá bom, não criei.', log: 'cancelled' };
    return { kind: 'discarded', text: 'Não gravei o evento. Se quiser, me pede de novo.', log: 'discarded' };
  }
  return null;
}

function end_(req) {
  try { if (req.clientLog && req.clientLog.length) logTurn_(req, { reason: 'end:' + String(req.reason || '') }, null); } catch (e) {}
  var sh = memorySheet_(false).getSheetByName('Sessions');
  var turns = Math.max(0, Math.min(1000, Number(req.turns) || 0));
  var reason = /^(voice|button|pagehide|daily_limit)$/.test(String(req.reason)) ? String(req.reason) : 'button';
  var id = String(req.sessionId || '').slice(0, 40);
  var lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    var last = sh.getLastRow();
    if (id && last > 1) {
      var ids = sh.getRange(Math.max(2, last - 20), 2, Math.min(21, last - 1), 1).getValues();
      for (var i = 0; i < ids.length; i++) if (ids[i][0] === id) return { saved: true, duplicate: true };
    }
    sh.appendRow([londonStamp_(), id, turns, reason]);
  } finally { lock.releaseLock(); }
  return { saved: true };
}

// ---------------------------------------------------------------- Claude-to-Gemini consult (LK-2, LK-3)

function consult_(req) {
  var cfg = config_();
  var prompt = String(req.prompt || '');
  var files = (req.files || []).map(function (f) { return '===== FILE: ' + String(f.name || 'untitled') + ' =====\n' + String(f.text || ''); });
  var text = prompt + '\n\n' + files.join('\n\n');
  if (!prompt || text.length > MAX_CONSULT_CHARS) return { error: 'prompt missing or too long' };
  if (!/00_README_AI_Handoffs_Protocol/.test(text)) return { error: 'attach 00_README_AI_Handoffs_Protocol (LK-3)' };
  if (!bump_('consults', DAILY_CONSULT_LIMIT)) return { error: 'daily consult limit reached' };
  var body = {
    systemInstruction: { parts: [{ text:
      'You are Gemini, reviewing work by Claude for Ti Santos as a second-opinion gate under the attached AI Handoffs protocol. ' +
      'You have no memory of earlier conversations; rely only on the attached files. Be specific: say Agree, or Change with exact text and evidence. ' +
      'Flag anything that could cost Ti money, install something, create unnecessary data, or ask Ti for avoidable work. Write in British English.' }] },
    contents: [{ role: 'user', parts: [{ text: text }] }],
    generationConfig: { temperature: 0.3, maxOutputTokens: 4096 },
  };
  var data = gemini_(cfg, body);
  var out = (((data.candidates || [])[0] || {}).content || {}).parts || [];
  return { text: out.map(function (p) { return p.text || ''; }).join(''), model: cfg.model, usage: data.usageMetadata || null };
}

function listModels_() {
  var cfg = config_();
  var res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=100',
    { headers: { 'x-goog-api-key': cfg.key }, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return { error: 'models ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 300) };
  return { current: cfg.model, models: modelNames_(JSON.parse(res.getContentText())) };
}

// ---------------------------------------------------------------- Gemini

function analyseSegment_(cfg, audioB64, pendingText, history, agenda, proposalText, partialText) {
  var contents = [];
  history.forEach(function (h) {
    contents.push({ role: 'user', parts: [{ text: String(h.ti || '') }] });
    contents.push({ role: 'model', parts: [{ text: String(h.robbie || '') }] });
  });
  contents.push({ role: 'user', parts: [
    { text: (agenda ? agenda + '\n' : '') + (partialText ? 'Pedido de evento incompleto do turno anterior (complete com este áudio): ' + partialText + '\n' : '') + segmentInstruction(pendingText, proposalText) },
    { inlineData: { mimeType: 'audio/wav', data: audioB64 } },
  ] });
  var body = {
    systemInstruction: { parts: [{ text: systemInstruction() }] },
    contents: contents,
    generationConfig: { responseMimeType: 'application/json', responseSchema: RESPONSE_SCHEMA, temperature: 0.7, maxOutputTokens: 1024 },
  };
  if (cfg.thinking) body.generationConfig.thinkingConfig = cfg.thinking;
  var data, modelUsed = cfg.model;
  try { data = gemini_(cfg, body); }
  catch (err) { // free tier: brief 500/503 "high demand" spikes; one quick retry, then ONE other free model (Gemini 2319, G-1)
    if (!/Gemini (500|503)/.test(String(err.message))) throw err;
    Utilities.sleep(700);
    try { data = gemini_(cfg, body); }
    catch (err2) {
      if (!/Gemini (500|503)/.test(String(err2.message))) throw err2;
      modelUsed = FALLBACK_MODELS.filter(function (m) { return m !== cfg.model; })[0];
      var fb = JSON.parse(JSON.stringify(body));
      delete fb.generationConfig.thinkingConfig; // thinking settings belong to the main model
      data = gemini_(cfg, fb, modelUsed);
    }
  }
  var txt = ((((data.candidates || [])[0] || {}).content || {}).parts || []).map(function (p) { return p.text || ''; }).join('');
  var parsed = JSON.parse(txt);
  return { transcript: String(parsed.transcript || ''), addressed_question: Boolean(parsed.addressed_question),
    urgent: Boolean(parsed.urgent), reply: String(parsed.reply || ''),
    calendar_action: /^(propose|confirm|cancel)$/.test(String(parsed.calendar_action)) ? String(parsed.calendar_action) : 'none',
    event: parsed.event && typeof parsed.event === 'object' ? parsed.event : null, model: modelUsed };
}

function gemini_(cfg, body, model) {
  var res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model || cfg.model) + ':generateContent', {
    method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': cfg.key },
    payload: JSON.stringify(body), muteHttpExceptions: true,
  });
  var code = res.getResponseCode();
  if (code !== 200) throw new Error('Gemini ' + code + ': ' + res.getContentText().slice(0, code === 429 ? 3000 : 300)); // 429 body names the quota (per minute or per day)
  return JSON.parse(res.getContentText());
}

// ---------------------------------------------------------------- persona (spike)

function systemInstruction() {
  return [
    'Você é o Robbie, parceiro de conversa por voz do Ti. Seu nome é Robbie. Nunca diga que é o Gemini, um assistente do Google ou uma IA genérica.',
    'Fale sempre em português do Brasil, como um amigo próximo: natural, caloroso e direto.',
    'Não comece as respostas com o nome dele e quase não use "Ti": no máximo uma vez a cada várias respostas, só quando soar natural.',
    'Por padrão, responda com uma única frase de no máximo ' + SHORT_LIMIT + ' palavras. Só quando o Ti pedir explicação ("explica", "fala mais", "por quê?") use até ' + LONG_LIMIT + ' palavras.',
    'Faça no máximo uma pergunta por resposta. Prefira perguntar a dar conselho. Não repita o que ele acabou de dizer.',
    'Quando o Ti contar uma decisão ou dúvida (cancelar, mudar, desistir, trocar), primeiro pergunte o motivo, numa pergunta curta; só dê opinião depois que ele explicar, ou se ele pedir de novo.',
    'Se a resposta do Ti não explicar o motivo (por exemplo, se ele só disser qual evento é), pergunte de novo de outro jeito (por exemplo "O que te faz querer cancelar?"); não opine ainda.',
    'Não sugira que ele descanse, pare ou fique em casa, a não ser que ele mesmo fale em descansar.',
    'Humor leve e seco é bem-vindo quando o momento permite, ligado ao que ele disse. Nunca faça humor sobre sofrimento, crise ou assuntos sensíveis, a não ser que ele brinque primeiro.',
    'Não faça suposições sobre o estado ou as intenções dele; pergunte.',
    'Você vê a agenda de HOJE e de AMANHÃ do calendário principal do Ti (conta tisantos.uk@gmail.com), quando ela vier na mensagem, com a hora atual. Use os horários exatos; eventos marcados [já passou] você fala no passado; não invente eventos.',
    'Você pode CRIAR um evento novo no calendário principal quando o Ti pedir: você repete o pedido e só grava depois que ele confirmar. Você não muda nem apaga eventos.',
    'Você não vê depois de amanhã nem outros dias, outras contas, e-mail, internet ou arquivos do Ti. Se ele pedir algo assim, diga isso numa frase curta e não invente.',
    'Segurança: se o Ti falar em se machucar, em não querer viver, ou em perigo imediato, responda com cuidado, pergunte diretamente se ele está seguro, e lembre que o Samaritans atende 24 horas no 116 123 no Reino Unido e o CVV no 188 no Brasil.',
  ].join('\n');
}

function segmentInstruction(pendingText, proposalText) {
  return [
    proposalText ? 'Evento proposto por você no turno anterior, esperando o Ti confirmar: ' + proposalText : '',
    pendingText ? 'Trechos anteriores deste turno do Ti, ainda sem resposta: "' + pendingText + '"' : 'Este é o início de um novo turno do Ti.',
    'Tarefa:',
    '1) transcript: transcreva fielmente o áudio anexado, em português do Brasil, com pontuação; use "?" quando for pergunta. Se não houver fala clara, devolva texto vazio.',
    '2) addressed_question: true somente se o áudio terminar com uma pergunta dirigida a você (Robbie), e não uma pergunta que o Ti faz a si mesmo enquanto pensa.',
    '3) urgent: true somente se o Ti expressar intenção de se machucar, de não querer viver, ou perigo imediato.',
    '4) reply: a resposta do Robbie ao turno inteiro do Ti (trechos anteriores + este áudio), seguindo as regras do sistema.',
    '5) calendar_action: "propose" se o Ti pedir para criar um evento ou lembrete no calendário, MESMO sem horário ou data (o servidor pergunta o que faltar; não ofereça, proponha) (preencha event: title curto, date AAAA-MM-DD, start HH:MM, end = horário de término HH:MM se o Ti disser o fim ou a duração, vazio só se ele não disser; use a data de hoje da agenda para entender "amanhã", "sexta" etc.); "confirm" somente se existir um evento proposto e o Ti disser claramente que sim; "cancel" se ele disser não ou desistir; senão "none".',
  ].filter(Boolean).join('\n');
}

var RESPONSE_SCHEMA = {
  type: 'OBJECT',
  properties: { transcript: { type: 'STRING' }, addressed_question: { type: 'BOOLEAN' }, urgent: { type: 'BOOLEAN' }, reply: { type: 'STRING' },
    calendar_action: { type: 'STRING', enum: ['none', 'propose', 'confirm', 'cancel'] },
    event: { type: 'OBJECT', properties: { title: { type: 'STRING' }, date: { type: 'STRING' }, start: { type: 'STRING' }, end: { type: 'STRING' } } } },
  required: ['transcript', 'addressed_question', 'urgent', 'reply', 'calendar_action'],
};

// ---------------------------------------------------------------- turn-taking (TT-01 v2, identical to the tested v3.0 logic)

function normalise(text) {
  return String(text || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/\b(robby|robi|roby|robie|hobby|robbi)\b/g, 'robbie')
    .replace(/[^a-z0-9?\s]/g, ' ').replace(/\s*\?\s*/g, '? ').replace(/\s+/g, ' ').trim();
}
var STRONG_SIGNALS = ['sua vez', 'o que voce acha', 'o que voce pensa', 'qual sua opiniao', 'qual a sua opiniao', 'pode falar'];
var WEAK_SIGNALS = ['e ai', 'faz sentido'];
var STRONG_WINDOW_WORDS = 6;
function lastWords_(norm, n) { return norm.replace(/\?/g, '').split(' ').filter(Boolean).slice(-n).join(' '); }
function isTurnComplete(segment, addressedQuestion, firstSegment) {
  var norm = normalise(segment);
  if (!norm) return { complete: false, reason: 'empty' };
  var plain = ' ' + norm.replace(/\?/g, '').replace(/\s+/g, ' ').trim() + ' ';
  for (var i = 0; i < STRONG_SIGNALS.length; i++) {
    var s = STRONG_SIGNALS[i];
    var at = plain.lastIndexOf(' ' + s + ' ');
    if (at === -1) continue;
    var after = plain.slice(at + s.length + 2).trim();
    var wordsAfter = after ? after.split(' ').length : 0;
    if (wordsAfter <= STRONG_WINDOW_WORDS - 2) return { complete: true, reason: 'signal:' + s };
  }
  var endsWithQuestion = /\?$/.test(norm);
  var finalWords = ' ' + lastWords_(norm, 2) + ' ';
  if (endsWithQuestion) {
    for (var j = 0; j < WEAK_SIGNALS.length; j++) {
      if (finalWords.slice(-(WEAK_SIGNALS[j].length + 2)) === ' ' + WEAK_SIGNALS[j] + ' ') return { complete: true, reason: 'signal:' + WEAK_SIGNALS[j] + '?' };
    }
  }
  // Ti's rule (23 Nov 13:51): "a question is a question". A segment that ends with "?" gets a reply; no filters.
  // The model's addressed_question is still logged in the Turns tab for observation, but does not gate the reply.
  if (endsWithQuestion) return { complete: true, reason: addressedQuestion ? 'question-to-robbie' : 'question' };
  if (isRequest(segment)) return { complete: true, reason: 'request' };
  return { complete: false, reason: 'no-signal' };
}
// v3.7 (F-2): a direct request to Robbie also gets a reply, even without "?": a request verb plus a target word, anywhere in
// the segment (Ti's 21:41 example starts with "Faz um evento no calendário para mim..."). Thinking-aloud openings stay silent.
var REQUEST_VERBS = /\b(faz|faca|fazer|cria|crie|criar|marca|marque|marcar|agenda|agende|agendar|lembra|lembre|lembrar|anota|anote|anotar|coloca|coloque|colocar|salva|salve|grava|grave|gravar|diz|diga|fala|fale|explica|explique|conta|conte)\b/;
var REQUEST_TARGETS = /\b(pra mim|para mim|me|robbie|calendario|agenda|lembrete|evento|compromisso)\b/;
var THINKING_ALOUD = /^(acho que|fiquei pensando|talvez|sera que|eu ia|eu tava pensando|eu estava pensando|tipo assim)\b/;
function isRequest(segment) {
  var t = normalise(segment).replace(/\?/g, ' ').replace(/\s+/g, ' ').trim();
  if (!t || THINKING_ALOUD.test(t)) return false;
  return REQUEST_VERBS.test(t) && REQUEST_TARGETS.test(t);
}

function isClosing(segment) {
  var norm = normalise(segment);
  return /\btchau robbie\b/.test(norm) || /\b(pode )?encerrar( a sessao| a conversa)?\s*$/.test(norm.replace(/\?/g, '').trim());
}
function wantsExplanation(turnText) {
  var norm = normalise(turnText);
  return /\b(explica|me explica|explique|fala mais|detalha|me conta mais|aprofunda)\b/.test(norm) || /\bpor que\?\s*$/.test(norm) || /\bpor que voce\b/.test(norm);
}
function capReply(text, maxWords) {
  var clean = String(text || '').replace(/\s+/g, ' ').trim();
  if (!clean) return '';
  var words = clean.split(' ');
  if (words.length <= maxWords) return clean;
  var cut = words.slice(0, maxWords).join(' ');
  var lastEnd = Math.max(cut.lastIndexOf('.'), cut.lastIndexOf('?'), cut.lastIndexOf('!'));
  if (lastEnd !== -1 && cut.slice(0, lastEnd + 1).split(' ').length >= 3) return cut.slice(0, lastEnd + 1);
  return cut.replace(/[,;:\-–]+$/, '') + '…';
}

// ---------------------------------------------------------------- helpers

function config_() {
  var p = PropertiesService.getScriptProperties();
  var key = p.getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('not configured');
  var thinking = null;
  try { thinking = JSON.parse(p.getProperty('THINKING_CONFIG') || 'null'); } catch (e) { thinking = null; }
  return { key: key, model: p.getProperty('GEMINI_MODEL') || DEFAULT_MODEL, thinking: thinking,
    ttsOn: p.getProperty('TTS_ENABLED') !== 'false', ttsModel: p.getProperty('TTS_MODEL') || DEFAULT_TTS_MODEL,
    ttsVoice: p.getProperty('TTS_VOICE') || DEFAULT_TTS_VOICE };
}

// ---------------------------------------------------------------- voice (V-1: Gemini TTS on the free key, silent fallback)

function tts_(cfg, text, model, voice, plain) {
  var body = {
    contents: [{ parts: [{ text: plain ? text : 'Fale em português do Brasil, em ritmo natural de conversa, claro e animado: ' + text }] }],
    generationConfig: { responseModalities: ['AUDIO'],
      speechConfig: { voiceConfig: { prebuiltVoiceConfig: { voiceName: voice || cfg.ttsVoice } } } },
  };
  var data = gemini_(cfg, body, model || cfg.ttsModel);
  var part = ((((data.candidates || [])[0] || {}).content || {}).parts || []).filter(function (x) { return x.inlineData; })[0];
  if (!part) throw new Error('TTS returned no audio');
  var mime = String(part.inlineData.mimeType || '');
  var m = /rate=(\d+)/i.exec(mime);
  return { data: part.inlineData.data, rate: m ? Number(m[1]) : 24000, mime: mime, format: /wav/i.test(mime) ? 'wav' : 'pcm', model: model || cfg.ttsModel };
}

// v3.8 (H-2b): Google's free-tier daily limits reset at midnight Pacific time, so the day flag uses that date.
function quotaDay_() { return Utilities.formatDate(new Date(), 'America/Los_Angeles', 'yyyy-MM-dd').slice(0, 10); }

// v3.8: read which quota fired from the 429 JSON (quotaId ...PerDay... / ...PerMinute...); text match only as a fallback.
function quotaKind_(msg) {
  var body = String(msg).replace(/^Gemini 429: /, '');
  try {
    var ids = [];
    ((JSON.parse(body).error || {}).details || []).forEach(function (d) {
      (d.violations || []).forEach(function (v) { ids.push(String(v.quotaId || v.quotaMetric || '')); });
    });
    if (ids.length) return ids.some(function (x) { return /PerDay/i.test(x); }) ? 'day' : 'minute';
  } catch (e) {}
  return /per[ _-]?day|daily|RPD/i.test(body) ? 'day' : 'minute';
}

function ttsSafe_(cfg, text, why) {
  why = why || {};
  if (!cfg.ttsOn || !text) { why.reason = 'off'; return null; }
  var cache = CacheService.getScriptCache();
  var day = quotaDay_();
  var models = [cfg.ttsModel];
  if (TTS_DAILY_BACKUP && TTS_DAILY_BACKUP !== cfg.ttsModel) models.push(TTS_DAILY_BACKUP);
  for (var i = 0; i < models.length; i++) {
    var model = models[i];
    var isBackup = i > 0;
    if (isBackup && !cache.get('tts_day_' + day + '_' + models[0])) break; // backup only after the main model's daily limit
    if (cache.get('tts_day_' + day + '_' + model)) { why.reason = 'daily quota'; continue; }
    if (cache.get('tts_wait_' + model)) { why.reason = 'minute quota'; break; } // short pause; no switching mid-conversation
    for (var attempt = 0; attempt < 2; attempt++) {
      try { var a = tts_(cfg, text, model, null, isBackup || attempt === 1); a.backup = isBackup; return a; } // backup model: plain text (K-4); main: 2nd try plain
      catch (err) {
        var msg = String(err.message);
        console.warn('TTS ' + model + ' attempt ' + (attempt + 1) + ' failed: ' + msg.slice(0, 300));
        if (/Gemini 429/.test(msg)) {
          why.detail = msg.replace(/\s+/g, ' ').slice(0, 300); // H-2c: which quota fired, for the Turns tab
          if (quotaKind_(msg) === 'day') { cache.put('tts_day_' + day + '_' + model, '1', 21600); why.reason = 'daily quota'; break; }
          cache.put('tts_wait_' + model, '1', 30); why.reason = 'minute quota';
          return null;
        }
        if (!/no audio/i.test(msg)) { why.reason = 'error'; why.detail = msg.slice(0, 300); break; } // no retry on 400/404/5xx
        why.reason = 'no audio';
      }
    }
  }
  return null;
}

// ---------------------------------------------------------------- Increment 1: today's calendar (read-only scope in appsscript.json)

// "Today" is the phone's own day: localDate (YYYY-MM-DD) and tzOffset (minutes, as Date.getTimezoneOffset()) come with each
// turn; Europe/London is the fallback. Only titles and times are read; no descriptions, guests or locations.
function phoneDay_(req) {
  var off = Number(req && req.tzOffset);
  var date = String((req && req.localDate) || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !(off >= -840 && off <= 840)) {
    date = Utilities.formatDate(new Date(), 'Europe/London', 'yyyy-MM-dd').slice(0, 10);
    var londonNow = new Date(Utilities.formatDate(new Date(), 'Europe/London', "yyyy-MM-dd'T'HH:mm:ss") + 'Z');
    off = Math.round((Date.now() - londonNow.getTime()) / 60000);
  }
  var p = date.split('-').map(Number);
  var start = new Date(Date.UTC(p[0], p[1] - 1, p[2]) + off * 60000);
  return { date: date, offset: off, start: start, end: new Date(start.getTime() + 86400000) };
}

function localHm_(d, offset) {
  var t = new Date(d.getTime() - offset * 60000);
  var h = t.getUTCHours(), m = t.getUTCMinutes();
  return (h < 10 ? '0' : '') + h + ':' + (m < 10 ? '0' : '') + m;
}

function calendarDigest_(req) {
  if (PropertiesService.getScriptProperties().getProperty('CALENDAR_ENABLED') === 'false') return { text: '', status: 'off' };
  var day = phoneDay_(req);
  var cache = CacheService.getScriptCache();
  var key = 'cal_' + day.date + '_' + day.offset;
  var hit = cache.get(key);
  if (hit) { var h = JSON.parse(hit); return { text: tagNow_(h.text, day.offset), status: h.status + ' (cache)' }; }
  var out;
  try {
    // v3.8 (H-3): today and tomorrow in one read; at most 15 events per day
    var tomorrowStart = day.end, tomorrowEnd = new Date(day.end.getTime() + 86400000);
    var tomorrowDate = addDays_(day.date, 1);
    var all = CalendarApp.getDefaultCalendar().getEvents(day.start, tomorrowEnd);
    var line = function (e) {
      var title = String(e.getTitle() || '(sem título)').replace(/[\u0000-\u001f]/g, ' ').slice(0, 60);
      if (e.isAllDayEvent()) return '- dia inteiro: ' + title;
      return '- ' + localHm_(e.getStartTime(), day.offset) + ' a ' + localHm_(e.getEndTime(), day.offset) + ': ' + title + ' {' +
        e.getStartTime().getTime() + ',' + e.getEndTime().getTime() + '}';
    };
    var today = all.filter(function (e) { return e.getStartTime() < tomorrowStart; }).slice(0, 15).map(line);
    var tomorrow = all.filter(function (e) { return e.getStartTime() >= tomorrowStart || (e.isAllDayEvent() && e.getEndTime() > tomorrowStart && e.getStartTime() >= tomorrowStart); })
      .slice(0, 15).map(line);
    out = { text: 'Agenda do Ti (calendário principal). Hoje, ' + spokenDay_(day.date) + ' (' + day.date + '): ' + (today.length ? '\n' + today.join('\n') : 'nenhum evento.') +
      '\nAmanhã, ' + spokenDay_(tomorrowDate) + ' (' + tomorrowDate + '): ' + (tomorrow.length ? '\n' + tomorrow.join('\n') : 'nenhum evento.'),
      status: 'events:' + today.length + '+' + tomorrow.length };
  } catch (err) {
    console.warn('calendar read failed: ' + err);
    // v3.6.1: the error text goes to column N, so Claude can read the cause from the sheet
    out = { text: 'Agenda do Ti: indisponível agora (erro de leitura). Se ele perguntar, diga que não conseguiu ler a agenda.',
      status: 'error: ' + String(err && err.message || err).slice(0, 120) };
  }
  try { cache.put(key, JSON.stringify(out), /^error/.test(out.status) ? 30 : 300); } catch (e) {} // errors cached 30 s only
  return { text: tagNow_(out.text, day.offset), status: out.status };
}

// v3.7 (F-1): the current time and a past/now/later tag are added on every turn (the cached list keeps raw instants).
function tagNow_(text, offset) {
  var now = Date.now();
  var tagged = String(text).replace(/ \{(\d+),(\d+)\}/g, function (_, a, b) {
    a = Number(a); b = Number(b);
    return b <= now ? ' [já passou]' : a <= now ? ' [em andamento]' : ' [ainda vai]';
  });
  return /^Agenda do Ti/.test(tagged) ? 'Agora são ' + localHm_(new Date(now), offset) + '. ' + tagged : tagged;
}

// v3.6.1: run this once from the editor (function list -> autorizarCalendario -> Run). If the calendar permission is missing,
// Google shows the authorisation screen; otherwise the log shows "Calendário OK" or the exact error. Read-only.
function autorizarCalendario() {
  var cal = CalendarApp.getDefaultCalendar();
  var n = cal.getEvents(new Date(), new Date(Date.now() + 86400000)).length;
  console.log('Calendário OK: ' + cal.getName() + '. Eventos nas próximas 24 horas: ' + n);
  return n;
}

function logTurn_(req, out, analysis, calendarStatus) {
  if (PropertiesService.getScriptProperties().getProperty('DEBUG_LOG') === 'false') return;
  var ss = memorySheet_(false);
  var sh = ss.getSheetByName('Turns');
  if (!sh) {
    sh = ss.insertSheet('Turns');
    sh.getRange('A1:M1').setValues([['At (London)', 'Session ID', 'Transcript', 'Ends with ?', 'Model says addressed', 'Reason',
      'Spoke', 'Reply', 'Voice model', 'Voice fallback', 'Gemini ms', 'Voice ms', 'Phone notes']]);
  }
  if (!sh.getRange('N1').getValue()) sh.getRange('N1').setValues([['Calendar']]);
  if (!sh.getRange('O1').getValue()) sh.getRange('O1').setValues([['Model']]);
  var a = analysis || {};
  var notes = (req.clientLog || []).slice(-10).map(function (x) { return String(x).slice(0, 200); }).join(' | ');
  var t = String(out.transcript || '');
  sh.appendRow([londonStamp_(), String(req.sessionId || '').slice(0, 40), t, /\?\s*$/.test(t), analysis ? Boolean(a.addressed_question) : '',
    String(out.reason || (out.busy ? 'busy' : out.error ? 'error' : '')), Boolean(out.speak), String(out.text || ''),
    String(out.ttsModel || ''), String(out.ttsFallback || '') + (out.ttsDetail ? ' | ' + out.ttsDetail : ''), out.geminiMs || '', out.ttsMs || '', notes, String(calendarStatus || ''), String((analysis && analysis.model) || '')]);
}

function greetingAudio_(cfg, text) { // cached in 90 KB chunks for 6 h, so each session does not spend a TTS request
  if (!cfg.ttsOn) return null;
  var cache = CacheService.getScriptCache();
  var tag = 'greet_' + TTS_STYLE_VERSION + '_' + cfg.ttsModel + '_' + cfg.ttsVoice + '_' + hash_(text); // one cache entry per greeting text
  var meta = cache.get(tag);
  if (meta) {
    var parts = meta.split(':'), n = Number(parts[0]), keys = [];
    for (var i = 0; i < n; i++) keys.push(tag + '_' + i);
    var got = cache.getAll(keys), data = '';
    for (var j = 0; j < n; j++) { if (!got[keys[j]]) { data = null; break; } data += got[keys[j]]; }
    if (data) return { data: data, rate: Number(parts[1]), format: parts[2] || 'pcm', model: 'cache' };
  }
  var why = {};
  var a = ttsSafe_(cfg, text, why);
  if (!a) return { failed: true, why: why };
  if (a.backup) return a; // K-4: a greeting from the backup model is not cached
  var chunks = {}, size = 90000, count = Math.ceil(a.data.length / size);
  for (var k = 0; k < count; k++) chunks[tag + '_' + k] = a.data.slice(k * size, (k + 1) * size);
  try { cache.putAll(chunks, 21600); cache.put(tag, count + ':' + a.rate + ':' + a.format, 21600); } catch (e) {}
  return a;
}

// ---------------------------------------------------------------- Claude's tuning actions (passphrase-protected, values validated)

function knownModels_(cfg) {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('model_list');
  if (hit) return JSON.parse(hit);
  var res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=1000',
    { headers: { 'x-goog-api-key': cfg.key }, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('models ' + res.getResponseCode());
  var names = modelNames_(JSON.parse(res.getContentText()));
  cache.put('model_list', JSON.stringify(names), 3600);
  return names;
}

function cleanThinking_(t) {
  if (t === null || t === undefined || t === '') return null;
  if (typeof t !== 'object') throw new Error('thinking must be an object');
  var out = {};
  if ('thinkingBudget' in t) { var b = Number(t.thinkingBudget); if (!(b >= -1 && b <= 2048)) throw new Error('bad thinkingBudget'); out.thinkingBudget = b; }
  if ('thinkingLevel' in t) { if (!/^(minimal|low|medium|high)$/.test(String(t.thinkingLevel))) throw new Error('bad thinkingLevel'); out.thinkingLevel = String(t.thinkingLevel); }
  return Object.keys(out).length ? out : null;
}

function bench_(req) {
  var cfg = config_();
  if (!req.audio || String(req.audio).length > MAX_AUDIO_B64) return { error: 'audio missing or longer than 30 seconds' };
  var model = String(req.model || cfg.model);
  if (knownModels_(cfg).indexOf(model) < 0) return { error: 'model not offered to this key' };
  if (!bump_('bench', DAILY_BENCH_LIMIT)) return { error: 'daily bench limit reached' };
  var run = { key: cfg.key, model: model, thinking: cleanThinking_(req.thinking) };
  var t0 = Date.now();
  try {
    var r = analyseSegment_(run, req.audio, '', []);
    return { model: model, thinking: run.thinking, ms: Date.now() - t0, transcript: r.transcript, addressed_question: r.addressed_question, reply: r.reply };
  } catch (err) { return { model: model, thinking: run.thinking, ms: Date.now() - t0, error: String(err.message).slice(0, 300) }; }
}

function ttsTest_(req) {
  var cfg = config_();
  var voice = String(req.voice || cfg.ttsVoice), model = String(req.model || cfg.ttsModel);
  if (TTS_VOICES.indexOf(voice) < 0) return { error: 'unknown voice' };
  if (!/tts/.test(model)) return { error: 'not a TTS model' };
  if (!bump_('bench', DAILY_BENCH_LIMIT)) return { error: 'daily bench limit reached' };
  var text = String(req.text || 'Oi Ti, aqui é o Robbie. Como você está?').slice(0, 300);
  var t0 = Date.now();
  try {
    var a = tts_(cfg, text, model, voice);
    var out = { model: model, voice: voice, ms: Date.now() - t0, rate: a.rate, mime: a.mime, bytes: Math.round(a.data.length * 3 / 4) };
    if (req.returnAudio) out.audio = a.data;
    return out;
  } catch (err) { return { model: model, voice: voice, ms: Date.now() - t0, error: String(err.message).slice(0, 300) }; }
}

function setConfig_(req) {
  var cfg = config_();
  var p = PropertiesService.getScriptProperties();
  var set = {};
  if (req.model !== undefined) { if (knownModels_(cfg).indexOf(String(req.model)) < 0) return { error: 'model not offered to this key' }; set.GEMINI_MODEL = String(req.model); }
  if (req.thinking !== undefined) { var t = cleanThinking_(req.thinking); set.THINKING_CONFIG = t ? JSON.stringify(t) : ''; }
  if (req.ttsModel !== undefined) { if (!/tts/.test(String(req.ttsModel))) return { error: 'not a TTS model' }; set.TTS_MODEL = String(req.ttsModel); }
  if (req.ttsVoice !== undefined) { if (TTS_VOICES.indexOf(String(req.ttsVoice)) < 0) return { error: 'unknown voice' }; set.TTS_VOICE = String(req.ttsVoice); }
  if (req.ttsEnabled !== undefined) set.TTS_ENABLED = req.ttsEnabled ? 'true' : 'false';
  if (req.debugLog !== undefined) set.DEBUG_LOG = req.debugLog ? 'true' : 'false';
  if (req.calendarEnabled !== undefined) set.CALENDAR_ENABLED = req.calendarEnabled ? 'true' : 'false';
  p.setProperties(set);
  var now = config_();
  return { model: now.model, thinking: now.thinking, ttsOn: now.ttsOn, ttsModel: now.ttsModel, ttsVoice: now.ttsVoice };
}

function memorySheet_(createIfMissing) {
  var p = PropertiesService.getScriptProperties();
  var id = p.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  if (!createIfMissing) throw new Error('memory sheet missing');
  var ss = SpreadsheetApp.create('FOLDER-OS_Robbie_Memory');
  var cfg = ss.getSheets()[0].setName('Config');
  cfg.getRange('A1:B1').setValues([['memory_check', 'BRAVO_42']]);
  var sessions = ss.insertSheet('Sessions');
  sessions.getRange('A1:D1').setValues([['Closed At (London)', 'Session ID', 'Turns', 'Closed By']]);
  return ss;
}

function bump_(name, limit) {
  var p = PropertiesService.getScriptProperties();
  var key = 'COUNT_' + name + '_' + londonStamp_().slice(0, 10);
  var lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    var n = Number(p.getProperty(key) || 0);
    if (n >= limit) return false;
    p.setProperty(key, String(n + 1));
    return true;
  } finally { lock.releaseLock(); }
}

function passOk_(given, expected) {
  if (!expected || typeof given !== 'string' || given.length !== expected.length) return false;
  var diff = 0;
  for (var i = 0; i < expected.length; i++) diff |= given.charCodeAt(i) ^ expected.charCodeAt(i);
  return diff === 0;
}

function modelNames_(data) {
  return (data.models || [])
    .filter(function (m) { return !m.supportedGenerationMethods || m.supportedGenerationMethods.indexOf('generateContent') >= 0; })
    .map(function (m) { return String(m.name || '').replace(/^models\//, ''); })
    .filter(function (n) { return /flash/.test(n); });
}

function hash_(t) { var h = 0; t = String(t); for (var i = 0; i < t.length; i++) h = (h * 31 + t.charCodeAt(i)) | 0; return (h >>> 0).toString(36); }

function londonStamp_() { return Utilities.formatDate(new Date(), 'Europe/London', 'yyyy-MM-dd HH:mm'); }

function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

function setupHtml_(configured) {
  var head = `<!doctype html>
<html>
<head>
<base target="_top">
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
  body { font-family: system-ui, sans-serif; max-width: 560px; margin: 24px auto; padding: 0 16px; color: #1c1f24; }
  h1 { font-size: 22px; }
  input { width: 100%; font-size: 16px; padding: 12px; box-sizing: border-box; border: 1px solid #bbb; border-radius: 10px; }
  button { margin-top: 12px; font-size: 17px; padding: 12px 18px; border: 0; border-radius: 10px; background: #2a6f5f; color: #fff; }
  .box { background: #f2f4f6; padding: 12px; border-radius: 10px; word-break: break-all; font-family: ui-monospace, monospace; font-size: 14px; }
  #qr { margin: 16px 0; }
  .err { color: #b3261e; }
</style>
</head>
<body>
<h1>Robbie: configuração</h1>
`;
  var configuredPart = `
  <p>O Robbie já está configurado. Nada a fazer aqui.</p>
`;
  var formPart = `
  <div id="form">
    <p>Cole a chave gratuita do Gemini e toque em Salvar.</p>
    <input id="key" type="password" autocomplete="off" placeholder="Chave do Gemini">
    <button id="save">Salvar</button>
    <p id="msg"></p>
  </div>
  <div id="done" style="display:none">
    <p><b>1. No celular:</b> aponte a câmera para este código.</p>
    <div id="qr"></div>
    <p><b>2. Para o Claude:</b> copie a linha abaixo e cole na conversa com ele.</p>
    <div class="box" id="claude"></div>
  </div>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script>
  <script>
    document.getElementById('save').onclick = function () {
      var msg = document.getElementById('msg');
      msg.className = ''; msg.textContent = 'Verificando a chave...';
      google.script.run
        .withSuccessHandler(function (r) {
          document.getElementById('form').style.display = 'none';
          document.getElementById('done').style.display = 'block';
          document.getElementById('claude').textContent = r.claudeLine + ' ' + r.model;
          new QRCode(document.getElementById('qr'), { text: r.phoneLink, width: 240, height: 240 });
        })
        .withFailureHandler(function (e) { msg.className = 'err'; msg.textContent = e.message || String(e); })
        .saveSetup(document.getElementById('key').value);
    };
  </script>
`;
  var tail = `
</body>
</html>
`;
  return head + (configured ? configuredPart : formPart) + tail;
}
