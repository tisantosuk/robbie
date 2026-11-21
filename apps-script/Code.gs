/**
 * Robbie v3.1 backend (Google Apps Script web app). Zero cost: no billing account anywhere.
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
var DEFAULT_TTS_VOICE = 'Aoede';
var TTS_STYLE_VERSION = 's1';    // change when the TTS style text below changes
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
      case 'start': return json_(start_());
      case 'turn': return json_(turn_(req));
      case 'end': return json_(end_(req));
      case 'consult': return json_(consult_(req));
      case 'models': return json_(listModels_());
      case 'bench': return json_(bench_(req));       // Claude: measure one model setting on a recording
      case 'tts_test': return json_(ttsTest_(req));  // Claude: measure one voice setting
      case 'set_config': return json_(setConfig_(req)); // Claude: apply the agreed model/voice settings
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

function start_() {
  var cfg = config_();
  var memory = '';
  var cache = CacheService.getScriptCache();
  memory = cache.get('memory_check');
  if (memory === null) {
    memory = String(memorySheet_(false).getSheetByName('Config').getRange('B1').getValue() || '');
    cache.put('memory_check', memory, 600);
  }
  var greeting = 'Oi Ti, aqui é o Robbie. Como você está?';
  var out = { greeting: greeting, memoryValue: memory, model: cfg.model };
  var audio = greetingAudio_(cfg, greeting);
  if (audio) { out.audio = audio.data; out.audioRate = audio.rate; }
  return out;
}

function turn_(req) {
  var cfg = config_();
  var pending = (req.pending || []).map(String).slice(-20);
  var history = (req.history || []).slice(-20);
  if (!req.audio || String(req.audio).length > MAX_AUDIO_B64) return { error: 'audio missing or longer than 30 seconds' };

  if (!bump_('turns', DAILY_TURN_LIMIT)) {
    return { transcript: '', speak: true, closed: true, reason: 'daily_limit',
      text: 'Ti, chegamos ao limite de hoje. A gente continua amanhã.', pending: [], history: history };
  }

  var t0 = Date.now();
  var r = analyseSegment_(cfg, req.audio, pending.join(' '), history);
  var geminiMs = Date.now() - t0;
  var transcript = String(r.transcript || '').trim();
  if (!transcript) return { transcript: '', speak: false, reason: 'no-speech', pending: pending, history: history, geminiMs: geminiMs };

  if (isClosing(transcript)) {
    return { transcript: transcript, speak: true, closed: true, reason: 'voice', text: 'Tchau, Ti. Guardei nossa conversa.',
      pending: [], history: history, geminiMs: geminiMs };
  }

  var firstSegment = pending.length === 0;
  pending.push(transcript);
  var decision = r.urgent ? { complete: true, reason: 'urgent' } : isTurnComplete(transcript, r.addressed_question, firstSegment);
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
  var out = { transcript: transcript, speak: true, reason: decision.reason, text: reply, pending: [], history: history.slice(-20), geminiMs: geminiMs };
  var t1 = Date.now();
  var audio = ttsSafe_(cfg, reply); // V-1: same round trip; on any failure the phone uses its own voice
  if (audio) { out.audio = audio.data; out.audioRate = audio.rate; out.ttsMs = Date.now() - t1; }
  return out;
}

function end_(req) {
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

function analyseSegment_(cfg, audioB64, pendingText, history) {
  var contents = [];
  history.forEach(function (h) {
    contents.push({ role: 'user', parts: [{ text: String(h.ti || '') }] });
    contents.push({ role: 'model', parts: [{ text: String(h.robbie || '') }] });
  });
  contents.push({ role: 'user', parts: [
    { text: segmentInstruction(pendingText) },
    { inlineData: { mimeType: 'audio/wav', data: audioB64 } },
  ] });
  var body = {
    systemInstruction: { parts: [{ text: systemInstruction() }] },
    contents: contents,
    generationConfig: { responseMimeType: 'application/json', responseSchema: RESPONSE_SCHEMA, temperature: 0.7, maxOutputTokens: 1024 },
  };
  if (cfg.thinking) body.generationConfig.thinkingConfig = cfg.thinking;
  var data = gemini_(cfg, body);
  var txt = ((((data.candidates || [])[0] || {}).content || {}).parts || []).map(function (p) { return p.text || ''; }).join('');
  var parsed = JSON.parse(txt);
  return { transcript: String(parsed.transcript || ''), addressed_question: Boolean(parsed.addressed_question),
    urgent: Boolean(parsed.urgent), reply: String(parsed.reply || '') };
}

function gemini_(cfg, body, model) {
  var res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model || cfg.model) + ':generateContent', {
    method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': cfg.key },
    payload: JSON.stringify(body), muteHttpExceptions: true,
  });
  var code = res.getResponseCode();
  if (code !== 200) throw new Error('Gemini ' + code + ': ' + res.getContentText().slice(0, 300));
  return JSON.parse(res.getContentText());
}

// ---------------------------------------------------------------- persona (spike)

function systemInstruction() {
  return [
    'Você é o Robbie, parceiro de conversa por voz do Ti. Seu nome é Robbie. Nunca diga que é o Gemini, um assistente do Google ou uma IA genérica.',
    'Fale sempre em português do Brasil, como um amigo próximo: natural, caloroso e direto. Chame-o de Ti.',
    'Por padrão, responda com uma única frase de no máximo ' + SHORT_LIMIT + ' palavras. Só quando o Ti pedir explicação ("explica", "fala mais", "por quê?") use até ' + LONG_LIMIT + ' palavras.',
    'Faça no máximo uma pergunta por resposta. Prefira perguntar a dar conselho. Não repita o que ele acabou de dizer.',
    'Humor leve e seco é bem-vindo quando o momento permite, ligado ao que ele disse. Nunca faça humor sobre sofrimento, crise ou assuntos sensíveis, a não ser que ele brinque primeiro.',
    'Não faça suposições sobre o estado ou as intenções dele; pergunte.',
    'Segurança: se o Ti falar em se machucar, em não querer viver, ou em perigo imediato, responda com cuidado, pergunte diretamente se ele está seguro, e lembre que o Samaritans atende 24 horas no 116 123 no Reino Unido e o CVV no 188 no Brasil.',
  ].join('\n');
}

function segmentInstruction(pendingText) {
  return [
    pendingText ? 'Trechos anteriores deste turno do Ti, ainda sem resposta: "' + pendingText + '"' : 'Este é o início de um novo turno do Ti.',
    'Tarefa:',
    '1) transcript: transcreva fielmente o áudio anexado, em português do Brasil, com pontuação; use "?" quando for pergunta. Se não houver fala clara, devolva texto vazio.',
    '2) addressed_question: true somente se o áudio terminar com uma pergunta dirigida a você (Robbie), e não uma pergunta que o Ti faz a si mesmo enquanto pensa.',
    '3) urgent: true somente se o Ti expressar intenção de se machucar, de não querer viver, ou perigo imediato.',
    '4) reply: a resposta do Robbie ao turno inteiro do Ti (trechos anteriores + este áudio), seguindo as regras do sistema.',
  ].join('\n');
}

var RESPONSE_SCHEMA = {
  type: 'OBJECT',
  properties: { transcript: { type: 'STRING' }, addressed_question: { type: 'BOOLEAN' }, urgent: { type: 'BOOLEAN' }, reply: { type: 'STRING' } },
  required: ['transcript', 'addressed_question', 'urgent', 'reply'],
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
  var sentences = norm.split(/(?<=\?)\s/).filter(Boolean);
  var lastSentence = sentences[sentences.length - 1] || '';
  var addressed = /\b(voce|robbie)\b/.test(lastSentence);
  if (addressedQuestion && endsWithQuestion && addressed) return { complete: true, reason: 'question-to-robbie' };
  // V-2: a whole turn that is only a short check-in question ("oi, tudo bem?"), never a self-directed one.
  var words = plain.trim().split(' ').filter(Boolean);
  var selfDirected = /^(por que|porque|sera que|como eu|o que eu|eu nao sei)\b/.test(plain.trim()) || /\b(eu|me|mim|meu|minha|comigo)\b/.test(plain);
  if (firstSegment && addressedQuestion && endsWithQuestion && sentences.length === 1 && words.length <= 5 && !selfDirected) {
    return { complete: true, reason: 'short-check-in' };
  }
  return { complete: false, reason: 'no-signal' };
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

function tts_(cfg, text, model, voice) {
  var body = {
    contents: [{ parts: [{ text: 'Fale em português do Brasil, com voz natural, calorosa e calma: ' + text }] }],
    generationConfig: { responseModalities: ['AUDIO'],
      speechConfig: { voiceConfig: { prebuiltVoiceConfig: { voiceName: voice || cfg.ttsVoice } } } },
  };
  var data = gemini_(cfg, body, model || cfg.ttsModel);
  var part = ((((data.candidates || [])[0] || {}).content || {}).parts || []).filter(function (x) { return x.inlineData; })[0];
  if (!part) throw new Error('TTS returned no audio');
  var m = /rate=(\d+)/.exec(part.inlineData.mimeType || '');
  return { data: part.inlineData.data, rate: m ? Number(m[1]) : 24000, mime: part.inlineData.mimeType };
}

function ttsSafe_(cfg, text) {
  if (!cfg.ttsOn || !text) return null;
  var cache = CacheService.getScriptCache();
  var flag = 'tts_off_' + londonStamp_().slice(0, 10);
  if (cache.get(flag)) return null; // quota used up today: skip the slow failing call
  try { return tts_(cfg, text); }
  catch (err) {
    if (/Gemini 429/.test(String(err.message))) cache.put(flag, '1', 21600);
    console.warn('TTS fallback: ' + String(err.message).slice(0, 200));
    return null;
  }
}

function greetingAudio_(cfg, text) { // cached in 90 KB chunks for 6 h, so each session does not spend a TTS request
  if (!cfg.ttsOn) return null;
  var cache = CacheService.getScriptCache();
  var tag = 'greet_' + TTS_STYLE_VERSION + '_' + cfg.ttsModel + '_' + cfg.ttsVoice; // Gemini note 1: new style text -> new cache
  var meta = cache.get(tag);
  if (meta) {
    var parts = meta.split(':'), n = Number(parts[0]), keys = [];
    for (var i = 0; i < n; i++) keys.push(tag + '_' + i);
    var got = cache.getAll(keys), data = '';
    for (var j = 0; j < n; j++) { if (!got[keys[j]]) { data = null; break; } data += got[keys[j]]; }
    if (data) return { data: data, rate: Number(parts[1]) };
  }
  var a = ttsSafe_(cfg, text);
  if (!a) return null;
  var chunks = {}, size = 90000, count = Math.ceil(a.data.length / size);
  for (var k = 0; k < count; k++) chunks[tag + '_' + k] = a.data.slice(k * size, (k + 1) * size);
  try { cache.putAll(chunks, 21600); cache.put(tag, count + ':' + a.rate, 21600); } catch (e) {}
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
