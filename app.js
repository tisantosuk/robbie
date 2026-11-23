'use strict';
// Robbie v3.1 phone page. No secrets in this file: the backend URL and passphrase arrive once in the link (#u=...&p=...)
// and are kept in this phone's browser storage.

const $ = (id) => document.getElementById(id);
const SAMPLE_RATE = 16000;
const MAX_SEGMENT_S = 30;
const state = {
  cfg: null, sessionId: null, vad: null, speaking: false, closed: true,
  queue: Promise.resolve(), wakeLock: null, pending: [], history: [], turns: 0,
  voices: [], voiceIndex: 0, cancelSpeech: null, audioCtx: null, clientLog: [],
};

function note(msg) { // v3.4 diagnostics: sent with the next request and written to the Turns tab
  state.clientLog.push(new Date().toISOString().slice(11, 19) + ' ' + msg);
  if (state.clientLog.length > 10) state.clientLog.shift();
}
function takeNotes() { const n = state.clientLog; state.clientLog = []; return n; }

function setMode(mode, status, hint) {
  document.body.className = mode;
  if (status !== undefined) $('status').textContent = status;
  if (hint !== undefined) $('hint').textContent = hint;
}

// ---------- configuration ----------
function loadConfig() {
  let h = null;
  try { h = new URLSearchParams(location.hash.slice(1)); } catch {}
  const u = h && h.get('u'); const p = h && h.get('p');
  if (u && p && /^(https:\/\/|http:\/\/localhost[:/])/.test(u)) {
    try { localStorage.setItem('robbie_cfg', JSON.stringify({ url: u, pass: p })); } catch {}
  }
  if (location.hash) history.replaceState(null, '', location.pathname + location.search); // keep any query string
  try { return JSON.parse(localStorage.getItem('robbie_cfg') || 'null'); } catch { return null; }
}

async function api(action, data = {}) {
  const res = await fetch(state.cfg.url, {
    method: 'POST', redirect: 'follow',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' }, // no CORS preflight (Apps Script cannot answer one)
    body: JSON.stringify({ action, passphrase: state.cfg.pass, ...data }),
  });
  if (!res.ok) throw new Error('HTTP ' + res.status);
  const out = await res.json();
  if (out.error) throw new Error(out.error);
  return out;
}

// ---------- voice out (phone's own pt-BR voice) ----------
function loadVoices() {
  const all = window.speechSynthesis ? speechSynthesis.getVoices() : [];
  const br = all.filter((v) => /^pt[-_]BR/i.test(v.lang));
  state.voices = br.length ? br : all.filter((v) => /^pt/i.test(v.lang));
  try { state.voiceIndex = Math.min(Number(localStorage.getItem('robbie_voice') || 0), Math.max(0, state.voices.length - 1)); } catch {}
}
if (window.speechSynthesis) { loadVoices(); speechSynthesis.onvoiceschanged = loadVoices; }

function unlockSpeech() { // must run inside the start tap
  try { // Gemini voice playback (V-1): the audio context must be unlocked by a tap on Android
    if (!state.audioCtx) state.audioCtx = new (window.AudioContext || window.webkitAudioContext)();
    if (state.audioCtx.state === 'suspended') state.audioCtx.resume();
  } catch {}
  if (!window.speechSynthesis) return;
  speechSynthesis.cancel();
  const u = new SpeechSynthesisUtterance(' '); u.volume = 0; u.lang = 'pt-BR';
  speechSynthesis.speak(u);
}

function stopPlayback() {
  if (window.speechSynthesis) speechSynthesis.cancel();
  if (state.cancelSpeech) state.cancelSpeech(); // Android may not fire onend after cancel()
}

function pcmBuffer(b64, rate, format) { // 16-bit signed little-endian mono PCM (or a WAV file of it) -> AudioBuffer
  const bin = atob(b64);
  let off = 0;
  if (format === 'wav' || bin.slice(0, 4) === 'RIFF') { const d = bin.indexOf('data', 12); off = d >= 0 ? d + 8 : 44; }
  const n = (bin.length - off) >> 1;
  const buf = state.audioCtx.createBuffer(1, n, rate || 24000);
  const ch = buf.getChannelData(0);
  for (let i = 0; i < n; i++) {
    let v = bin.charCodeAt(off + 2 * i) | (bin.charCodeAt(off + 2 * i + 1) << 8);
    if (v >= 32768) v -= 65536;
    ch[i] = v / 32768;
  }
  return buf;
}

// Gemini review 1053, Defect 1: resolves true when the Gemini audio has played (or was stopped by a tap),
// false when it could not start, so say() can use the phone voice instead of a silent turn.
async function playGemini(b64, rate, format) {
  try {
    const ctx = state.audioCtx;
    if (!ctx) { note('gemini-audio: no audio context'); return false; }
    if (ctx.state !== 'running') {
      note('audio context ' + ctx.state + ', resuming');
      try { await Promise.race([ctx.resume(), new Promise((ok) => setTimeout(ok, 800))]); } catch {}
      if (ctx.state !== 'running') { note('audio context still ' + ctx.state + ': phone voice'); return false; }
    }
    const buf = pcmBuffer(b64, rate, format);
    const src = ctx.createBufferSource();
    src.buffer = buf; src.connect(ctx.destination);
    note('gemini-audio start ' + buf.duration.toFixed(1) + 's');
    return await new Promise((resolve) => {
      let ended = false;
      const timer = setTimeout(() => finish('timer'), buf.duration * 1000 + 2000);
      function finish(why) { if (ended) return; ended = true; clearTimeout(timer); state.cancelSpeech = null; note('gemini-audio end ' + why); resolve(why !== 'start-error'); }
      src.onended = () => finish('ended');
      state.cancelSpeech = () => { try { src.stop(); } catch {} finish('tap'); };
      try { src.start(); } catch (e) { note('gemini-audio start-error ' + e.message); finish('start-error'); }
    });
  } catch (e) { note('gemini-audio error ' + e.message); return false; }
}

async function say(text, audio, rate, format) {
  if (!text) return;
  state.speaking = true;
  if (state.vad) await state.vad.pause(); // microphone off while Robbie speaks: no self-interruption
  setMode('speaking', 'Robbie falando', 'Toque na tela para interromper.');
  $('last').textContent = 'Robbie: ' + text;
  try {
    const played = audio ? await playGemini(audio, rate, format) : false;
    if (!played && !state.interrupted) await phoneVoice(text);
  } finally {
    state.speaking = false; state.interrupted = false;
    if (!state.closed && state.vad) { await state.vad.start(); setMode('listening', 'Ouvindo', 'Quando terminar, diga "o que você acha?" ou "sua vez".'); }
  }
}

function phoneVoice(text) {
  return new Promise((resolve) => {
    const timer = setTimeout(() => done('timer'), 3000 + text.split(' ').length * 600); // safety: some phones skip onend
    function done(why) { clearTimeout(timer); state.cancelSpeech = null; note('phone-voice end ' + why); resolve(); }
    state.cancelSpeech = () => done('tap');
    if (!window.speechSynthesis) return done('no speechSynthesis');
    const u = new SpeechSynthesisUtterance(text);
    u.lang = 'pt-BR';
    if (state.voices[state.voiceIndex]) u.voice = state.voices[state.voiceIndex];
    u.onstart = () => note('phone-voice start');
    u.onend = () => done('ended'); u.onerror = (e) => done('error ' + (e.error || ''));
    note('phone-voice speak voices=' + state.voices.length + ' pending=' + speechSynthesis.pending + ' speaking=' + speechSynthesis.speaking);
    if (speechSynthesis.paused) speechSynthesis.resume();
    speechSynthesis.speak(u);
  });
}

document.addEventListener('click', (e) => {
  if (state.speaking && !e.target.closest('footer')) { state.interrupted = true; stopPlayback(); } // tap anywhere to interrupt
});

// ---------- voice in ----------
function chunks(audio) { // CH-31-A: at most 30 s per request
  const max = SAMPLE_RATE * MAX_SEGMENT_S;
  const out = [];
  for (let i = 0; i < audio.length; i += max) out.push(audio.subarray(i, i + max));
  return out;
}

function sendSegment(audio) {
  const utils = window.vad && window.vad.utils; // Gemini review 1445, Defect 2: no implicit global
  if (!utils) { setMode('listening', 'Ouvindo', 'Falha ao preparar o áudio. Recarregue a página.'); return; }
  for (const part of chunks(audio)) {
    const b64 = utils.arrayBufferToBase64(utils.encodeWAV(part, 1, SAMPLE_RATE, 1, 16));
    state.queue = state.queue.then(() => processSegment(b64));
  }
}

async function processSegment(b64) {
  if (state.closed) return;
  setMode('thinking', 'Pensando', undefined);
  const t0 = performance.now();
  try {
    let r = await api('turn', { audio: b64, pending: state.pending, history: state.history, sessionId: state.sessionId, clientLog: takeNotes() });
    if (r.busy) { // Gemini briefly overloaded: one more try with the same recording
      await new Promise((ok) => setTimeout(ok, 1500));
      r = await api('turn', { audio: b64, pending: state.pending, history: state.history, sessionId: state.sessionId, clientLog: takeNotes() });
      if (r.busy) { setMode('listening', 'Ouvindo', 'O Gemini está sobrecarregado agora. Pode repetir a última frase?'); return; }
    }
    const ms = Math.round(performance.now() - t0);
    $('timing').textContent = `Ida e volta: ${(ms / 1000).toFixed(1)} s` + (r.geminiMs ? ` (Gemini ${(r.geminiMs / 1000).toFixed(1)} s` + (r.ttsMs ? `, voz ${(r.ttsMs / 1000).toFixed(1)} s` : '') + ')' : '') + (r.ttsFallback ? ' · voz reserva' : '');
    if (r.ttsFallback) console.warn('Robbie voice fallback:', r.ttsFallback);
    state.pending = r.pending || [];
    state.history = r.history || state.history;
    if (r.transcript) $('last').textContent = 'Você: ' + r.transcript;
    if (r.closed) state.closed = true; // keep the microphone off after the goodbye
    if (r.speak) { state.turns += 1; await say(r.text, r.audio, r.audioRate, r.audioFormat); }
    else if (!state.closed) setMode('listening', 'Ouvindo', 'Pode continuar. Robbie espera o seu sinal.');
    if (r.closed) await finish(r.reason === 'daily_limit' ? 'daily_limit' : 'voice');
  } catch (err) {
    setMode('listening', 'Ouvindo', 'Falha: ' + err.message + '. Pode repetir.');
  }
}

async function keepAwake() {
  try { if ('wakeLock' in navigator) state.wakeLock = await navigator.wakeLock.request('screen'); } catch {}
}
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible' && !state.closed) keepAwake(); });

$('start').addEventListener('click', async () => {
  unlockSpeech();
  $('start').classList.add('hidden');
  setMode('thinking', 'Preparando', 'Permita o uso do microfone.');
  try {
    await keepAwake();
    if (!(window.vad && window.vad.MicVAD && window.vad.utils)) throw new Error('áudio não carregou; recarregue a página');
    if (!state.vad) {
      state.vad = await window.vad.MicVAD.new({
        model: 'v5', baseAssetPath: new URL('vad/', location.href).href, onnxWASMBasePath: new URL('ort/', location.href).href,
        positiveSpeechThreshold: 0.5, negativeSpeechThreshold: 0.35,
        redemptionMs: 1400, preSpeechPadMs: 300, minSpeechMs: 400, startOnLoad: false,
        getStream: () => navigator.mediaDevices.getUserMedia({
          audio: { channelCount: 1, echoCancellation: true, noiseSuppression: true, autoGainControl: true },
        }),
        onSpeechStart: () => { if (!state.speaking) setMode('speech', 'Ouvindo você', undefined); },
        onVADMisfire: () => { if (!state.speaking) setMode('listening', 'Ouvindo', undefined); },
        onSpeechEnd: (audio) => sendSegment(audio),
      });
    }
    const t0 = performance.now();
    const s = await api('start', { hour: new Date().getHours() }); // greeting by the phone's own time of day
    $('timing').textContent = `Início: ${((performance.now() - t0) / 1000).toFixed(1)} s`;
    state.sessionId = (crypto.randomUUID && crypto.randomUUID()) || String(Date.now());
    state.pending = []; state.history = []; state.turns = 0; state.closed = false;
    $('memory').textContent = s.memoryValue ? 'Memória: ' + s.memoryValue : '';
    $('controls').classList.remove('hidden');
    await say(s.greeting, s.audio, s.audioRate, s.audioFormat);
  } catch (err) {
    setMode('', 'Não consegui começar', String(err.message || err));
    $('start').classList.remove('hidden');
  }
});

async function finish(reason) {
  state.closed = true;
  if (state.vad) { await state.vad.pause(); }
  if (state.wakeLock) { state.wakeLock.release().catch(() => {}); state.wakeLock = null; }
  $('controls').classList.add('hidden');
  let msg = 'Conversa salva. Até a próxima.';
  try { await api('end', { sessionId: state.sessionId, turns: state.turns, reason, clientLog: takeNotes() }); }
  catch { msg = 'Encerrado (não consegui confirmar o salvamento).'; }
  setMode('', msg, '');
  $('start').textContent = 'Começar de novo';
  $('start').classList.remove('hidden');
}

$('end').addEventListener('click', () => { stopPlayback(); finish('button'); });

$('voice').addEventListener('click', async () => {
  if (!state.voices.length) { await say('Este celular tem só uma voz em português.'); return; }
  state.voiceIndex = (state.voiceIndex + 1) % state.voices.length;
  try { localStorage.setItem('robbie_voice', String(state.voiceIndex)); } catch {}
  await say(`Esta é a voz ${state.voiceIndex + 1} de ${state.voices.length}.`);
});

window.addEventListener('pagehide', () => {
  if (!state.closed && state.cfg && navigator.sendBeacon) {
    navigator.sendBeacon(state.cfg.url, new Blob([JSON.stringify({
      action: 'end', passphrase: state.cfg.pass, sessionId: state.sessionId, turns: state.turns, reason: 'pagehide',
    })], { type: 'text/plain;charset=utf-8' }));
  }
});

// ---------- boot ----------
state.cfg = loadConfig();
if (!state.cfg) setMode('', 'Robbie ainda não configurado', 'Abra o link do QR code da página de configuração.');
else { $('start').classList.remove('hidden'); setMode('', 'Pronto', 'Toque em Começar. Depois, é só conversar.'); }
