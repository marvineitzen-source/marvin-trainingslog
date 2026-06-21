/**
 * Hyper-Coaching · Trainingslog — Sheet-Sync + Oura-Pull + Coach-Chat (Version 7)
 *
 * Tab 1 (Untitled): Trainingssätze (append).
 * Tab "Koerper": 1 Zeile/Tag (Upsert by Datum) mit Gewicht, Schlaf, Tiefschlaf, REM,
 *   Effizienz, HRV, Ruhepuls, Atemfrequenz, Temp-Abweichung, Readiness, Stress-Min,
 *   Recovery-Min, Schritte, TDEE, Trainingsdauer, Trainings-kcal, Energie, Notiz.
 *
 * pullOura (Trigger 09:00): Recovery-Marker letzte Nacht + Aktivitaet/Stress je Tag.
 *
 * Coach-Chat: doGet liefert eine same-origin Chat-Seite (HtmlService). Die App auf
 *   GitHub Pages bindet sie als iframe ein. coachChat() baut System-Kontext aus
 *   COACH_WISSEN + letzten Tagen Sheet-Daten und ruft OpenAI GPT-5.5.
 *
 * Setup: Skripteigenschaften OURA_TOKEN, OPENAI_API_KEY und COACH_PIN setzen, dann `setup` einmal ausfuehren.
 */

const SHEET_ID = '1TKdLxnRlmNg0YYr5bubD5HByE3IXw2EsO8fKq9eXfJ8';
const TOKEN = 'hc-9f3a7c21';

const KOERPER_HEADER = ['Datum','Koerpergewicht (kg)','Schlaf (h)','Tiefschlaf (h)','REM (h)','Effizienz (%)','HRV','Ruhepuls','Atemfrequenz','Temp-Abw (C)','Readiness','Stress-Min','Recovery-Min','Schritte','TDEE (kcal)','Trainingsdauer (min)','Trainings-kcal','Energie','Notiz'];
const COL = { kg:2, schlaf:3, tiefschlaf:4, rem:5, effizienz:6, hrv:7, rhr:8, atem:9, temp:10, readiness:11, stress:12, recovery:13, schritte:14, tdee:15, dauer:16, trainingskcal:17, energie:18, notiz:19 };

// ============================ SYNC (doPost) ============================

function doPost(e) {
  if (!e || !e.postData) { pullOura(); return json({ ok: true }); }
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.token !== TOKEN) return json({ ok: false, error: 'auth' });
    if (d.type === 'body') {
      upsertKoerper(d.datum, { kg:d.gewicht, schlaf:d.schlaf, energie:d.energie, dauer:d.dauer, trainingskcal:d.trainingskcal, notiz:d.notiz });
    } else {
      SpreadsheetApp.openById(SHEET_ID).getSheets()[0]
        .appendRow([d.datum, d.tag, d.uebung, d.muskel, d.satz, d.gewicht, d.reps, d.rir, d.e1rm, d.notiz || '']);
    }
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function upsertKoerper(datum, v) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sb = ss.getSheetByName('Koerper');
  if (!sb) sb = ss.insertSheet('Koerper', ss.getNumSheets());
  if (sb.getLastRow() === 0) sb.appendRow(KOERPER_HEADER);
  const dates = sb.getRange(1, 1, sb.getLastRow(), 1).getValues();
  let row = -1;
  for (let i = 1; i < dates.length; i++) { if (dstr(dates[i][0]) === String(datum)) { row = i + 1; break; } }
  if (row === -1) { row = sb.getLastRow() + 1; sb.getRange(row, 1).setValue(datum); }
  Object.keys(COL).forEach(function (k) {
    if (v[k] !== undefined && v[k] !== null && v[k] !== '') sb.getRange(row, COL[k]).setValue(v[k]);
  });
}

function dstr(x) {
  if (x instanceof Date) return Utilities.formatDate(x, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(x);
}

function h(sec) { return sec ? Math.round(sec / 3600 * 10) / 10 : ''; }
function m(sec) { return (sec === undefined || sec === null) ? '' : Math.round(sec / 60); }

// ============================ OURA-PULL ============================

function pullOura() {
  const token = PropertiesService.getScriptProperties().getProperty('OURA_TOKEN');
  if (!token) throw new Error('OURA_TOKEN fehlt.');
  const tz = Session.getScriptTimeZone();
  const end = Utilities.formatDate(new Date(Date.now() + 86400000), tz, 'yyyy-MM-dd');
  const from = Utilities.formatDate(new Date(Date.now() - 3 * 86400000), tz, 'yyyy-MM-dd');
  const base = 'https://api.ouraring.com/v2/usercollection/';
  const opt = { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true };
  function get(ep) { return JSON.parse(UrlFetchApp.fetch(base + ep + '?start_date=' + from + '&end_date=' + end, opt).getContentText()); }

  const sleep = get('sleep');
  let periods = (sleep.data || []).filter(function (p) { return p.type === 'long_sleep'; });
  if (!periods.length) periods = sleep.data || [];
  const s = periods[periods.length - 1];
  const readiness = get('daily_readiness').data || [];
  if (s) {
    const rd = readiness.filter(function (x) { return x.day === s.day; }).pop() || {};
    upsertKoerper(s.day, {
      schlaf: h(s.total_sleep_duration), tiefschlaf: h(s.deep_sleep_duration), rem: h(s.rem_sleep_duration),
      effizienz: s.efficiency || '', hrv: s.average_hrv || '', rhr: s.lowest_heart_rate || '', atem: s.average_breath || '',
      temp: (rd.temperature_deviation === undefined || rd.temperature_deviation === null) ? '' : rd.temperature_deviation,
      readiness: rd.score || ''
    });
  }

  (get('daily_activity').data || []).forEach(function (a) {
    if (a.day) upsertKoerper(a.day, { schritte: a.steps || '', tdee: a.total_calories || '' });
  });
  (get('daily_stress').data || []).forEach(function (st) {
    if (st.day) upsertKoerper(st.day, { stress: m(st.stress_high), recovery: m(st.recovery_high) });
  });
}

function installOuraTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'pullOura') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('pullOura').timeBased().everyDays(1).atHour(9).create();
}

function setup() { pullOura(); installOuraTrigger(); }

// ============================ COACH-CHAT (GPT-5.5) ============================

// Verdichtete Wissensbasis = "Brain" des Coaches (Quelle: Coach_Wissen.md).
const COACH_WISSEN = [
'Du bist Marvins persoenlicher, evidenzbasierter Natural-Bodybuilding-Coach. Antworte direkt, konkret,',
'auf seine Daten bezogen, ohne Bro-Science. Quelle der Prinzipien: Schoenfeld, Israetel/RP, Wolf, Kassiano,',
'Beardsley, Helms, Altini, Stevenson, Walker/Parsley.',
'',
'## Marvins Profil & Ziel',
'- Fortgeschrittener Natural (~12-13 J. Training). Ausdauer-Vergangenheit (Rennrad) -> gute Regeneration/NEAT, Cardio via Oura.',
'- Ziel: Hypertrophie-Spezialisierung auf Arme & Schultern. Rest auf Erhalt. Beine 1x/Woche (Quad/Beinstrecker + Waden = Schwachpunkt).',
'- Schulter = hartnaeckigster Schwachpunkt. Bekommt NIE Muskelkater in der Schulter -> mechanisch erwartbar, KEIN Defizit-Marker.',
'  Steuerung ueber Leistung/e1RM + Pump + Massband, NICHT ueber Kater.',
'- Koerpergewicht ~63-64 kg. Selbststaendig (Zeit-/Stress-Faktor). Trainiert auch am WE; ein WE-Tag lockeres Rennrad (Z2).',
'- Modus: mit bestem Muskelerhalt diaeten, so lean wie moeglich aufbauen.',
'',
'## Hypertrophie-Kern (Schoenfeld, Wolf)',
'- Mechanische Spannung = primaerer Treiber. Kein Pump/Kater-Jagen.',
'- Stretch-Mediated Hypertrophy: Spannung in gedehnter Position ueberproportional wirksam (Wolf 2023; Maeo Overhead-Trizeps +29% langer Kopf). Lengthened Partials bei Isolation/Maschinen.',
'- Uebungsauswahl (Kassiano): verschiedene Winkel pro Woche; Pool aber 4-12 Wochen stabil halten (progressiver Overload).',
'- Pro Kopf: Trizeps langer Kopf = Ueberkopf; lateral = enge Drueck/Pushdown. Bizeps langer Kopf = Incline; kurzer = Preacher. Brachialis = Hammer.',
'',
'## Volumen-Landmarks (Israetel/RP) — Saetze/Woche (MV/MEV/MAV/MRV, Freq.)',
'- Seitl. Schulter: 6/8/9-24/25-40, 3+x',
'- Hint. Schulter: 0/6/7-17/18-35, 2-5x',
'- Vord. Schulter: gedeckt durch Druecken -> kein Frontheben',
'- Bizeps: 4/8/9-19/20-35, 2-3x   Trizeps: 4/6/7-19/20+, 2-6x',
'- Brust: 4/6/7-19/20-35, 2-3x    Ruecken: 6/10/11-19/20-35, 2-4x',
'- Quad: 6/8/9-17/18-30, 2-3x     Waden: 0/2-8/9-19/20+, 2-6x',
'- Pro Einheit ~6-8 harte Saetze/Muskel (2-min-Pausen), dann Junk Volume. Volume Cycling: Start MEV, +1-2 Saetze/Wo bis MAV, bei Leistungseinbruch (MRV) -> Deload.',
'',
'## Intensitaet / Naehe zum Versagen (Beardsley, Helms, Refalo)',
'- Sweet Spot 1-2 RIR; Isolation/leichte Lasten naeher 0 RIR. Schwer (>80% 1RM) braucht kein Versagen.',
'- Trainierte profitieren oft davon, vor dem Versagen zu stoppen (weniger ZNS-Ermuedung). 6-12 WH effizient, 5-30 alle wirksam nah am Versagen.',
'',
'## Recovery (Altini) — Abweichung von der persoenlichen 7-Tage-Baseline, nicht Absolutwerte',
'- GRUEN trainieren: HRV stabil/steigend, Ruhepuls normal/sinkend, Schlaf tief, keine Infektzeichen.',
'- GELB deload: HRV leicht runter, Ruhepuls hoch, CV hoch, Schlaf unruhig, Leistung stagniert -> Volumen/Intensitaet runter.',
'- ROT pause: HRV dauerhaft niedrig + Ruhepuls hoch + CV niedrig, Infekte, Kraftverlust.',
'- HRV reagiert extrem sensibel auf Atemwegsinfekt & Alkohol. Temp-Abw. hoch + Atemfrequenz hoch = Fruehwarnung.',
'',
'## Schlaf (Walker, Parsley)',
'- Tiefschlaf = ~70% der GH-Ausschuettung + Testosteron -> wichtigster muskelrelevanter Marker.',
'- 8h werktags = Floor fuer das Spezialisierungs-Volumen. Unter 8h: bis -10-30% Leistung. Unter 6h: Muskelabbau, Testosteron runter.',
'- Banking Sleep vor harten Bloecken; Naps <30 min. Eisbaeder direkt nach Kraft NICHT (daempfen Hypertrophie-Signal).',
'',
'## Ernaehrung & Raten',
'- Protein 1,6-2,2 g/kg/Tag, verteilt auf 3-5 Mahlzeiten.',
'- Lean-Aufbau +0,1-0,25 kg/Woche; Diaet mit Muskelerhalt -0,3-0,7 kg/Woche (~0,5-1%/Wo).',
'- TDEE-Abgleich: Oura-TDEE + Trainings-kcal (MET) vs. Kalorien -> erwartete Aenderung; Wochentrend kalibriert nach. Schaetzung muss nur konsistent sein.',
'- Gewichtsschwankungen = Wasser/Glykogen/Darm; nur 7-Tage-Schnitt bewerten, nie Einzeltage. 3-Tage-Regel vor Anpassungen.',
'',
'## Tracking & Entscheidungen (Stevenson) — emotionsfrei',
'- Lead Indicators (steuerbar): Kalorien, Protein, Volumen/RIR, Schlaf. Lag (Resultat): Gewicht, Masse, Kraft, Spiegel.',
'- Trainings-kcal-MET: Beine 6,0; Arme/Schulter 4,0-4,5. kcal = MET x kg x h.',
'- Bei Stagnation Kette: 1) Erholung optimieren 2) Deload 3) Volumen anpassen 4) Uebung rotieren 5) erst dann Programm wechseln.',
'',
'## Antwort-Stil',
'- Beziehe dich konkret auf die Live-Daten (HRV/Readiness/Tiefschlaf/Gewichtstrend/TDEE), wenn vorhanden.',
'- Klare, umsetzbare Empfehlung + kurze Begruendung (welches Prinzip). Keine medizinischen Diagnosen; bei Schmerzen/Krankheit zum Arzt verweisen.',
'- Bei Schulter: Kater ist KEIN Marker; Stretch-Loading (Kabel-Seitheben Lean-away) + Volumen-Rampe + Leistung steuern.',
'- Antworte auf Deutsch, kompakt (mobil), mit konkreten Zahlen statt Floskeln.'
].join('\n');

// Live-Kontext aus dem Sheet (letzte Tage Koerper + letzte Trainingssaetze).
function buildCoachContext() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const out = [];
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  out.push('Heutiges Datum: ' + today);

  const sb = ss.getSheetByName('Koerper');
  if (sb && sb.getLastRow() > 1) {
    const last = sb.getLastRow();
    const n = Math.min(14, last - 1);
    const rows = sb.getRange(last - n + 1, 1, n, 19).getValues();
    out.push('');
    out.push('### Koerper/Recovery (letzte ' + n + ' Tage, neueste zuletzt):');
    rows.forEach(function (r) {
      const parts = [];
      parts.push(dstr(r[0]));
      if (r[1] !== '') parts.push(r[1] + 'kg');
      if (r[2] !== '') parts.push('Schlaf ' + r[2] + 'h (Tief ' + (r[3] || '?') + ')');
      if (r[6] !== '') parts.push('HRV ' + r[6]);
      if (r[7] !== '') parts.push('RHR ' + r[7]);
      if (r[9] !== '') parts.push('TempAbw ' + r[9]);
      if (r[10] !== '') parts.push('Readiness ' + r[10]);
      if (r[14] !== '') parts.push('TDEE ' + r[14]);
      if (r[16] !== '') parts.push('Train-kcal ' + r[16]);
      if (r[17] !== '') parts.push('Energie ' + r[17]);
      out.push('- ' + parts.join(' · '));
    });
  }

  const tr = ss.getSheets()[0];
  if (tr && tr.getLastRow() > 1) {
    const lastT = tr.getLastRow();
    const nt = Math.min(30, lastT - 1);
    const tv = tr.getRange(lastT - nt + 1, 1, nt, 9).getValues();
    out.push('');
    out.push('### Letzte Trainingssaetze (Datum Tag · Uebung (Muskel) Satz: Gewicht x WH @RIR e1RM):');
    tv.forEach(function (r) {
      out.push('- ' + dstr(r[0]) + ' ' + (r[1] || '') + ' · ' + r[2] + ' (' + r[3] + ') S' + r[4] + ': ' + r[5] + 'kg x ' + r[6] + ' @' + r[7] + 'RIR e1RM' + r[8]);
    });
  }
  return out.join('\n');
}

// Wird vom Chat (google.script.run) aufgerufen. PIN-Gate schuetzt vor fremder Nutzung.
function coachChat(message, history, pin) {
  const props = PropertiesService.getScriptProperties();
  const savedPin = props.getProperty('COACH_PIN');
  if (!savedPin) return 'Setup: Bitte Skripteigenschaft COACH_PIN setzen.';
  if (String(pin || '') !== savedPin) return '__PINFAIL__';
  const key = props.getProperty('OPENAI_API_KEY');
  if (!key) return 'Fehler: OPENAI_API_KEY fehlt in den Skripteigenschaften.';
  const sys = COACH_WISSEN + '\n\n## Aktuelle Live-Daten (aus dem Sheet)\n' + buildCoachContext();
  const msgs = [{ role: 'system', content: sys }];
  (history || []).forEach(function (mm) {
    if (mm && mm.role && mm.content) msgs.push({ role: mm.role, content: String(mm.content) });
  });
  msgs.push({ role: 'user', content: String(message || '') });

  const payload = {
    model: 'gpt-5.5',
    messages: msgs,
    max_completion_tokens: 900,
    reasoning_effort: 'medium'
  };
  const res = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + key },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  let data;
  try { data = JSON.parse(res.getContentText()); } catch (e) { return 'Fehler: ungueltige Antwort (' + code + ')'; }
  if (code !== 200) return 'Fehler ' + code + ': ' + ((data.error && data.error.message) || res.getContentText());
  return data.choices && data.choices[0] && data.choices[0].message ? data.choices[0].message.content : 'Keine Antwort.';
}

// ============================ WEB (doGet = Chat-Seite) ============================

function doGet() {
  return HtmlService.createHtmlOutput(CHAT_HTML)
    .setTitle('Hyper-Coaching · Coach')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

const CHAT_HTML = [
'<!DOCTYPE html><html lang="de"><head><meta charset="UTF-8"></head><body>',
'<style>',
'*{box-sizing:border-box;-webkit-tap-highlight-color:transparent}',
'html,body{margin:0;height:100%;background:#0B0E14;color:#EAecef;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Inter,Roboto,sans-serif;-webkit-font-smoothing:antialiased}',
'#wrap{display:flex;flex-direction:column;height:100vh;max-width:520px;margin:0 auto}',
'#log{flex:1;overflow-y:auto;padding:14px 12px 6px}',
'.msg{margin:8px 0;display:flex}',
'.msg.u{justify-content:flex-end}',
'.b{max-width:84%;padding:10px 12px;border-radius:14px;font-size:14px;line-height:1.5;white-space:pre-wrap;word-wrap:break-word}',
'.u .b{background:#7C83F7;color:#fff;border-bottom-right-radius:5px}',
'.a .b{background:#1B2230;border:1px solid rgba(255,255,255,.07);border-bottom-left-radius:5px}',
'.sys{font-size:12px;color:#8B93A7;text-align:center;padding:10px 16px;line-height:1.5}',
'.sugg{display:flex;flex-wrap:wrap;gap:6px;padding:0 12px 8px}',
'.sugg button{background:#141925;border:1px solid rgba(255,255,255,.12);color:#EAecef;font-size:12px;padding:7px 11px;border-radius:999px;cursor:pointer}',
'#bar{display:flex;gap:8px;padding:10px 12px;border-top:1px solid rgba(255,255,255,.07);background:rgba(11,14,20,.95)}',
'#in{flex:1;background:#1B2230;border:1px solid rgba(255,255,255,.12);border-radius:12px;color:#EAecef;padding:11px 12px;font-size:15px;outline:none;resize:none;max-height:120px}',
'#send{width:46px;border:0;border-radius:12px;background:#7C83F7;color:#fff;font-size:20px;cursor:pointer}',
'#send:disabled{opacity:.5}',
'.dot{display:inline-block;width:6px;height:6px;margin:0 1px;border-radius:50%;background:#8B93A7;animation:bl 1s infinite}',
'.dot:nth-child(2){animation-delay:.2s}.dot:nth-child(3){animation-delay:.4s}',
'@keyframes bl{0%,80%,100%{opacity:.3}40%{opacity:1}}',
'#lock{position:fixed;inset:0;background:#0B0E14;z-index:50;display:none;flex-direction:column;align-items:center;justify-content:center;padding:24px;text-align:center}',
'#lock input{width:200px;text-align:center;background:#1B2230;border:1px solid rgba(255,255,255,.12);border-radius:12px;color:#EAecef;padding:12px;font-size:18px;letter-spacing:3px;outline:none}',
'#lock button{margin-top:12px;width:200px;padding:12px;border:0;border-radius:12px;background:#7C83F7;color:#fff;font-size:15px;font-weight:500;cursor:pointer}',
'</style>',
'<div id="wrap">',
'<div id="log">',
'<div class="sys">Dein Coach · Zugriff auf dein Wissens-Brain + Live-Daten (HRV, Schlaf, Gewicht, Training). Frag alles zu Training, Recovery, Ernaehrung.</div>',
'<div class="sugg">',
'<button onclick="quick(this)">Wie soll ich heute trainieren?</button>',
'<button onclick="quick(this)">Passt mein Gewichtstrend?</button>',
'<button onclick="quick(this)">Was esse ich heute?</button>',
'<button onclick="quick(this)">Schulter-Volumen ok?</button>',
'</div>',
'</div>',
'<div id="bar">',
'<textarea id="in" rows="1" placeholder="Frag deinen Coach..." oninput="grow()" onkeydown="kd(event)"></textarea>',
'<button id="send" onclick="sendMsg()">&#10148;</button>',
'</div>',
'</div>',
'<div id="lock"><div style="font-size:34px;margin-bottom:6px;">🔒</div><div style="font-size:16px;font-weight:600;margin-bottom:4px;">Hyper-Coaching</div><div id="lockmsg" style="font-size:13px;color:#8B93A7;margin-bottom:16px;">PIN eingeben</div><input id="pin" type="password" inputmode="numeric" placeholder="PIN" onkeydown="if(event.key===\'Enter\')unlock()"/><button onclick="unlock()">Entsperren</button></div>',
'<script>',
'var hist=[];var busy=false;',
'var logEl=document.getElementById("log");var inEl=document.getElementById("in");var sendEl=document.getElementById("send");',
'function getPin(){return localStorage.getItem("hc_pin")||"";}',
'function showLock(msg){document.getElementById("lockmsg").textContent=msg||"PIN eingeben";document.getElementById("lock").style.display="flex";setTimeout(function(){var p=document.getElementById("pin");if(p)p.focus();},100);}',
'function unlock(){var p=document.getElementById("pin").value.trim();if(!p)return;localStorage.setItem("hc_pin",p);document.getElementById("pin").value="";document.getElementById("lock").style.display="none";}',
'function grow(){inEl.style.height="auto";inEl.style.height=Math.min(inEl.scrollHeight,120)+"px";}',
'function kd(e){if(e.key==="Enter"&&!e.shiftKey){e.preventDefault();sendMsg();}}',
'function quick(b){inEl.value=b.textContent;sendMsg();}',
'function add(role,text){var d=document.createElement("div");d.className="msg "+(role==="user"?"u":"a");var b=document.createElement("div");b.className="b";b.textContent=text;d.appendChild(b);logEl.appendChild(d);logEl.scrollTop=logEl.scrollHeight;return b;}',
'function sendMsg(){if(busy)return;var t=inEl.value.trim();if(!t)return;inEl.value="";grow();add("user",t);busy=true;sendEl.disabled=true;',
'var th=document.createElement("div");th.className="msg a";th.innerHTML="<div class=\\"b\\"><span class=\\"dot\\"></span><span class=\\"dot\\"></span><span class=\\"dot\\"></span></div>";logEl.appendChild(th);logEl.scrollTop=logEl.scrollHeight;',
'google.script.run.withSuccessHandler(function(r){th.remove();busy=false;sendEl.disabled=false;if(r==="__PINFAIL__"){localStorage.removeItem("hc_pin");showLock("PIN falsch — bitte erneut eingeben");return;}add("assistant",r);hist.push({role:"user",content:t});hist.push({role:"assistant",content:r});if(hist.length>12)hist=hist.slice(hist.length-12);}).withFailureHandler(function(e){th.remove();add("assistant","Fehler: "+e.message);busy=false;sendEl.disabled=false;}).coachChat(t,hist,getPin());}',
'if(!getPin())showLock();',
'<\/script>',
'</body></html>'
].join('\n');

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
