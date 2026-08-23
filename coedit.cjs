// CO-EDITING IRAS: tartalom beirasa egy MEGNYITOTT Nextcloud Office sessionbe.
//
// MIERT EZ AZ UT (a tulajdonos rendelkezese, 2026-08-15 10:2x): a DocBuilder a fajlt a lemezen irja,
// session nelkul -- ha kozben valaki nyitva tartja, az irasunk elveszik vagy utkozik, es ezt
// ELORE NEM TUDJUK MEGNEZNI. A co-editing ut ugyanabba a sessionbe ir, amit a felhasznalo lat.
// A DocBuilder hatokore ezzel egyetlen lepesre szukult: LETREHOZAS. Ami mar letezik, ide tartozik.
//
// AMI MERT, ES AMIRE EZ A MODUL EPUL (megmerve, 2026-08-15):
//   mag                  az API objektum       PasteHtml   GetFileHTML
//   documenteditor       window.editor            van         van
//   spreadsheeteditor    window.Asc.editor        van         NINCS
//   presentationeditor   window.editor            van         NINCS
// A `PasteHtml` mindharom magban LANDOL, formazassal egyutt (tablazat, felkover, dolt, szin,
// lista, hatterszin -- a mentett OOXML-bol visszaolvasva, neg. kontrollal).
//
// *** AZ API OBJEKTUM HELYET FUTASIDOBEN KERESSUK, NEM EGETJUK BE: *** egy fix `window.editor`
// a tablazat-savon CSENDBEN undefined-ot adna -- vagyis a hiba nem hibakent jelentkezne.

const fs = require('fs')
const os = require('os')
const path = require('path')
const { execFileSync } = require('child_process')

// same pattern as lib.cjs's own checkFreshness() -- a SEPARATE
// mtime for THIS file, not a shared threshold with lib.cjs/euro-mcp.cjs. Measured (card's own
// HOL): the three files' mtimes diverge on a real fleet checkout (lib.cjs and euro-mcp.cjs
// combed at the same moment, coedit.cjs over an hour earlier) -- a single "any file changed"
// flag would either miss a stale coedit.cjs next to a fresh lib.cjs, or false-positive a fresh
// coedit.cjs just because lib.cjs happened to comb later.
const __COEDIT_LOAD_MTIME_MS = fs.statSync(__filename).mtimeMs
function checkFreshness() {
  const diskMtimeMs = fs.statSync(__filename).mtimeMs
  if (diskMtimeMs <= __COEDIT_LOAD_MTIME_MS) return { fresh: true }
  return {
    fresh: false,
    file: 'coedit.cjs',
    message:
      `coedit.cjs: a betoltott kod regebbi, mint a lemezen levo (betoltve: ${new Date(__COEDIT_LOAD_MTIME_MS).toISOString()}, ` +
      `lemezen: ${new Date(diskMtimeMs).toISOString()}) -- inditsd ujra a sessiont`,
  }
}

// --- A BRIDGE SAJAT .env-je -------------------------------------------------------------------
// Az `euro-mcp.cjs` eddig NEM toltotte be (megmerve: 0 hivatkozas benne), mert a DocBuilder-utnak
// nem kellett -- a JWT-t a box-helper olvassa a MASIK gepen. A co-editing utnak viszont a
// PER-AGENS NC-hitelesito kell, es az ITT all. A mintat a talk-bridge mcp.js-etol vesszuk at
// SZO SZERINT, egy fontos reszlettel egyutt: *** a mar beallitott kornyezeti valtozot NEM irjuk
// felul *** -- kulonben egy teszt vagy egy hivo altal szandekosan adott ertek csendben eltunne,
// es a teszt a .env-et merne a sajat bemenete helyett.
function loadBridgeEnv(env = process.env, dir = __dirname) {
  const envPath = path.join(dir, '.env')
  if (!fs.existsSync(envPath)) return { ok: false, indok: `nincs .env a bridge konyvtaraban (${envPath})` }
  let db = 0
  for (const line of fs.readFileSync(envPath, 'utf8').split('\n')) {
    const t = line.trim()
    if (!t || t.startsWith('#')) continue
    const pos = t.indexOf('=')
    if (pos < 1) continue
    const key = t.slice(0, pos).trim()
    if (!(key in env)) { env[key] = t.slice(pos + 1).trim(); db += 1 }
  }
  return { ok: true, betoltve: db }
}


// EGY kulcs kiolvasasa a bridge .env-jebol, a KORNYEZET ERINTESE NELKUL.
// Miert nem a `loadBridgeEnv` ehhez: az a TELJES fajlt betolti a process kornyezetebe, titkokkal
// egyutt. Van olyan kodut (pl. egy megtagado valasz kiegeszitese), aminek EGY nem-titkos kulcs
// kell -- ott egy teljes betoltes annyit jelentene, hogy a titkok olyan uton is beallnak, aminek
// semmi koze hozzajuk. *A titok akkor a legbiztonsagosabb, ha a legkevesebb kodut latja.*
function bridgeEnvKulcs(nev, dir = __dirname) {
  const envPath = path.join(dir, '.env')
  if (!fs.existsSync(envPath)) return null
  for (const line of fs.readFileSync(envPath, 'utf8').split('\n')) {
    const t = line.trim()
    if (!t || t.startsWith('#')) continue
    const pos = t.indexOf('=')
    if (pos < 1) continue
    if (t.slice(0, pos).trim() === nev) return t.slice(pos + 1).trim()
  }
  return null
}

// A CO-EDITING UT ALLAPOTA EGY ADOTT HIVORA -- valasz arra a kerdesre, hogy ha ide kuldjuk,
// egyaltalan be tud-e lepni. *Egy megtagadas, ami egy MASIK, szinten zart ajtora mutat, rosszabb
// a semminel: a hivo azt hiszi, van hova mennie.*
function coeditUtAllapota(env = process.env, cwd = process.cwd(), dir = __dirname) {
  const allowlist = bridgeEnvKulcs('EURO_COEDIT_AGENTS', dir)
  const r = detectCallerId({ ...env, EURO_COEDIT_AGENTS: allowlist || '' }, cwd)
  if (r.ok) return { elerheto: true, indok: `a co-editing ut nyitva all ennek a hivonak (${r.id})` }
  return { elerheto: false, indok: r.indok }
}

// --- KI A HIVO? -----------------------------------------------------------------------------
// A mintat egy mar mukodo, megmert megoldasbol vesszuk at (detectCallerId): a session cwd-je AZ
// AGENS OTTHONA (`<valamilyen root>/agents/<nev>`).
// KET PONTON SZIGORUBB VAGYOK NALA, es mindketto szandekos:
//   1. NINCS FALLBACK. A Talk-ut vissza tud esni a szoba tagsagara; egy dokumentum-szerkesztesnel
//      a visszaeses azt jelentene, hogy MASVALAKI nevében irunk egy KOZOS fajlba. Azonossag
//      nelkul ez a modul MEGTAGADJA a muveletet.
//   2. A `EURO_AGENT_ID` explicit valtozo NEM eleg onmagaban: a flotta egy unix-useren fut, tehat
//      barmelyik agens beallithatna barmit. Ezert az explicit erteknek IS at kell mennie az
//      allowlisten. *Egy parancssori/kornyezeti kapcsolo, ami a hivot nevezi meg, nem azonositas.*
//
// A .env betoltes IDE kerult be, a fuggveny SAJAT elso soraba -- nem a 10+ csupasz hivasi hely
// ele, egyenkent. MERT ok: a `coedit_write`/
// `coedit_write_operations` kezelok korabban KULON hivtak `loadBridgeEnv()`-et a
// detectCallerId() elott, es EZERT mukodtek -- de csak akkor, ha a folyamatban MAR futott ilyen
// hivas ("felmelegedett"). A tobbi ~10 hivasi hely csupaszon hivta meg, es process.env-re esett
// vissza -- ami az EXEC-IDEJU kornyezetet tukrozi, nem a .env fajl mai tartalmat. Egy helyessegi
// feltetel, amit MINDEN hivasi helynek kulon kellene emlekeznie, defektus-generator (a 11.
// hivasi hely holnap ujra elfelejtene) -- ezert a betoltes a KOZOS belepesi pontba kerul, ahol a
// sorrendi fuggoseg SZERKEZETILEG nem letezhet tobbe.
//
// *** ELSO VALTOZAT `loadBridgeEnv(env)`-et hivta itt -- ez a TELJES .env-et (titkokkal egyutt)
// tolti a process.env-be MINDEN hivasnal, azt is beleertve a 9 CSAK-OLVASO eszkozt es a
// coeditUtAllapota() statusz-hivast, amiknek soha nem kellett volna titkot latniuk. Ez pontosan
// azt a szandekot sertette meg, amit a fajl teteje ("Miert nem a loadBridgeEnv ehhez") kimond:
// "a titok akkor a legbiztonsagosabb, ha a legkevesebb kodut latja". Fuggetlen meres: egy elutasitott,
// nem-allowlistelt hivo hivasa utan is 28 kulcs allt az env-ben, koztuk NC-jelszavak. ***
// MASODIK VALTOZAT a MAR LETEZO `bridgeEnvKulcs()`-et hivta, es EGY kulcsot (EURO_COEDIT_AGENTS)
// IRT VISSZA az `env`-be -- kevesebb kitettseg, de meg mindig MUTALTA a hivo altal atadott
// objektumot. Egy harmadik, EGYSZERUBB alakot javasolt: ne irjunk semmit az
// `env`-be, csak OLVASSUK a hianyzo erteket lokalis valtozoba -- igy a fuggveny NULLA uj kulcsot
// lat a folyamat env-jeben, meg az allowlist-sztringet sem. Ugyanaz a `coeditUtAllapota()` (70.
// sor) mar igy jart el korabban is -- ez a ket ut EGYSEGESITESE, nem uj mechanizmus.
//
// *** A `||` ALAKNAK (egy korabbi javaslatnak) VOLT EGY SAJAT hibaja: egy EXPLICIT URES sztring
// (`env.EURO_COEDIT_AGENTS = ''`, amit a test_coedit_identity.cjs sajat izolacios tesztje ad,
// "URES allowlist -> senki nem megy at") falsy, tehat `||` a MASODIK agra esett volna -- a VALODI
// .env-bol olvasva -- es a teszt hamisan ATENGEDTE volna azt, akinek explicit ki kellett volna
// zarva lennie. `in` az egyetlen operator, ami megkulonbozteti "a kulcs JELEN VAN, de ures"-t
// "a kulcs HIANYZIK"-tol -- csak az utobbi esetben olvasunk a fajlbol. ***
function detectCallerId(env = process.env, cwd = process.cwd()) {
  const allowlist = 'EURO_COEDIT_AGENTS' in env ? env.EURO_COEDIT_AGENTS : (bridgeEnvKulcs('EURO_COEDIT_AGENTS') || '')
  const allowed = String(allowlist)
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean)

  const explicit = String(env.EURO_AGENT_ID || '').trim()
  const fromCwd = (cwd.match(/\/agents\/([^/]+)(?:\/|$)/) || [])[1] || ''
  const jelolt = explicit || fromCwd
  if (!jelolt) return { id: '', ok: false, indok: 'a hivo nem azonosithato: sem EURO_AGENT_ID, sem /agents/<nev> a munkakonyvtarban' }
  if (!allowed.includes(jelolt)) {
    return {
      id: '',
      ok: false,
      // A korabbi szoveg ("a co-editing iras per-agens engedelyezett") tobbet allitott, mint
      // amit ez a kapu kikenyszerit -- MAGAT a
      // co-editing irast barmely sajat NC-fiokkal rendelkezo agens elerhetne bongeszovel,
      // enelkul a kod-ut nelkul is. Amit ez a kapu TENYLEG ved: hogy senki ne irjon MAS agens
      // neveben, a bridge altal feloldott flotta-identitassal/hitelesitovel. A helyes ut annak,
      // aki nincs a listan, KIMONDVA: sajat NC-fiokkal ez a megtagadas nem all utjaba.
      indok: `a hivo (${jelolt}) nincs az EURO_COEDIT_AGENTS allowlisten -- EZT A KOD-UTAT tagadjuk meg (a bridge altal feloldott flotta-identitassal/hitelesitovel valo iras), hogy senki ne irhasson MAS agens neveben; sajat NC-fiokkal, sajat munkameneteben a szerkeszto ettol fuggetlenul elerheto`,
    }
  }
  // A forras megnevezese nem diszites: egy kesobbi auditban ez mondja meg, MIBOL kovetkezett
  // az identitas, es hogy nem egy allitott nevet fogadtunk el.
  return { id: jelolt, ok: true, forras: explicit ? 'EURO_AGENT_ID (allowlistelve)' : 'munkakonyvtar (allowlistelve)' }
}

// --- A CELPELDANY ----------------------------------------------------------------------------
// *** A `NEXTCLOUD_URL` ITT SZANDEKOSAN NEM HASZNALHATO. *** A bridge kozos .env-jeben az az ELES
// peldanyra mutat (pl.: `https://cloud.example.com`), es ezt a tobbi ut ORÖKLI. Egy
// dokumentum-IRO eszkoznel az oroklodo alapertelmezes azt jelentene, hogy aki elfelejt celt adni,
// az az ELES Nextcloudba ir -- es ez a tevedes nem hibauzenetkent jelentkezne, hanem egy idegen
// fajl megvaltozasakent. Ezert SAJAT kulcsot kerunk, es a hianya MEGTAGADAS, nem visszaeses.
const CEL_KULCS = 'EURO_COEDIT_NC_URL'

// --- A HITELESITO ---------------------------------------------------------------------------
// KET FORRAS, EBBEN A SORRENDBEN -- es a valasz MEGNEVEZI, melyikbol jott:
//   1. a bridge kozos .env-je:  <AGENS>_NEXTCLOUD_USER + <AGENS>_NEXTCLOUD_APP_PASSWORD
//      (ma 9 agensnek all -- ezert erintetlen ag: aki eddig atment, ezutan is ugyanott megy at)
//   2. a VAULT:  `nc-<agens>-pass`, az agens SAJAT dashboard-tokenjevel olvasva
//
// *** MIERT A VAULT, ES MIERT NEM A .env-BE MASOLAS (a dontes indoka, 2026-08-15): *** a hianyzo
// agenst be lehetne irni a kozos .env-be is -- de az egy UJABB hitelesito-masolatot tenne abba az
// EGY fajlba, amit a flotta barmelyik MCP-gyereke elolvas. A vault-ut nem hoz letre uj masolatot,
// es a `nc-<agens>-pass` konvencio MAR ALL (megmerve: 5 agensnek van ilyen titka).
// *** AZ ISMERT AR, KIMONDVA: *** a flotta EGY unix-felhasznalokent fut, tehat egy agens MCP-gyereke
// fizikailag a tobbiek tokenjehez is hozzafer. Ezt egyik ut sem szunteti meg; a rekeszeles helye a
// kulon kartya (E2), nem ez a fajl.
// EGY kulcs erteke: ha a HIVO ATADTA `env`-ben (akar explicit URES stringgel is), az szamit --
// MASKULONBEN a bridge SAJAT .env-jebol, a folyamat env-jenek mutacioja NELKUL. Ugyanaz a mintat
// koveti, mint `detectCallerId` az EURO_COEDIT_AGENTS-nel (117. sor): az `in` operator
// kulonbozteti meg "a kulcs JELEN VAN, de ures"-t "a kulcs HIANYZIK"-tol -- enelkul egy teszt
// altal szandekosan URES-re allitott kulcs csendben visszaesne a fajlra (lasd test_coedit_identity
// [4b], ami pontosan ezt vizsgalja EURO_COEDIT_NC_URL-re).
// A KOZVETLEN `env[kulcs]` olvasas csak akkor
// latott erteket, ha egy KORABBI hivas (euro-mcp.cjs `loadBridgeEnv()`-et hivo 2 helye) mar
// mutalta a process.env-et -- a tobbi (9) csupasz hivasi hely hidegen megtagadott, holott a .env
// tartalmazta a kulcsot. Ugyanaz a defektus-generator alak, amit a 86-95. sor mar leirt
// `detectCallerId`-re: a helyesseget NEM egyenkent a hivasi helyeknek kell emlekezniuk, hanem
// SZERKEZETILEG kizarni a sorrendi fuggoseget.
function envVagyFajl(env, kulcs, dir = __dirname) {
  return kulcs in env ? env[kulcs] : (bridgeEnvKulcs(kulcs, dir) || '')
}

async function credentialsFor(agentId, env = process.env, opts = {}) {
  const elotag = agentId.toUpperCase()
  const userKey = `${elotag}_NEXTCLOUD_USER`
  const passKey = `${elotag}_NEXTCLOUD_APP_PASSWORD`
  const url = String(envVagyFajl(env, CEL_KULCS)).replace(/\/$/, '')
  const user = String(envVagyFajl(env, userKey)).trim() || agentId
  const envPass = String(envVagyFajl(env, passKey))

  if (!url) {
    // A HIANYZO KULCS NEVET megnevezzuk, az ERTEKET soha. Egy "nem sikerult" uzenet, ami nem mondja
    // meg MI hianyzik, ugyanannyit er, mint a csendes null, amit ez a kartya epp megszuntetett.
    return { ok: false, indok: `hianyzo konfiguracio: ${CEL_KULCS} (a NEXTCLOUD_URL-t ez az ut SZANDEKOSAN nem orokli: az az eles peldany)` }
  }
  if (envPass) return { ok: true, url, user, pass: envPass, forras: `bridge .env (${passKey})` }

  // Only one credential source in this project: the bridge's own .env. (An earlier internal
  // deployment had a second, vault-service-backed fallback here; that service is not part of
  // this repository, so a public build that tried to fall back to it would fail for every user
  // who does not happen to run that exact service. See README/CONTRIBUTING.)
  return { ok: false, indok: `hianyzo konfiguracio: ${passKey} nincs a bridge .env-jeben` }
}

// --- A PLAYWRIGHT FELOLDASA ------------------------------------------------------------------
// Ez a projekt NEM listazza a playwright-ot sajat fuggosegkent (megmerve: a package.json ket
// csomagot listaz) -- LUSTAN toltjuk be, csak amikor ezt az eszkozt tenylegesen hivjak, kulonben
// minden inditas fizetne egy bongeszo-konyvtar betolteseert, hogy aztan sose hasznalja.
// *Eloszor a NORMALIS feloldas fut (ha a hivo kornyezet mar telepitette playwright-ot); ha nincs,
// az `EURO_PLAYWRIGHT_PATH` env-valtozo egy MASIK, kulon telepitett peldanyra mutathat -- nincs
// beegetett, host-specifikus alapertelmezes. Ha egyik sem megy, ez NEM-MERT es kimondjuk -- nem
// osszeomlas.*
function loadPlaywright() {
  try {
    return { ok: true, pw: require('playwright'), honnan: 'normal feloldas' }
  } catch { /* nincs telepitve -- probaljuk a kulon megadott utat, ha van */ }
  const tartalek = process.env.EURO_PLAYWRIGHT_PATH
  if (!tartalek) {
    return { ok: false, indok: 'a playwright nincs telepitve, es az EURO_PLAYWRIGHT_PATH env-valtozo sincs beallitva' }
  }
  try {
    return { ok: true, pw: require(path.resolve(tartalek)), honnan: `tartalek ut (${tartalek})` }
  } catch (err) {
    return { ok: false, indok: `a playwright nem toltheto be (sem normal feloldassal, sem ${tartalek}): ${err.message}` }
  }
}

// --- A MAGONKENTI API-FELDERITES ------------------------------------------------------------
// Ez a fuggveny a BONGESZOBEN fut, ezert onalloan kell allnia (nincs kulso hivatkozasa).
// Visszaadja, HOL talalta az objektumot -- a hivo ezt naplozza, mert egy kesobbi verzio
// athelyezheti, es akkor a hiba a NEVROL fog szolni, nem a kepessegrol.
const BONGESZO_FELDERITES = `(() => {
  const jeloltek = [
    ['window.editor', window.editor],
    ['window.Asc.editor', (window.Asc || {}).editor],
  ];
  for (const [ut, obj] of jeloltek) {
    if (obj && typeof obj.pluginMethod_PasteHtml === 'function') {
      return { ut, vanOlvaso: typeof obj.pluginMethod_GetFileHTML === 'function' };
    }
  }
  return { ut: null, vanOlvaso: false };
})()`

// Ugyanaz a felderito minta, mint a PasteHtml-e, de a `callCommand`-ot keresi -- az a
// PROTOTIPUSON ul (`Object.keys()`-szel szurve nem latszik), a
// `typeof obj.callCommand === 'function'` viszont a prototipusra is lat, ugyanugy, mint a mar
// meglevo `pluginMethod_PasteHtml` ellenorzes.
const BONGESZO_FELDERITES_CALLCOMMAND = `(() => {
  const jeloltek = [
    ['window.editor', window.editor],
    ['window.Asc.editor', (window.Asc || {}).editor],
  ];
  for (const [ut, obj] of jeloltek) {
    if (obj && typeof obj.callCommand === 'function') return { ut };
  }
  return { ut: null };
})()`

// (getSelection/replaceSelection): a `callMethod` felderitese -- ugyanaz a mintazat, mint a
// fenti kettoe, csak a `GetSelectedText`/`PasteText` MASIK aljzatot hasznal (megmerve: ez az ut
// mindharom magban -- docx/xlsx/pptx -- mukodik, a `callCommand`-tol elteroen, ami cellak/
// dia-tartalom eseten nem mindig ugyanugy viselkedik).
const BONGESZO_FELDERITES_CALLMETHOD = `(() => {
  const jeloltek = [
    ['window.editor', window.editor],
    ['window.Asc.editor', (window.Asc || {}).editor],
  ];
  for (const [ut, obj] of jeloltek) {
    if (obj && typeof obj.callMethod === 'function') return { ut };
  }
  return { ut: null };
})()`

// A fileid -> DAV-relativ ut feloldasa a hivo SAJAT WebDAV-gyokerebol, PROPFIND Depth:infinity
// (nextcloud-talk-file-exchange skill fallback-mintaja, csak a /Talk mappa helyett a teljes
// gyoker). NEM a bongeszo-sessiont hasznalja: sima HTTP Basic Authcal, mert ez a mentett CSOMAG
// bizonyitasahoz kell, nem a szerkeszto UI-hoz -- ket kulon dolog, ket kulon ut.
async function findPathByFileId({ url, user, pass, fileId, timeoutMs = 20000, fetchImpl = fetch }) {
  const auth = Buffer.from(`${user}:${pass}`, 'utf8').toString('base64')
  const body = '<?xml version="1.0"?><d:propfind xmlns:d="DAV:" xmlns:oc="http://owncloud.org/ns"><d:prop><oc:fileid/></d:prop></d:propfind>'
  const controller = new AbortController()
  const timer = setTimeout(() => controller.abort(), timeoutMs)
  let res
  try {
    res = await fetchImpl(`${url}/remote.php/dav/files/${encodeURIComponent(user)}/`, {
      method: 'PROPFIND',
      headers: { Authorization: `Basic ${auth}`, Depth: 'infinity', 'Content-Type': 'application/xml' },
      body,
      signal: controller.signal,
    })
  } catch (err) {
    return { ok: false, indok: `PROPFIND nem valaszolt: ${err.message}` }
  } finally {
    clearTimeout(timer)
  }
  if (!res.ok) return { ok: false, indok: `PROPFIND HTTP ${res.status}` }
  const xml = await res.text()
  // Egy <d:response> blokk tartalmazza a href-et ES a fileid-et is -- a ket regexet a
  // BLOKKON belul parositjuk, nem a teljes szovegen vegigfutva (kulonben az N-edik href a
  // (N+1)-edik fileid-del parosodna, ha barmelyik hianyzik egy blokkbol).
  const blokkok = xml.split('<d:response>').slice(1)
  for (const blokk of blokkok) {
    const hrefM = /<d:href>([^<]+)<\/d:href>/.exec(blokk)
    const idM = /<oc:fileid>([^<]+)<\/oc:fileid>/.exec(blokk)
    if (hrefM && idM && idM[1] === String(fileId)) {
      return { ok: true, href: decodeURIComponent(hrefM[1]) }
    }
  }
  return { ok: false, indok: `a fileid ${fileId} nem talalhato a PROPFIND valaszban (${blokkok.length} bejegyzes atnezve)` }
}

// Egy zip-resz kiolvasasa, VAGY null ha nincs (nem hiba -- pl. egy komment nelkuli dokumentumnak
// nincs word/comments.xml resze SEM, ez a MEGLEVO all). Csak
// OPCIONALIS reszekre hasznald; egy KOTELEZO resz hianya (a fo dokumentum-XML) a hivo oldalan
// 'csomag-olvashatatlan'-kent kezelendo, nem null-kent.
function unzipEntryOrNull(tmpPath, entryName) {
  try {
    return execFileSync('unzip', ['-p', tmpPath, entryName], { encoding: 'utf8' })
  } catch {
    return null
  }
}

// word/document.xml + word/comments.xml -> {bekezdesek, kommentek}. A docx ag,
// kulon fuggvenybe emelve, hogy a pptx-ag (lasd lent) NE ugyanabba a szal-agba nojon -- ket
// teljesen mas OOXML-sema, kozos kod itt csak a hivo elrejtett komplexitasat noveIne.
function readDocxContent(tmpPath) {
  const lib = require('./lib.cjs')
  let documentXml
  try {
    documentXml = execFileSync('unzip', ['-p', tmpPath, 'word/document.xml'], { encoding: 'utf8' })
  } catch (err) {
    return { ok: false, outcome: 'csomag-olvashatatlan', indok: `word/document.xml nem olvashato ki: ${err.message}` }
  }
  const commentsXml = unzipEntryOrNull(tmpPath, 'word/comments.xml')
  // footnotes.xml/endnotes.xml are OPTIONAL parts (a document with
  // no footnote/endnote at all has neither), same reasoning as commentsXml above -- unzipEntryOrNull
  // returns null, and the lib.cjs parsers already treat a null/missing input as "zero items" (they
  // String(...ml ?? '') internally), so no extra null-check is needed here.
  const footnotesXml = unzipEntryOrNull(tmpPath, 'word/footnotes.xml')
  const endnotesXml = unzipEntryOrNull(tmpPath, 'word/endnotes.xml')
  return {
    ok: true,
    outcome: 'olvasva',
    bekezdesek: lib.paragraphTextsFromDocumentXml(documentXml),
    kommentek: lib.commentsFromCommentsXml(commentsXml),
    labjegyzetek: lib.footnotesFromFootnotesXml(footnotesXml),
    vegjegyzetek: lib.endnotesFromEndnotesXml(endnotesXml),
    konyvjelzok: lib.bookmarksFromDocumentXml(documentXml),
    tablazatok: lib.tablesFromDocumentXml(documentXml),
  }
}

// ppt/presentation.xml (+ rels) -> a dia-sorrend -> minden dia sajat slideN.xml-je (szoveg) + a
// hozza tartozo commentM.xml (ha van, a slide-rels "…/relationships/comments" kapcsolatan at, ld.
// lib.cjs pptxCommentsPartFromSlideRelsXml fejlec-kommentje) -- egyenkent, MERT alak (korabbi
// lelet: a resz-nev NEM a megjelenites sorrendje, ld. a MoveTo-muvelet korul). commentAuthors.xml
// egyszer olvasva, minden diahoz megosztva -- a szerzo-nevek dokumentum-szintuek, nem
// dia-szintuek.
//
// A KIMENET KET ALAKBAN adja a szoveget: `bekezdesek` LAPOSAN, minden dian at, docx-szal
// KONZISZTENS alakban (ugyanaz a mezonev, ugyanaz a tipus -- egy meglevo office_find hivo
// valtoztatas nelkul mukodik pptx-en is), ES `diak` per-dia bontasban -- mert egy pptx-hivonak
// szinte mindig szamit, MELYIK dian van a szoveg, es a lapos alak ezt eldobna. Ugyanez a ketto-
// alak all a "dia-tartalom visszaolvasasara" is (a hatokor kesobb bovult): `diak[].tartalom`
// per-dia (slideContentSummaryFromSlideXml), `tartalomOsszesen` LAPOSAN minden dian at, mindegyik
// elem sajat `slideIndex`-szel megjelolve -- ugyanaz a minta, mint a kommentek listajae, nem egy
// harmadik alak.
// "Elrendezes/tema lekerdezese" -- the slide -> layout -> master ->
// theme rels chain (lib.cjs pptxLayoutAndThemeFromSlideLayoutXml's own header comment explains
// WHY the master hop cannot be skipped), pulled out of readPptxContent's own per-slide loop so
// that function's complexity does not grow with a four-file resolution chain inline. Any missing
// hop (a rels entry not found, a part not readable) stops the chain and returns null fields --
// NOT a thrown error, since a slide missing its own layout/theme link is a real, if unusual,
// package state, not this reader's failure.
function resolvePptxLayoutAndTheme(tmpPath, slidePart, slideRelsXml) {
  const lib = require('./lib.cjs')
  const slideDir = slidePart.replace(/\/[^/]+$/, '')
  const layoutTargetRaw = lib.ooxmlRelationshipTargetByType(slideRelsXml, '/slideLayout')
  if (!layoutTargetRaw) return { layoutName: null, layoutType: null, themeName: null }
  const layoutPath = lib.resolveOoxmlRelativeTarget(slideDir, layoutTargetRaw)
  const layoutXml = unzipEntryOrNull(tmpPath, layoutPath)
  if (!layoutXml) return { layoutName: null, layoutType: null, themeName: null }

  const layoutDir = layoutPath.replace(/\/[^/]+$/, '')
  const layoutFile = layoutPath.split('/').pop()
  const layoutRelsXml = unzipEntryOrNull(tmpPath, `${layoutDir}/_rels/${layoutFile}.rels`)
  const masterTargetRaw = lib.ooxmlRelationshipTargetByType(layoutRelsXml, '/slideMaster')
  if (!masterTargetRaw) return lib.pptxLayoutAndThemeFromSlideLayoutXml(layoutXml, null)
  const masterPath = lib.resolveOoxmlRelativeTarget(layoutDir, masterTargetRaw)

  const masterDir = masterPath.replace(/\/[^/]+$/, '')
  const masterFile = masterPath.split('/').pop()
  const masterRelsXml = unzipEntryOrNull(tmpPath, `${masterDir}/_rels/${masterFile}.rels`)
  const themeTargetRaw = lib.ooxmlRelationshipTargetByType(masterRelsXml, '/theme')
  if (!themeTargetRaw) return lib.pptxLayoutAndThemeFromSlideLayoutXml(layoutXml, null)
  const themePath = lib.resolveOoxmlRelativeTarget(masterDir, themeTargetRaw)
  const themeXml = unzipEntryOrNull(tmpPath, themePath)

  return lib.pptxLayoutAndThemeFromSlideLayoutXml(layoutXml, themeXml)
}

function readPptxContent(tmpPath) {
  const lib = require('./lib.cjs')
  let presentationXml
  let relsXml
  try {
    presentationXml = execFileSync('unzip', ['-p', tmpPath, 'ppt/presentation.xml'], { encoding: 'utf8' })
    relsXml = execFileSync('unzip', ['-p', tmpPath, 'ppt/_rels/presentation.xml.rels'], { encoding: 'utf8' })
  } catch (err) {
    return { ok: false, outcome: 'csomag-olvashatatlan', indok: `ppt/presentation.xml (+ rels) nem olvashato ki: ${err.message}` }
  }
  const slideParts = lib.pptxSlideOrderFromPresentationXml(presentationXml, relsXml)
  if (!slideParts.length) {
    return { ok: false, outcome: 'csomag-olvashatatlan', indok: 'ppt/presentation.xml sldIdLst-je 0 diat sorolt fel -- vagy ures a bemutato, vagy a resz nem parszolhato a mert alakkal' }
  }
  const commentAuthorsXml = unzipEntryOrNull(tmpPath, 'ppt/commentAuthors.xml')
  const diak = []
  const bekezdesek = []
  const kommentek = []
  const tartalomOsszesen = { shapes: [], images: [], tables: [], charts: [], groups: [] }
  slideParts.forEach((slidePart, index) => {
    const slideXml = unzipEntryOrNull(tmpPath, slidePart)
    const sajatBekezdesek = lib.paragraphTextsFromSlideXml(slideXml)
    const sajatTartalom = lib.slideContentSummaryFromSlideXml(slideXml)
    bekezdesek.push(...sajatBekezdesek)
    Object.keys(tartalomOsszesen).forEach((kulcs) => {
      sajatTartalom[kulcs].forEach((elem) => tartalomOsszesen[kulcs].push({ ...elem, slideIndex: index }))
    })

    // slidePart pl. "ppt/slides/slide3.xml" -> a sajat rels-je "ppt/slides/_rels/slide3.xml.rels"
    const slideFile = slidePart.split('/').pop()
    const relsPart = slidePart.replace(/\/[^/]+$/, `/_rels/${slideFile}.rels`)
    const slideRelsXml = unzipEntryOrNull(tmpPath, relsPart)
    const elrendezesEsTema = resolvePptxLayoutAndTheme(tmpPath, slidePart, slideRelsXml)
    diak.push({ index, bekezdesek: sajatBekezdesek, tartalom: sajatTartalom, elrendezesEsTema })

    const commentsPart = lib.pptxCommentsPartFromSlideRelsXml(slideRelsXml)
    if (!commentsPart) return
    const commentXml = unzipEntryOrNull(tmpPath, commentsPart)
    const sajatKommentek = lib.commentsFromPptxCommentXml(commentXml, commentAuthorsXml)
    sajatKommentek.forEach((k) => kommentek.push({ ...k, slideIndex: index }))
  })
  // "Dokumentum-metaadat" -- docProps/ is a package-level path (not under ppt/), shared unchanged
  // by docx/xlsx/pptx alike, so documentMetadataFromCoreAndAppXml itself is core-agnostic; only
  // wired here (pptx) for now, not because the parser could not also serve docx. Both parts are
  // OPTIONAL reads (unzipEntryOrNull)
  // -- a document missing docProps/app.xml still returns its core.xml fields, and vice versa.
  const metaadat = lib.documentMetadataFromCoreAndAppXml(
    unzipEntryOrNull(tmpPath, 'docProps/core.xml'),
    unzipEntryOrNull(tmpPath, 'docProps/app.xml'),
  )
  return { ok: true, outcome: 'olvasva', bekezdesek, diak, kommentek, tartalomOsszesen, metaadat }
}

// (office_get_text / office_get_comments): OLVASAS a MENTETT csomagbol, NEM a
// callCommand visszateresebol -- lasd writeOperationsToDocument fejlec-kommentjet arrol, hogy
// EZ a visszateres semmit nem bizonyit. A megoldas ITT egyszerubb, mint az irasnal: nincs
// szukseg bongeszore/Playwrightra/elo sessionre EGYALTALAN -- a WEBDAV-on letoltott fajl a
// mentett allapotot MAR hordozza (ugyanaz a letoltes, amit writeOperationsToDocument a sajat
// irasi bizonyitasara hasznal), es azt egy `unzip -p`-vel kiolvasott XML-en at, KIZAROLAG
// PARANCS-VEGREHAJTAS NELKULI, tiszta fuggvenyekkel (lib.cjs-beli parszolokkal) dolgozzuk fel.
// Ezert ez a fuggveny NEM igenyel playwright-ot, NEM igenyel EURO_COEDIT_AGENTS-tagsagot -- ez a
// ket muvelet EZERT nem szamit a kijeloles-fuggo, elo-proba-var halmazba.
//
// *** HATAR, MERT CSAK ANNYIT ALLITUNK, AMENNYIT MERTUNK: a csomag-alapu olvasas docx-re ES
// pptx-re bizonyitott (docx es pptx eseten is, elo
// Document Server ellen mert sema -- ld. lib.cjs pptxSlideOrderFromPresentationXml/
// commentsFromPptxCommentXml fejlec-kommentjei). xlsx-en a szoveg/komment MEG MINDIG mas
// utvonalon/semaban el (xl/worksheets/sheetN.xml, sejt-szintu, nem bekezdes-szintu) -- arra EZ a
// fuggveny tovabbra is outcome: 'core-nem-tamogatott'-tal ter vissza, MIELOTT barmit probalna. ***
async function readDocumentContent({ url, user, pass, fileId, core, timeoutMs = 20000, fetchImpl = fetch }) {
  if (core !== 'docx' && core !== 'pptx') {
    return {
      ok: false,
      outcome: 'core-nem-tamogatott',
      indok: `document.xml/comments.xml-alapu olvasas ma docx-re es pptx-re bizonyitott, xlsx-re nem (kaptam: ${core}) -- xlsx sejt-szinten, mas utvonalon tartja a szoveget, azt ez a fuggveny meg nem ismeri`,
    }
  }
  const utvonal = await findPathByFileId({ url, user, pass, fileId, timeoutMs, fetchImpl })
  if (!utvonal.ok) return { ok: false, outcome: 'fileid-nem-oldhato-fel', indok: utvonal.indok }

  const auth = Buffer.from(`${user}:${pass}`, 'utf8').toString('base64')
  const controller = new AbortController()
  const timer = setTimeout(() => controller.abort(), timeoutMs)
  let res
  try {
    res = await fetchImpl(`${url}${utvonal.href}`, { headers: { Authorization: `Basic ${auth}` }, signal: controller.signal })
  } catch (err) {
    return { ok: false, outcome: 'letoltes-sikertelen', indok: `letoltes nem valaszolt: ${err.message}` }
  } finally {
    clearTimeout(timer)
  }
  if (!res.ok) return { ok: false, outcome: 'letoltes-sikertelen', indok: `letoltes HTTP ${res.status}` }
  const buf = Buffer.from(await res.arrayBuffer())

  // A TRANZIT-fajl a `Tranzit-fájl higiénia` szabaly szerint: csak addig el, amig a zip-bejegyzeseket
  // ki nem olvassuk belole, utana AZONNAL torlodik -- nem tart meg tartos masolatot.
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'euro-mcp-read-'))
  const tmpPath = path.join(tmpDir, `dokumentum-${fileId}.${core}`)
  try {
    fs.writeFileSync(tmpPath, buf)
    const eredmeny = core === 'docx' ? readDocxContent(tmpPath) : readPptxContent(tmpPath)
    if (!eredmeny.ok) return eredmeny
    return { ...eredmeny, bytesOlvasva: buf.length }
  } finally {
    fs.rmSync(tmpDir, { recursive: true, force: true })
  }
}

module.exports = {
  checkFreshness,
  detectCallerId, credentialsFor, loadPlaywright, loadBridgeEnv,
  bridgeEnvKulcs, coeditUtAllapota, CEL_KULCS, BONGESZO_FELDERITES,
  BONGESZO_FELDERITES_CALLCOMMAND, BONGESZO_FELDERITES_CALLMETHOD, findPathByFileId,
  readDocumentContent,
}

// --- AZ IRAS MAGA ---------------------------------------------------------------------------
// Egy MEGNYITOTT dokumentumba ir a hivo agens SAJAT NC-azonossagaval.
//
// *** A VISSZAOLVASAS MAGONKENT ELTER, ES EZT A VALASZ KIMONDJA (a kartya 3. kikotese): ***
// a `GetFileHTML` CSAK a szoveges magban letezik. Ahol nincs, ott NEM allitunk tartalmi
// igazolast -- a valasz `tartalomIgazolva: null` es egy megnevezett indok. *Egy "nem tudom
// visszaolvasni" elfogadhato valasz; a csendes null nem az -- pont az szulte ezt a kartyat.*
async function writeToDocument({ url, user, pass, fileId, html, timeoutMs = 60000 }) {
  const betoltes = loadPlaywright()
  if (!betoltes.ok) return { ok: false, outcome: 'nem-mert', indok: betoltes.indok }
  const { chromium } = betoltes.pw

  const bongeszo = await chromium.launch()
  try {
    const ctx = await bongeszo.newContext({ viewport: { width: 1400, height: 900 } })
    const lap = await ctx.newPage()

    await lap.goto(`${url}/login`, { waitUntil: 'domcontentloaded', timeout: timeoutMs })
    await lap.fill('#user', user)
    await lap.fill('#password', pass)
    await Promise.all([
      lap.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: timeoutMs }).catch(() => null),
      lap.click('button[type=submit], input[type=submit]'),
    ])
    await lap.waitForTimeout(2500)
    // A bejelentkezes SIKERE nem a HTTP-kod: egy rossz jelszo ugyanugy 200-at ad, csak a
    // login-lapot adja vissza. Az URL az, ami elvalasztja a kettot.
    if (/\/login/.test(lap.url())) {
      return { ok: false, outcome: 'auth', indok: 'a bejelentkezes nem ment at (a login-lapon maradtunk)' }
    }

    await lap.goto(`${url}/index.php/apps/eurooffice/${fileId}`, { waitUntil: 'domcontentloaded', timeout: timeoutMs })
    await lap.waitForTimeout(22000)
    const frame = lap.frames().find((f) => /documenteditor|spreadsheeteditor|presentationeditor/.test(f.url()))
    if (!frame) return { ok: false, outcome: 'nem-nyilt-meg', indok: 'a szerkeszto-frame nem jott fel' }
    const mag = (frame.url().match(/(documenteditor|spreadsheeteditor|presentationeditor)/) || [])[1] || 'ismeretlen'

    const felderites = await frame.evaluate(BONGESZO_FELDERITES)
    if (!felderites.ut) {
      return { ok: false, outcome: 'nincs-api', mag, indok: 'a PasteHtml egyik ismert helyen sem talalhato (window.editor / window.Asc.editor)' }
    }

    const elotte = felderites.vanOlvaso
      ? await frame.evaluate((u) => {
        const e = u === 'window.editor' ? window.editor : window.Asc.editor
        return String(e.pluginMethod_GetFileHTML() || '').length
      }, felderites.ut).catch(() => null)
      : null

    const hivas = await frame.evaluate(({ u, h }) => {
      try {
        const e = u === 'window.editor' ? window.editor : window.Asc.editor
        e.pluginMethod_PasteHtml(h)
        return { hivva: true }
      } catch (err) { return { hivva: false, hiba: String(err && err.message).slice(0, 160) } }
    }, { u: felderites.ut, h: html })
    if (!hivas.hivva) return { ok: false, outcome: 'a-hivas-dobott', mag, apiHely: felderites.ut, indok: hivas.hiba }

    await lap.waitForTimeout(6000)

    let tartalomIgazolva = null
    let igazolasIndok = null
    if (felderites.vanOlvaso) {
      const utana = await frame.evaluate((u) => {
        const e = u === 'window.editor' ? window.editor : window.Asc.editor
        return String(e.pluginMethod_GetFileHTML() || '')
      }, felderites.ut).catch(() => null)
      if (utana === null) { igazolasIndok = 'a visszaolvasas dobott' } else {
        // A HOSSZ-NOVEKEDES onmagaban gyenge; a KERESETT SZOVEG jelenlete a bizonyitek.
        // A hivo altal adott HTML-bol a nyers szoveget vesszuk, mert a szerkeszto sajat
        // jelolest general -- a betuk maradnak, a tagek nem.
        const nyers = String(html).replace(/<[^>]*>/g, ' ').split(/\s+/).filter((w) => w.length > 6)
        tartalomIgazolva = nyers.length ? nyers.some((w) => utana.includes(w)) : null
        if (tartalomIgazolva === null) igazolasIndok = 'a beirt HTML-bol nem volt eleg hosszu szo a kereseshez'
      }
    } else {
      igazolasIndok = `ebben a magban (${mag}) NINCS pluginMethod_GetFileHTML, ezert a szerkesztobol nem olvashato vissza`
    }

    return {
      ok: true,
      outcome: tartalomIgazolva === true ? 'ok' : 'beirva-igazolas-nelkul',
      mag,
      apiHely: felderites.ut,
      tartalomIgazolva,
      igazolasIndok,
      hosszElotte: elotte,
    }
  } finally {
    await bongeszo.close().catch(() => {})
  }
}

// Az `operations` lista beirasa egy MAR NYITOTT dokumentumba, `editor.callCommand(...)`-on at --
// lib.buildCoeditScript() forditja a torzset, ez a fuggveny futtatja le a bongeszoben, es
// EGYEDUL a MENTETT CSOMAGOT fogadja el bizonyitéknak (megmerve: "a callCommand VISSZATÉRÉSE
// SEMMIT NEM BIZONYÍT" -- szándékosan hibás törzs is "sikerült" választ ad).
//
// A verifikacio EZERT NEM a szerkesztobol olvas vissza (mint writeToDocument GetFileHTML-je),
// hanem a fajlt magat tolti le WebDAV-on, KETSZER: a callCommand ELOTT es UTAN. A MENTES A
// SESSION ZARASA UTAN KESIK kb. egy percet (megmerve) -- ezert a bongeszo bezarasa es a
// masodik letoltes koze egy nevesitett varakozas kerul, nem azonnal olvasunk.
// Split out of writeOperationsToDocument purely to keep ITS complexity down (qlty smells) --
// the workbook/rels/sheet-part I/O is a single-purpose
// step (resolve each xlsx op's OWN target sheet, then read each needed part exactly once, cached
// by name since several ops can share a sheet). Returns an array aligned to `operations` (same
// index), each entry the resolved sheet's XML text or null (unresolvable/unreadable -- the pure
// `lib.xlsxCellVerificationReport` reports that as NEM-MERT, not as a failure).
function buildXlsxPerOpSheetXml(lib, tmpFile, operations) {
  let workbookXml = null
  let relsXml = null
  try { workbookXml = execFileSync('unzip', ['-p', tmpFile, 'xl/workbook.xml'], { encoding: 'utf8' }) } catch { workbookXml = null }
  try { relsXml = execFileSync('unzip', ['-p', tmpFile, 'xl/_rels/workbook.xml.rels'], { encoding: 'utf8' }) } catch { relsXml = null }

  const sheetXmlCache = new Map() // part-name -> xml string | null (dedupe: several ops can target the same sheet)
  const readSheetPart = (partName) => {
    if (sheetXmlCache.has(partName)) return sheetXmlCache.get(partName)
    let xml = null
    try { xml = execFileSync('unzip', ['-p', tmpFile, partName], { encoding: 'utf8' }) } catch { xml = null }
    sheetXmlCache.set(partName, xml)
    return xml
  }

  return (operations || []).map((op) => {
    const partName = lib.resolveXlsxSheetFile(workbookXml, relsXml, op && op.sheet)
    // "5. pont": a kicsomagolas/feloldas hibaja NEM az iras hibaja -- a pure fuggveny ezt (null)
    // NEM-MERT-kent jelenti, nem "sikertelen"-kent.
    return partName === null ? null : readSheetPart(partName)
  })
}

async function writeOperationsToDocument({ url, user, pass, fileId, core, operations, timeoutMs = 60000, postSaveWaitMs = 70000 }) {
  const betoltes = loadPlaywright()
  if (!betoltes.ok) return { ok: false, outcome: 'nem-mert', indok: betoltes.indok }
  const { chromium } = betoltes.pw

  let built
  try {
    built = require('./lib.cjs').buildCoeditScript({ core, operations })
  } catch (err) {
    // This route is all-or-nothing -- buildCoeditScript only throws once the translator ran and named EVERY
    // operation, and on ANY failure here NONE of them were applied (not just the fully-degenerate
    // all-refused case create_document also has). err.report travels with the failure regardless,
    // so a caller can still see WHICH operations were asked for and WHY each was refused or
    // withheld, not just a single collapsed message.
    return { ok: false, outcome: 'fordito-megtagadta', indok: err.message, report: err.report ?? null }
  }

  const utvonal = await findPathByFileId({ url, user, pass, fileId, timeoutMs })
  if (!utvonal.ok) return { ok: false, outcome: 'fileid-nem-oldhato-fel', indok: utvonal.indok }

  const auth = Buffer.from(`${user}:${pass}`, 'utf8').toString('base64')
  async function letoltMeret() {
    const res = await fetch(`${url}${utvonal.href}`, { headers: { Authorization: `Basic ${auth}` } })
    if (!res.ok) return { ok: false, indok: `letoltes HTTP ${res.status}` }
    const buf = Buffer.from(await res.arrayBuffer())
    return { ok: true, bytes: buf.length, buf }
  }
  const elotte = await letoltMeret()
  if (!elotte.ok) return { ok: false, outcome: 'letoltes-elotte-sikertelen', indok: elotte.indok }

  const bongeszo = await chromium.launch()
  try {
    const ctx = await bongeszo.newContext({ viewport: { width: 1400, height: 900 } })
    const lap = await ctx.newPage()

    await lap.goto(`${url}/login`, { waitUntil: 'domcontentloaded', timeout: timeoutMs })
    await lap.fill('#user', user)
    await lap.fill('#password', pass)
    await Promise.all([
      lap.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: timeoutMs }).catch(() => null),
      lap.click('button[type=submit], input[type=submit]'),
    ])
    await lap.waitForTimeout(2500)
    if (/\/login/.test(lap.url())) {
      return { ok: false, outcome: 'auth', indok: 'a bejelentkezes nem ment at (a login-lapon maradtunk)' }
    }

    await lap.goto(`${url}/index.php/apps/eurooffice/${fileId}`, { waitUntil: 'domcontentloaded', timeout: timeoutMs })
    await lap.waitForTimeout(22000)
    const frame = lap.frames().find((f) => /documenteditor|spreadsheeteditor|presentationeditor/.test(f.url()))
    if (!frame) return { ok: false, outcome: 'nem-nyilt-meg', indok: 'a szerkeszto-frame nem jott fel' }
    const mag = (frame.url().match(/(documenteditor|spreadsheeteditor|presentationeditor)/) || [])[1] || 'ismeretlen'

    const felderites = await frame.evaluate(BONGESZO_FELDERITES_CALLCOMMAND)
    if (!felderites.ut) {
      return { ok: false, outcome: 'nincs-api', mag, indok: 'a callCommand egyik ismert helyen sem talalhato (window.editor / window.Asc.editor)' }
    }

    // A hivas SAJAT visszaterese NEM bizonyitek (lasd fenti fejlecmegjegyzes) -- csak azt
    // rogzitjuk, hogy dobott-e kivetelt, mert az legalabb a "meg sem probalta" esetet zarja ki.
    const hivas = await frame.evaluate(({ u, body }) => {
      try {
        const e = u === 'window.editor' ? window.editor : window.Asc.editor
        const fv = new Function(body)
        e.callCommand(fv)
        return { hivva: true }
      } catch (err) { return { hivva: false, hiba: String(err && err.message).slice(0, 200) } }
    }, { u: felderites.ut, body: built.script })
    if (!hivas.hivva) return { ok: false, outcome: 'a-hivas-dobott', mag, apiHely: felderites.ut, indok: hivas.hiba }

    await lap.waitForTimeout(6000)
    await bongeszo.close().catch(() => {})

    // A mentes kesik a session zarasa utan (megmerve) -- itt varunk, MIELOTT masodszor
    // letoltjuk a fajlt. Enelkul egy meg-nem-mentett allapotot merunk, es azt hinnenk, nem
    // tortent semmi.
    await new Promise((resolve) => setTimeout(resolve, postSaveWaitMs))

    const utana = await letoltMeret()
    if (!utana.ok) return { ok: false, outcome: 'letoltes-utana-sikertelen', mag, apiHely: felderites.ut, indok: utana.indok }

    // Tartalom-ellenorzes docx-en: a `text` muveletek sajat szovege kereshetó a MENTETT
    // csomagban (ugyanaz a minta, mint lib.cjs markerInDocumentXml()-je -- itt kulon
    // implementalva, mert ez a modul nem importalja lib.cjs-t a fenti buildCoeditScript-hivason
    // kivul, es a fajl a lemezen all, nem egy JS-string). `unzip -p` a rendszer eszkoze (mar
    // hasznalja package-consistency.cjs is), zip-fuggoseg nelkul.
    let tartalomIgazolva = null
    let igazolasIndok = 'ehhez a maghoz nincs szoveg-alapu tartalom-ellenorzes -- csak a byte-meret hasonlitva'
    const kertSzovegek = (Array.isArray(operations) ? operations : [])
      .filter((op) => op && op.type === 'text' && typeof op.text === 'string' && op.text.trim())
      .map((op) => op.text)
    if (core === 'docx' && kertSzovegek.length) {
      const tmpFile = path.join(os.tmpdir(), `coedit-verify-${process.pid}-${Date.now()}.docx`)
      fs.writeFileSync(tmpFile, utana.buf)
      try {
        const xml = execFileSync('unzip', ['-p', tmpFile, 'word/document.xml'], { encoding: 'utf8' })
        const szoveg = xml.replace(/<[^>]*>/g, '')
        tartalomIgazolva = kertSzovegek.every((t) => szoveg.includes(t))
        igazolasIndok = tartalomIgazolva
          ? 'mind a kert szoveg megtalalhato a mentett word/document.xml-ben'
          : `nem talalhato mind a kert szoveg: ${JSON.stringify(kertSzovegek.filter((t) => !szoveg.includes(t)))}`
      } catch (err) {
        igazolasIndok = `a letoltott csomag nem volt kicsomagolhato: ${err.message}`
      } finally {
        fs.rmSync(tmpFile, { force: true })
      }
    }

    // Az xlsx `table`/`formula` muveletek a MENTETT csomagban a SAJAT op.sheet celjuknak megfelelo
    // lapon vannak visszaigazolva, PER MUVELET (lib.cjs `xlsxCellVerificationReport` -- ez itt
    // csak az I/O: kicsomagolja a szukseges lapo(ka)t, a DONTES a pure fuggvenyben all, kulon
    // tesztelheto bongeszo nelkul). A gyoker ok (lib.cjs `formula` bejegyzese, komment ott):
    // `SetValue` egy szintaktikailag "="-jellel kezdodo, de szemantikailag ervenytelen kepletre
    // CSENDBEN nem hoz letre cellat -- se kivetel, se ertek, es a fenti `report` ezt
    // "alkalmazva"-nak jelenti, mert a KLIENS-oldali validalas (buildCoeditScript) csak a vezeto
    // "="-t nezi.
    //
    // *** MEASURED ROOT CAUSE (2026-08-17): ez a resz korabban MINDIG 'xl/worksheets/sheet1.xml'-t
    // olvasta, fuggetlenul attol, melyik lapra celzott egy adott muvelet -- azota, hogy `op.sheet`
    // valodi per-operacio lap-cimzest ad, ez routinszeru hamis `hianyzoCellak` jelzest termelt
    // egy MASODIK/HARMADIK lapra celzott, BIZONYITOTTAN helyesen landolt irasra (elo proba,
    // fileId 33316: sheet:1-es formula a mentett sheet2.xml-en landolt, de az ELOZO kod-alak
    // sheet1.xml-en kereste -> hamis hiany). ***
    // The SAME per-operation sheet-XML resolution the
    // cell check above needs, reused for the autoFilter "third case" (coedit-iras-elveszi-az-
    // autofiltert01's runtime guard throws on the Document Server, AFTER buildCoeditScript's own
    // client-side report already marked the operation 'alkalmazva' -- see that function's own
    // header comment in lib.cjs). Unzipped ONCE, shared by whichever check(s) this batch needs.
    let cellaEllenorzes = null
    const lib = require('./lib.cjs')
    const kellCellaEllenorzes = core === 'xlsx' && (operations || []).some((op) => lib.xlsxRequestedCellRefs(op).length)
    const kellAutoFilterEllenorzes = core === 'xlsx' && (operations || []).some((op) => lib.xlsxRequestsAutoFilterVerification(op))
    if (kellCellaEllenorzes || kellAutoFilterEllenorzes) {
      const tmpFile = path.join(os.tmpdir(), `coedit-verify-xlsx-${process.pid}-${Date.now()}.xlsx`)
      fs.writeFileSync(tmpFile, utana.buf)
      try {
        const perOpSheetXml = buildXlsxPerOpSheetXml(lib, tmpFile, operations)
        if (kellCellaEllenorzes) {
          cellaEllenorzes = lib.xlsxCellVerificationReport(operations, perOpSheetXml)
          // Promote the verdict into `report[i].outcome`
          // itself, not just the `cellaEllenorzes` side-channel -- see that function's own header
          // comment in lib.cjs for why a caller checking `outcome` alone should not be misled.
          lib.applyCellVerificationToReport(built.report, cellaEllenorzes)
        }
        if (kellAutoFilterEllenorzes) {
          // The BEFORE snapshot too (same per-op sheet
          // resolution, on the file downloaded before the write) -- this is what lets the
          // verification tell "the wanted range already existed before this call" (a genuine
          // idempotent no-op) apart from "this call created it", which an AFTER-only read cannot.
          const tmpFileElotte = path.join(os.tmpdir(), `coedit-verify-xlsx-elotte-${process.pid}-${Date.now()}.xlsx`)
          fs.writeFileSync(tmpFileElotte, elotte.buf)
          let perOpSheetXmlElotte
          try {
            perOpSheetXmlElotte = buildXlsxPerOpSheetXml(lib, tmpFileElotte, operations)
          } finally {
            fs.rmSync(tmpFileElotte, { force: true })
          }
          const autoFilterVerification = lib.xlsxAutoFilterVerificationReport(operations, perOpSheetXml, perOpSheetXmlElotte)
          lib.applyAutoFilterVerificationToReport(built.report, autoFilterVerification)
        }
      } finally {
        fs.rmSync(tmpFile, { force: true })
      }
    }

    return {
      ok: true,
      outcome: elotte.bytes !== utana.bytes ? 'meret-valtozott' : 'meret-valtozatlan',
      mag,
      apiHely: felderites.ut,
      // Threaded through from buildCoeditScript, unchanged --
      // this function does not recompute or trust anything beyond what the translator itself
      // reported.
      applied: built.applied,
      report: built.report,
      bytesElotte: elotte.bytes,
      bytesUtana: utana.bytes,
      savedBefore: elotte.buf.toString('base64'),
      savedAfter: utana.buf.toString('base64'),
      tartalomIgazolva,
      igazolasIndok,
      cellaEllenorzes,
    }
  } finally {
    await bongeszo.close().catch(() => {})
  }
}

module.exports.writeToDocument = writeToDocument
module.exports.writeOperationsToDocument = writeOperationsToDocument

// --- GETSELECTION / REPLACESELECTION (the owner's office.get_selection /
// office.replace_selection) ---------------------------------------------------------------------
//
// Shared login+navigate+frame-detection for the two functions below -- the SAME sequence
// writeToDocument/writeOperationsToDocument above run inline; split out here (not touching
// those two, out of scope for this card) because getSelection and replaceSelection would
// otherwise duplicate it a second and third time, and qlty smells already flags the two
// existing inline copies as high-complexity -- a third copy would only make that worse.
// Returns either an already-shaped error result (ok:false) or {browser, frame, mag, apiHely}.
async function nyitottCallMethodFrame({ url, user, pass, fileId, timeoutMs }) {
  const betoltes = loadPlaywright()
  if (!betoltes.ok) return { ok: false, outcome: 'nem-mert', indok: betoltes.indok }
  const { chromium } = betoltes.pw
  const bongeszo = await chromium.launch()

  const ctx = await bongeszo.newContext({ viewport: { width: 1400, height: 900 } })
  const lap = await ctx.newPage()
  await lap.goto(`${url}/login`, { waitUntil: 'domcontentloaded', timeout: timeoutMs })
  await lap.fill('#user', user)
  await lap.fill('#password', pass)
  await Promise.all([
    lap.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: timeoutMs }).catch(() => null),
    lap.click('button[type=submit], input[type=submit]'),
  ])
  await lap.waitForTimeout(2500)
  if (/\/login/.test(lap.url())) {
    await bongeszo.close().catch(() => {})
    return { ok: false, outcome: 'auth', indok: 'a bejelentkezes nem ment at (a login-lapon maradtunk)' }
  }

  await lap.goto(`${url}/index.php/apps/eurooffice/${fileId}`, { waitUntil: 'domcontentloaded', timeout: timeoutMs })
  await lap.waitForTimeout(22000)
  const frame = lap.frames().find((f) => /documenteditor|spreadsheeteditor|presentationeditor/.test(f.url()))
  if (!frame) {
    await bongeszo.close().catch(() => {})
    return { ok: false, outcome: 'nem-nyilt-meg', indok: 'a szerkeszto-frame nem jott fel' }
  }
  const mag = (frame.url().match(/(documenteditor|spreadsheeteditor|presentationeditor)/) || [])[1] || 'ismeretlen'

  const felderites = await frame.evaluate(BONGESZO_FELDERITES_CALLMETHOD)
  if (!felderites.ut) {
    await bongeszo.close().catch(() => {})
    return { ok: false, outcome: 'nincs-api', mag, indok: 'a callMethod egyik ismert helyen sem talalhato (window.editor / window.Asc.editor)' }
  }

  return { ok: true, bongeszo, frame, mag, apiHely: felderites.ut }
}

// Runs `Asc.editor.callMethod(name, args, callback)` in the frame and normalises the three
// outcomes (call threw / callback never fired / callback fired) that both functions below need.
function frameCallMethod(frame, apiHely, name, args) {
  return frame.evaluate(({ u, n, a }) => new Promise((resolve) => {
    let cb = false
    try {
      const e = u === 'window.editor' ? window.editor : window.Asc.editor
      e.callMethod(n, a, (v) => { cb = true; resolve({ hivva: true, ertek: v === null || v === undefined ? null : String(v) }) })
    } catch (err) { return resolve({ hivva: false, hiba: String(err && err.message).slice(0, 200) }) }
    setTimeout(() => { if (!cb) resolve({ hivva: false, hiba: 'a callback 6 mp alatt nem hivodott' }) }, 6000)
  }), { u: apiHely, n: name, a: args })
}

// READS whatever is CURRENTLY selected in the live editor -- it does not select anything itself.
// A caller wanting to read a specific range has to select it first (a separate concern, not this
// tool's job); an empty result here legitimately means "nothing is selected right now", not an
// error. Proven cross-core (docx/xlsx/pptx) via `Asc.editor.callMethod('GetSelectedText', ...)`
// -- deliberately NOT `callCommand`, which is the document-editing surface (Api.*), not this
// plugin-method one.
async function getSelectionFromDocument({ url, user, pass, fileId, timeoutMs = 60000 }) {
  const belepett = await nyitottCallMethodFrame({ url, user, pass, fileId, timeoutMs })
  if (!belepett.ok) return belepett
  const { bongeszo, frame, mag, apiHely } = belepett
  try {
    const eredmeny = await frameCallMethod(frame, apiHely, 'GetSelectedText', [])
    if (!eredmeny.hivva) return { ok: false, outcome: 'a-hivas-dobott', mag, apiHely, indok: eredmeny.hiba }
    return { ok: true, outcome: 'ok', mag, apiHely, text: eredmeny.ertek ?? '' }
  } finally {
    await bongeszo.close().catch(() => {})
  }
}

// WRITES at the current selection via `PasteText` (proven: a masik kliens ujratoltes nelkul,
// 1 mp-en belul latja). Same verification discipline as writeOperationsToDocument above: the
// callMethod's own callback is NOT proof (a sandboxed write from inside it is unreadable from
// outside) -- only the file re-downloaded via WebDAV after the post-save wait counts.
async function replaceSelectionInDocument({ url, user, pass, fileId, text, timeoutMs = 60000, postSaveWaitMs = 70000 }) {
  const utvonal = await findPathByFileId({ url, user, pass, fileId, timeoutMs })
  if (!utvonal.ok) return { ok: false, outcome: 'fileid-nem-oldhato-fel', indok: utvonal.indok }
  const elotte = await letoltVeMeret(url, utvonal.href, user, pass)
  if (!elotte.ok) return { ok: false, outcome: 'letoltes-elotte-sikertelen', indok: elotte.indok }

  const belepett = await nyitottCallMethodFrame({ url, user, pass, fileId, timeoutMs })
  if (!belepett.ok) return belepett
  const { bongeszo, frame, mag, apiHely } = belepett
  try {
    const hivas = await frameCallMethod(frame, apiHely, 'PasteText', [text])
    if (!hivas.hivva) return { ok: false, outcome: 'a-hivas-dobott', mag, apiHely, indok: hivas.hiba }

    await frame.page().waitForTimeout(6000)
    await bongeszo.close().catch(() => {})
    await new Promise((resolve) => setTimeout(resolve, postSaveWaitMs))

    const utana = await letoltVeMeret(url, utvonal.href, user, pass)
    if (!utana.ok) return { ok: false, outcome: 'letoltes-utana-sikertelen', mag, apiHely, indok: utana.indok }

    const igazolas = ellenorizDocxSzoveget(mag, text, utana.buf)
    return {
      ok: true,
      outcome: elotte.bytes !== utana.bytes ? 'meret-valtozott' : 'meret-valtozatlan',
      mag,
      apiHely,
      bytesElotte: elotte.bytes,
      bytesUtana: utana.bytes,
      tartalomIgazolva: igazolas.tartalomIgazolva,
      igazolasIndok: igazolas.igazolasIndok,
    }
  } finally {
    await bongeszo.close().catch(() => {})
  }
}

async function letoltVeMeret(url, href, user, pass) {
  const auth = Buffer.from(`${user}:${pass}`, 'utf8').toString('base64')
  const res = await fetch(`${url}${href}`, { headers: { Authorization: `Basic ${auth}` } })
  if (!res.ok) return { ok: false, indok: `letoltes HTTP ${res.status}` }
  const buf = Buffer.from(await res.arrayBuffer())
  return { ok: true, bytes: buf.length, buf }
}

// docx-en a mentett word/document.xml-ben keresi a beirt szoveget (ugyanaz a modszer, mint
// writeOperationsToDocument sajat, inline verzioja) -- xlsx/pptx eseten NEM-MERT marad, nem
// allitjuk sem igennek, sem nemnek.
function ellenorizDocxSzoveget(mag, text, buf) {
  if (mag !== 'documenteditor' || !text.trim()) {
    return { tartalomIgazolva: null, igazolasIndok: 'ehhez a maghoz nincs szoveg-alapu tartalom-ellenorzes -- csak a byte-meret hasonlitva' }
  }
  const tmpFile = path.join(os.tmpdir(), `coedit-replacesel-verify-${process.pid}-${Date.now()}.docx`)
  fs.writeFileSync(tmpFile, buf)
  try {
    const xml = execFileSync('unzip', ['-p', tmpFile, 'word/document.xml'], { encoding: 'utf8' })
    const szoveg = xml.replace(/<[^>]*>/g, '')
    const tartalomIgazolva = szoveg.includes(text)
    return {
      tartalomIgazolva,
      igazolasIndok: tartalomIgazolva
        ? 'a beirt szoveg megtalalhato a mentett word/document.xml-ben'
        : 'a beirt szoveg NEM talalhato a mentett word/document.xml-ben',
    }
  } catch (err) {
    return { tartalomIgazolva: null, igazolasIndok: `a letoltott csomag nem volt kicsomagolhato: ${err.message}` }
  } finally {
    fs.rmSync(tmpFile, { force: true })
  }
}

module.exports.getSelectionFromDocument = getSelectionFromDocument
module.exports.replaceSelectionInDocument = replaceSelectionInDocument
