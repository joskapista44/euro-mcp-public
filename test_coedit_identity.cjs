// A CO-EDITING IRAS AZONOSSAG- ES HITELESITO-LOGIKAJA -- bongeszo nelkul, mindket iranyban.
//
// MIERT KULON TESZT: ez a modul BIZTONSAGI hatart huz. Egy tevedes itt nem hibauzenetkent
// jelentkezik, hanem ugy, hogy EGY AGENS MASVALAKI NEVEBEN ir egy KOZOS dokumentumba -- es a
// dokumentumon az fog latszani, hogy a masik irta. *Ez pont az a hibaosztaly, amit egy zold
// futas nem mutat meg.*
//
// A ZOLD AGAK ONMAGUKBAN NEM ERNEK SEMMIT ITT: minden engedelyezo eset melle tartozik egy
// MEGTAGADO, ugyanabbol a savbol.

const assert = require('assert')
const fs = require('fs')
const os = require('os')
const path = require('path')
const { detectCallerId, credentialsFor, loadPlaywright } = require('./coedit.cjs')

let osszes = 0
const hibak = []
function check(cimke, felteteles, reszlet = '') {
  osszes += 1
  if (felteteles) { console.log(`  ok    ${cimke}`) } else { hibak.push(cimke); console.log(`  BUKAS ${cimke}${reszlet ? ' -- ' + reszlet : ''}`) }
}

console.log('\n[1] A HIVO AZONOSITASA a munkakonyvtarbol')
const env1 = { EURO_COEDIT_AGENTS: 'alpha,gamma' }
const r1 = detectCallerId(env1, '/srv/example-deployment/agents/alpha')
check('allowlistelt agens a cwd-bol -> azonositva', r1.ok === true && r1.id === 'alpha', JSON.stringify(r1))
check('es megnevezi, MIBOL kovetkezett az identitas', /munkakonyvtar/.test(r1.forras || ''), r1.forras)

console.log('\n[2] MEGTAGADAS -- es ezek a fontosabbak')
const r2 = detectCallerId(env1, '/srv/example-deployment/agents/beta')
check('NEM allowlistelt agens -> megtagadva', r2.ok === false && r2.id === '', JSON.stringify(r2))
check('  es az indok MEGNEVEZI a hivot es az allowlistet', /beta/.test(r2.indok) && /EURO_COEDIT_AGENTS/.test(r2.indok), r2.indok)
// A denial message must name the ACTUAL scope this gate enforces (the code path is denied, not
// co-editing writes in general) -- asserted explicitly here, not just that it denied, so a future
// wording change cannot silently regress into a broader, overstated claim.
check('  es a TENYLEGES hatokort nevezi meg (a kod-utat tagadjuk meg, nem a co-editing irast magat)',
  /kod-utat/i.test(r2.indok) && /MAS agens neveben/.test(r2.indok) && /sajat NC-fiokkal/.test(r2.indok), r2.indok)
check('  es NEM allitja tobbet, mint amit kikenyszerit ("per-agens engedelyezett" a co-editing irasra -- eltavolitva)',
  !/co-editing iras per-agens engedelyezett/.test(r2.indok), r2.indok)

const r3 = detectCallerId(env1, '/tmp/valahol/maswhere')
check('az agens-fan KIVULI cwd -> megtagadva (fail-closed)', r3.ok === false && r3.id === '', JSON.stringify(r3))

const r4 = detectCallerId({ EURO_COEDIT_AGENTS: '' }, '/srv/example-deployment/agents/alpha')
check('URES allowlist -> senki nem megy at (nem "mindenki")', r4.ok === false, JSON.stringify(r4))

console.log('\n[3] *** AZ EXPLICIT VALTOZO SEM KERULHETI MEG AZ ALLOWLISTET ***')
// A flotta EGY unix-useren fut: barmelyik agens beallithatna barmilyen EURO_AGENT_ID-t.
// Ha ez megkerulne az allowlistet, az egesz kapu diszites lenne.
const hamis = detectCallerId({ EURO_COEDIT_AGENTS: 'alpha', EURO_AGENT_ID: 'someone-else' }, '/srv/example-deployment/agents/alpha')
check('EURO_AGENT_ID=someone-else, de o NINCS az allowlisten -> MEGTAGADVA', hamis.ok === false, JSON.stringify(hamis))
const jo = detectCallerId({ EURO_COEDIT_AGENTS: 'someone-else', EURO_AGENT_ID: 'someone-else' }, '/srv/example-deployment/agents/alpha')
check('EURO_AGENT_ID=someone-else ES allowlisten van -> atmegy (poz. kontroll)', jo.ok === true && jo.id === 'someone-else', JSON.stringify(jo))

console.log('\n[5] A PLAYWRIGHT FELOLDASA -- harom allapot, nem ketto')
const pwr = loadPlaywright()
// Nem allitjuk, hogy MEGVAN: a bridge-nek nincs ilyen fuggosege, es a tartalek ut kornyezetfuggo.
// Amit allitunk: a valasz MINDIG megmondja, melyik allapotban van, es sosem omlik ossze.
check('a feloldas nem dob kivetelt, es allapotot ad', typeof pwr.ok === 'boolean', JSON.stringify({ ok: pwr.ok, honnan: pwr.honnan }))
check('  ha nem megy, MEGNEVEZI az okot (nem csak false)', pwr.ok === true || Boolean(pwr.indok), pwr.indok || '(ok)')
if (pwr.ok) console.log(`     (most feloldodott -- ${pwr.honnan})`)
else console.log(`     (most NEM oldodott fel -- ez NEM-MERT allapot, nem hiba: ${pwr.indok})`)

// A hitelesito-agak ASZINKRONOK (a vault-ut halozatot hasznal), ezert kulon blokkban allnak.
// A vault-hivast INJEKTALT `fetchFn` adja: egy valodi halozati hivas ettol a teszttol fuggetlenul
// is elbukhatna, es akkor nem a LOGIKA-t merne. A poz. kontroll ugyanabbol a savbol jon: ugyanaz
// a fake fetch adja a sikeres es a 404-es agat is.
async function hitelesitoAgak() {
  console.log('\n[4] A HITELESITO -- a kulcs NEVE megjelenik, az ERTEKE soha')
  const TITOK_ENV = 'TITKOS-ERTEK-NEM-JELENHET-MEG'
  const teljesEnv = {
    EURO_COEDIT_NC_URL: 'https://pelda.hu/',
    ALPHA_NEXTCLOUD_USER: 'alpha',
    ALPHA_NEXTCLOUD_APP_PASSWORD: TITOK_ENV,
  }
  const c1 = await credentialsFor('alpha', teljesEnv)
  check('teljes .env-konfig -> ok', c1.ok === true && c1.user === 'alpha', JSON.stringify({ ...c1, pass: '<elrejtve>' }))
  check('  a zaro / levagva az URL-rol', c1.url === 'https://pelda.hu', c1.url)
  check('  es MEGNEVEZI a hitelesito forrasat (.env)', /\.env/.test(c1.forras || ''), c1.forras)

  // A masik agens kulcsai hianyoznak a .env-bol -> megtagadas (nincs masodik forras).
  const c2 = await credentialsFor('gamma', teljesEnv)
  check('hianyzo agens-kulcsok -> nem ok', c2.ok === false, JSON.stringify(c2))
  check('  es MEGNEVEZI a hianyzo .env-kulcsot',
    /GAMMA_NEXTCLOUD_APP_PASSWORD/.test(c2.indok), c2.indok)

  console.log('\n[4b] *** A CELPELDANY NEM OROKLODIK AZ ELES NEXTCLOUD_URL-BOL ***')
  // Ez a fajl legfontosabb NEGATIV kontrollja a cel-savon: a kozos .env-ben a NEXTCLOUD_URL az
  // ELES peldanyra mutat. Ha ez az ut visszaesne ra, egy elfelejtett kulcs eles fajlba iratna.
  const c3 = await credentialsFor('alpha', { ...teljesEnv, EURO_COEDIT_NC_URL: '', NEXTCLOUD_URL: 'https://eles.pelda.hu' })
  check('van NEXTCLOUD_URL, de nincs EURO_COEDIT_NC_URL -> MEGTAGADVA (nincs visszaeses)',
    c3.ok === false && /EURO_COEDIT_NC_URL/.test(c3.indok), c3.indok)
  check('  es a megtagadas NEM az eles cimet fogadja el', !c3.url, String(c3.url))

  // *** A LEGFONTOSABB ELLENORZES EBBEN A FAJLBAN: *** egy hibauzenet, ami a jelszot is kiirja,
  // pont abbol csinal naplo-bejegyzest, amit vedeni akarunk.
  const mindenUzenet = [c1.indok, c2.indok, c3.indok, JSON.stringify(c2)].join(' ')
  check('*** a TITOK SOHA nem jelenik meg egyetlen uzenetben sem ***',
    !mindenUzenet.includes(TITOK_ENV), mindenUzenet.slice(0, 120))
}

// A KAPU VAKFOLTJA A HIVOHOZ TARTOZIK: a DocBuilder-tilalom egy MASIK utra kuldi a hivot, es ha
// az szinten zart, azt a megtagadasnak ki kell mondania. Ezek a fuggvenyek adjak hozza a valaszt.
function utAllapotAgak() {
  console.log('\n[6] A MASIK UT ALLAPOTA -- amit a megtagadas magaval visz')
  const { bridgeEnvKulcs, coeditUtAllapota } = require('./coedit.cjs')

  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'coedit-env-'))
  fs.writeFileSync(path.join(dir, '.env'),
    '# megjegyzes\nEURO_COEDIT_AGENTS=alpha,gamma\nTITKOS_KULCS=NEM-KERHETO-LE-VELETLENUL\n')

  check('bridgeEnvKulcs kiolvassa a kert kulcsot', bridgeEnvKulcs('EURO_COEDIT_AGENTS', dir) === 'alpha,gamma', String(bridgeEnvKulcs('EURO_COEDIT_AGENTS', dir)))
  check('  nem letezo kulcsra null (nem ures sztring)', bridgeEnvKulcs('NINCS_ILYEN', dir) === null)
  // *** A LENYEG: ez az ut NEM tolti be a kornyezetbe a fajlt. *** Kulonben egy megtagado valasz
  // kiegeszitese titkokat allitana be olyan kodutakon, aminek semmi koze hozzajuk.
  check('*** NEM allitja be a process.env-et (a titok nem szivarog at) ***',
    !('TITKOS_KULCS' in process.env) && !('EURO_COEDIT_AGENTS' in process.env))

  const nyitva = coeditUtAllapota({}, '/srv/example-deployment/agents/alpha', dir)
  check('allowlisten levo hivo -> a masik ut ELERHETO', nyitva.elerheto === true, JSON.stringify(nyitva))
  const zarva = coeditUtAllapota({}, '/srv/example-deployment/agents/beta', dir)
  check('NEM allowlistelt hivo -> NEM elerheto, es megnevezi az okot',
    zarva.elerheto === false && /EURO_COEDIT_AGENTS/.test(zarva.indok), JSON.stringify(zarva))
  check('  coeditUtAllapota UGYANAZT a javitott szoveget adja tovabb (nem sajat, kulon uzenetet)',
    /kod-utat/i.test(zarva.indok) && !/co-editing iras per-agens engedelyezett/.test(zarva.indok), zarva.indok)
  // A mai elo allapot alakja: a .env letezik, de a kulcs nincs benne.
  const uresDir = fs.mkdtempSync(path.join(os.tmpdir(), 'coedit-env-ures-'))
  fs.writeFileSync(path.join(uresDir, '.env'), 'MAS_KULCS=1\n')
  const ures = coeditUtAllapota({}, '/srv/example-deployment/agents/alpha', uresDir)
  check('nincs allowlist-kulcs a .env-ben -> NEM elerheto (nem "mindenki")', ures.elerheto === false, JSON.stringify(ures))

  fs.rmSync(dir, { recursive: true, force: true })
  fs.rmSync(uresDir, { recursive: true, force: true })
}

hitelesitoAgak().then(() => {
  utAllapotAgak()
  console.log(`\nellenorzesek: ${osszes - hibak.length} ok, ${hibak.length} bukas`)
  assert.strictEqual(hibak.length, 0, `bukott: ${hibak.join(' | ')}`)
})
