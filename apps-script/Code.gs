/**
 * Foglio "Invitati e Tavoli" nel formato di La Stella + ricezione delle conferme dal sito.
 *
 * Uso (una volta sola):
 *  1. Apri il Google Sheet -> Estensioni -> Apps Script, incolla questo file.
 *  2. Esegui setup() e autorizza: crea/formatta i fogli "Invitati" e "Tavoli".
 *  3. Distribuisci -> Nuova distribuzione -> App web
 *     (Esegui come: me, Chi ha accesso: Chiunque) e copia l'URL.
 *  4. Incolla l'URL in SHEET_URL dentro index.html e index-con-gioco.html.
 */

var SHEET_INV = 'Invitati';
var SHEET_TAV = 'Tavoli';
var MAX_ROWS_INV = 300;
var MAX_TAVOLI = 40;
var MAX_PERSONE_PER_INVIO = 40;

var ALLERGIE = ['Glutine (cereali)', 'Crostacei', 'Uova', 'Pesce', 'Arachidi', 'Soia', 'Latte e derivati',
  'Frutta a guscio', 'Sedano', 'Senape', 'Sesamo', 'Solfiti', 'Lupini', 'Molluschi'];
var INTOLLERANZE = ['Celiachia', 'Lattosio', 'Nichel', 'Istamina', 'Fruttosio', 'Favismo', 'Gravidanza', 'Funghi'];
var DIETE = ['Vegetariano', 'Vegano', 'Fruttariano', 'Pescetariano', 'Crudista', 'Halal', 'Kosher'];

function intestazioni_() {
  return ['Nome', 'Cognome', 'Fascia', 'Dieta', 'Seggiolone']
    .concat(ALLERGIE.map(function (a) { return 'Allergia – ' + a; }))
    .concat(INTOLLERANZE.map(function (a) { return 'Intolleranza – ' + a; }))
    .concat(['Altre allergie / intolleranze (scrivere)', 'Note per la cucina', 'Nome completo', 'Stato tavolo']);
}

function setup() {
  var ss = SpreadsheetApp.getActive();
  var inv = ss.getSheetByName(SHEET_INV) || ss.getSheets()[0].setName(SHEET_INV);
  var head = intestazioni_();
  var nCols = head.length;

  inv.getRange(1, 1, 1, nCols).setValues([head]).setFontWeight('bold').setFontColor('#FFFFFF')
    .setBackground('#2E3540').setWrap(true).setVerticalAlignment('middle');
  inv.getRange(1, 6, 1, ALLERGIE.length).setBackground('#C46A4F');
  inv.getRange(1, 6 + ALLERGIE.length, 1, INTOLLERANZE.length).setBackground('#D9A441');
  inv.setFrozenRows(1);
  inv.setFrozenColumns(2);
  inv.setColumnWidths(1, 2, 140);
  inv.setColumnWidth(3, 90);
  inv.setColumnWidth(4, 120);
  inv.setColumnWidths(5, 23, 55);
  inv.setColumnWidths(28, 2, 220);
  inv.setColumnWidths(30, 2, 150);

  var rows = MAX_ROWS_INV;
  inv.getRange(2, 3, rows, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Adulto', 'Bambino', 'Neonato', 'Staff'], true).setAllowInvalid(false).build());
  inv.getRange(2, 4, rows, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(DIETE, true).setAllowInvalid(false).build());
  inv.getRange(2, 5, rows, 1 + ALLERGIE.length + INTOLLERANZE.length).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['X'], true).setAllowInvalid(false).build())
    .setHorizontalAlignment('center');

  var colNome = 'AD', colStato = 'AE';
  var fNome = [], fStato = [];
  for (var r = 2; r <= rows + 1; r++) {
    fNome.push(['=IF(OR(A' + r + '<>"",B' + r + '<>""),TRIM(A' + r + '&" "&B' + r + '),"")']);
    fStato.push(['=IF(' + colNome + r + '="","",IF(COUNTIF(Tavoli!$E$2:$AB$41,' + colNome + r + ')=0,"Da assegnare",' +
      'IF(COUNTIF(Tavoli!$E$2:$AB$41,' + colNome + r + ')>1,"DUPLICATO","Tavolo "&SUMPRODUCT((Tavoli!$E$2:$AB$41=' +
      colNome + r + ')*Tavoli!$A$2:$A$41))))']);
  }

  var tav = ss.getSheetByName(SHEET_TAV) || ss.insertSheet(SHEET_TAV);
  var thead = ['N. tavolo', 'Nome tavolo (facoltativo)', 'Forma tavolo', 'N. posti'];
  for (var p = 1; p <= 24; p++) thead.push('Posto ' + p);
  thead.push('Occupati', 'Controllo');
  tav.getRange(1, 1, 1, thead.length).setValues([thead]).setFontWeight('bold').setFontColor('#FFFFFF')
    .setBackground('#2E3540').setWrap(true);
  tav.setFrozenRows(1);
  tav.setFrozenColumns(1);
  var nums = [], cCount = [], cCheck = [];
  for (var t = 2; t <= MAX_TAVOLI + 1; t++) {
    nums.push([t - 1]);
    cCount.push(['=COUNTA(E' + t + ':AB' + t + ')']);
    cCheck.push(['=IF(AND(C' + t + '="",D' + t + '="",AC' + t + '=0),"",IF(OR(C' + t + '="",D' + t + '=""),"Indicare forma e posti",' +
      'IF(AC' + t + '>D' + t + ',"Troppi invitati",IF(AC' + t + '<D' + t + ',"Posti liberi: "&(D' + t + '-AC' + t + '),"OK"))))']);
  }
  tav.getRange(2, 1, MAX_TAVOLI, 1).setValues(nums);
  tav.getRange(2, 29, MAX_TAVOLI, 1).setFormulas(cCount);
  tav.getRange(2, 30, MAX_TAVOLI, 1).setFormulas(cCheck);
  tav.getRange(2, 3, MAX_TAVOLI, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Tondo', 'Rettangolare'], true).setAllowInvalid(false).build());
  tav.getRange(2, 4, MAX_TAVOLI, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThanOrEqualTo(1).setAllowInvalid(false).build());
  tav.getRange(2, 5, MAX_TAVOLI, 24).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(inv.getRange('AD2:AD' + (rows + 1)), true).setAllowInvalid(false).build());

  // Le formule di "Invitati" si scrivono dopo la creazione di "Tavoli" per evitare #REF!
  inv.getRange(2, 30, rows, 1).setFormulas(fNome);
  inv.getRange(2, 31, rows, 1).setFormulas(fStato);
}

/** Riceve le conferme dal sito e aggiunge una riga per persona in "Invitati". */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var data = JSON.parse(e.postData.contents);
    var persone = (data.persone || []).slice(0, MAX_PERSONE_PER_INVIO);
    if (!persone.length) return risposta_({ ok: false, errore: 'nessuna persona' });

    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_INV);
    var colA = sh.getRange(2, 1, MAX_ROWS_INV, 2).getValues();
    var riga = 2;
    while (riga - 2 < colA.length && (colA[riga - 2][0] !== '' || colA[riga - 2][1] !== '')) riga++;
    if (riga - 1 + persone.length > MAX_ROWS_INV + 1) return risposta_({ ok: false, errore: 'foglio pieno' });

    var righe = persone.map(function (p) {
      var all = p.allergie || [], int = p.intolleranze || [];
      var out = [pulisci_(p.nome), pulisci_(p.cognome),
        p.fascia === 'Bambino' ? 'Bambino' : 'Adulto',
        DIETE.indexOf(p.dieta) > -1 ? p.dieta : '',
        p.seggiolone ? 'X' : ''];
      ALLERGIE.forEach(function (a) { out.push(all.indexOf(a) > -1 ? 'X' : ''); });
      INTOLLERANZE.forEach(function (a) { out.push(int.indexOf(a) > -1 ? 'X' : ''); });
      out.push(pulisci_(p.altre), '');
      return out;
    });
    sh.getRange(riga, 1, righe.length, righe[0].length).setValues(righe);
    return risposta_({ ok: true, righe: righe.length });
  } catch (err) {
    return risposta_({ ok: false, errore: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Evita che testo libero venga interpretato come formula nel foglio.
function pulisci_(v) {
  var s = String(v == null ? '' : v).trim().slice(0, 200);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function risposta_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
