/**
 * Recebe os leads do Gerador de Copy Mantora e grava na planilha.
 *
 * Como usar:
 * 1. Crie uma planilha no Google Sheets (ex: "Leads - Gerador de Copy").
 * 2. Extensões > Apps Script. Apague o código e cole este arquivo inteiro.
 * 3. Implantar > Nova implantação > Tipo: App da Web
 *      Executar como: Eu  |  Quem tem acesso: Qualquer pessoa
 * 4. Copie a URL que termina em /exec e cole em CONFIG.ENDPOINT no index.html.
 */
const ABA = 'Leads';
const COLUNAS = ['data', 'nome', 'whatsapp', 'email', 'instagram', 'area', 'anuncia', 'template', 'origem', 'campanha', 'consentimento'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(ABA) || ss.insertSheet(ABA);
    if (sh.getLastRow() === 0) {
      sh.appendRow(COLUNAS.map(c => c.toUpperCase()));
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, COLUNAS.length).setFontWeight('bold').setBackground('#744dff').setFontColor('#ffffff');
    }
    // whatsapp como texto para não perder zeros / virar notação científica
    const linha = COLUNAS.map(c => c === 'whatsapp' ? "'" + (p[c] || '') : (p[c] === undefined ? '' : p[c]));
    sh.appendRow(linha);
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, erro: String(err) })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
