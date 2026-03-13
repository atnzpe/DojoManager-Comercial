class DBQuery {

 constructor(table){
   this.data = getTable(table);
 }

 where(field, value){
   this.data = this.data.filter(r => r[field] == value);
   return this;
 }

 get(){
   return this.data;
 }

 first(){
   return this.data[0] || null;
 }

}

/**
 * ============================================================================
 * MOTOR DINÂMICO DE GRADUAÇÕES (O "PROCV" DO BACKEND)
 * ============================================================================
 * Lê a aba 'GRADUACAO' e cria um mapa para busca instantânea.
 */
function getMapaGraduacoes() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // Tenta encontrar a aba com acento ou sem acento
    const sheet = ss.getSheetByName("GRADUACAO") || ss.getSheetByName("Graduação") || ss.getSheetByName("Graduacao"); 
    if (!sheet) return {};

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());

    const colGrad = headers.indexOf("graduação") !== -1 ? headers.indexOf("graduação") : headers.indexOf("faixa / nivel");
    const colId = headers.indexOf("id");
    const colModalidade = headers.indexOf("modalidade");
    const colNivel = headers.indexOf("nivel");

    if (colGrad === -1 || colId === -1) return {};

    const mapa = {};
    for (let i = 1; i < data.length; i++) {
      const nomeFaixa = String(data[i][colGrad]).trim().toLowerCase();
      if(nomeFaixa) {
        mapa[nomeFaixa] = {
          nome: String(data[i][colGrad]).trim(),
          id: parseInt(data[i][colId]) || 0,
          modalidade: colModalidade > -1 ? String(data[i][colModalidade]).trim() : "Geral",
          nivel: colNivel > -1 ? String(data[i][colNivel]).trim().toUpperCase() : "ALUNO"
        };
      }
    }
    return mapa;
  } catch (e) {
    console.error("[DBQuery] Erro ao ler aba GRADUACAO: ", e);
    return {};
  }
}