/**
 * ARQUIVO: Auth.gs
 * DESCRIÇÃO: Backend de Autenticação.
 */

const SHEET_ID = "11uTU-BIiLvvXWD1rDmff-lcP3vHYKCr_1cZJUM-nL7U";
const COLUNA_LOGIN_CADASTRO = "LOGIN"; // Corrigido (adicionado ;)

function logarUsuario(login_attempt, senha) {
  Logger.log(`[AUTH] Tentativa: ${login_attempt}`);

  if (!login_attempt || !senha) return null;

  const inputLogin = String(login_attempt).trim().toLowerCase();
  const inputSenha = String(senha).trim();

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    
    // 1. Validação em 'Acessos'
    const accessSheet = ss.getSheetByName("Acessos");
    const data = accessSheet.getDataRange().getValues();
    const headers = data.shift();

    const idxLogin = headers.indexOf("LOGIN");
    const idxSenha = headers.indexOf("Senha");
    const idxStatus = headers.indexOf("STATUS");

    if (idxLogin === -1) return null;

    let autenticado = false;
    for (let row of data) {
      if (String(row[idxLogin]).trim().toLowerCase() === inputLogin &&
          String(row[idxSenha]).trim() === inputSenha) {
        if (String(row[idxStatus]).trim() === "Ativo") {
          autenticado = true;
          break;
        } else {
          Logger.log("[AUTH] Usuário Inativo");
          return null;
        }
      }
    }

    if (!autenticado) return null;

    // 2. Dados em 'cadastro_de_alunos'
    const userSheet = ss.getSheetByName("cadastro_de_alunos");
    const userData = userSheet.getDataRange().getValues();
    const userHeaders = userData.shift();
    const idxUserLogin = userHeaders.indexOf(COLUNA_LOGIN_CADASTRO);

    if (idxUserLogin === -1) {
      // Retorna dados básicos se não achar a coluna na outra aba
      return JSON.stringify({ "LOGIN": inputLogin, "STATUS": "Ativo" });
    }

    let perfil = {};
    for (let row of userData) {
      if (String(row[idxUserLogin]).trim().toLowerCase() === inputLogin) {
        for (let i = 0; i < userHeaders.length; i++) {
            perfil[userHeaders[i]] = row[i];
        }
        perfil["STATUS"] = "Ativo";
        return JSON.stringify(perfil);
      }
    }

    return JSON.stringify({ "LOGIN": inputLogin, "STATUS": "Ativo" });

  } catch (err) {
    Logger.log(`[AUTH ERROR] ${err}`);
    return null;
  }
}