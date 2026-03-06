/**
 * ============================================================================
 * ARQUIVO: Auth.gs
 * DESCRIÇÃO: Backend de Autenticação (Arquitetura SaaS / White-Label).
 * ============================================================================
 */

const COLUNA_LOGIN_CADASTRO = "LOGIN";

function logarUsuario(login_attempt, senha) {
  console.log(`[AUTH] Tentativa de login: ${login_attempt}`);

  if (!login_attempt || !senha) {
    console.warn("[AUTH] Login ou senha vazios.");
    return null;
  }

  const inputLogin = String(login_attempt).trim().toLowerCase();
  const inputSenha = String(senha).trim();

  try {
    // 🛡️ CORREÇÃO CRÍTICA PARA SAAS: Usa a planilha ativa em vez de um ID fixo.
    // Quando o cliente copiar a planilha, o código apontará automaticamente para a cópia dele!
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Validação na aba 'Acessos'
    const accessSheet = ss.getSheetByName("Acessos");

    if (!accessSheet) {
      console.error("[AUTH ERROR] Aba 'Acessos' não encontrada na planilha.");
      return null;
    }

    const data = accessSheet.getDataRange().getValues();
    if (data.length < 2) return null; // Planilha vazia

    // Pega os cabeçalhos e limpa espaços
    const headers = data.shift().map(h => String(h).trim());

    const idxLogin = headers.indexOf("LOGIN");
    const idxSenha = headers.indexOf("Senha");
    const idxStatus = headers.indexOf("STATUS");

    if (idxLogin === -1 || idxSenha === -1) {
      console.error("[AUTH ERROR] Colunas 'LOGIN' ou 'Senha' não encontradas na aba Acessos.");
      return null;
    }

    let autenticado = false;
    for (let row of data) {
      if (String(row[idxLogin]).trim().toLowerCase() === inputLogin &&
        String(row[idxSenha]).trim() === inputSenha) {

        // Verifica o Status (se a coluna não existir, assume como Ativo por padrão)
        const statusUser = idxStatus !== -1 ? String(row[idxStatus]).trim() : "Ativo";

        if (statusUser.toLowerCase() === "ativo") {
          autenticado = true;
          break;
        } else {
          console.warn(`[AUTH] Usuário Inativo tentou logar: ${inputLogin}`);
          return null;
        }
      }
    }

    if (!autenticado) {
      console.warn(`[AUTH] Credenciais inválidas para: ${inputLogin}`);
      return null;
    }

    // 2. Coleta de Dados em 'cadastro_de_alunos'
    const userSheet = ss.getSheetByName("cadastro_de_alunos");
    if (!userSheet) {
      // Se a aba de cadastro não existir, retorna dados mínimos para não quebrar a sessão
      return JSON.stringify({ "LOGIN": inputLogin, "STATUS": "Ativo" });
    }

    const userData = userSheet.getDataRange().getValues();
    if (userData.length < 2) return JSON.stringify({ "LOGIN": inputLogin, "STATUS": "Ativo" });

    const userHeaders = userData.shift().map(h => String(h).trim());
    const idxUserLogin = userHeaders.indexOf(COLUNA_LOGIN_CADASTRO);

    if (idxUserLogin === -1) {
      return JSON.stringify({ "LOGIN": inputLogin, "STATUS": "Ativo" });
    }

    let perfil = {};
    for (let row of userData) {
      if (String(row[idxUserLogin]).trim().toLowerCase() === inputLogin) {
        for (let i = 0; i < userHeaders.length; i++) {
          perfil[userHeaders[i]] = row[i];
        }
        perfil["STATUS"] = "Ativo";
        console.log(`[AUTH] Sucesso. Perfil carregado para: ${inputLogin}`);
        return JSON.stringify(perfil);
      }
    }

    // Fallback: Autenticado na aba Acesso, mas sem ficha no Cadastro
    return JSON.stringify({ "LOGIN": inputLogin, "STATUS": "Ativo" });

  } catch (err) {
    console.error(`[AUTH FATAL ERROR] ${err.message}`);
    return null;
  }
}