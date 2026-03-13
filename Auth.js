/**
 * ============================================================================
 * ARQUIVO: Auth.js
 * DESCRIÇÃO: Backend de Autenticação (Arquitetura SaaS / White-Label).
 * ============================================================================
 */

const COLUNA_LOGIN_CADASTRO = "LOGIN";

function logarUsuario(login_attempt, senha) {
  if (!login_attempt || !senha) return null;

  const inputLogin = String(login_attempt).trim().toLowerCase();
  const inputSenha = String(senha).trim();

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Validação na aba 'Acessos'
    const accessSheet = ss.getSheetByName("Acessos");
    if (!accessSheet) return null;

    const data = accessSheet.getDataRange().getValues();
    if (data.length < 2) return null;

    const headers = data.shift().map(h => String(h).trim());
    const idxLogin = headers.indexOf("LOGIN");
    const idxSenha = headers.indexOf("Senha");
    const idxStatus = headers.indexOf("STATUS");

    if (idxLogin === -1 || idxSenha === -1) return null;

    let autenticado = false;
    for (let row of data) {
      if (String(row[idxLogin]).trim().toLowerCase() === inputLogin && String(row[idxSenha]).trim() === inputSenha) {
        const statusUser = idxStatus !== -1 ? String(row[idxStatus]).trim() : "Ativo";
        if (statusUser.toLowerCase() === "ativo") { autenticado = true; break; }
        else { return null; }
      }
    }

    if (!autenticado) return null;

    // 2. Coleta de Dados e PROCV Interno
    const userSheet = ss.getSheetByName("cadastro_de_alunos");
    if (!userSheet) return JSON.stringify({ "LOGIN": inputLogin, "STATUS": "Ativo" });

    const userData = userSheet.getDataRange().getValues();
    const userHeaders = userData.shift().map(h => String(h).trim());
    const idxUserLogin = userHeaders.indexOf(COLUNA_LOGIN_CADASTRO);

    if (idxUserLogin === -1) return JSON.stringify({ "LOGIN": inputLogin, "STATUS": "Ativo" });

    let perfil = {};
    for (let row of userData) {
      if (String(row[idxUserLogin]).trim().toLowerCase() === inputLogin) {
        for (let i = 0; i < userHeaders.length; i++) {
          perfil[userHeaders[i]] = row[i];
        }
        perfil["STATUS"] = "Ativo";

        // 🧠 PROCV MÁGICO ACONTECE AQUI
        const mapaGrad = getMapaGraduacoes();
        const userGradStr = String(perfil["GRADUACAO_ATUAL"] || perfil["Graduação"] || "Iniciante").trim().toLowerCase();
        const gradInfo = mapaGrad[userGradStr] || { id: 0, nivel: "ALUNO", modalidade: "Geral" };

        perfil["Nível do Praticante"] = gradInfo.nivel;
        perfil["Modalidade"] = gradInfo.modalidade;
        perfil["idGraduacao"] = gradInfo.id;

        // CRIA FLAGS DINÂMICAS PARA O DASHBOARD (Acaba com o Hardcode)
        const nivelOficial = String(gradInfo.nivel).toUpperCase();
        perfil.isInstrutor = nivelOficial.includes("INSTRUTOR") || nivelOficial.includes("PROFESSOR") || nivelOficial.includes("MESTRE");
        perfil.isMestre = nivelOficial.includes("MESTRE");
        perfil.isAdmin = nivelOficial.includes("ADMIN") || nivelOficial.includes("MESTRE");

        return JSON.stringify(perfil);
      }
    }
    return JSON.stringify({ "LOGIN": inputLogin, "STATUS": "Ativo" });
  } catch (err) { return null; }
}