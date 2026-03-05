/**
 * 
 * Este arquivo contém todos os gatilhos (triggers) e automações do projeto.
 * ATUALIZADO (11/11/2025): Esta versão foi simplificada para remover
 * a dependência da aba "Lista_Alunos" (que foi excluída).
 */

// ATUALIZADO: ID da sua nova planilha "BaseDadosFBKMKLN"
const SS_ID = "11uTU-BIiLvvXWD1rDmff-lcP3vHYKCr_1cZJUM-nL7U";

/**
 * Esta função é o gatilho principal.
 * Ela é executada AUTOMATICAMENTE toda vez que o Google Form é enviado
 * e uma nova linha é adicionada à aba "cadastro_de_alunos".
 *
 * @param {object} e - O objeto de evento 'onFormSubmit' fornecido pelo Google.
 */
function processarNovoCadastro(e) {
  try {
    // 1. O 'e.values' é um array com os dados da linha que acabou de ser adicionada.
    const values = e.values;
    Logger.log(`Gatilho disparado. Processando: ${values}`);

    // 2. Mapeamento das colunas com base no seu CSV (Índice começa em 0)
    // Coluna N = Login do App (Índice 13)
    // Coluna O = Senha do App (Índice 14)
    const login = values[13];
    const senha = values[14];
    
    // 3. Definimos o status padrão para todo novo usuário.
    // Isso garante que o administrador precise aprovar manualmente.
    const status = "Inativo"; 

    // 4. Verifica se o login e a senha não estão vazios
    if (!login || !senha) {
      Logger.log("Falha na automação: Login ou Senha não preenchidos no formulário.");
      return; // Interrompe a execução
    }
    
    // 5. Acessa a aba "Acessos"
    const ss = SpreadsheetApp.openById(SS_ID);
    const accessSheet = ss.getSheetByName("Acessos");

    // 6. Adiciona os novos dados como uma nova linha estática na aba "Acessos"
    // Esta aba agora é usada *apenas* para autenticação.
    accessSheet.appendRow([
      login,    // Coluna A (LOGIN)
      senha,    // Coluna B (Senha)
      status    // Coluna C (STATUS)
    ]);

    // Log de Sucesso
    Logger.log(`Novo usuário '${login}' processado e adicionado à aba 'Acessos' como 'Inativo'.`);

  } catch (err) {
    // Log de Erro
    Logger.log(`ERRO CRÍTICO em 'processarNovoCadastro': ${err}`);
  }
}

/**
 * FUNÇÃO DE INSTALAÇÃO (Executar 1 vez)
 * Esta é uma função manual. Você precisa executá-la UMA ÚNICA VEZ
 * para "instalar" o gatilho 'processarNovoCadastro' na sua planilha nova.
 */
function instalarGatilhoOnFormSubmit() {
  // Deleta gatilhos antigos para evitar duplicação
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === "processarNovoCadastro") {
      ScriptApp.deleteTrigger(trigger);
      Logger.log("Gatilho antigo 'processarNovoCadastro' deletado.");
    }
  }

  // Abre a planilha pelo ID
  const ss = SpreadsheetApp.openById(SS_ID);
  
  // Cria o novo gatilho
  ScriptApp.newTrigger("processarNovoCadastro")
    .forSpreadsheet(ss) // Vinculado a esta planilha
    .onFormSubmit()     // No evento de envio de formulário (quando "cadastro_de_alunos" recebe dados)
    .create();          // Cria o gatilho

  Logger.log("Gatilho 'processarNovoCadastro' instalado com sucesso!");
}