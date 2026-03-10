

/**
 * ============================================================================
 * ARQUIVO: Código.gs 
 *  
 * DESCRIÇÃO: Controller Principal e API de Dados (Versão 7.0 - CRUDs Totais)
 * AUTOR: Gleyson Atanazio  *
 * ESTE ARQUIVO É O CÉREBRO DA APLICAÇÃO.
 * ELE GERENCIA:
 * 1. Conexão com o Banco de Dados (Google Sheets).
 * 2. Autenticação de Usuários.
 * 3. Regras de Negócio (Quem vê o quê).
 * 4. Rotas de Navegação (Qual página HTML mostrar).
 * ============================================================================
 */

// --- CONSTANTES GERAIS (Configurações Globais) ---
// Nomes das abas nas planilhas. Centralizados aqui para facilitar mudanças futuras.
const NOME_ABA_ALUNOS = "cadastro_de_alunos";     // Tabela principal de usuários
const NOME_ABA_CURSOS = "Cursos";                 // Tabela de eventos/cursos
const NOME_ABA_VIDEOTECA = "Config_Videoteca";    // Tabela de links do YouTube
const NOME_ABA_LOCAIS = "Locais_de_treino";       // Tabela de academias
const NOME_ABA_PROGRAMAS = "Config_Programas";    // Tabela de PDFs técnicos
const NOME_ABA_BIBLIOTECA = "Config_Biblioteca";  // Tabela de livros extras
const NOME_ABA_AULAS = "Aulas_Em_Andamento";      // Tabela temporária para check-ins ativos
const ABA_LOGS = "Logs_Auditoria";                // Tabela de logs de segurança

// URL do script publicada (usada para redirecionamentos e links internos)
const SCRIPT_URL = ScriptApp.getService().getUrl();

// IDs de arquivos no Google Drive para assets padrão (Logo e Fundo)
const DEFAULT_ASSETS = {
  BACKGROUND_ID: "1JfeThTR7oe4w7fhg7XFk0ImdIRDOR1KF",
  ICON_ID: "1mVV2idWbyfoOP4EV76I1fHV_PaoEUcdv",
  PRIMARY_COLOR: "#FFFFFF",
  SECONDARY_COLOR: "#1e1e1e",
  TEXT_COLOR: "#ffffff",
  BTN_TEXT_COLOR: "#000000",
  BG_COLOR: "#121212" // <-- NOVO: 5ª Cor Padrão (Fundo Geral)
};



// ============================================================================
// 🧠 HELPERS DE LEITURA POR NOME DE COLUNA
// ============================================================================

/**
 * Lê uma aba inteira e retorna um array de objetos JSON baseados no cabeçalho.
 * CORREÇÃO: Converte Datas para String IMEDIATAMENTE para evitar erro "Uncaught ns"
 */
function lerTabelaDinamica(nomeAba) {
  try {
    const sheet = getSheet(nomeAba);
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    const headers = data[0].map(h => String(h).trim().toLowerCase().replace(/\s+/g, "_"));

    return data.slice(1).map((row, i) => {
      let item = { "_linha": i + 2 };
      headers.forEach((h, colIndex) => {
        let valor = row[colIndex];

        // --- BLINDAGEM DE DATA ---
        if (valor instanceof Date) {
          try {
            valor = Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd/MM/yyyy");
          } catch (e) {
            valor = String(valor);
          }
        }
        // -------------------------

        item[h] = valor;
      });
      return item;
    });
  } catch (e) {
    console.warn(`Aba ${nomeAba} vazia ou inexistente.`);
    return [];
  }
}

/**
 * Salva dados usando o nome da coluna como referência (já existente no seu código, 
 * mas reforçada aqui para o contexto financeiro).
 */
function salvarFinanceiroSeguro(nomeAba, dadosObjeto, linha = null) {
  // Chama sua função existente salvarDadosSeguro, que já é excelente para isso.
  // Apenas um wrapper para manter o contexto semântico.
  return salvarDadosSeguro(nomeAba, dadosObjeto, linha);
}

// ============================================================================
// 💰 CONTROLLER FINANCEIRO (SaaS)
// ============================================================================

// ============================================================================
// 🏗️ MÓDULO FINANCEIRO: AUTO-SETUP E RELATÓRIOS (Adicionar ao final do Código.gs)
// ============================================================================

function verificarCriarAbasFinanceiras() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const estrutura = [
    { nome: "Fin_Transacoes", colunas: ["ID_Transacao", "Data_Registro", "Tipo", "Categoria", "Descricao", "Valor", "Forma_Pagto", "Responsavel", "Login_Aluno", "Academia_Ref", "Status", "Comprovante_Url"] },
    { nome: "Fin_Pacotes", colunas: ["Nome_Pacote", "Valor_Padrao", "Duracao_Dias", "Academias_Permitidas", "Status_Pacote", "Descricao"] },
    { nome: "Fin_Assinaturas", colunas: ["Login_Aluno", "Pacote_Atual", "Data_Inicio", "Data_Fim", "Status_Assinatura", "ID_Ultima_Transacao"] },
    {
      nome: "Config_App",
      // <-- NOVO: "Cor_Fundo" adicionado no final do array de colunas
      colunas: ["Logo_URL", "Fundo_URL", "Cor_Primaria", "Cor_Secundaria", "Cor_Texto", "Cor_Texto_Botao", "Cor_Fundo"]
    }
  ];

  estrutura.forEach(aba => {
    let sheet = ss.getSheetByName(aba.nome);
    if (!sheet) {
      sheet = ss.insertSheet(aba.nome);
      sheet.appendRow(aba.colunas);
      sheet.getRange(1, 1, 1, aba.colunas.length).setFontWeight("bold").setBackground("#2c3e50").setFontColor("#ffffff");
      sheet.setFrozenRows(1);

      if (aba.nome === "Config_App") {
        // <-- NOVO: 7º item adicionado para respeitar a nova coluna
        sheet.appendRow(["", "", "#FFD700", "#1e1e1e", "#ffffff", "#000000", "#121212"]);
      }
    }
  });
}
/**
 * RELATÓRIO ADMIN/INSTRUTOR (COM FILTROS E PERMISSÕES)
 */
function getRelatorioFinanceiroAdmin(loginSolicitante, filtros = {}) {
  try {
    const transacoes = lerTabelaDinamica("Fin_Transacoes");
    const alunos = lerTabelaDinamica("cadastro_de_alunos"); // Para checar nível do solicitante

    // 1. Verifica Nível do Solicitante
    const solicitante = alunos.find(a => String(a.login).toLowerCase() === String(loginSolicitante).toLowerCase());
    const nivel = solicitante ? String(solicitante.graduacao_atual || "").toUpperCase() : "";
    const isSuperAdmin = LISTA_ADMINS_GERAL.some(role => nivel.includes(role));

    // Se não for Admin Geral, assume que é Instrutor e força filtro pela academia dele
    const academiaInstrutor = solicitante ? String(solicitante.academia_vinculada).trim() : "";

    let ent = 0, sai = 0;

    // 2. Filtragem
    const dadosFiltrados = transacoes.filter(t => {
      let valido = true;
      const tAcad = String(t.academia_ref || "").trim();
      const tData = parseDataSegura(t.data_registro);

      // REGRA 1: Hierarquia
      if (!isSuperAdmin) {
        // Instrutor só vê a SUA academia
        if (tAcad.toLowerCase() !== academiaInstrutor.toLowerCase()) return false;
      } else {
        // Admin vê o que quiser (Filtro de Tela)
        if (filtros.academia && filtros.academia !== "TODAS") {
          if (tAcad.toLowerCase() !== String(filtros.academia).toLowerCase()) return false;
        }
      }

      // REGRA 2: Filtro de Data (Se fornecido)
      if (filtros.dataInicio && tData) {
        const dIni = new Date(filtros.dataInicio);
        if (tData < dIni) return false;
      }
      if (filtros.dataFim && tData) {
        const dFim = new Date(filtros.dataFim);
        // Ajuste para pegar até o final do dia
        dFim.setHours(23, 59, 59);
        if (tData > dFim) return false;
      }

      return true;
    });

    // 3. Cálculos
    dadosFiltrados.forEach(t => {
      let vStr = String(t.valor || "0").replace("R$", "").trim();
      if (vStr.includes(",") && !vStr.includes(".")) vStr = vStr.replace(/\./g, "").replace(",", ".");
      const val = parseFloat(vStr);

      const tipo = String(t.tipo).toLowerCase();
      if (!isNaN(val)) {
        if (tipo.includes('receita')) ent += val;
        else if (tipo.includes('despesa')) sai += val;
      }
    });

    const resultado = {
      saldo: ent - sai,
      entradas: ent,
      saidas: sai,
      historico: dadosFiltrados.reverse(), // Retorna todos filtrados
      isSuperAdmin: isSuperAdmin // Informa ao frontend o nível
    };

    return JSON.stringify(resultado);

  } catch (e) {
    console.error("Erro Admin Fin:", e);
    return JSON.stringify({ saldo: 0, entradas: 0, saidas: 0, historico: [] });
  }
}

/**
 * 🛠️ HELPER DE FORMULÁRIO (Menus Dinâmicos)
 * Busca listas de Academias e Alunos para preencher os selects do Financeiro.
 */
function getDadosParaLancamento() {
  // 1. Busca Locais
  const locaisRaw = lerTabelaDinamica("Locais_de_treino");
  const listaLocais = locaisRaw.map(l => String(l.nome_do_local).trim()).filter(n => n !== "");

  // 2. Busca Alunos (Login e Nome)
  const alunosRaw = lerTabelaDinamica("cadastro_de_alunos");
  // Ordena por nome para facilitar a busca
  const listaAlunos = alunosRaw
    .filter(a => String(a.status).toLowerCase() === "ativo") // Só ativos
    .map(a => ({
      login: String(a.login).trim(),
      nome: String(a.nome_completo).trim()
    }))
    .sort((a, b) => a.nome.localeCompare(b.nome));

  return {
    locais: listaLocais,
    alunos: listaAlunos
  };
}

// ============================================================================
// 🔒 MÓDULO FINANCEIRO V11.0 (Hierarquia & Filtros)
// ============================================================================

const LISTA_ADMINS_GERAL = ["MESTRE", "ADMIN", "PRETA E BRANCA"]; // Níveis que veem TUDO

/**
 * RELATÓRIO DO ALUNO (SEM HISTÓRICO + CONTATO WHATSAPP)
 */
function getFinanceiroPessoal(login) {
  try {
    const assinaturas = lerTabelaDinamica("Fin_Assinaturas");
    const alunos = lerTabelaDinamica("cadastro_de_alunos");
    const locais = lerTabelaDinamica("Locais_de_treino");

    const loginBusca = String(login).toLowerCase().trim();

    // 1. Busca Dados da Assinatura
    const assinatura = assinaturas.find(a => String(a.login_aluno).toLowerCase().trim() === loginBusca);

    // 2. Busca Dados do Aluno para saber a Academia
    const alunoPerfil = alunos.find(a => String(a.login).toLowerCase().trim() === loginBusca);
    const academiaAluno = alunoPerfil ? alunoPerfil.academia_vinculada : "";

    // 3. Busca Contato do Responsável pela Academia
    let contatoResp = "5581997629232"; // Fallback (Número do Mestre/Geral)
    if (academiaAluno) {
      const localInfo = locais.find(l => String(l.nome_do_local).toLowerCase() === String(academiaAluno).toLowerCase());
      if (localInfo && localInfo.contato) {
        // Limpa o telefone para formato WhatsApp (apenas números)
        contatoResp = String(localInfo.contato).replace(/\D/g, "");
      }
    }

    const resultado = {
      plano: "Sem Plano",
      status: "Inativo",
      vencimento: null,
      diasRestantes: -999,
      contatoWhatsApp: contatoResp
    };

    if (assinatura) {
      resultado.plano = assinatura.pacote_atual;
      resultado.status = assinatura.status_assinatura;

      const dtFim = parseDataSegura(assinatura.data_fim);
      if (dtFim) {
        resultado.vencimento = formatDate(dtFim);
        // Calcula dias para vencer
        const hoje = new Date();
        hoje.setHours(0, 0, 0, 0);
        const diffTime = dtFim - hoje;
        resultado.diasRestantes = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      }
    }

    // OBS: Removemos o 'historico' daqui conforme solicitado
    return JSON.stringify(resultado);

  } catch (e) {
    console.error("Erro Aluno Fin:", e);
    return JSON.stringify({ erro: e.message });
  }
}

/**
 * 4. LISTA DE PACOTES DISPONÍVEIS (Para venda)
 */
function getPacotesDisponiveis() {
  const pacotes = lerTabelaDinamica("Fin_Pacotes");
  return pacotes.filter(p => String(p.status_pacote).toLowerCase() === "ativo");
}

/**
 * Registra Entradas (Aulas, Vendas) ou Saídas (Despesas).
 * Se for Mensalidade, atualiza automaticamente a assinatura do aluno.
 */
function registrarMovimentacao(dados) {
  try {
    const novaTransacao = {
      "ID_Transacao": "TRX-" + Date.now(),
      "Data_Registro": new Date(),
      "Tipo": dados.tipo,
      "Categoria": dados.categoria,
      "Descricao": dados.descricao,
      "Valor": dados.valor,
      "Forma_Pagto": dados.forma,
      "Responsavel": dados.responsavel,
      "Login_Aluno": dados.alunoLogin || "",
      "Academia_Ref": dados.academia || "Matriz",
      "Status": "Concluido"
    };

    salvarFinanceiroSeguro("Fin_Transacoes", novaTransacao);

    // Se for Mensalidade, renova assinatura
    if (dados.tipo === "Receita" && (dados.categoria === "Mensalidade" || dados.categoria === "Pacote")) {
      processarRenovacaoAssinatura(dados.alunoLogin, dados.nomePacote);
    }

    return { success: true, msg: "Movimentação registrada com sucesso!" };
  } catch (e) {
    return { success: false, msg: "Erro financeiro: " + e.message };
  }
}

/**
 * Atualiza a tabela de Assinaturas baseada no pacote comprado.
 */
function processarRenovacaoAssinatura(login, nomePacote) {
  if (!login) return;

  // 🛡️ CORREÇÃO: Se nomePacote vier vazio, força um nome padrão
  const pacoteSalvar = nomePacote || "Mensalidade Padrão";

  const pacotes = lerTabelaDinamica("Fin_Pacotes");
  const pacoteRef = pacotes.find(p => String(p.nome_pacote).toLowerCase() === String(pacoteSalvar).toLowerCase());
  const dias = pacoteRef ? parseInt(pacoteRef.duracao_dias) : 30;

  const assinaturas = lerTabelaDinamica("Fin_Assinaturas");
  const assinaturaAtual = assinaturas.find(a => String(a.login_aluno).toLowerCase() == String(login).toLowerCase());

  const hoje = new Date();
  let novaDataInicio = hoje;
  let novaDataFim = new Date();

  if (assinaturaAtual && assinaturaAtual.data_fim) {
    const fimAtual = parseDataSegura(assinaturaAtual.data_fim);
    if (fimAtual && fimAtual > hoje) { // Ainda válido
      novaDataInicio = fimAtual;
      novaDataFim = new Date(fimAtual);
      novaDataFim.setDate(novaDataFim.getDate() + dias);
    } else { // Vencido
      novaDataFim.setDate(hoje.getDate() + dias);
    }
  } else {
    novaDataFim.setDate(hoje.getDate() + dias);
  }

  const dadosAssinatura = {
    "Login_Aluno": login,
    "Pacote_Atual": pacoteSalvar, // Usa o nome corrigido
    "Data_Inicio": formatDate(novaDataInicio),
    "Data_Fim": formatDate(novaDataFim),
    "Status_Assinatura": "Ativo"
  };

  if (assinaturaAtual) {
    salvarFinanceiroSeguro("Fin_Assinaturas", dadosAssinatura, assinaturaAtual._linha);
  } else {
    salvarFinanceiroSeguro("Fin_Assinaturas", dadosAssinatura);
  }
}

// ============================================================================
// 1. UTILITÁRIOS E HELPERS (Ferramentas Genéricas)
// ============================================================================

/**
 * 🛡️ FUNÇÃO MESTRA DE GRAVAÇÃO (Anti-Block Writing)
 * Função genérica para salvar ou editar dados em QUALQUER aba.
 * É "blindada" porque mapeia colunas pelo nome, evitando erros se colunas mudarem de lugar.
 *
 * @param {string} nomeAba - Nome da aba onde salvar.
 * @param {object} mapaDados - Objeto { "NomeColuna": "Valor" }.
 * @param {number|null} idLinha - Se fornecido, EDITA essa linha. Se null, CRIA nova linha.
 */
function salvarDadosSeguro(nomeAba, mapaDados, idLinha = null) {
  const sheet = getSheet(nomeAba);
  const headers = getHeaders(sheet); // Pega cabeçalhos da linha 1
  const mapaColunas = {};

  // Cria um mapa { "nome_coluna_lower": indice } para busca rápida
  headers.forEach((h, i) => mapaColunas[String(h).trim().toLowerCase()] = i);

  // MODO CRIAÇÃO (Nova Linha)
  if (!idLinha || isNaN(idLinha)) {
    const novaLinhaArray = new Array(headers.length).fill(""); // Array vazio do tamanho da planilha

    // Preenche o array apenas nas posições corretas
    for (const [colunaAlvo, valor] of Object.entries(mapaDados)) {
      const idx = mapaColunas[String(colunaAlvo).trim().toLowerCase()];
      if (idx !== undefined) novaLinhaArray[idx] = valor;
    }
    sheet.appendRow(novaLinhaArray); // Adiciona ao final
    return "✅ Registro criado com sucesso!";
  }
  // MODO EDIÇÃO (Atualizar Linha Existente)
  else {
    for (const [colunaAlvo, valor] of Object.entries(mapaDados)) {
      const idx = mapaColunas[String(colunaAlvo).trim().toLowerCase()];
      if (idx !== undefined) {
        // Atualiza célula específica: Linha X, Coluna Y
        sheet.getRange(idLinha, idx + 1).setValue(valor);
      }
    }
    return "✅ Registro atualizado com sucesso!";
  }
}

/**
 * Inclui conteúdo HTML de outro arquivo. Usado para modularizar o Frontend.
 * Ex: <?!= include('CSS'); ?> dentro do index.html
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Retorna a URL base do script (útil para o frontend saber para onde voltar)
function getScriptUrl() { return SCRIPT_URL; }

/**
 * Transforma links brutos do Google Drive em links diretos de imagem.
 * Resolve problemas de visualização de fotos de perfil, logos e fundos.
 */
function padronizarLinkDrive(input) {
  if (!input) return "";

  // Se já for um link blindado (já foi convertido), apenas retorna ele mesmo
  if (input.includes("googleusercontent.com")) return input;

  let id = "";
  // Tenta extrair o ID do link original do Drive (padrão share link ou open link)
  const matchId = input.match(/id=([a-zA-Z0-9_-]+)/) || input.match(/\/d\/([a-zA-Z0-9_-]+)/);

  if (matchId) {
    id = matchId[1];
  } else if (input.length > 20 && !input.includes("/")) {
    id = input; // Se o usuário colou APENAS o ID solto
  } else {
    return input; // Fallback: retorna o original se não conseguir extrair
  }

  // Retorna a URL blindada que sempre renderiza a imagem no HTML
  return "https://lh3.googleusercontent.com/d/" + id;
}

// Wrappers para compatibilidade com chamadas antigas
function getDriveImageUrl(id) { return padronizarLinkDrive(id); }
function parseDriveLink(url) { return padronizarLinkDrive(url); }

/**
 * Formata datas para o padrão brasileiro (DD/MM/YYYY)
 * Essencial para exibir datas corretamente no frontend vindo do Sheets.
 */
function formatDate(date) {
  if (!date) return "";
  try {
    if (typeof date === 'string') return date; // Se já for texto, devolve
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy");
  } catch (e) { return date; }
}

/**
 * Obtém objeto Sheet pelo nome de forma segura. Lança erro se não existir.
 */
function getSheet(nomeAba) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeAba);
  if (!sheet) throw new Error(`Aba '${nomeAba}' não foi encontrada.`);
  return sheet;
}

/**
 * Lê a primeira linha da aba (cabeçalhos) e retorna como array de strings.
 */
function getHeaders(sheet) {
  if (sheet.getLastRow() < 1) return [];
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
}

// ============================================================================
// 2. LOGS E AUDITORIA (Segurança e Rastreabilidade)
// ============================================================================

/**
 * Logger interno para debug no console do Apps Script.
 * Usa ícones para facilitar leitura visual.
 */
function logMilitar(modulo, tipo, mensagem, dados = null) {
  const icones = { 'INFO': 'ℹ️', 'SUCESSO': '✅', 'ALERTA': '⚠️', 'ERRO': '❌', 'CRITICO': '🚨' };
  let logLine = `[${modulo}] ${icones[tipo] || '📝'} ${mensagem}`;
  if (dados) { try { logLine += ` | DADOS: ${JSON.stringify(dados)}`; } catch (e) { } }
  console.log(logLine);
}

/**
 * Grava ações importantes na aba 'Logs_Auditoria'.
 * Ex: "Instrutor Fulano alterou a nota do aluno Ciclano".
 */
function registrarLogAuditoria(usuario, acao, alvo, detalhes) {
  try {
    const mapa = { "Data/Hora": new Date(), "Usuario_Editor": usuario, "Acao": acao, "Alvo": alvo, "Detalhes": detalhes };
    salvarDadosSeguro(ABA_LOGS, mapa);
  } catch (e) { console.error("Falha auditoria: " + e.message); }
}

/**
 * Registra tentativas de login (Sucesso ou Falha).
 */
function registrarLogLogin(user, status) {
  try {
    const mapa = { "Data": new Date(), "User": user, "Status": status };
    salvarDadosSeguro("Logs_Login", mapa);
  } catch (e) { console.error("Erro log login: " + e.message); }
}

/**
 * Verifica se uma determinada faixa/graduação é considerada "Instrutor".
 * Usado para definir permissões de visualização no Dashboard.
 */
function checkInstrutor(faixa) {
  if (!faixa) return false;
  const niveis = ["VERDE ESCURO", "AZUL ESCURO", "MARROM", "PRETA E BRANCA", "PRETA"];
  return niveis.includes(String(faixa).toUpperCase().trim());
}

/**
 * Helper: Converte qualquer coisa (String ou Objeto) em Data JS válida
 */
function parseDataSegura(input) {
  if (!input) return null;
  if (input instanceof Date) return input; // Já é data

  // Tenta converter string YYYY-MM-DD ou DD/MM/YYYY
  const str = String(input).trim();

  if (str.includes('/')) {
    // Formato BR: 05/03/2026
    const partes = str.split('/');
    // new Date(ano, mês-1, dia)
    return new Date(partes[2], partes[1] - 1, partes[0]);
  }

  if (str.includes('-')) {
    // Formato ISO: 2026-03-05
    const partes = str.split('-');
    // Previne problemas de fuso horário definindo hora 12:00
    return new Date(partes[0], partes[1] - 1, partes[2], 12, 0, 0);
  }

  // Tenta parser nativo
  const d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

// ============================================================================
// 3. ROTEADOR (Gerenciamento de Navegação)
// ============================================================================

function doGet(e) {
  verificarCriarAbasFinanceiras();
  const page = e.parameter.page || 'login';
  let htmlFile;

  switch (page) {
    case 'agendar': htmlFile = 'Agendar'; break;
    case 'cursos': htmlFile = 'Cursos'; break;
    case 'dashboard': htmlFile = 'Dashboard'; break;
    case 'locais': htmlFile = 'Locais_treino'; break;
    case 'instrutor': htmlFile = 'Instrutor'; break;
    case 'recuperar': htmlFile = 'RecuperarSenha'; break;
    case 'login': htmlFile = 'Login'; break;
    case 'ajuda': htmlFile = 'Ajuda'; break;
    default: htmlFile = 'Login';
  }

  try {
    const template = HtmlService.createTemplateFromFile(htmlFile);

    // 🛡️ BLINDAGEM DE VARIÁVEIS (Valores Default Seguros)
    template.scriptUrl = SCRIPT_URL;
    template.bgUrl = ""; template.logoUrl = "";
    template.primaryColor = "#FFD700";
    template.secondaryColor = "#1e1e1e";
    template.textColor = "#ffffff";
    template.btnTextColor = "#000000";
    template.bgColor = "#121212"; // <-- NOVO: Blindagem da 5ª cor

    try {
      const config = getAppConfig();
      if (config) {
        template.bgUrl = getDriveImageUrl(config.bgId) || template.bgUrl;
        template.logoUrl = getDriveImageUrl(config.iconId) || template.logoUrl;
        template.primaryColor = config.primaryColor || template.primaryColor;
        template.secondaryColor = config.secondaryColor || template.secondaryColor;
        template.textColor = config.textColor || template.textColor;
        template.btnTextColor = config.btnTextColor || template.btnTextColor;
        template.bgColor = config.bgColor || template.bgColor; // <-- NOVO: Injeção da 5ª cor dinâmica
      }
    } catch (e) { console.warn("Erro no tema", e); }

    return template.evaluate()
      .setTitle("DojoManager - Portal do Aluno")
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (err) {
    return HtmlService.createHtmlOutput(`<h3>Erro Crítico</h3><p>${err.message}</p>`);
  }
}

// ============================================================================
// 4. API DE LEITURA (Dados para o App Frontend)
// ============================================================================

/**
 * Lê a aba de Locais de Treino para exibir no mapa/lista do aluno.
 */
function getLocaisTreino() {
  try {
    const sheet = getSheet("Locais_de_treino");
    const data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return []; // Sem dados

    // Mapeamento dinâmico de colunas
    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const findCol = (terms) => { for (let t of terms) { const i = headers.indexOf(t); if (i > -1) return i; } return -1; };

    const col = {
      nome: findCol(["nome do local", "nome"]),
      end: findCol(["endereço"]),
      cidade: findCol(["cidade/estado"]),
      contato: findCol(["contato"]),
      dias: findCol(["dias"]),
      horarios: findCol(["horários"]),
      maps: findCol(["link_google_maps"]),
      iframe: findCol(["html_mapa_off_lline"]),
      resp: findCol(["responsavel"]),
      status: findCol(["status"])
    };

    if (col.nome === -1) return [];

    // Transforma linhas em objetos JSON
    return data.slice(1).map(row => ({
      nome: row[col.nome],
      endereco: row[col.end],
      cidade: row[col.cidade],
      contato: row[col.contato],
      dias: row[col.dias],
      horarios: row[col.horarios],
      linkMaps: row[col.maps],
      iframeHtml: row[col.iframe],
      responsavel: row[col.resp],
      status: row[col.status] || "Ativo"
    })).filter(l => l.nome); // Remove vazios
  } catch (e) { return []; }
}

/**
 * Lê a aba de Cursos para exibir eventos disponíveis.
 */
function getCursosList() {
  try {
    const sheet = getSheet(NOME_ABA_CURSOS);
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());

    // Mapeamento fixo (baseado em nomes exatos)
    const col = { nome: headers.indexOf("Nome do Curso"), data: headers.indexOf("Data"), desc: headers.indexOf("Descrição"), vagas: headers.indexOf("Número de Vagas"), img: headers.indexOf("Imagem"), status: headers.indexOf("Status"), link: headers.indexOf("Link da Inscrição") };

    return data.slice(1).map(row => ({
      nome: row[col.nome],
      data: formatDate(row[col.data]), // Formata data
      descricao: row[col.desc],
      vagas: row[col.vagas],
      imagem: padronizarLinkDrive(row[col.img]),
      status: row[col.status],
      linkInscricao: row[col.link]
    })).filter(c => c.nome);
  } catch (e) { return []; }
}

/**
 * Wrapper para obter dados do usuário logado (usado no carregamento do Dashboard).
 */
function getDashboardData(loginUser) { return verificarCredenciais({ login: loginUser, checkOnly: true }); }

// Hierarquia para controle de acesso a vídeos
const HIERARQUIA_FAIXAS = ["Iniciante", "Branca", "Amarela", "Laranja", "Verde", "Azul", "Verde Escuro", "Azul Escuro", "Marrom", "Preta e Branca", "Preta"];

/**
 * 🛡️ Lista vídeos da Videoteca com lógica de PERMISSÃO DE FAIXA.
 * Um aluno só vê vídeos até a sua faixa atual + 1 nível.
 * Admins veem tudo.
 */
function listarVideoteca(user) {
  try {
    const sheet = getSheet(NOME_ABA_VIDEOTECA);
    const data = sheet.getDataRange().getDisplayValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());

    // Mapeamento de colunas
    const col = {
      faixa: headers.indexOf("faixa"),
      titulo: headers.indexOf("titulo"),
      link: headers.indexOf("youtube_link"),
      desc: headers.indexOf("descricao") !== -1 ? headers.indexOf("descricao") : headers.indexOf("descrição")
    };

    // Compatibilidade com nomes antigos
    if (col.link === -1) col.link = headers.indexOf("link");
    if (col.faixa === -1 || col.titulo === -1) throw new Error("Colunas obrigatórias 'Faixa' ou 'Titulo' não encontradas na aba Videoteca.");

    // Define índice da graduação do usuário
    let grad = String(user.graduacao || "Iniciante").trim();
    let index = HIERARQUIA_FAIXAS.findIndex(f => f.toLowerCase() === grad.toLowerCase());
    if (index === -1) index = 0;

    // Regra de "Deus" (Admin/Mestre vê tudo)
    const isGodMode = String(user.nivel).toLowerCase().includes("admin") || index === 10;

    // Define quais faixas o usuário pode ver
    let faixasPermitidas = isGodMode ? HIERARQUIA_FAIXAS.slice(1) : HIERARQUIA_FAIXAS.slice(1, index + 2); // Slice 1 ignora "Iniciante"

    const acervo = {};

    // Itera vídeos
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const fx = String(row[col.faixa]).trim();
      const tit = String(row[col.titulo]).trim();

      let descricao = "";
      if (col.desc > -1) { descricao = String(row[col.desc]).trim(); }
      if (!descricao) descricao = "Sem descrição disponível.";

      if (!fx || !tit) continue;

      const fxNorm = fx.charAt(0).toUpperCase() + fx.slice(1).toLowerCase(); // Normaliza texto (Primeira Maiúscula)

      // Verifica se a faixa do vídeo está na lista permitida
      if (faixasPermitidas.some(f => f.toLowerCase() === fxNorm.toLowerCase())) {
        const key = faixasPermitidas.find(f => f.toLowerCase() === fxNorm.toLowerCase()) || fx;

        if (!acervo[key]) acervo[key] = [];

        // Extrai ID do Youtube
        let link = row[col.link] || "";
        let vidId = "";
        if (link.length > 5) {
          try {
            if (link.includes("v=")) vidId = link.split("v=")[1].split("&")[0];
            else if (link.includes("youtu.be/")) vidId = link.split("youtu.be/")[1].split("?")[0];
            else if (link.length === 11) vidId = link;
          } catch (e) { }
        }

        acervo[key].push({ titulo: tit, youtubeId: vidId, faixa: key, descricao: descricao });
      }
    }

    return { faixasPermitidas: faixasPermitidas, acervo: acervo };

  } catch (e) {
    console.error("Erro Crítico Videoteca:", e);
    throw new Error("Erro Videoteca: " + e.message);
  }
}

/**
 * Lista PDFs do Programa Técnico (similar à videoteca, com permissões).
 */
function listarProgramasTecnicos(user) {
  try {
    const sheet = getSheet(NOME_ABA_PROGRAMAS);
    const data = sheet.getDataRange().getDisplayValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const col = { faixa: headers.indexOf("faixa"), id: headers.indexOf("id_arquivo"), desc: headers.indexOf("descricao") };

    // Lógica de Permissão (repetida da videoteca)
    let grad = String(user.graduacao || "Iniciante").trim();
    let index = HIERARQUIA_FAIXAS.findIndex(f => f.toLowerCase() === grad.toLowerCase());
    if (index === -1) index = 0;

    const isGodMode = String(user.nivel).toLowerCase().includes("admin") || index === 10;
    let faixasPermitidas = isGodMode ? HIERARQUIA_FAIXAS : HIERARQUIA_FAIXAS.slice(0, index + 2);

    const lista = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const fx = String(row[col.faixa]).trim();
      const id = row[col.id] || "";

      if (fx && id && faixasPermitidas.some(f => f.toLowerCase() === fx.toLowerCase())) {
        // Constrói link de download direto
        let url = id.startsWith("http") ? id : "https://drive.google.com/uc?export=download&id=" + id;
        lista.push({ faixa: fx, url: url, descricao: row[col.desc] || "", disponivel: true });
      }
    }
    return lista.reverse();
  } catch (e) { throw new Error("Erro PDF: " + e.message); }
}

/**
 * Lista livros da Biblioteca Digital (acesso livre).
 */
function listarBiblioteca() {
  try {
    const sheet = getSheet(NOME_ABA_BIBLIOTECA);
    const data = sheet.getDataRange().getValues();
    const headers = getHeaders(sheet);
    const col = { tit: headers.indexOf("Titulo"), desc: headers.indexOf("Descricao"), link: headers.indexOf("Link_Arquivo"), capa: headers.indexOf("Link_Capa") };

    return data.slice(1).map(row => {
      if (!row[col.tit] || !row[col.link]) return null;

      let l = String(row[col.link]);
      if (l.includes("drive") && !l.includes("export")) {
        const m = l.match(/id=([a-zA-Z0-9_-]+)/);
        if (m) l = "https://drive.google.com/uc?export=download&id=" + m[1];
      }
      return { titulo: row[col.tit], descricao: row[col.desc], link: l, capa: row[col.capa] };
    }).filter(b => b);
  } catch (e) { return []; }
}

// ============================================================================
// 5. API DE ESCRITA (ADMIN CRUD)
// ============================================================================

/**
 * [ADMIN] Retorna lista completa de alunos para o painel de gestão.
 */
function listarAlunosAdmin() {
  try {
    const sheet = getSheet("cadastro_de_alunos");
    const data = sheet.getDataRange().getDisplayValues();
    const h = data[0].map(c => String(c).trim()); // Headers

    // Mapeamento exaustivo de colunas para edição
    const col = {
      nome: h.indexOf("Nome Completo"), nasc: h.indexOf("Data de Nascimento"), tel: h.indexOf("Telefone"),
      cpf: h.indexOf("CPF"), pai: h.indexOf("Nome do Pai"), mae: h.indexOf("Nome da Mãe"),
      end: h.indexOf("Endereço"), acad: h.indexOf("Academia Vinculada"), email: h.indexOf("E-mail"),
      login: h.indexOf("LOGIN"), senha: h.indexOf("Senha"), grad: h.indexOf("GRADUACAO_ATUAL"),
      foto: h.indexOf("Foto 3x4 (para a carteirinha)"), ultCarteira: h.indexOf("Data Ultima Carteirinha"),
      status: h.indexOf("STATUS"), proxGrad: h.indexOf("PROX_GRADUACAO"), nivel: h.indexOf("Nível do Praticante"),
      exame: h.indexOf("Data Próximo Exame")
    };

    const safeGet = (row, idx) => (idx > -1 && row[idx]) ? row[idx] : "";

    return data.slice(1).map((row, i) => ({
      id: i + 2, // Número da linha (1-based) para edição futura
      nome: safeGet(row, col.nome), nasc: safeGet(row, col.nasc), tel: safeGet(row, col.tel),
      cpf: safeGet(row, col.cpf), pai: safeGet(row, col.pai), mae: safeGet(row, col.mae),
      endereco: safeGet(row, col.end), academia: safeGet(row, col.acad), email: safeGet(row, col.email),
      login: safeGet(row, col.login), senha: safeGet(row, col.senha), graduacao: safeGet(row, col.grad),
      foto: padronizarLinkDrive(safeGet(row, col.foto)), dataCarteira: safeGet(row, col.ultCarteira),
      status: safeGet(row, col.status) || "Ativo", proxGrad: safeGet(row, col.proxGrad),
      nivel: safeGet(row, col.nivel), proxExame: safeGet(row, col.exame)
    }));
  } catch (e) {
    console.error(e);
    return [];
  }
}

function buscarAlunoPorLogin(login) {
  const res = verificarCredenciais({ login: login, checkOnly: true });
  return res.success ? res.user : null;
}

/**
 * [ADMIN] Cria ou Atualiza um Aluno.
 */
function salvarAluno(form) {
  try {
    const dados = {
      "Nome Completo": form.aluno_nome, "Data de Nascimento": form.aluno_nasc, "Telefone": form.aluno_tel,
      "CPF": form.aluno_cpf, "Nome do Pai": form.aluno_pai, "Nome da Mãe": form.aluno_mae,
      "Endereço": form.aluno_end, "Academia Vinculada": form.aluno_acad, "E-mail": form.aluno_email,
      "LOGIN": form.aluno_login, "Senha": form.aluno_senha, "GRADUACAO_ATUAL": form.aluno_grad,
      "Foto 3x4 (para a carteirinha)": padronizarLinkDrive(form.aluno_foto), "Data Ultima Carteirinha": form.aluno_data_cart,
      "STATUS": form.aluno_status, "PROX_GRADUACAO": form.aluno_prox_grad, "Nível do Praticante": form.aluno_nivel,
      "Data Próximo Exame": form.aluno_exame
    };

    const idLinha = parseInt(form.aluno_id);

    // Se tem ID válido, edita. Se não, cria.
    if (!isNaN(idLinha) && idLinha > 1) {
      salvarDadosSeguro("cadastro_de_alunos", dados, idLinha);
      return "✅ Aluno atualizado com sucesso!";
    } else {
      salvarDadosSeguro("cadastro_de_alunos", dados);
      return "✅ Novo aluno cadastrado!";
    }
  } catch (e) {
    return "❌ Erro ao salvar: " + e.message;
  }
}

// --- GESTÃO DE LOCAIS ---
function listarLocaisAdmin() {
  const locais = getLocaisTreino();
  return locais.map((l, i) => ({ ...l, id: i + 2 }));
}

function salvarLocal(form) {
  try {
    const dados = { "Nome do Local": form.loc_nome, "Responsavel": form.loc_resp, "Status": form.loc_status, "Cidade/Estado": form.loc_cidade, "Endereço": form.loc_end, "Contato": form.loc_contato, "Dias": form.loc_dias, "Horários": form.loc_horas, "Link_Google_Maps": form.loc_maps, "html_mapa_off_lline": form.loc_iframe };
    const idLinha = parseInt(form.loc_id);
    salvarDadosSeguro(NOME_ABA_LOCAIS, dados, isNaN(idLinha) ? null : idLinha);
    registrarLogAuditoria(form.editor_login || "Admin", isNaN(idLinha) ? "CRIAR_LOCAL" : "EDITAR_LOCAL", form.loc_nome, "");
    return isNaN(idLinha) ? "✅ Local criado!" : "✅ Local atualizado!";
  } catch (e) { return "❌ Erro: " + e.message; }
}

// --- GESTÃO DE CURSOS ---
function listarCursosAdmin() {
  try {
    const sheet = getSheet(NOME_ABA_CURSOS);
    const data = sheet.getDataRange().getDisplayValues();
    const headers = getHeaders(sheet);
    const col = { nome: headers.indexOf("Nome do Curso"), data: headers.indexOf("Data"), desc: headers.indexOf("Descrição"), vagas: headers.indexOf("Número de Vagas"), img: headers.indexOf("Imagem"), status: headers.indexOf("Status"), link: headers.indexOf("Link da Inscrição") };
    if (col.nome === -1) return [];
    return data.slice(1).map((row, index) => ({ id: index + 2, nome: row[col.nome], data: row[col.data], descricao: row[col.desc], vagas: row[col.vagas], imagem: padronizarLinkDrive(row[col.img]), status: row[col.status] || "Abertas", link: row[col.link] })).filter(c => c.nome);
  } catch (e) { return []; }
}

function salvarCurso(form) {
  try {
    let dataF = form.curso_data; if (dataF.includes("-")) { const p = dataF.split("-"); dataF = `${p[2]}/${p[1]}/${p[0]}`; }
    const dados = { "Nome do Curso": form.curso_nome, "Data": dataF, "Descrição": form.curso_desc, "Número de Vagas": form.curso_vagas, "Imagem": padronizarLinkDrive(form.curso_img), "Status": form.curso_status, "Link da Inscrição": form.curso_link };
    const idLinha = parseInt(form.curso_id);
    salvarDadosSeguro(NOME_ABA_CURSOS, dados, isNaN(idLinha) ? null : idLinha);
    return isNaN(idLinha) ? "✅ Curso criado!" : "✅ Curso atualizado!";
  } catch (e) { return "❌ Erro: " + e.message; }
}

// --- GESTÃO DE VÍDEOS ---
function listarVideosAdmin() {
  try {
    const sheet = getSheet(NOME_ABA_VIDEOTECA);
    const data = sheet.getDataRange().getValues();
    const headers = getHeaders(sheet);
    const col = { faixa: headers.indexOf("Faixa"), tit: headers.indexOf("Titulo"), link: headers.indexOf("Youtube_Link"), stat: headers.indexOf("Status"), desc: headers.indexOf("Descricao") };
    if (col.tit === -1) return [];
    return data.slice(1).map((row, i) => ({ id: i + 2, faixa: row[col.faixa], titulo: row[col.tit], link: row[col.link], status: row[col.stat] || "Ativo", desc: row[col.desc] })).filter(v => v.titulo);
  } catch (e) { return []; }
}

function salvarVideo(form) {
  try {
    let link = form.vid_link; if (link.includes("v=")) link = link.split("v=")[1].split("&")[0];
    const dados = { "Faixa": form.vid_faixa, "Titulo": form.vid_titulo, "Descricao": form.vid_desc, "Youtube_Link": link, "Status": form.vid_status };
    const idLinha = parseInt(form.vid_id);
    salvarDadosSeguro(NOME_ABA_VIDEOTECA, dados, isNaN(idLinha) ? null : idLinha);
    return isNaN(idLinha) ? "✅ Vídeo criado!" : "✅ Vídeo atualizado!";
  } catch (e) { return "❌ Erro: " + e.message; }
}

// --- GESTÃO DE BIBLIOTECA ---
function listarBibliotecaAdmin() {
  try {
    const sheet = getSheet(NOME_ABA_BIBLIOTECA);
    const data = sheet.getDataRange().getValues();
    const headers = getHeaders(sheet);
    const col = { tit: headers.indexOf("Titulo"), desc: headers.indexOf("Descricao"), link: headers.indexOf("Link_Arquivo"), capa: headers.indexOf("Link_Capa") };
    if (col.tit === -1) return [];
    return data.slice(1).map((row, i) => ({ id: i + 2, titulo: row[col.tit], descricao: row[col.desc], link: row[col.link], capa: row[col.capa] })).filter(b => b.titulo);
  } catch (e) { return []; }
}

function salvarLivroBiblioteca(form) {
  try {
    const dados = { "Titulo": form.lib_titulo, "Descricao": form.lib_desc, "Link_Arquivo": padronizarLinkDrive(form.lib_link), "Link_Capa": padronizarLinkDrive(form.lib_capa) };
    const idLinha = parseInt(form.lib_id);
    salvarDadosSeguro(NOME_ABA_BIBLIOTECA, dados, isNaN(idLinha) ? null : idLinha);
    return isNaN(idLinha) ? "✅ Livro adicionado!" : "✅ Livro atualizado!";
  } catch (e) { return "❌ Erro: " + e.message; }
}

function excluirLivro(id) {
  try {
    const sheet = getSheet(NOME_ABA_BIBLIOTECA);
    sheet.deleteRow(parseInt(id));
    registrarLogAuditoria("Admin", "EXCLUIR_LIVRO", `ID ${id}`, "Removido da biblioteca");
    return "✅ Livro excluído!";
  } catch (e) { return "❌ Erro: " + e.message; }
}

// --- GESTÃO DE CERTIFICADOS ---
function listarCertificadosAdmin() {
  try {
    const sheet = getSheet("Certificados");
    const data = sheet.getDataRange().getDisplayValues(); // Pega valores visuais
    const h = getHeaders(sheet).map(x => x.toLowerCase());

    const colCpf = h.indexOf("cpf");
    const colCurso = h.indexOf("curso");
    const colData = h.indexOf("data_emissao");
    const colLink = h.indexOf("link_pdf");

    if (colCpf === -1) return [];

    return data.slice(1).map((row, i) => ({
      id: i + 2,
      cpf: row[colCpf],
      curso: row[colCurso],
      data: row[colData],
      link: row[colLink]
    })).filter(c => c.cpf);
  } catch (e) { return []; }
}

function salvarCertificado(form) {
  try {
    const dados = {
      "CPF": form.cert_cpf,
      "Curso": form.cert_curso,
      "Data_Emissao": form.cert_data, // yyyy-mm-dd
      "Link_PDF": padronizarLinkDrive(form.cert_link)
    };

    const idLinha = parseInt(form.cert_id);

    // Ajuste de data
    if (dados.Data_Emissao.includes('-')) {
      const p = dados.Data_Emissao.split('-');
      dados.Data_Emissao = `${p[2]}/${p[1]}/${p[0]}`; // Salva dd/mm/yyyy
    }

    if (isNaN(idLinha)) {
      salvarDadosSeguro("Certificados", dados);
      return "✅ Certificado criado!";
    } else {
      salvarDadosSeguro("Certificados", dados, idLinha);
      return "✅ Certificado atualizado!";
    }
  } catch (e) { return "Erro: " + e.message; }
}

function excluirCertificado(id) {
  try {
    const sheet = getSheet("Certificados");
    sheet.deleteRow(parseInt(id));
    return "✅ Certificado excluído!";
  } catch (e) { return "Erro: " + e.message; }
}

// --- GESTÃO DE PROGRAMAS TÉCNICOS ---
function listarProgramasAdmin() {
  try {
    const sheet = getSheet(NOME_ABA_PROGRAMAS);
    const data = sheet.getDataRange().getValues();
    const headers = getHeaders(sheet);
    const col = { faixa: headers.indexOf("Faixa"), id: headers.indexOf("ID_Arquivo"), link: headers.indexOf("Link_Original"), desc: headers.indexOf("Descricao") };
    if (col.faixa === -1) return [];
    return data.slice(1).map((row, i) => ({ id: i + 2, faixa: row[col.faixa], id_arquivo: row[col.id], link_original: row[col.link], descricao: row[col.desc] })).filter(p => p.faixa);
  } catch (e) { return []; }
}

function salvarProgramaTecnico(form) {
  try {
    let idArq = "";
    if (form.prog_link) { const m = form.prog_link.match(/id=([a-zA-Z0-9_-]+)/); if (m) idArq = m[1]; }
    const dados = { "Faixa": form.prog_faixa, "ID_Arquivo": idArq, "Link_Original": form.prog_link, "Descricao": form.prog_desc };
    const idLinha = parseInt(form.prog_id);
    salvarDadosSeguro(NOME_ABA_PROGRAMAS, dados, isNaN(idLinha) ? null : idLinha);
    return isNaN(idLinha) ? "✅ Programa criado!" : "✅ Programa atualizado!";
  } catch (e) { return "❌ Erro: " + e.message; }
}

function getAppConfig() {
  const configBase = {
    bgId: DEFAULT_ASSETS.BACKGROUND_ID,
    iconId: DEFAULT_ASSETS.ICON_ID,
    primaryColor: DEFAULT_ASSETS.PRIMARY_COLOR,
    secondaryColor: DEFAULT_ASSETS.SECONDARY_COLOR,
    textColor: DEFAULT_ASSETS.TEXT_COLOR,
    btnTextColor: DEFAULT_ASSETS.BTN_TEXT_COLOR,
    bgColor: DEFAULT_ASSETS.BG_COLOR // <-- NOVO
  };

  try {
    const configData = lerTabelaDinamica("Config_App");

    if (configData && configData.length > 0) {
      const configRow = configData[0];

      if (configRow.logo_url) configBase.iconId = padronizarLinkDrive(configRow.logo_url);
      if (configRow.fundo_url) configBase.bgId = padronizarLinkDrive(configRow.fundo_url);
      if (configRow.cor_primaria) configBase.primaryColor = configRow.cor_primaria;
      if (configRow.cor_secundaria) configBase.secondaryColor = configRow.cor_secundaria;
      if (configRow.cor_texto) configBase.textColor = configRow.cor_texto;
      if (configRow.cor_texto_botao) configBase.btnTextColor = configRow.cor_texto_botao;
      // <-- NOVO: Lendo a Cor_Fundo do banco de dados
      if (configRow.cor_fundo) configBase.bgColor = configRow.cor_fundo;
    }

    return configBase;
  } catch (e) {
    return configBase;
  }
}

function salvarConfig(form) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Config_App");

    if (!sheet) {
      verificarCriarAbasFinanceiras();
      sheet = ss.getSheetByName("Config_App");
    }

    const configsToSave = {
      "Logo_URL": padronizarLinkDrive(form.cfg_icon) || "",
      "Fundo_URL": padronizarLinkDrive(form.cfg_bg) || "",
      "Cor_Primaria": form.cfg_color || "#FFD700",
      "Cor_Secundaria": form.cfg_color_sec || "#1e1e1e",
      "Cor_Texto": form.cfg_color_text || "#ffffff",
      "Cor_Texto_Botao": form.cfg_color_btn_text || "#000000",
      "Cor_Fundo": form.cfg_color_bg || "#121212" // <-- NOVO: Salva a 5ª cor
    };

    if (sheet.getLastRow() < 2) {
      sheet.appendRow(["", "", "#FFD700", "#1e1e1e", "#ffffff", "#000000", "#121212"]);
    }

    salvarDadosSeguro("Config_App", configsToSave, 2);

    return "✅ Tema atualizado com sucesso!";
  } catch (e) {
    return "❌ Erro ao salvar configuração: " + e.message;
  }
}

function trocarSenhaUsuario(form) {
  try {
    const loginAlvo = String(form.cpf).trim().toLowerCase();
    const sheet = getSheet(NOME_ABA_ALUNOS);
    const colLogin = getHeaders(sheet).indexOf("LOGIN");
    const colSenha = getHeaders(sheet).indexOf("Senha");
    const data = sheet.getDataRange().getValues();

    let linha = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][colLogin]).trim().toLowerCase() === loginAlvo) {
        if (String(data[i][colSenha]).trim() !== String(form.senha_antiga).trim()) throw new Error("Senha antiga incorreta.");
        linha = i + 1;
        break;
      }
    }
    if (linha === -1) throw new Error("Usuário não encontrado.");
    salvarDadosSeguro(NOME_ABA_ALUNOS, { "Senha": form.senha_nova }, linha);
    registrarLogAuditoria(loginAlvo, "TROCA_SENHA", "Propria", "Sucesso");
    return "✅ Senha alterada!";
  } catch (e) { throw e; }
}

function atualizarFotoPerfil(form) {
  try {
    const loginAlvo = String(form.aluno_login).trim().toLowerCase();
    const sheet = getSheet(NOME_ABA_ALUNOS);
    const colLogin = getHeaders(sheet).indexOf("LOGIN");
    let linha = -1;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { if (String(data[i][colLogin]).trim().toLowerCase() === loginAlvo) { linha = i + 1; break; } }
    if (linha === -1) throw new Error("Usuário não encontrado.");
    salvarDadosSeguro(NOME_ABA_ALUNOS, { "Foto 3x4 (para a carteirinha)": padronizarLinkDrive(form.nova_foto) }, linha);
    registrarLogAuditoria(loginAlvo, "UPDATE_FOTO", "Propria", "Sucesso");
    return "✅ Foto atualizada!";
  } catch (e) { throw new Error(e.message); }
}

function salvarTicketSuporte(form) {
  try {
    const dados = { "Data/Hora": new Date(), "Login": "'" + form.sup_login, "Nome": form.sup_nome, "Tipo": form.sup_tipo, "Assunto": form.sup_assunto, "Mensagem": form.sup_msg, "Status": "Aberto" };
    salvarDadosSeguro("Suporte", dados);
    return "✅ Mensagem enviada!";
  } catch (e) { throw new Error(e.message); }
}

// ============================================================================
// 6. AUTENTICAÇÃO E CHECAGENS (Login)
// ============================================================================

function verificarCredenciais(formObject) {
  const loginInput = String(formObject.login).trim().toLowerCase();
  const isCheckOnly = formObject.checkOnly || false;
  try {
    const sheet = getSheet(NOME_ABA_ALUNOS);
    const data = sheet.getDataRange().getDisplayValues();
    const headers = data[0].map(h => String(h).trim());
    const col = { login: headers.indexOf("LOGIN"), senha: headers.indexOf("Senha"), nome: headers.indexOf("Nome Completo"), nivel: headers.indexOf("Nível do Praticante"), foto: headers.indexOf("Foto 3x4 (para a carteirinha)"), grad: headers.indexOf("GRADUACAO_ATUAL"), acad: headers.indexOf("Academia Vinculada"), exame: headers.indexOf("Data Próximo Exame"), status: headers.indexOf("STATUS") };

    if (col.login === -1) throw new Error("Coluna LOGIN não encontrada.");

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[col.login]).trim().toLowerCase() === loginInput) {
        const statusUser = (col.status > -1) ? String(row[col.status]).trim().toLowerCase() : "ativo";
        if (statusUser !== "ativo") return { success: false, message: "Cadastro inativo." };
        if (!isCheckOnly) { if (String(formObject.senha).trim() !== String(row[col.senha]).trim()) return { success: false, message: "Senha incorreta." }; }

        const userPayload = {
          LOGIN: loginInput,
          nomeCompleto: row[col.nome],
          graduacao: (col.grad > -1) ? row[col.grad] : "Iniciante",
          nivel: (col.nivel > -1) ? row[col.nivel] : "Aluno",
          fotoUrl: padronizarLinkDrive(row[col.foto]),
          academia: (col.acad > -1) ? row[col.acad] : "---",
          proximoExame: (col.exame > -1) ? row[col.exame] : "A definir",
          isInstrutor: checkInstrutor(row[col.grad])
        };

        if (!isCheckOnly) registrarLogLogin(loginInput, "SUCESSO");
        return { success: true, user: userPayload, redirectUrl: SCRIPT_URL + "?page=dashboard" };
      }
    }
    return { success: false, message: "Usuário não encontrado." };
  } catch (e) { return { success: false, message: e.message }; }
}

function getListaAcademias() { return getLocaisTreino().map(l => l.nome); }
function getAlunosPorAcademia(academiaNome) { return []; } // Stub

// ============================================================================
// 7. MÓDULO AVANÇADO: GESTÃO DE PRESENÇA E RANKING
// ============================================================================

/**
 * [RANKING] Calcula o ranking de presença (VERSÃO CORRIGIDA - CASE INSENSITIVE)
 */
function getRankingPresenca(mes, ano) {
  console.time("RankingTimer");
  console.log(`🛡️ [INIT RANKING] Iniciando cálculo para: Mês=${mes}, Ano=${ano}`);

  try {
    const sheetChamada = getSheet("Registro_Chamada");
    const sheetAlunos = getSheet("cadastro_de_alunos");

    // --- ETAPA 1: MAPEAMENTO DE COLUNAS (CHAMADA) ---
    const dadosChamada = sheetChamada.getDataRange().getDisplayValues();
    const headChamada = dadosChamada[0].map(h => String(h).trim().toLowerCase());

    const colData = headChamada.indexOf("data_treino");
    const colIds = headChamada.indexOf("lista_alunos_ids");

    if (colData === -1 || colIds === -1) {
      throw new Error("🚨 Colunas 'Data_Treino' ou 'Lista_Alunos_IDs' não encontradas na aba Registro_Chamada.");
    }

    // --- ETAPA 2: MAPEAMENTO DE ALUNOS (NORMALIZADO) ---
    const dadosAlunos = sheetAlunos.getDataRange().getDisplayValues();
    const headAlunos = dadosAlunos[0].map(h => String(h).trim());

    // Busca dinâmica de colunas
    const findCol = (headers, names) => {
      const lowerNames = names.map(n => n.toLowerCase());
      return headers.findIndex(h => lowerNames.includes(h.toLowerCase().trim()));
    };

    const colLogin = findCol(headAlunos, ["LOGIN", "Login", "ID"]);
    const colNome = findCol(headAlunos, ["Nome Completo", "Nome"]);
    const colAcad = findCol(headAlunos, ["Academia Vinculada", "Academia"]);
    const colFoto = findCol(headAlunos, ["Foto 3x4 (para a carteirinha)", "Foto"]);

    const mapAlunos = {};
    let alunosMapeadosCount = 0;

    for (let i = 1; i < dadosAlunos.length; i++) {
      const loginRaw = String(dadosAlunos[i][colLogin]).trim();

      // BLINDAGEM: Converte tudo para minúsculo para garantir o match
      const loginKey = loginRaw.toLowerCase();

      if (loginKey && loginKey !== "undefined") {
        mapAlunos[loginKey] = {
          nome: dadosAlunos[i][colNome] || "Sem Nome",
          academia: dadosAlunos[i][colAcad] || "---",
          foto: padronizarLinkDrive(dadosAlunos[i][colFoto]),
          loginOriginal: loginRaw // Guarda o original para exibição bonita se precisar
        };
        alunosMapeadosCount++;
      }
    }
    console.log(`✅ [ALUNOS] Total mapeados (Normalizados): ${alunosMapeadosCount}`);

    // --- ETAPA 3: PROCESSAMENTO DO RANKING ---
    const ranking = {};
    let linhasNoPeriodo = 0;

    for (let i = 1; i < dadosChamada.length; i++) {
      const row = dadosChamada[i];
      let dataTreino = String(row[colData]).trim();

      // Parser de Data Robusto
      let mesPlanilha = 0, anoPlanilha = 0;
      if (dataTreino.includes('-')) { // YYYY-MM-DD
        const partes = dataTreino.split("-");
        anoPlanilha = parseInt(partes[0]);
        mesPlanilha = parseInt(partes[1]);
      } else if (dataTreino.includes('/')) { // DD/MM/YYYY
        const partes = dataTreino.split("/");
        mesPlanilha = parseInt(partes[1]);
        anoPlanilha = parseInt(partes[2]);
      }

      // Filtra pelo Mês/Ano
      if (mesPlanilha === parseInt(mes) && anoPlanilha === parseInt(ano)) {
        linhasNoPeriodo++;
        const rawIds = String(row[colIds]);
        const ids = rawIds.split(",");

        ids.forEach(idRaw => {
          // BLINDAGEM: Normaliza o ID da chamada também
          const loginKey = idRaw.trim().toLowerCase();

          if (loginKey) {
            if (mapAlunos[loginKey]) {
              // Se encontrou no mapa (agora insensível a maiúsculas/minúsculas)
              if (!ranking[loginKey]) {
                ranking[loginKey] = { count: 0, ...mapAlunos[loginKey], login: mapAlunos[loginKey].loginOriginal };
              }
              ranking[loginKey].count++;
            } else {
              console.warn(`⚠️ [FANTASMA] ID "${loginKey}" (Raw: ${idRaw}) não encontrado no cadastro.`);
            }
          }
        });
      }
    }

    // --- ETAPA 4: RETORNO ---
    const resultadoFinal = Object.values(ranking).sort((a, b) => b.count - a.count);

    console.log(`📊 [RESULTADO] Linhas no Período: ${linhasNoPeriodo}`);
    console.log(`📊 [RESULTADO] Alunos Ranqueados: ${resultadoFinal.length}`);

    if (resultadoFinal.length > 0) {
      console.log(`🏆 [LÍDER] ${resultadoFinal[0].nome} (${resultadoFinal[0].count} aulas)`);
    } else {
      console.warn("⚠️ Nenhum aluno pontuou neste período.");
    }

    console.timeEnd("RankingTimer");
    return resultadoFinal;

  } catch (e) {
    console.error("🚨 [ERRO FATAL RANKING]: " + e.message);
    return [];
  }
}

// ============================================================================
// 8. MÓDULO: CHAMADA DINÂMICA (Smart Check-in v2.0)
// ============================================================================

// --- NOVO: FUNÇÕES PARA ADIÇÃO MANUAL DE ALUNOS ---

/**
 * Busca alunos por nome ou login para o instrutor selecionar
 */
function buscarAlunosParaManual(termo) {
  if (!termo || termo.length < 3) return [];
  const termoLimpo = termo.toLowerCase();

  // Reutiliza sua função de leitura otimizada
  const todos = lerTabelaDinamica("cadastro_de_alunos");

  return todos
    .filter(a => {
      const nome = String(a.nome_completo || "").toLowerCase();
      const login = String(a.login || "").toLowerCase();
      const status = String(a.status || "Ativo").toLowerCase();
      return status === "ativo" && (nome.includes(termoLimpo) || login.includes(termoLimpo));
    })
    .slice(0, 10) // Retorna no máximo 10 para não pesar
    .map(a => ({
      login: a.login,
      nome: a.nome_completo,
      graduacao: a.graduacao_atual,
      foto: padronizarLinkDrive(a['foto 3x4 (para a carteirinha)'])
    }));
}

/**
 * Adiciona um aluno diretamente à aula.
 * Se o ID começar com "VISITANTE_", cria um registro temporário.
 */
function adicionarAlunoManual(idAula, loginOuNome) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSheet(NOME_ABA_AULAS);
    const data = sheet.getDataRange().getValues();
    let linhaAlvo = -1;
    let checkinsAtuais = [];
    let dadosAula = {};

    // 1. Acha a aula
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === idAula) {
        linhaAlvo = i + 1;
        checkinsAtuais = JSON.parse(data[i][8] || "[]");
        dadosAula = { academia: data[i][4] };
        break;
      }
    }

    if (linhaAlvo === -1) return { success: false, msg: "Aula não encontrada." };

    // 2. Prepara o Objeto do Aluno
    let novoCheckin = null;

    // CASO A: É UM VISITANTE / NOME AVULSO (Não existe no banco)
    if (String(loginOuNome).startsWith("VISITANTE_")) {
      const nomeLimpo = loginOuNome.replace("VISITANTE_", "");

      // Verifica duplicidade visual
      if (checkinsAtuais.some(c => c.nome.toLowerCase() === nomeLimpo.toLowerCase())) {
        return { success: false, msg: "Nome já está na lista." };
      }

      novoCheckin = {
        login: "GUEST-" + Date.now(), // Gera ID único temporário
        nome: nomeLimpo,
        foto: "https://via.placeholder.com/150?text=VIS", // Foto genérica
        graduacao: "Visitante",
        academia_origem: "Externo",
        status: "Confirmado",
        tipo: "Visitante",
        hora: new Date().toLocaleTimeString()
      };
    }
    // CASO B: É UM ALUNO CADASTRADO (Busca no banco)
    else {
      if (checkinsAtuais.some(c => c.login === loginOuNome)) {
        return { success: false, msg: "Aluno já está na lista." };
      }

      const aluno = buscarAlunoPorLogin(loginOuNome);
      if (!aluno) return { success: false, msg: "Aluno não encontrado no banco." };

      const acadAula = String(dadosAula.academia).trim().toUpperCase();
      const acadAluno = String(aluno.academia).trim().toUpperCase();

      novoCheckin = {
        login: aluno.LOGIN || aluno.login,
        nome: aluno["Nome Completo"] || aluno.nomeCompleto,
        foto: padronizarLinkDrive(aluno["Foto 3x4 (para a carteirinha)"] || aluno.fotoUrl),
        graduacao: aluno["GRADUACAO_ATUAL"] || aluno.graduacao,
        academia_origem: aluno.academia || aluno["Academia Vinculada"],
        status: "Confirmado",
        tipo: acadAula !== acadAluno ? "Visitante" : "Regular",
        hora: new Date().toLocaleTimeString()
      };
    }

    // 3. Salva
    checkinsAtuais.push(novoCheckin);
    sheet.getRange(linhaAlvo, 9).setValue(JSON.stringify(checkinsAtuais));

    return { success: true, msg: "Adicionado com sucesso!" };

  } catch (e) {
    return { success: false, msg: "Erro: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

function formatarHoraSimples(dado) {
  if (!dado) return "--:--";
  if (dado instanceof Date) return Utilities.formatDate(dado, Session.getScriptTimeZone(), "HH:mm");
  return String(dado).substring(0, 5);
}

function iniciarAulaDinamica(dados) {
  limparAulasExpiradas();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(NOME_ABA_AULAS);
    if (!sheet) {
      sheet = ss.insertSheet(NOME_ABA_AULAS);
      sheet.appendRow(["ID_Aula", "Data_Aula", "Hora_Inicio", "Hora_Fim", "Academia", "Instrutor", "PIN", "Status", "Checkins_JSON"]);
    }
    const idAula = "AULA-" + Date.now();
    const pin = Math.floor(1000 + Math.random() * 9000);

    // Data sempre como string YYYY-MM-DD
    let dataStr = dados.data;
    if (dados.data instanceof Date) dataStr = Utilities.formatDate(dados.data, Session.getScriptTimeZone(), "yyyy-MM-dd");

    sheet.appendRow([idAula, dataStr, String(dados.inicio), String(dados.fim), dados.academia, dados.instrutor, pin, "ABERTA", "[]"]);
    return { success: true, idAula: idAula, pin: pin };
  } catch (e) { return { success: false, erro: e.message }; } finally { lock.releaseLock(); }
}

/**
 * [ALUNO] Busca aulas disponíveis
 * Se academiaAluno == "TODOS", retorna todas as aulas ativas do dia (para o Dashboard).
 */
function buscarAulasDisponiveisAluno(academiaAluno) {
  try {
    const sheet = getSheet(NOME_ABA_AULAS);
    const data = sheet.getDataRange().getValues();
    const agora = new Date();
    const aulas = [];
    const minhaAcademia = String(academiaAluno).trim().toUpperCase();
    const buscarTudo = (minhaAcademia === "TODOS"); // Flag nova
    const hojeStr = Utilities.formatDate(agora, Session.getScriptTimeZone(), "yyyy-MM-dd");

    for (let i = 1; i < data.length; i++) {
      const status = String(data[i][7]);
      const acadAula = String(data[i][4]).trim().toUpperCase();

      // Lógica atualizada: Aceita se for a academia do aluno OU se estiver buscando tudo
      if (status === "ABERTA" && (buscarTudo || acadAula === minhaAcademia)) {

        let diaPlanilha = data[i][1];
        let diaStr = "";
        if (diaPlanilha instanceof Date) {
          diaStr = Utilities.formatDate(diaPlanilha, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else {
          diaStr = String(diaPlanilha).split("T")[0];
        }

        if (diaStr === hojeStr) {
          aulas.push({
            id: data[i][0],
            instrutor: data[i][5],
            academia: data[i][4], // Importante retornar o nome para o Dashboard
            horario: `${formatarHoraSimples(data[i][2])} às ${formatarHoraSimples(data[i][3])}`,
            pin: data[i][6]
          });
        }
      }
    }
    return aulas;
  } catch (e) {
    console.error("Erro busca aluno: " + e.message);
    return [];
  }
}

function realizarCheckinAluno(login, pinDigitado) {
  console.log("incia a funçõa realizarCheckinAluno")
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSheet(NOME_ABA_AULAS);
    const data = sheet.getDataRange().getValues();
    let linhaAlvo = -1;
    let checkinsAtuais = [];
    let dadosAula = {};
    //validarAcessoFinanceiro();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][6]).trim() == String(pinDigitado).trim() && data[i][7] == "ABERTA") {
        linhaAlvo = i + 1;
        checkinsAtuais = JSON.parse(data[i][8] || "[]");
        dadosAula = { academia: data[i][4] };
        break;
      }
    }

    if (linhaAlvo === -1) return { success: false, msg: "PIN inválido ou aula encerrada." };
    if (checkinsAtuais.some(c => c.login === login)) return { success: false, msg: "Você já realizou check-in." };

    const aluno = buscarAlunoPorLogin(login);
    if (!aluno) return { success: false, msg: "Aluno não identificado." };

    const nomeAluno = aluno["Nome Completo"] || aluno.nomeCompleto || aluno.nome || "Aluno";
    const fotoAluno = aluno["Foto 3x4 (para a carteirinha)"] || aluno.fotoUrl || "";
    const gradAluno = aluno["GRADUACAO_ATUAL"] || aluno.graduacao || "Iniciante";

    const acadAula = String(dadosAula.academia).trim().toUpperCase();
    const acadAluno = String(aluno.academia).trim().toUpperCase();
    const tipoAluno = acadAula !== acadAluno ? "Visitante" : "Regular";

    checkinsAtuais.push({
      login: login, nome: nomeAluno, foto: padronizarLinkDrive(fotoAluno),
      graduacao: gradAluno, academia_origem: aluno.academia,
      status: "Pendente", tipo: tipoAluno, hora: new Date().toLocaleTimeString()
    });

    sheet.getRange(linhaAlvo, 9).setValue(JSON.stringify(checkinsAtuais));
    return { success: true, msg: "Check-in realizado! Aguarde o instrutor." };
  } catch (e) { return { success: false, msg: "Erro: " + e.message }; } finally { lock.releaseLock(); }
}

function buscarAulasAtivasInstrutor(loginInstrutor) {
  console.log("incia funçõa buscarAulasAtivasInstrutor")
  try {
    const sheet = getSheet(NOME_ABA_AULAS);
    const data = sheet.getDataRange().getValues();
    const aulas = [];
    const nomeInst = String(loginInstrutor).trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      if (data[i][7] === "ABERTA") {
        const donoAula = String(data[i][5]).trim().toLowerCase();
        // Permite se o nome bater ou se for admin
        if (donoAula.includes(nomeInst.split(' ')[0].toLowerCase()) || true) {
          let dataVisual = data[i][1];
          if (dataVisual instanceof Date) dataVisual = Utilities.formatDate(dataVisual, Session.getScriptTimeZone(), "dd/MM/yyyy");
          else if (typeof dataVisual === 'string' && dataVisual.includes('-')) {
            const p = dataVisual.split('-'); dataVisual = `${p[2]}/${p[1]}/${p[0]}`;
          }

          aulas.push({
            id: data[i][0], data: dataVisual, academia: data[i][4],
            horario: `${formatarHoraSimples(data[i][2])} - ${formatarHoraSimples(data[i][3])}`,
            pin: data[i][6], qtd: JSON.parse(data[i][8] || "[]").length
          });
        }
      }
    }
    return aulas;
  } catch (e) { return []; }
}

function buscarCheckinsAula(idAula) {
  console.log("incia funçõa buscarCheckinsAula")
  try {
    const sheet = getSheet(NOME_ABA_AULAS);
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === idAula) {
        return { success: true, status: data[i][7], checkins: JSON.parse(data[i][8] || "[]") };
      }
    }
    return { success: false };
  } catch (e) { return { success: false }; }
}

function finalizarAulaDinamica(idAula, listaFinalAlunos) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const sheetTemp = getSheet(NOME_ABA_AULAS);
    const dataTemp = sheetTemp.getDataRange().getValues();
    let linhaAlvo = -1;
    let dadosAula = null;

    for (let i = dataTemp.length - 1; i >= 1; i--) {
      if (String(dataTemp[i][0]) === idAula) {
        linhaAlvo = i + 1;
        let dataF = dataTemp[i][1];
        // Converte data para string YYYY-MM-DD
        if (dataF instanceof Date) dataF = Utilities.formatDate(dataF, Session.getScriptTimeZone(), "yyyy-MM-dd");
        else if (typeof dataF === 'string' && dataF.includes('/')) { const p = dataF.split('/'); dataF = `${p[2]}-${p[1]}-${p[0]}`; }

        dadosAula = { data: dataF, instrutor: dataTemp[i][5], local: dataTemp[i][4] };
        break;
      }
    }

    if (!dadosAula) return { success: false, msg: "Aula não encontrada." };

    const presentes = listaFinalAlunos.filter(a => a.status === "Confirmado");

    if (presentes.length > 0) {
      // Salva no histórico oficial para o Ranking ler
      salvarDadosSeguro("Registro_Chamada", {
        "ID_Chamada": idAula,
        "Data_Registro": new Date(),
        "Data_Treino": dadosAula.data, // Data limpa yyyy-mm-dd
        "Hora_Treino": new Date().toLocaleTimeString(),
        "Instrutor_Logado": dadosAula.instrutor,
        "Local_Treino": dadosAula.local,
        "Qtd_Presentes": presentes.length,
        "Lista_Alunos_IDs": presentes.map(a => a.login).join(", "),
        "Lista_Nomes": presentes.map(a => a.nome + (a.tipo === 'Visitante' ? ' [VIS]' : '')).join(", "),
        "Observacoes": "Chamada Dinâmica App"
      });
    }

    sheetTemp.getRange(linhaAlvo, 8).setValue("ENCERRADA");
    return { success: true, msg: "Aula finalizada e Ranking atualizado!" };

  } catch (e) { return { success: false, msg: "Erro: " + e.message }; } finally { lock.releaseLock(); }
}

function limparAulasExpiradas() {
  try {
    const sheet = getSheet(NOME_ABA_AULAS);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    const agora = new Date();
    for (let i = data.length - 1; i >= 1; i--) {
      let d = new Date(data[i][1]);
      if ((agora - d) > (24 * 3600 * 1000) && data[i][7] === "ABERTA") { sheet.deleteRow(i + 1); }
    }
  } catch (e) { }
}

/**
 * Verifica se o aluno pode treinar NESTA academia específica.
 * Lógica para SaaS Multi-Tenancy.
 */
function validarAcessoFinanceiro(login, academiaOndeEstaTentandoCheckin) {
  const assinaturas = lerTabelaDinamica("Fin_Assinaturas");
  // Normaliza o login para busca
  const assinaturaUser = assinaturas.find(a =>
    String(a.login_aluno).toLowerCase().trim() === String(login).toLowerCase().trim()
  );

  // 1. Verifica existência
  if (!assinaturaUser) return { ok: false, msg: "🚫 Nenhum plano ativo encontrado. Fale com a recepção." };

  // 2. Verifica Status Textual
  if (String(assinaturaUser.status_assinatura).toLowerCase() !== "ativo") {
    return { ok: false, msg: "🚫 Sua assinatura está Inativa/Cancelada." };
  }

  // 3. Verifica Data (Vigência)
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0); // Zera hora para comparar apenas dias

  const dataFim = parseDataSegura(assinaturaUser.data_fim);

  // Se a data for inválida, BLOQUEIA por segurança
  if (!dataFim || isNaN(dataFim.getTime())) {
    return { ok: false, msg: "🚫 Erro no cadastro da data. Contate o suporte." };
  }

  // Comparações de Data
  if (dataFim < hoje) {
    return { ok: false, msg: `⛔ Plano vencido em ${formatDate(dataFim)}.` };
  }

  // 4. Verifica Multi-Academia
  const pacotes = lerTabelaDinamica("Fin_Pacotes");
  const pacoteInfo = pacotes.find(p => p.nome_pacote === assinaturaUser.pacote_atual);

  if (pacoteInfo) {
    const permitidas = String(pacoteInfo.academias_permitidas).toUpperCase();
    const localAtual = String(academiaOndeEstaTentandoCheckin).toUpperCase();

    // Se o pacote diz "TODAS" ou contém o nome da academia atual
    if (permitidas.includes("TODAS") || permitidas.includes(localAtual)) {
      return { ok: true, msg: "Acesso Autorizado" };
    } else {
      return { ok: false, msg: `⛔ Seu plano (${pacoteInfo.nome_pacote}) não cobre esta unidade.` };
    }
  }

  // Se não achou regra de pacote, libera (Fallback) ou Bloqueia (depende da sua política). 
  // Vou deixar liberado para não travar se o pacote foi deletado, mas ideal seria revisar.
  return { ok: true, msg: "Acesso Básico Liberado" };
}

function getListaGraduacoes() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Graduacao") || ss.getSheetByName("Config_Graduacoes");
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      const lista = data.slice(1).map(r => String(r[0]).trim()).filter(g => g !== "");
      if (lista.length > 0) return lista;
    }
    return ["Iniciante", "Branca", "Amarela", "Laranja", "Verde", "Azul", "Marrom", "Preta"];
  } catch (e) {
    return ["Branca", "Amarela", "Laranja", "Verde", "Azul", "Marrom", "Preta"];
  }
}