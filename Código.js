

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
  BG_COLOR: "#121212",
  // <-- NOVOS LINKS PADRÕES DO SISTEMA (FALLBACK)
  LINK_LOJA: "https://wa.me/5581997629232",
  LINK_INSTA: "https://www.instagram.com/fbkmklnoficial/",
  LINK_YT: "https://www.youtube.com/@KMLNDEFENSE",
  LINK_CAD: "https://forms.gle/3vXp86VCmuLzHrk46"
};


// ============================================================================
// 🧠 HELPERS DE LEITURA POR NOME DE COLUNA
// ============================================================================

/**
 * Lê uma aba inteira e retorna um array de objetos JSON baseados no cabeçalho.
 * FIX: Inteligência para diferenciar Data (dd/MM/yyyy) de Hora (HH:mm) baseada na "Época" do Sheets (1899).
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

        // --- BLINDAGEM DE DATA E HORA ---
        if (valor instanceof Date) {
          try {
            // Se o ano for 1899, o Sheets está enviando APENAS uma Hora (ex: 20:30)
            if (valor.getFullYear() === 1899) {
              valor = Utilities.formatDate(valor, Session.getScriptTimeZone(), "HH:mm");
            } else {
              // Caso contrário, é uma Data normal
              valor = Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd/MM/yyyy");
            }
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

/**
 * 🏗️ MÓDULO DE AUTO-SETUP E SELF-HEALING (BANCO DE DADOS DINÂMICO)
 * Verifica todas as 16 abas do sistema e TODAS as suas respectivas colunas. 
 * Se alguém apagar uma coluna ou aba sem querer, o sistema recria automaticamente.
 */
function verificarCriarAbasSistema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 🗺️ O Dicionário da Verdade (Single Source of Truth)
  const estrutura = [
    { nome: "cadastro_de_alunos", colunas: ["Carimbo de data/hora", "Endereço de e-mail", "Nome Completo", "Data de Nascimento", "Peso", "Altura", "Telefone", "CPF", "Nome do Pai", "Nome da Mãe", "Endereço", "Turma Vinculada", "Academia Vinculada", "LOGIN", "Senha", "GRADUACAO_ATUAL", "Foto 3x4 (para a carteirinha)", "Data Ultima Carteirinha", "STATUS", "PROX_GRADUACAO", "Nível do Praticante", "NivelAdministrativo", "Modalidade", "Data Próximo Exame"] },
    { nome: "Locais_de_treino", colunas: ["Nome do Local", "Endereço", "Cidade/Estado", "Contato", "Link_Google_Maps", "html_mapa_off_lline", "Responsavel", "Status", "Pix_Chave_Local_de_Treino", "Pix_Nome_Local_de_Treino", "Banco_Local_de_Treino", "btn_pix_copia_e_cola_Local_de_Treino", "Ativo", "Pix_Cidade_Local_de_Treino"] },
    { nome: "Config_Turmas", colunas: ["ID_Turma", "Nome da Turma", "Modalidade", "Faixa Etária", "Local Vinculado", "Dias da Semana", "Horário Início", "Horário Fim", "Status", "Responsável", "Telefone"] },
    { nome: "Config_App", colunas: ["Logo_URL", "Fundo_URL", "Cor_Fundo", "Cor_Primaria", "Cor_Secundaria", "Cor_Texto", "Cor_Texto_Botao", "Link_Loja", "Link_Instagram", "Link_YouTube", "Link_Cadastro", "Nome_Academia", "Pix_Global_Ativo", "Pix_Chave_Global", "Pix_Nome", "Nome_Banco_PIX_Global", "Pix_Cidade"] },
    { nome: "Fin_Transacoes", colunas: ["ID_Transacao", "Data_Registro", "Tipo", "Categoria", "Descricao", "Valor", "Forma_Pagto", "Responsavel", "Login_Aluno", "Academia_Ref", "Status", "Comprovante_Url", "Modalidade"] },
    { nome: "Fin_Pacotes", colunas: ["Nome_Pacote", "Valor_Padrao", "Duracao_Dias", "Academias_Permitidas", "Status_Pacote", "Descricao"] },
    { nome: "Fin_Assinaturas", colunas: ["Login_Aluno", "Pacote_Atual", "Data_Inicio", "Data_Fim", "Status_Assinatura", "ID_Ultima_Transacao"] },
    { nome: "Aulas_Em_Andamento", colunas: ["ID_Aula", "Data_Aula", "Hora_Inicio", "Hora_Fim", "Academia", "Turma", "Instrutor", "PIN", "Status", "Checkins_JSON"] },
    { nome: "Registro_Chamada", colunas: ["ID_Chamada", "Data_Registro", "Data_Treino", "Hora_Treino", "Instrutor_Logado", "Local_Treino", "Qtd_Presentes", "Lista_Alunos_IDs", "Lista_Nomes", "Observacoes"] },
    { nome: "Certificados", colunas: ["CPF", "Curso", "Modalidade", "Data_Emissao", "Link_PDF"] },
    { nome: "Suporte", colunas: ["Data/Hora", "Login", "Nome", "Tipo", "Assunto", "Mensagem", "Status"] },
    { nome: "Config_Programas", colunas: ["Faixa", "Modalidade", "ID_Arquivo", "Link_Original", "Descricao"] },
    { nome: "Cursos", colunas: ["Nome do Curso", "Data", "Descrição", "Número de Vagas", "Imagem", "Status", "Link da Inscrição"] },
    { nome: "Config_Biblioteca", colunas: ["Titulo", "Descricao", "Link_Arquivo", "Link_Capa"] },
    { nome: "Config_Videoteca", colunas: ["Faixa", "Modalidade", "Titulo", "Youtube_Link", "Status", "Descricao"] },
    { nome: "GRADUACAO", colunas: ["Faixa / Nivel", "Observacao", "ID", "Modalidade", "Nivel"] }
  ];

  estrutura.forEach(aba => {
    let sheet = ss.getSheetByName(aba.nome);
    let isNewSheet = false;

    // 1. Se a aba não existir no Google Sheets, o sistema CRIA a aba e as colunas do zero
    if (!sheet) {
      sheet = ss.insertSheet(aba.nome);
      sheet.appendRow(aba.colunas);
      isNewSheet = true;

      // Injeta valores de segurança padrão no Config_App se a aba for recém-criada
      if (aba.nome === "Config_App") {
        sheet.appendRow(["", "", "#121212", "#FFD700", "#1e1e1e", "#ffffff", "#000000", "", "", "", "", "DojoManager SaaS", "Nao", "", "", "", ""]);
      }
    }

    // 2. 🛡️ MOTOR SELF-HEALING: Se a aba já existe, confere se alguém apagou alguma coluna
    const headersAtuais = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0].map(h => String(h).trim().toLowerCase());
    let precisaAtualizarHeader = false;

    aba.colunas.forEach(colunaEsperada => {
      // Se a coluna esperada não estiver na planilha...
      if (!headersAtuais.includes(String(colunaEsperada).trim().toLowerCase())) {
        // Encontrou uma coluna em falta! Vai adicioná-la na última posição disponível.
        const ultimaCol = sheet.getLastColumn();
        sheet.getRange(1, ultimaCol + 1).setValue(colunaEsperada);
        precisaAtualizarHeader = true;
      }
    });

    // 3. Aplica o UI Style "Banco de Dados" (Negrito, Fundo Escuro, Congelar Linha 1)
    if (isNewSheet || precisaAtualizarHeader) {
      sheet.getRange(1, 1, 1, sheet.getLastColumn())
        .setFontWeight("bold")
        .setBackground("#2c3e50")
        .setFontColor("#ffffff");
      sheet.setFrozenRows(1);
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
      "Status": "Concluido",
      "Modalidade": dados.modalidade || "Geral" // <-- INJETANDO A MODALIDADE NO CAIXA
    };

    salvarFinanceiroSeguro("Fin_Transacoes", novaTransacao);

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
  verificarCriarAbasSistema();
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

    template.scriptUrl = SCRIPT_URL;
    template.bgUrl = ""; template.logoUrl = "";
    template.primaryColor = "#FFD700";
    template.secondaryColor = "#1e1e1e";
    template.textColor = "#ffffff";
    template.btnTextColor = "#000000";
    template.bgColor = "#121212";

    template.linkLoja = "https://wa.me/5581997629232";
    template.linkInstagram = "https://www.instagram.com/";
    template.linkYouTube = "https://www.youtube.com/";
    template.linkCadastro = "https://forms.gle/3vXp86VCmuLzHrk46";

    let tituloNavegador = "DojoManager SaaS"; // Fallback do Titulo

    try {
      const config = getAppConfig();
      if (config) {
        template.bgUrl = getDriveImageUrl(config.bgId) || template.bgUrl;
        template.logoUrl = getDriveImageUrl(config.iconId) || template.logoUrl;
        template.primaryColor = config.primaryColor || template.primaryColor;
        template.secondaryColor = config.secondaryColor || template.secondaryColor;
        template.textColor = config.textColor || template.textColor;
        template.btnTextColor = config.btnTextColor || template.btnTextColor;
        template.bgColor = config.bgColor || template.bgColor;

        template.linkLoja = config.linkLoja || template.linkLoja;
        template.linkInstagram = config.linkInstagram || template.linkInstagram;
        template.linkYouTube = config.linkYouTube || template.linkYouTube;
        template.linkCadastro = config.linkCadastro || template.linkCadastro;

        tituloNavegador = config.nomeAcademia || tituloNavegador; // Puxa o nome
      }
    } catch (e) { console.warn("Erro ao carregar configurações", e); }

    return template.evaluate()
      // <-- INJETANDO O NOME DA ACADEMIA NA ABA DO NAVEGADOR
      .setTitle(tituloNavegador + " | Portal do Aluno")
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
 * 🛡️ Lista vídeos da Videoteca com Lógica de ID Numérico e MULTIMODALIDADE.
 */
function listarVideoteca(user) {
  try {
    const sheet = getSheet(NOME_ABA_VIDEOTECA);
    const data = sheet.getDataRange().getDisplayValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());

    const col = {
      faixa: headers.indexOf("faixa"), titulo: headers.indexOf("titulo"), link: headers.indexOf("youtube_link"),
      desc: headers.indexOf("descricao") !== -1 ? headers.indexOf("descricao") : headers.indexOf("descrição"),
      modalidade: headers.indexOf("modalidade")
    };

    if (col.link === -1) col.link = headers.indexOf("link");
    if (col.faixa === -1 || col.titulo === -1) throw new Error("Colunas obrigatórias 'Faixa' ou 'Titulo' não encontradas.");

    const mapaGrad = getMapaGraduacoes();

    const graduacaoAluno = user.graduacao || user.GRADUACAO_ATUAL || user["Graduação"] || "Iniciante";
    const modalidadeAluno = user.modalidade || user.Modalidade || "Geral";

    const userGradId = mapaGrad[String(graduacaoAluno).trim().toLowerCase()]?.id || 0;

    // Transforma "Krav Maga, Muay Thay" em Array: ['krav maga', 'muay thay']
    const userMods = String(modalidadeAluno).toLowerCase().split(',').map(m => m.trim());
    const isGodMode = user.isAdmin || user.isMestre;
    const acervo = {};
    const faixasPermitidasSet = new Set();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const fxRaw = String(row[col.faixa]).trim();
      const tit = String(row[col.titulo]).trim();
      const modRaw = col.modalidade > -1 ? String(row[col.modalidade]).trim() : "Geral";

      if (!fxRaw || !tit) continue;

      const fxLower = fxRaw.toLowerCase();
      const videoMods = modRaw.toLowerCase().split(',').map(m => m.trim()); // Array de mods do vídeo

      const videoGradInfo = mapaGrad[fxLower] || { id: 999 };

      // 🚨 A MÁGICA AQUI: O Mestre NÃO ignora mais a modalidade.
      // O vídeo tem que ser da modalidade dele ou 'Geral'
      const isSameMod = videoMods.includes("geral") || videoMods.includes("") || videoMods.some(vm => userMods.includes(vm));

      // O Mestre continua a ver TODAS as faixas (ignora o ID Numérico +1)
      const isAllowedLevel = isGodMode || videoGradInfo.id <= (userGradId + 1);

      if (isSameMod && isAllowedLevel) {
        const fxNorm = fxRaw.charAt(0).toUpperCase() + fxRaw.slice(1).toLowerCase();
        if (!acervo[fxNorm]) acervo[fxNorm] = [];
        faixasPermitidasSet.add(fxNorm);

        let link = row[col.link] || "";
        let vidId = "";
        if (link.length > 5) {
          try {
            if (link.includes("v=")) vidId = link.split("v=")[1].split("&")[0];
            else if (link.includes("youtu.be/")) vidId = link.split("youtu.be/")[1].split("?")[0];
            else if (link.length === 11) vidId = link;
          } catch (e) { }
        }

        acervo[fxNorm].push({ titulo: tit, youtubeId: vidId, faixa: fxNorm, descricao: row[col.desc] || "Sem descrição" });
      }
    }

    const faixasPermitidas = Array.from(faixasPermitidasSet).sort((a, b) => {
      const idA = mapaGrad[a.toLowerCase()] ? mapaGrad[a.toLowerCase()].id : 999;
      const idB = mapaGrad[b.toLowerCase()] ? mapaGrad[b.toLowerCase()].id : 999;
      return idA - idB;
    });

    return { faixasPermitidas: faixasPermitidas, acervo: acervo };
  } catch (e) { throw new Error("Erro Videoteca: " + e.message); }
}

/**
 * Lista PDFs do Programa Técnico (Com Lógica Numérica e MULTIMODALIDADE).
 */
function listarProgramasTecnicos(user) {
  try {
    const sheet = getSheet(NOME_ABA_PROGRAMAS);
    const data = sheet.getDataRange().getDisplayValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());

    const col = {
      faixa: headers.indexOf("faixa"), id: headers.indexOf("id_arquivo"),
      desc: headers.indexOf("descricao"), modalidade: headers.indexOf("modalidade")
    };

    const mapaGrad = getMapaGraduacoes();

    const graduacaoAluno = user.graduacao || user.GRADUACAO_ATUAL || user["Graduação"] || "Iniciante";
    const modalidadeAluno = user.modalidade || user.Modalidade || "Geral";

    const userGradId = mapaGrad[String(graduacaoAluno).trim().toLowerCase()]?.id || 0;
    const userMods = String(modalidadeAluno).toLowerCase().split(',').map(m => m.trim());

    const isGodMode = user.isAdmin || user.isMestre;

    const lista = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const fxRaw = String(row[col.faixa]).trim();
      const modRaw = col.modalidade > -1 ? String(row[col.modalidade]).trim() : "Geral";

      if (!fxRaw || !row[col.id]) continue;

      const pdfGradInfo = mapaGrad[fxRaw.toLowerCase()] || { id: 999 };
      const pdfMods = modRaw.toLowerCase().split(',').map(m => m.trim());

      // 🚨 Igual aos vídeos: Mestre respeita a modalidade dele.
      const isSameMod = pdfMods.includes("geral") || pdfMods.includes("") || pdfMods.some(vm => userMods.includes(vm));

      // Mestre vê todas as faixas da modalidade dele
      const isAllowedLevel = isGodMode || pdfGradInfo.id <= (userGradId + 1);

      if (isSameMod && isAllowedLevel) {
        let url = row[col.id].startsWith("http") ? row[col.id] : "https://drive.google.com/uc?export=download&id=" + row[col.id];
        lista.push({ faixa: fxRaw, url: url, descricao: row[col.desc] || "", disponivel: true });
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
 * 🥋 Busca todas as Modalidades únicas cadastradas na aba GRADUACAO (Coluna D)
 */
function getListaModalidades() {
  try {
    const sheet = getSheet("GRADUACAO"); // Tente também "Graduação" se falhar
    const data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return ["Geral"];

    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const colMod = headers.indexOf("modalidade");
    if (colMod === -1) return ["Geral"];

    const modsSet = new Set();
    for (let i = 1; i < data.length; i++) {
      const m = String(data[i][colMod]).trim();
      if (m && m.toLowerCase() !== "geral") modsSet.add(m); // Exclui vazios e "Geral"
    }
    return Array.from(modsSet);
  } catch (e) { return ["Geral"]; }
}

/**
 * ⏱️ Calcula o tempo desde a data de cadastro do aluno
 */
function calcularTempoCadastro(dataStr) {
  if (!dataStr) return "Recente";
  let d = parseDataSegura(dataStr);
  if (!d) return "Recente";

  let diffMs = new Date() - d;
  if (diffMs < 0) return "Recente";

  let diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));
  if (diffDays < 30) return `${diffDays} dia(s)`;
  let diffMonths = Math.floor(diffDays / 30);
  if (diffMonths < 12) return `${diffMonths} mês(es)`;
  let diffYears = Math.floor(diffMonths / 12);
  let remainMonths = diffMonths % 12;
  if (remainMonths === 0) return `${diffYears} ano(s)`;
  return `${diffYears} ano(s) e ${remainMonths} mês(es)`;
}

// ============================================================================
// 🧬 MÓDULO FISIOLÓGICO E CÁLCULO DE IDADE
// ============================================================================

/**
 * Retorna a idade exata em: "X anos, Y meses, Z dias"
 */
function calcularIdadeExata(dataNascStr) {
  if (!dataNascStr) return "N/A";
  let nasc = parseDataSegura(dataNascStr);
  if (!nasc) return "N/A";

  let hoje = new Date();
  let anos = hoje.getFullYear() - nasc.getFullYear();
  let meses = hoje.getMonth() - nasc.getMonth();
  let dias = hoje.getDate() - nasc.getDate();

  if (dias < 0) {
    meses--;
    let ultimoDiaMesAnterior = new Date(hoje.getFullYear(), hoje.getMonth(), 0).getDate();
    dias += ultimoDiaMesAnterior;
  }
  if (meses < 0) {
    anos--;
    meses += 12;
  }
  return `${anos}a, ${meses}m, ${dias}d`;
}

/**
 * Retorna apenas o número inteiro de anos (Para o filtro do Super Relatório)
 */
function calcularIdadeApenasAnos(dataNascStr) {
  if (!dataNascStr) return 0;
  let nasc = parseDataSegura(dataNascStr);
  if (!nasc) return 0;
  let hoje = new Date();
  let idade = hoje.getFullYear() - nasc.getFullYear();
  let m = hoje.getMonth() - nasc.getMonth();
  if (m < 0 || (m === 0 && hoje.getDate() < nasc.getDate())) idade--;
  return idade;
}

/**
 * 🩺 FIX: ANTI-APAGÃO - LISTAR ALUNOS COM DADOS PRESERVADOS, TRATAMENTO DE INDEFINIDOS E TURMA
 */
function listarAlunosAdmin() {
  try {
    const alunos = lerTabelaDinamica(NOME_ABA_ALUNOS);
    const chamadas = lerTabelaDinamica("Registro_Chamada");

    return alunos.map(a => {
      const dataNasc = a.data_de_nascimento_ || a.data_de_nascimento;
      const idadeCalculada = calcularIdadeExata(dataNasc);
      const loginBusca = String(a.login || "").toLowerCase().trim();

      const totalAulas = chamadas.filter(c => String(c.lista_alunos_ids || "").toLowerCase().includes(loginBusca)).length;

      return {
        id: a._linha,
        nome: a.nome_completo || a.nome || "Sem Nome",
        cpf: a.cpf || "",
        login: a.login || "---",
        senha: a.senha || "",
        tel: a.telefone || "",
        email: a.endereço_de_e_mail || a.e_mail || "",
        pai: a.nome_do_pai || "",
        mae: a.nome_da_mãe || "",
        endereco: a.endereço || "",
        foto: a["foto_3x4_(para_a_carteirinha)"] || "",
        academia: a.academia_vinculada || "",
        turma: a.turma_vinculada || "Sem Turma", // <-- 🚀 NOVA PROPRIEDADE DE TURMA
        graduacao: a.graduacao_atual || "Iniciante",
        proxGrad: a.prox_graduacao || "",
        modalidade: a.modalidade || "Geral",
        status: a.status || "Ativo",
        nivel: a["nível_do_praticante"] || a["nivel_do_praticante"] || "Aluno",
        peso: a.peso || "",
        altura: a.altura || "",
        proxExame: a.data_próximo_exame || "",
        dataCarteira: a.data_ultima_carteirinha || "",
        nasc: dataNasc || "",
        idadeExata: idadeCalculada,
        totalAulas: totalAulas,
        carimbo: a.carimbo_de_data_hora || a["carimbo_de_data/hora"] || "",
        statusAssinatura: a.status_assinatura || "Inativo"
      };
    });
  } catch (e) { console.error("ERRO LISTAR ALUNOS: " + e.message); return []; }
}

function buscarAlunoPorLogin(login) {
  const res = verificarCredenciais({ login: login, checkOnly: true });
  return res.success ? res.user : null;
}

/**
 * 🛠️ FIX: MAPEAMENTO DINÂMICO PARA EVITAR CHAVE 'UNDEFINED'
 */
function getDadosPagamentoAluno(login) {
  try {
    const configApp = lerTabelaDinamica("Config_App")[0];
    const alunos = lerTabelaDinamica(NOME_ABA_ALUNOS);
    const locais = lerTabelaDinamica(NOME_ABA_LOCAIS);
    const assinaturas = lerTabelaDinamica("Fin_Assinaturas");
    const pacotes = lerTabelaDinamica("Fin_Pacotes");

    const aluno = alunos.find(a => String(a.login).toLowerCase() === String(login).toLowerCase());
    const assinatura = assinaturas.find(as => String(as.login_aluno).toLowerCase() === String(login).toLowerCase());
    const pacote = pacotes.find(p => p.nome_pacote === assinatura?.pacote_atual);

    let info = {
      chave: configApp.pix_chave,
      beneficiario: configApp.pix_nome || "Dojo Manager",
      cidade: configApp.pix_cidade || "RECIFE",
      valor: pacote ? pacote.valor_padrao : 0,
      contatoWhatsApp: ""
    };

    // REGRA: Se não for Global, busca o PIX da Unidade
    if (configApp.pix_global_ativo !== "Sim") {
      const local = locais.find(l => l.nome_do_local === aluno.academia_vinculada);
      if (local && local.pix_chave) {
        info.chave = local.pix_chave;
        info.beneficiario = local.pix_nome || info.beneficiario;
        info.cidade = local.pix_cidade || info.cidade;
      }
    }

    // Busca o WhatsApp do local para o envio do comprovante
    const localResp = locais.find(l => l.nome_do_local === aluno.academia_vinculada);
    info.contatoWhatsApp = localResp?.contato ? String(localResp.contato).replace(/\D/g, "") : "5581997629232";

    return info;
  } catch (e) {
    console.error("Erro fatal no servidor ao buscar PIX: " + e.message);
    return null;
  }
}

/**
 * ============================================================================
 * 🛡️ FIX DEFINITIVO: ANTI-APAGÃO (HYDRATION V5 NÍVEL DEUS)
 * MODIFICADO POR: Arquitetura de Dados / QA
 * PORQUE MUDOU: O método antigo apagava dados se o formulário HTML não 
 * enviasse todos os campos (Ex: Painel reduzido do Instrutor) ou se houvesse 
 * divergência de acentos/espaços no cabeçalho.
 * O QUE FAZ AGORA: Lê a linha original, cruza as 26 colunas e "hidrata" 
 * os dados faltantes antes de salvar, além de atualizar as colunas duplicadas.
 * ============================================================================
 */
function salvarAluno(form) {
  try {
    const sheet = getSheet(NOME_ABA_ALUNOS);
    // 1. Puxa todos os cabeçalhos diretamente da linha 1 para saber a verdade absoluta
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // 2. Normaliza para garantir o "match" (ignora acentos, espaços extras, maiúsculas)
    const normalize = (str) => String(str).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, "");

    const headerMap = {};
    headers.forEach((h) => {
      if (h) headerMap[normalize(h)] = h; // Mapeia a string normalizada para o NOME REAL e exato da sua planilha
    });

    const idLinha = parseInt(form.aluno_id);
    let linhaExistente = [];

    // 3. Se for um aluno que já existe, pega a linha INTEIRA antes de fazer qualquer coisa
    if (!isNaN(idLinha) && idLinha > 1 && idLinha <= sheet.getLastRow()) {
      linhaExistente = sheet.getRange(idLinha, 1, 1, headers.length).getValues()[0];
    }

    const getOldVal = (nomeRealColuna) => {
      if (linhaExistente.length === 0) return "";
      const idx = headers.indexOf(nomeRealColuna);
      return idx !== -1 ? linhaExistente[idx] : "";
    };

    // 🚀 O MOTOR DE BLINDAGEM (O Coração do Sistema DCU)
    const resolveField = (formKey, colPlanilha1, colPlanilha2 = "") => {
      const realCol1 = headerMap[normalize(colPlanilha1)];
      const realCol2 = colPlanilha2 ? headerMap[normalize(colPlanilha2)] : null;

      let valorParaSalvar = "";
      let campoFoiEnviadoPelaTela = Object.prototype.hasOwnProperty.call(form, formKey);

      // Se o front-end (HTML) mandou o campo, nós usamos o valor que veio (mesmo que seja vazio, pois o Admin pode querer apagar)
      if (campoFoiEnviadoPelaTela) {
        valorParaSalvar = form[formKey];
      } else {
        // Se a tela NÃO enviou o campo, nós resgatamos a fotocópia da planilha para NÃO APAGAR!
        let old1 = realCol1 ? getOldVal(realCol1) : "";
        let old2 = realCol2 ? getOldVal(realCol2) : "";
        valorParaSalvar = old1 || old2 || "";
      }

      const results = [];
      if (realCol1) results.push({ col: realCol1, val: valorParaSalvar });
      // Se houver coluna duplicada (ex: E-mail e Endereço de e-mail), salva nas duas para garantir a sincronia!
      if (realCol2 && realCol2 !== realCol1) results.push({ col: realCol2, val: valorParaSalvar });

      return results;
    };

    // 🚀 LÓGICA DE NÍVEL DE ACESSO (Com Proteção)
    let nivelValido = "";
    if (Object.prototype.hasOwnProperty.call(form, 'aluno_nivel')) {
      const n = String(form.aluno_nivel).toUpperCase();
      if (n.includes("MESTRE")) nivelValido = "Mestre";
      else if (n.includes("PROFESSOR")) nivelValido = "Professor";
      else if (n.includes("N1")) nivelValido = "Instrutor N1";
      else if (n.includes("N2")) nivelValido = "Instrutor N2";
      else if (n.includes("N3")) nivelValido = "Instrutor N3";
      else if (n.includes("ADMIN")) nivelValido = "Admin";
      else nivelValido = "Aluno";
    } else {
      // Se o formulário não tem o campo nível, puxa o antigo
      nivelValido = getOldVal(headerMap[normalize("Nível do Praticante")]) || getOldVal(headerMap[normalize("NivelAdministrativo")]) || "Aluno";
    }

    // 🚀 MAPEAMENTO EXATO DAS SUAS 26 COLUNAS!
    const campos = [
      ...resolveField("aluno_carimbo", "Carimbo de data/hora"),
      ...resolveField("aluno_email", "Endereço de e-mail", "E-mail"),
      ...resolveField("aluno_nome", "Nome Completo"),
      ...resolveField("aluno_nasc", "Data de Nascimento"),
      ...resolveField("aluno_peso", "Peso"),
      ...resolveField("aluno_altura", "Altura"),
      ...resolveField("aluno_tel", "Telefone"),
      ...resolveField("aluno_cpf", "CPF"),
      ...resolveField("aluno_pai", "Nome do Pai"),
      ...resolveField("aluno_mae", "Nome da Mãe"),
      ...resolveField("aluno_end", "Endereço"),
      ...resolveField("aluno_turma", "Turma Vinculada"),
      ...resolveField("aluno_acad", "Academia Vinculada"),
      ...resolveField("aluno_login", "LOGIN"),
      ...resolveField("aluno_senha", "Senha"),
      ...resolveField("aluno_grad", "Graduação", "GRADUACAO_ATUAL"),
      ...resolveField("aluno_foto", "Foto 3x4 (para a carteirinha)"),
      ...resolveField("aluno_data_cart", "Data Ultima Carteirinha"),
      ...resolveField("aluno_status", "STATUS"),
      ...resolveField("aluno_prox_grad", "PROX_GRADUACAO"),
      ...resolveField("aluno_modalidade", "Modalidade"),
      ...resolveField("aluno_exame", "Data Próximo Exame")
    ];

    const dadosFinais = {};
    campos.forEach(item => {
      dadosFinais[item.col] = item.val;
    });

    // Injeta os níveis nas duas colunas administrativas (Retrocompatibilidade)
    const colNiv1 = headerMap[normalize("Nível do Praticante")];
    const colNiv2 = headerMap[normalize("NivelAdministrativo")];
    if (colNiv1) dadosFinais[colNiv1] = nivelValido;
    if (colNiv2) dadosFinais[colNiv2] = nivelValido;

    // Se for aluno novo (sem carimbo), cria a data de ingresso
    const colCarimbo = headerMap[normalize("Carimbo de data/hora")];
    if (colCarimbo && (!dadosFinais[colCarimbo] || dadosFinais[colCarimbo] === "")) {
      dadosFinais[colCarimbo] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    }

    // DISPARO SEGURO PRO BANCO DE DADOS
    salvarDadosSeguro(NOME_ABA_ALUNOS, dadosFinais, isNaN(idLinha) ? null : idLinha);

    // Sincronização de Pacote Financeiro (Protegida)
    if (Object.prototype.hasOwnProperty.call(form, 'aluno_pacote')) {
      let loginFinanceiro = dadosFinais[headerMap[normalize("LOGIN")]] || form.aluno_login;
      const dadosAssin = {
        "Login_Aluno": loginFinanceiro,
        "Pacote_Atual": form.aluno_pacote,
        "Data_Fim": form.aluno_vencimento,
        "Status_Assinatura": form.aluno_status_assinatura
      };
      const assinaturas = lerTabelaDinamica("Fin_Assinaturas");
      const aAtual = assinaturas.find(a => String(a.login_aluno).toLowerCase() === String(loginFinanceiro).toLowerCase());
      salvarDadosSeguro("Fin_Assinaturas", dadosAssin, aAtual ? aAtual._linha : null);
    }

    return "✅ Aluno salvo com Sucesso! (Blindagem contra apagão ativada)";
  } catch (e) {
    return "❌ Erro ao salvar: " + e.message;
  }
}

/**
 * ============================================================================
 * MOTOR DINÂMICO DE GRADUAÇÕES (O "PROCV" DO BACKEND)
 * ============================================================================
 */
function getMapaGraduacoes() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
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
      if (nomeFaixa) {
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
    return {};
  }
}

// ============================================================================
// 🌍 CRUD ADMIN DE LOCAIS (Removido Dias e Horas)
// ============================================================================

function listarLocaisAdmin() {
  try {
    const data = lerTabelaDinamica("Locais_de_treino");
    return data.map(row => ({
      id: row._linha,
      nome: row.nome_do_local || "",
      endereco: row['endereço'] || "",
      cidade: row['cidade/estado'] || "",
      contato: row.contato || "",
      linkMaps: row.link_google_maps || "",
      iframeHtml: row.html_mapa_off_lline || "",
      responsavel: row.responsavel || "",
      status: row.status || "Ativo",
      pixChave: row.pix_chave_local_de_treino || "",
      pixNome: row.pix_nome_local_de_treino || "",
      pixBanco: row.banco_local_de_treino || "",
      pixCidade: row.pix_cidade_local_de_treino || ""
    })).filter(l => l.nome !== "");
  } catch (e) { return []; }
}

function salvarLocal(form) {
  try {
    const dados = {
      "Nome do Local": form.loc_nome,
      "Responsavel": form.loc_resp,
      "Status": form.loc_status,
      "Cidade/Estado": form.loc_cidade,
      "Endereço": form.loc_end,
      "Contato": form.loc_contato,
      "Link_Google_Maps": form.loc_maps,
      "html_mapa_off_lline": form.loc_iframe,
      "Pix_Chave_Local_de_Treino": form.loc_pix_chave, // <-- Mapeado Novo Cabeçalho
      "Pix_Nome_Local_de_Treino": form.loc_pix_nome,   // <-- Mapeado Novo Cabeçalho
      "Banco_Local_de_Treino": form.loc_pix_banco,     // <-- Mapeado Novo Cabeçalho
      "Pix_Cidade_Local_de_Treino": form.loc_pix_cidade // <-- Mapeado Novo Cabeçalho
    };
    const idLinha = parseInt(form.loc_id);
    salvarDadosSeguro(NOME_ABA_LOCAIS, dados, isNaN(idLinha) ? null : idLinha);
    registrarLogAuditoria(form.editor_login || "Admin", isNaN(idLinha) ? "CRIAR_LOCAL" : "EDITAR_LOCAL", form.loc_nome, "");
    return isNaN(idLinha) ? "✅ Local criado com sucesso!" : "✅ Local atualizado com sucesso!";
  } catch (e) { return "❌ Erro ao salvar local: " + e.message; }
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
      "Modalidade": form.cert_mod, // INJETAR ESTA LINHA AQUI!
      "Data_Emissao": form.cert_data,
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
    bgColor: DEFAULT_ASSETS.BG_COLOR,
    linkLoja: DEFAULT_ASSETS.LINK_LOJA,
    linkInstagram: DEFAULT_ASSETS.LINK_INSTA,
    linkYouTube: DEFAULT_ASSETS.LINK_YT,
    linkCadastro: DEFAULT_ASSETS.LINK_CAD,
    nomeAcademia: "DojoManager SaaS" // <-- FALLBACK PADRÃO
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
      if (configRow.cor_fundo) configBase.bgColor = configRow.cor_fundo;

      if (configRow.link_loja) configBase.linkLoja = configRow.link_loja;
      if (configRow.link_instagram) configBase.linkInstagram = configRow.link_instagram;
      if (configRow.link_youtube) configBase.linkYouTube = configRow.link_youtube;
      if (configRow.link_cadastro) configBase.linkCadastro = configRow.link_cadastro;

      // <-- LENDO O NOME DA ACADEMIA DO BANCO
      if (configRow.nome_academia) configBase.nomeAcademia = configRow.nome_academia;
    }

    return configBase;
  } catch (e) {
    return configBase;
  }
}

// --- ATUALIZAÇÃO NO SETUP DO APP (Aba Config_App para o PIX Global) ---
function salvarConfig(form) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Config_App");

    const configsToSave = {
      "Logo_URL": padronizarLinkDrive(form.cfg_icon) || "",
      "Fundo_URL": padronizarLinkDrive(form.cfg_bg) || "",
      "Cor_Primaria": form.cfg_color || "#FFD700",
      "Cor_Secundaria": form.cfg_color_sec || "#1e1e1e",
      "Cor_Texto": form.cfg_color_text || "#ffffff",
      "Cor_Texto_Botao": form.cfg_color_btn_text || "#000000",
      "Cor_Fundo": form.cfg_color_bg || "#121212",
      "Link_Loja": form.cfg_link_loja || "",
      "Link_Instagram": form.cfg_link_insta || "",
      "Link_YouTube": form.cfg_link_yt || "",
      "Link_Cadastro": form.cfg_link_cad || "",
      "Nome_Academia": form.cfg_nome_acad || "DojoManager SaaS",
      "Pix_Global_Ativo": form.cfg_pix_global || "Nao",
      "Pix_Chave_Global": form.cfg_pix_chave,       // <-- Mapeado Novo Cabeçalho
      "Pix_Nome": form.cfg_pix_nome,
      "Nome_Banco_PIX_Global": form.cfg_pix_banco,  // <-- Mapeado Novo Cabeçalho
      "Pix_Cidade": form.cfg_pix_cidade
    };

    if (sheet.getLastRow() < 2) {
      sheet.appendRow(["", "", "#FFD700", "#1e1e1e", "#ffffff", "#000000", "#121212", "", "", "", "", "DojoManager SaaS"]);
    }

    salvarDadosSeguro("Config_App", configsToSave, 2);

    return "✅ Configurações atualizadas com sucesso!";
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
// 6. AUTENTICAÇÃO E A BALA DE PRATA (COM DADOS FISIOLÓGICOS E AULAS)
// ============================================================================

// ============================================================================
// 🔒 MÓDULO DE AUTENTICAÇÃO E RBAC (Role-Based Access Control)
// ============================================================================

function verificarCredenciais(formObject) {
  const loginInput = String(formObject.login).trim().toLowerCase();
  const isCheckOnly = formObject.checkOnly || false;

  try {
    const sheet = getSheet(NOME_ABA_ALUNOS);
    const data = sheet.getDataRange().getDisplayValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());

    const col = {
      login: headers.indexOf("login"),
      senha: headers.indexOf("senha"),
      nome: headers.indexOf("nome completo"),
      // 🥋 Poder do Tatame (Quem dá aula)
      nivelTatame: headers.indexOf("nível do praticante") !== -1 ? headers.indexOf("nível do praticante") : headers.indexOf("nivel do praticante"),
      // 🏢 Poder do Escritório (RBAC - ADM)
      nivelAdm: headers.indexOf("niveladministrativo") !== -1 ? headers.indexOf("niveladministrativo") : headers.indexOf("nivel administrativo"),
      foto: headers.indexOf("foto 3x4 (para a carteirinha)"),
      grad: headers.indexOf("graduacao_atual") !== -1 ? headers.indexOf("graduacao_atual") : headers.indexOf("graduação_atual"),
      acad: headers.indexOf("academia vinculada"),
      status: headers.indexOf("status"),
      modalidade: headers.indexOf("modalidade"),
      exame: headers.indexOf("data próximo exame"),
      nasc: headers.indexOf("data de nascimento"),
      peso: headers.indexOf("peso"),
      altura: headers.indexOf("altura"),
      carimbo: headers.findIndex(h => h.includes("carimbo"))
    };

    if (col.login === -1) throw new Error("Erro Crítico: Coluna LOGIN não encontrada.");

    let chamadas = [];
    try { chamadas = getSheet("Registro_Chamada").getDataRange().getDisplayValues(); } catch (e) { }

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[col.login]).trim().toLowerCase() === loginInput) {
        if (String(row[col.status] || "ativo").trim().toLowerCase() !== "ativo") return { success: false, message: "Cadastro inativo." };
        if (!isCheckOnly && String(formObject.senha).trim() !== String(row[col.senha]).trim()) return { success: false, message: "Senha incorreta." };

        const mapaGrad = getMapaGraduacoes();
        const gradStr = (col.grad > -1 && row[col.grad]) ? String(row[col.grad]).trim().toLowerCase() : "iniciante";
        const gradInfo = mapaGrad[gradStr] || { id: 0, nivel: "ALUNO", modalidade: "Geral" };

        const userModalidade = (col.modalidade > -1 && row[col.modalidade]) ? String(row[col.modalidade]).trim() : gradInfo.modalidade;

        // 🛡️ RBAC: SEPARAÇÃO DE PODERES
        const roleTatame = String((col.nivelTatame > -1 && row[col.nivelTatame]) ? row[col.nivelTatame] : gradInfo.nivel).toUpperCase();
        const roleOffice = String(col.nivelAdm > -1 ? row[col.nivelAdm] : "").toUpperCase();

        const isInstrutor = roleTatame.includes("INSTRUTOR") || roleTatame.includes("PROFESSOR") || roleTatame.includes("MESTRE");
        const isMestre = roleTatame.includes("MESTRE");

        // 🚨 A MÁGICA DO RBAC: Só é Admin do sistema se a coluna NivelAdministrativo disser "ADMIN" ou "ADM"
        const isAdmin = (roleOffice === "ADMIN" || roleOffice === "ADM");

        const dataNascimento = (col.nasc > -1) ? row[col.nasc] : "";
        const idadeCalculada = calcularIdadeExata(dataNascimento);

        let totalAulas = 0;
        if (chamadas.length > 1) {
          const colIdsChamada = chamadas[0].map(h => String(h).trim().toLowerCase()).indexOf("lista_alunos_ids");
          if (colIdsChamada !== -1) {
            for (let c = 1; c < chamadas.length; c++) {
              if (String(chamadas[c][colIdsChamada]).toLowerCase().includes(loginInput)) totalAulas++;
            }
          }
        }

        let dataIngresso = (col.carimbo > -1 && row[col.carimbo]) ? String(row[col.carimbo]).split(" ")[0] : "--/--/----";

        const userPayload = {
          LOGIN: loginInput,
          nomeCompleto: row[col.nome],
          graduacao: (col.grad > -1) ? row[col.grad] : "Iniciante",
          nivel: roleTatame,
          fotoUrl: padronizarLinkDrive(row[col.foto]),
          academia: (col.acad > -1) ? row[col.acad] : "---",
          proximoExame: (col.exame > -1) ? row[col.exame] : "A definir",
          modalidade: userModalidade,
          peso: (col.peso > -1) ? row[col.peso] : "",
          altura: (col.altura > -1) ? row[col.altura] : "",
          idadeExata: idadeCalculada,
          totalAulas: totalAulas,
          carimbo: dataIngresso,
          isInstrutor: isInstrutor,
          isMestre: isMestre,
          isAdmin: isAdmin // <-- Variável que trava o Financeiro e Admin no Frontend
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
 * 📝 QA Note: God Mode. O Mestre adiciona o aluno na mão, furando a trava de idade,
 * mas mantendo a TAG de idade para controle visual. Mapeamento dinâmico incluso.
 */
function adicionarAlunoManual(idAula, loginOuNome) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSheet("Aulas_Em_Andamento");
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: false, msg: "Nenhuma aula ativa." };

    // 🧠 MAPEAMENTO DINÂMICO DE COLUNAS
    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const colId = headers.indexOf("id_aula");
    const colJson = headers.indexOf("checkins_json");

    let linhaAlvo = -1; let checkinsAtuais = [];

    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][colId]) === idAula) {
        linhaAlvo = i + 1; checkinsAtuais = JSON.parse(data[i][colJson] || "[]"); break;
      }
    }
    if (linhaAlvo === -1) return { success: false, msg: "Aula não encontrada." };

    let novoCheckin = null;
    if (String(loginOuNome).startsWith("VISITANTE_")) {
      const nomeLimpo = loginOuNome.replace("VISITANTE_", "");
      if (checkinsAtuais.some(c => c.nome.toLowerCase() === nomeLimpo.toLowerCase())) return { success: false, msg: "Nome já está na lista." };
      novoCheckin = {
        login: "GUEST-" + Date.now(), nome: nomeLimpo, foto: "https://via.placeholder.com/150?text=VIS",
        graduacao: "Visitante", academia_origem: "Externo", status: "Confirmado",
        tipo: "Visitante", idade: "?", categoria: "Visitante", hora: new Date().toLocaleTimeString()
      };
    } else {
      if (checkinsAtuais.some(c => c.login === loginOuNome)) return { success: false, msg: "Aluno já está na lista." };

      const aluno = buscarAlunoPorLogin(loginOuNome);
      if (!aluno) return { success: false, msg: "Aluno não encontrado no banco." };

      let idadeAnos = 0;
      if (aluno.idadeExata && aluno.idadeExata !== "N/A") { idadeAnos = parseInt(aluno.idadeExata.split('a')[0]) || 0; }

      const IDADE_CORTE_ADULTO = 15;
      const categoriaIdadeAluno = idadeAnos >= IDADE_CORTE_ADULTO ? "Adulto" : "Infantil";

      novoCheckin = {
        login: aluno.LOGIN || loginOuNome, nome: aluno.nomeCompleto, foto: aluno.fotoUrl,
        graduacao: aluno.graduacao, academia_origem: String(aluno.academia), status: "Confirmado",
        tipo: "Regular", idade: idadeAnos, categoria: categoriaIdadeAluno, hora: new Date().toLocaleTimeString()
      };
    }

    checkinsAtuais.push(novoCheckin);
    sheet.getRange(linhaAlvo, colJson + 1).setValue(JSON.stringify(checkinsAtuais));
    return { success: true, msg: "Adicionado com sucesso (Modo Mestre)!" };
  } catch (e) { return { success: false, msg: "Erro: " + e.message }; } finally { lock.releaseLock(); }
}

// ============================================================================
// ⚙️ HELPER DE FORMATAÇÃO DE HORA
// ============================================================================
function formatarHoraSimples(hora) {
  if (!hora) return "--:--";
  if (hora instanceof Date) return Utilities.formatDate(hora, Session.getScriptTimeZone(), "HH:mm");
  let str = String(hora).trim();
  if (str.length >= 5 && str.includes(":")) return str.substring(0, 5); // Ex: "20:30:00" vira "20:30"
  return str;
}

// ============================================================================
// 🏫 MOTOR DE AULAS DINÂMICAS E CHECK-IN (ARQUITETURA DE 10 COLUNAS)
// ============================================================================

/**
 * 📝 QA Note: Cria uma nova aula. Utiliza a função 'salvarDadosSeguro' que 
 * internamente procura as colunas pelo nome exato, tornando a gravação blindada
 * contra mudanças na ordem das colunas da planilha.
 */
function iniciarAulaDinamica(dados) {
  limparAulasExpiradas();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Aulas_Em_Andamento");

    const idAula = "AULA-" + Date.now();
    const pin = Math.floor(1000 + Math.random() * 9000);

    let dataStr = dados.data;
    if (dados.data instanceof Date) dataStr = Utilities.formatDate(dados.data, Session.getScriptTimeZone(), "yyyy-MM-dd");

    const novaAula = {
      "ID_Aula": idAula,
      "Data_Aula": dataStr,
      "Hora_Inicio": String(dados.inicio),
      "Hora_Fim": String(dados.fim),
      "Academia": dados.local,
      "Turma": dados.turma,
      "Instrutor": dados.instrutor,
      "PIN": pin,
      "Status": "ABERTA",
      "Checkins_JSON": "[]" // AQUI ESTAVA O SEGREDO DO JSON INVALIDO
    };

    salvarDadosSeguro("Aulas_Em_Andamento", novaAula);
    return { success: true, idAula: idAula, pin: pin };
  } catch (e) { return { success: false, erro: e.message }; } finally { lock.releaseLock(); }
}
/**
 * 📝 QA Note: Busca as aulas ativas para o aluno fazer check-in.
 * Mapeia dinamicamente os cabeçalhos para ler as posições corretas.
 */
function buscarAulasDisponiveisAluno(academiaAluno) {
  try {
    const sheet = getSheet("Aulas_Em_Andamento");
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    const agora = new Date();
    const aulas = [];
    const minhaAcademia = String(academiaAluno).trim().toUpperCase();
    const buscarTudo = (minhaAcademia === "TODOS");
    const hojeStr = Utilities.formatDate(agora, Session.getScriptTimeZone(), "yyyy-MM-dd");

    // 🧠 MAPEAMENTO DINÂMICO DE COLUNAS
    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const colId = headers.indexOf("id_aula");
    const colData = headers.indexOf("data_aula");
    const colInicio = headers.indexOf("hora_inicio");
    const colFim = headers.indexOf("hora_fim");
    const colAcad = headers.indexOf("academia");
    const colTurma = headers.indexOf("turma");
    const colInst = headers.indexOf("instrutor");
    const colPin = headers.indexOf("pin");
    const colStatus = headers.indexOf("status");

    for (let i = 1; i < data.length; i++) {
      const status = String(data[i][colStatus]);
      const acadAula = String(data[i][colAcad]).trim().toUpperCase();

      if (status === "ABERTA" && (buscarTudo || acadAula === minhaAcademia)) {
        let diaPlanilha = data[i][colData];
        let diaStr = "";
        if (diaPlanilha instanceof Date) {
          diaStr = Utilities.formatDate(diaPlanilha, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else {
          diaStr = String(diaPlanilha).split("T")[0];
        }

        if (diaStr === hojeStr) {
          aulas.push({
            id: data[i][colId],
            instrutor: data[i][colInst],
            academia: data[i][colAcad],
            turma: data[i][colTurma],
            horario: `${formatarHoraSimples(data[i][colInicio])} às ${formatarHoraSimples(data[i][colFim])}`,
            pin: data[i][colPin]
          });
        }
      }
    }
    return aulas;
  } catch (e) { return []; }
}

// ============================================================================
// 🥋 MOTOR DE CHECK-IN E CONTROLE DE TATAME (COM TRAVA DE IDADE V3 - FINAL)
// ============================================================================

/**
 * 📝 QA Note: O aluno tenta fazer check-in. O sistema busca dinamicamente as colunas
 * para evitar quebra de índice, calcula a idade e barra se a turma não permitir.
 */
function realizarCheckinAluno(login, pinDigitado) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheetAulas = getSheet("Aulas_Em_Andamento");
    const dataAulas = sheetAulas.getDataRange().getValues();
    if (dataAulas.length < 2) return { success: false, msg: "Nenhuma aula ativa encontrada." };

    // 🧠 MAPEAMENTO DINÂMICO DE COLUNAS
    const headers = dataAulas[0].map(h => String(h).trim().toLowerCase());
    const colPin = headers.indexOf("pin");
    const colStatus = headers.indexOf("status");
    const colLocal = headers.indexOf("academia");
    const colTurma = headers.indexOf("turma");
    const colJson = headers.indexOf("checkins_json");

    let linhaAlvo = -1; let checkinsAtuais = []; let dadosAula = {};

    for (let i = dataAulas.length - 1; i >= 1; i--) {
      if (String(dataAulas[i][colPin]).trim() === String(pinDigitado).trim() && dataAulas[i][colStatus] === "ABERTA") {
        linhaAlvo = i + 1;
        checkinsAtuais = JSON.parse(dataAulas[i][colJson] || "[]");
        dadosAula = { local: dataAulas[i][colLocal], turma: dataAulas[i][colTurma] };
        break;
      }
    }

    if (linhaAlvo === -1) return { success: false, msg: "PIN inválido ou aula encerrada." };
    if (checkinsAtuais.some(c => c.login === login)) return { success: false, msg: "Você já realizou check-in." };

    const aluno = buscarAlunoPorLogin(login);
    if (!aluno) return { success: false, msg: "Aluno não identificado." };

    const turmas = lerTabelaDinamica("Config_Turmas");
    const turmaInfo = turmas.find(t => String(t.nome_da_turma).trim().toLowerCase() === String(dadosAula.turma).trim().toLowerCase());

    const faixaEtariaTurma = turmaInfo ? String(turmaInfo["faixa_etária"] || turmaInfo.faixa_etaria || "Misto").trim() : "Misto";
    const localDaTurma = dadosAula.local || "Matriz";

    let idadeAnos = 0;
    if (aluno.idadeExata && aluno.idadeExata !== "N/A") {
      idadeAnos = parseInt(aluno.idadeExata.split('a')[0]) || 0;
    }

    const IDADE_CORTE_ADULTO = 15;
    const categoriaIdadeAluno = idadeAnos >= IDADE_CORTE_ADULTO ? "Adulto" : "Infantil";

    // 🛑 A TRAVA DE IDADE BLINDADA
    if (faixaEtariaTurma.toLowerCase() === "infantil" && idadeAnos >= IDADE_CORTE_ADULTO) {
      return { success: false, msg: `🚫 Turma Infantil. Sua idade (${idadeAnos} anos) requer aprovação manual do Mestre.` };
    }
    if (faixaEtariaTurma.toLowerCase() === "adulto" && idadeAnos > 0 && idadeAnos < IDADE_CORTE_ADULTO) {
      return { success: false, msg: `🚫 Turma Adulto. Sua idade (${idadeAnos} anos) requer aprovação manual do Mestre.` };
    }

    const acadAluno = String(aluno.academia || "");
    const loc1 = String(localDaTurma).replace(/\s+/g, '').toUpperCase();
    const loc2 = acadAluno.replace(/\s+/g, '').toUpperCase();
    const tipoAluno = (loc1 !== loc2) ? "Visitante" : "Regular";

    checkinsAtuais.push({
      login: login, nome: aluno.nomeCompleto || "Aluno", foto: aluno.fotoUrl || "",
      graduacao: aluno.graduacao || "Iniciante", academia_origem: acadAluno,
      status: "Pendente", tipo: tipoAluno, idade: idadeAnos, categoria: categoriaIdadeAluno,
      hora: new Date().toLocaleTimeString()
    });

    sheetAulas.getRange(linhaAlvo, colJson + 1).setValue(JSON.stringify(checkinsAtuais));
    return { success: true, msg: "Check-in realizado! Aguarde o instrutor." };
  } catch (e) { return { success: false, msg: "Erro Crítico: " + e.message }; } finally { lock.releaseLock(); }
}


// ============================================================================
// 🔍 BUSCA DE AULAS PARA O INSTRUTOR E ALUNO
// ============================================================================

/**
 * 📝 QA Note: Retorna as aulas em andamento para o painel do instrutor.
 * Totalmente desacoplada de colunas engessadas.
 */
function buscarAulasAtivasInstrutor(loginInstrutor) {
  try {
    const sheet = getSheet("Aulas_Em_Andamento");
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    const aulas = [];
    const nomeInst = String(loginInstrutor).trim().toLowerCase();

    // 🧠 MAPEAMENTO DINÂMICO DE COLUNAS
    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const colId = headers.indexOf("id_aula");
    const colData = headers.indexOf("data_aula");
    const colInicio = headers.indexOf("hora_inicio");
    const colFim = headers.indexOf("hora_fim");
    const colAcad = headers.indexOf("academia");
    const colTurma = headers.indexOf("turma");
    const colInst = headers.indexOf("instrutor");
    const colPin = headers.indexOf("pin");
    const colStatus = headers.indexOf("status");
    const colJson = headers.indexOf("checkins_json");

    for (let i = 1; i < data.length; i++) {
      if (data[i][colStatus] === "ABERTA") {
        const donoAula = String(data[i][colInst]).trim().toLowerCase();
        if (donoAula.includes(nomeInst.split(' ')[0].toLowerCase()) || true) {
          let dataVisual = data[i][colData];
          if (dataVisual instanceof Date) dataVisual = Utilities.formatDate(dataVisual, Session.getScriptTimeZone(), "dd/MM/yyyy");
          else if (typeof dataVisual === 'string' && dataVisual.includes('-')) {
            const p = dataVisual.split('-'); dataVisual = `${p[2]}/${p[1]}/${p[0]}`;
          }

          aulas.push({
            id: data[i][colId],
            data: dataVisual,
            academia: data[i][colAcad],
            turma: data[i][colTurma],
            horario: `${formatarHoraSimples(data[i][colInicio])} - ${formatarHoraSimples(data[i][colFim])}`,
            pin: data[i][colPin],
            qtd: JSON.parse(data[i][colJson] || "[]").length
          });
        }
      }
    }
    return aulas;
  } catch (e) { return []; }
}

/**
 * 📝 QA Note: Lê os check-ins ao vivo (Polling) via mapeamento dinâmico.
 */
function buscarCheckinsAula(idAula) {
  try {
    const sheet = getSheet("Aulas_Em_Andamento");
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: false };

    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const colId = headers.indexOf("id_aula");
    const colStatus = headers.indexOf("status");
    const colJson = headers.indexOf("checkins_json");

    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][colId]) === idAula) {
        return { success: true, status: data[i][colStatus], checkins: JSON.parse(data[i][colJson] || "[]") };
      }
    }
    return { success: false };
  } catch (e) { return { success: false }; }
}

/**
 * 📝 QA Note: Transforma a aula provisória em um Histórico Oficial.
 * Mapeia os cabeçalhos para evitar reescrever o Status na coluna errada.
 */
function finalizarAulaDinamica(idAula, listaFinalAlunos) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const sheetTemp = getSheet("Aulas_Em_Andamento");
    const dataTemp = sheetTemp.getDataRange().getValues();
    if (dataTemp.length < 2) return { success: false, msg: "Tabela de aulas vazia." };

    // 🧠 MAPEAMENTO DINÂMICO DE COLUNAS
    const headers = dataTemp[0].map(h => String(h).trim().toLowerCase());
    const colId = headers.indexOf("id_aula");
    const colData = headers.indexOf("data_aula");
    const colAcad = headers.indexOf("academia");
    const colTurma = headers.indexOf("turma");
    const colInst = headers.indexOf("instrutor");
    const colStatus = headers.indexOf("status");

    let linhaAlvo = -1;
    let dadosAula = null;

    for (let i = dataTemp.length - 1; i >= 1; i--) {
      if (String(dataTemp[i][colId]) === idAula) {
        linhaAlvo = i + 1;
        let dataF = dataTemp[i][colData];
        if (dataF instanceof Date) dataF = Utilities.formatDate(dataF, Session.getScriptTimeZone(), "yyyy-MM-dd");
        else if (typeof dataF === 'string' && dataF.includes('/')) { const p = dataF.split('/'); dataF = `${p[2]}-${p[1]}-${p[0]}`; }

        dadosAula = { data: dataF, instrutor: dataTemp[i][colInst], turma: dataTemp[i][colTurma], local: dataTemp[i][colAcad] };
        break;
      }
    }

    if (!dadosAula) return { success: false, msg: "Aula não encontrada." };

    const localRelatorio = dadosAula.local + " | " + dadosAula.turma;
    const presentes = listaFinalAlunos.filter(a => a.status === "Confirmado");

    if (presentes.length > 0) {
      salvarDadosSeguro("Registro_Chamada", {
        "ID_Chamada": idAula,
        "Data_Registro": Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"),
        "Data_Treino": dadosAula.data,
        "Hora_Treino": new Date().toLocaleTimeString(),
        "Instrutor_Logado": dadosAula.instrutor,
        "Local_Treino": localRelatorio,
        "Qtd_Presentes": presentes.length,
        "Lista_Alunos_IDs": presentes.map(a => a.login).join(", "),
        "Lista_Nomes": presentes.map(a => a.nome + (a.tipo === 'Visitante' ? ' [VIS]' : '')).join(", "),
        "Observacoes": "Chamada Dinâmica App"
      });
    }

    // Atualiza a célula exata do Status, não importa onde ela esteja (+1 porque array começa no 0 e getRange no 1)
    sheetTemp.getRange(linhaAlvo, colStatus + 1).setValue("ENCERRADA");
    return { success: true, msg: "Aula finalizada e Histórico atualizado!" };
  } catch (e) { return { success: false, msg: "Erro: " + e.message }; } finally { lock.releaseLock(); }
}

/**
 * 📝 QA Note: Rotina de limpeza. Encontra colunas de data e status dinamicamente.
 */
function limparAulasExpiradas() {
  try {
    const sheet = getSheet("Aulas_Em_Andamento");
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;

    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const colData = headers.indexOf("data_aula");
    const colStatus = headers.indexOf("status");

    const agora = new Date();
    // Apaga de baixo para cima para não quebrar os índices
    for (let i = data.length - 1; i >= 1; i--) {
      let d = new Date(data[i][colData]);
      if ((agora - d) > (24 * 3600 * 1000) && data[i][colStatus] === "ABERTA") {
        sheet.deleteRow(i + 1);
      }
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

/// ============================================================================
// 🪪 MÓDULO: GERADOR DE CARTEIRINHA EM PDF (COM CONVERSÃO DE IMAGENS BASE64)
// ============================================================================

/**
 * Converte imagens do Drive para Base64.
 * Se falhar, retorna um SVG inline (texto puro) para não depender de APIs externas.
 */
function converterImagemParaBase64(input, isLogo = false) {
  // SVG nativo para Alunos sem foto (Rosto genérico)
  const defaultUserSvg = "data:image/svg+xml;base64," + Utilities.base64Encode('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="#666"><path d="M12 12c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4zm0 2c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4z"/></svg>');

  // SVG nativo para academias sem Logo (Escudo genérico)
  const defaultLogoSvg = "data:image/svg+xml;base64," + Utilities.base64Encode('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="#FFD700"><path d="M12 1L3 5v6c0 5.55 3.84 10.74 9 12 5.16-1.26 9-6.45 9-12V5l-9-4z"/></svg>');

  const fallback = isLogo ? defaultLogoSvg : defaultUserSvg;

  if (!input || String(input).trim() === "") {
    return fallback;
  }

  let id = "";
  const strInput = String(input).trim();

  const matchId = strInput.match(/id=([a-zA-Z0-9_-]+)/) || strInput.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (matchId) {
    id = matchId[1];
  } else if (strInput.includes("googleusercontent.com/profile/picture/0")) {
    id = strInput.split("picture/0")[1];
  } else if (strInput.length > 20 && !strInput.includes("/")) {
    id = strInput;
  }

  try {
    if (id) {
      const file = DriveApp.getFileById(id);
      const blob = file.getBlob();
      return "data:" + blob.getContentType() + ";base64," + Utilities.base64Encode(blob.getBytes());
    }
  } catch (e) {
    return fallback;
  }

  return fallback;
}

/** Helper para converter links da Web em Base64 */
function fetchBase64(url) {
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() === 200) {
      const blob = response.getBlob();
      return "data:" + blob.getContentType() + ";base64," + Utilities.base64Encode(blob.getBytes());
    }
  } catch (e) { }
  return ""; // Retorna vazio se tudo falhar, o CSS mascara a ausência
}

/**
 * Busca todos os dados brutos do aluno direto na planilha
 */
function getAlunoFullData(loginInput) {
  try {
    const sheet = getSheet("cadastro_de_alunos");
    const data = sheet.getDataRange().getDisplayValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());

    const colLogin = headers.indexOf("login");
    if (colLogin === -1) return null;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][colLogin]).trim().toLowerCase() === String(loginInput).trim().toLowerCase()) {
        let obj = {};
        headers.forEach((h, idx) => { obj[h] = data[i][idx]; });
        return obj;
      }
    }
    return null;
  } catch (e) { return null; }
}

/**
 * 🛠️ BACKEND: BUSCA DADOS PIX (Lógica Global vs Local)
 */
function getPixDataFromServer(login) {
  try {
    const configApp = lerTabelaDinamica("Config_App")[0];
    const alunos = lerTabelaDinamica(NOME_ABA_ALUNOS);
    const locais = lerTabelaDinamica(NOME_ABA_LOCAIS);
    const assinaturas = lerTabelaDinamica("Fin_Assinaturas");
    const pacotes = lerTabelaDinamica("Fin_Pacotes");

    const aluno = alunos.find(a => String(a.login).toLowerCase() === String(login).toLowerCase());
    if (!aluno) return null;

    const assinatura = assinaturas.find(as => String(as.login_aluno).toLowerCase() === String(login).toLowerCase());
    const pacote = pacotes.find(p => p.nome_pacote === assinatura?.pacote_atual);

    // 1. Inicia o payload puxando as colunas da CONTA GLOBAL (Aba Config_App)
    let info = {
      chave: configApp.pix_chave_global || "",
      beneficiario: configApp.pix_nome || "Dojo Manager",
      cidade: configApp.pix_cidade || "RECIFE",
      banco: configApp.nome_banco_pix_global || "Instituição Bancária",
      valor: pacote ? pacote.valor_padrao : 0, // Zera se não tiver pacote
      contatoWhatsApp: ""
    };

    // 2. 🛡️ A MÁGICA: Se o PIX Global estiver DESATIVADO, sobrepõe com a CONTA LOCAL
    if (String(configApp.pix_global_ativo).toLowerCase() !== "sim") {
      const local = locais.find(l => String(l.nome_do_local).toLowerCase() === String(aluno.academia_vinculada).toLowerCase());

      // Só substitui se o local existir e a aba status do local estiver 'Ativo'
      if (local && String(local.status).toLowerCase() === "ativo") {
        info.chave = local.pix_chave_local_de_treino || info.chave;
        info.beneficiario = local.pix_nome_local_de_treino || info.beneficiario;
        info.cidade = local.pix_cidade_local_de_treino || info.cidade;
        info.banco = local.banco_local_de_treino || info.banco;
      }
    }

    // Pega o contato WhatsApp do local de qualquer forma para o aluno enviar o comprovante
    const localResp = locais.find(l => String(l.nome_do_local).toLowerCase() === String(aluno.academia_vinculada).toLowerCase());
    info.contatoWhatsApp = localResp?.contato ? String(localResp.contato).replace(/\D/g, "") : "5581997629232";

    return info;
  } catch (e) {
    console.error("Erro fatal no servidor ao buscar PIX: " + e.message);
    return null;
  }
}

/**
 * Gera o HTML, converte para PDF e retorna codificado
 */
function gerarPDFCarteirinhaServer(login) {
  try {
    const alunoInfo = getAlunoFullData(login);
    if (!alunoInfo) throw new Error("Dados do aluno não encontrados.");

    const config = getAppConfig();
    const template = HtmlService.createTemplateFromFile('CardTemplate');

    // 1. Injeta Dados em Texto
    template.nome = alunoInfo["nome completo"] || "NÃO INFORMADO";
    template.cpf = alunoInfo["cpf"] || "000.000.000-00";
    template.graduacao = alunoInfo["graduacao_atual"] || "INICIANTE";
    template.academia = alunoInfo["academia vinculada"] || "MATRIZ";

    let filiacaoStr = (alunoInfo["nome da mãe"] || "") + " / " + (alunoInfo["nome do pai"] || "");
    if (filiacaoStr === " / ") filiacaoStr = "NÃO INFORMADA";
    template.filiacao = filiacaoStr;

    template.nascimento = alunoInfo["data de nascimento"] || "00/00/0000";

    // Validade: 1 ano a partir de hoje
    let hj = new Date();
    template.emissao = Utilities.formatDate(hj, Session.getScriptTimeZone(), "dd/MM/yyyy");
    hj.setFullYear(hj.getFullYear() + 1);
    template.validade = Utilities.formatDate(hj, Session.getScriptTimeZone(), "dd/MM/yyyy");

    // 2. Injeta as Cores
    template.primaryColor = config.primaryColor;
    template.secondaryColor = config.secondaryColor;
    template.textColor = config.textColor;
    template.btnTextColor = config.btnTextColor;

    // 3. INJETA AS IMAGENS "FISICAMENTE" (BASE64) NO HTML
    // 3. INJETA AS IMAGENS "FISICAMENTE" (BASE64) NO HTML
    template.fotoUrl = converterImagemParaBase64(alunoInfo["foto 3x4 (para a carteirinha)"], false);
    template.logoUrl = converterImagemParaBase64(config.iconId, true);

    // 4. Avalia e converte o motor
    const htmlEvaluated = template.evaluate();
    const blob = htmlEvaluated.getAs(MimeType.PDF);

    const fileName = "Carteira_" + template.nome.split(" ")[0] + ".pdf";
    blob.setName(fileName);

    return {
      success: true,
      base64: Utilities.base64Encode(blob.getBytes()),
      fileName: fileName
    };

  } catch (e) {
    return { success: false, msg: e.message };
  }
}

// ============================================================================
// 🥋 NOVOS CRUDS ADMIN (GRADUAÇÕES E MULTIMODALIDADE)
// ============================================================================

function listarGraduacoesAdminCompleto() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("GRADUACAO") || ss.getSheetByName("Graduação") || ss.getSheetByName("Graduacao");
    if (!sheet) return [];

    const data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return [];

    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const colGrad = headers.indexOf("graduação") !== -1 ? headers.indexOf("graduação") : headers.indexOf("faixa / nivel");
    const colObs = headers.indexOf("observação") !== -1 ? headers.indexOf("observação") : headers.indexOf("observacao");
    const colId = headers.indexOf("id");
    const colMod = headers.indexOf("modalidade");
    const colNivel = headers.indexOf("nivel");

    return data.slice(1).map((row, i) => ({
      linha: i + 2,
      graduacao: row[colGrad] || "",
      observacao: colObs > -1 ? row[colObs] : "",
      id_hierarquia: row[colId] || "0",
      modalidade: colMod > -1 ? row[colMod] : "Geral",
      nivel: colNivel > -1 ? row[colNivel] : "Aluno"
    })).filter(g => g.graduacao !== "");
  } catch (e) { return []; }
}

function salvarGraduacaoDefinitiva(form) {
  try {
    const dados = {
      "Graduação": form.grad_nome,
      "Observação": form.grad_obs,
      "ID": form.grad_id,
      "Modalidade": form.grad_mod,
      "Nivel": form.grad_nivel
    };
    const idLinha = parseInt(form.linha_id);
    salvarDadosSeguro("GRADUACAO", dados, isNaN(idLinha) ? null : idLinha);
    return isNaN(idLinha) ? "✅ Nova graduação criada com sucesso!" : "✅ Graduação atualizada com sucesso!";
  } catch (e) { return "❌ Erro ao salvar graduação: " + e.message; }
}

function excluirGraduacaoAdmin(linha) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("GRADUACAO") || ss.getSheetByName("Graduação");
    sheet.deleteRow(parseInt(linha));
    return "✅ Graduação excluída!";
  } catch (e) { return "❌ Erro: " + e.message; }
}

// ============================================================================
// 🏫 MÓDULO: GESTÃO DE TURMAS (GRADE DE HORÁRIOS)
// ============================================================================

/**
 * Lê todas as turmas cadastradas na aba Config_Turmas
 */
function listarTurmasAdmin() {
  try {
    const data = lerTabelaDinamica("Config_Turmas");
    return data.map(t => ({
      idLinha: t._linha,
      idTurma: t.id_turma || "",
      nome: t.nome_da_turma || "Sem Nome",
      modalidade: t.modalidade || "Geral",
      faixaEtaria: t["faixa_etária"] || t.faixa_etaria || "Misto",
      local: t.local_vinculado || "Matriz",
      dias: t.dias_da_semana || "",
      inicio: t["horário_início"] || t.horario_inicio || "",
      fim: t["horário_fim"] || t.horario_fim || "",
      status: t.status || "Ativa",
      responsavel: t.responsavel || t['responsável'] || "",
      telefone: t.telefone || ""
    }));
  } catch (e) {
    console.error("Erro ao listar turmas: " + e.message);
    return [];
  }
}

/**
 * ============================================================================
 * 🛡️ FIX: ANTI-APAGÃO PARA TURMAS (HYDRATION)
 * Garante que a edição de uma turma não apague colunas extras que o cliente 
 * possa adicionar no futuro na aba 'Config_Turmas'.
 * ============================================================================
 */
function salvarTurma(form) {
  try {
    const nomeAba = "Config_Turmas";
    const sheet = getSheet(nomeAba);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    const normalize = (str) => String(str).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, "");
    const headerMap = {};
    headers.forEach(h => { if (h) headerMap[normalize(h)] = h; });

    const idLinha = parseInt(form.turma_linha_id);
    let linhaExistente = [];

    if (!isNaN(idLinha) && idLinha > 1 && idLinha <= sheet.getLastRow()) {
      linhaExistente = sheet.getRange(idLinha, 1, 1, headers.length).getValues()[0];
    }

    const getOldVal = (nomeRealColuna) => {
      if (linhaExistente.length === 0) return "";
      const idx = headers.indexOf(nomeRealColuna);
      return idx !== -1 ? linhaExistente[idx] : "";
    };

    const resolveField = (formKey, colPlanilha) => {
      const realCol = headerMap[normalize(colPlanilha)];
      if (!realCol) return null;
      let valorParaSalvar = Object.prototype.hasOwnProperty.call(form, formKey) ? form[formKey] : getOldVal(realCol);
      return { col: realCol, val: valorParaSalvar };
    };

    const idTurmaGerado = form.turma_id || getOldVal(headerMap[normalize("ID_Turma")]) || ("TUR-" + Date.now());

    const campos = [
      { col: headerMap[normalize("ID_Turma")], val: idTurmaGerado },
      resolveField("turma_nome", "Nome da Turma"),
      resolveField("turma_mod", "Modalidade"),
      resolveField("turma_idade", "Faixa Etária"),
      resolveField("turma_local", "Local Vinculado"),
      resolveField("turma_dias", "Dias da Semana"),
      resolveField("turma_inicio", "Horário Início"),
      resolveField("turma_fim", "Horário Fim"),
      resolveField("turma_status", "Status"),
      resolveField("turma_resp", "Responsável"),
      resolveField("turma_tel", "Telefone")
    ];

    const dadosFinais = {};
    campos.forEach(item => { if (item && item.col) dadosFinais[item.col] = item.val; });

    salvarDadosSeguro(nomeAba, dadosFinais, isNaN(idLinha) ? null : idLinha);
    return isNaN(idLinha) ? "✅ Turma criada com sucesso!" : "✅ Turma atualizada com sucesso!";
  } catch (e) { return "❌ Erro ao salvar turma: " + e.message; }
}