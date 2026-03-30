
/**
 * ============================================================================
 * ARQUIVO: Código.gs 
 *  
 *
 * AUTOR: Gleyson Atanazio  *
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
const NOME_ABA_ALUNOS = "cadastro_de_alunos";     // Tabela principal de usuários
const NOME_ABA_CURSOS = "Cursos";                 // Tabela de eventos/cursos
const NOME_ABA_VIDEOTECA = "Config_Videoteca";    // Tabela de links do YouTube
const NOME_ABA_LOCAIS = "Locais_de_treino";       // Tabela de academias
const NOME_ABA_PROGRAMAS = "Config_Programas";    // Tabela de PDFs técnicos
const NOME_ABA_BIBLIOTECA = "Config_Biblioteca";  // Tabela de livros extras
const NOME_ABA_AULAS = "Aulas_Em_Andamento";      // Tabela temporária para check-ins ativos
const ABA_LOGS = "Logs_Auditoria";                // Tabela de logs de segurança

// URL do script publicada (usada para redirecionamentos e links internos)
const SCRIPT_URL = ScriptApp.getService().getUrl();

// IDs de arquivos no Google Drive para assets padrão (Logo e Fundo)
const DEFAULT_ASSETS = {
  BACKGROUND_ID: "",
  ICON_ID: "",
  PRIMARY_COLOR: "#FFFFFF",
  SECONDARY_COLOR: "#1e1e1e",
  TEXT_COLOR: "#ffffff",
  BTN_TEXT_COLOR: "#000000",
  BG_COLOR: "#121212",
  // <-- NOVOS LINKS PADRÕES DO SISTEMA (FALLBACK)
  LINK_LOJA: "",
  LINK_INSTA: "",
  LINK_YT: "",
  LINK_CAD: ""
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
 * ============================================================================
 * 🛡️ MOTOR DE AUTORIZAÇÃO MILITAR (MULTI-TENANCY)
 * Verifica se o usuário é Mono-Loja, Multi-Loja ou Franquia (Master).
 * ============================================================================
 */
function getPermissoesUsuario(login) {
  if (!login) return { isMaster: false, academias: [] };
  try {
    const alunos = lerTabelaDinamica("cadastro_de_alunos");
    const user = alunos.find(a => String(a.login).toLowerCase() === String(login).toLowerCase());
    if (!user) return { isMaster: false, academias: [] };

    // Pega as academias do usuário e transforma em um array em minúsculas
    const acads = String(user.academia_vinculada || "").split(',').map(a => a.trim().toLowerCase());

    // Se tiver "todas", ele tem passe livre no sistema inteiro (Dono da Franquia)
    const isMaster = acads.includes("todas");

    return { isMaster: isMaster, academias: acads };
  } catch (e) {
    return { isMaster: false, academias: [] };
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
    { nome: "GRADUACAO", colunas: ["Graduação", "Observacao", "ID", "Modalidade", "Nivel"] },
    { nome: "Categoria_financeira", colunas: ["Nome", "Tipo", "Local", "Exibe Aluno", "Status"] },
    { nome: "Forma_Pagamento", colunas: ["Nome", "Local", "Status"] }
  ];

  estrutura.forEach(aba => {
    let sheet = ss.getSheetByName(aba.nome);
    let mudancaDetectada = false;

    if (!sheet) {
      sheet = ss.insertSheet(aba.nome);
      sheet.appendRow(aba.colunas);
      mudancaDetectada = true;

      // Setup inicial para Config_App
      if (aba.nome === "Config_App") {
        sheet.appendRow(["", "", "#121212", "#FFD700", "#1e1e1e", "#ffffff", "#000000", "", "", "", "", "DojoManager SaaS", "Nao", "", "", "", ""]);
      }
    }

    const colunasAtuais = sheet.getLastColumn() > 0
      ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim().toLowerCase())
      : [];

    aba.colunas.forEach(col => {
      if (!colunasAtuais.includes(col.trim().toLowerCase())) {
        sheet.getRange(1, sheet.getLastColumn() + 1).setValue(col);
        mudancaDetectada = true;
      }
    });

    // 🎨 Só aplica estilo se houver mudança ou aba nova (Poupa tempo de execução)
    if (mudancaDetectada) {
      const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
      headerRange.setFontWeight("bold")
        .setBackground("#2c3e50")
        .setFontColor("#ffffff")
        .setHorizontalAlignment("center");
      sheet.setFrozenRows(1);
      // Auto-ajuste de largura para facilitar a leitura na planilha
      sheet.autoResizeColumns(1, aba.colunas.length);
    }
  });

  console.log("✅ Verificação de integridade do banco concluída.");
}

/**
 * RELATÓRIO ADMIN (FINANCEIRO TELA DE LANÇAMENTOS) - 100% BLINDADO PELA ACADEMIA VINCULADA
 */
function getRelatorioFinanceiroAdmin(loginSolicitante, filtros = {}) {
  try {
    const transacoes = lerTabelaDinamica("Fin_Transacoes");
    const alunos = lerTabelaDinamica("cadastro_de_alunos");

    // 1. Puxa o solicitante e aplica a Regra Mestra da Academia
    const solicitante = alunos.find(a => String(a.login).toLowerCase() === String(loginSolicitante).toLowerCase());
    const acadsSolicitante = solicitante ? String(solicitante.academia_vinculada || "").toLowerCase().split(',').map(a => a.trim()) : [];

    // O usuário só é Master/Deus se tiver "todas" escrito na Academia Vinculada!
    const isMaster = acadsSolicitante.includes("todas");

    let ent = 0, sai = 0;

    // 2. Filtragem
    const dadosFiltrados = transacoes.filter(t => {
      const tAcad = String(t.academia_ref || "").trim().toLowerCase();
      const tData = parseDataSegura(t.data_registro);

      // REGRA MESTRA: Se não for Master, a transação TEM que pertencer às academias dele
      if (!isMaster) {
        if (!acadsSolicitante.some(myAcad => tAcad.includes(myAcad))) return false;
      }

      // Filtro de Tela (O que o usuário selecionou no Dropdown)
      if (filtros.academia && filtros.academia !== "TODAS") {
        if (tAcad !== String(filtros.academia).toLowerCase()) return false;
      }

      // Filtro de Data
      if (filtros.dataInicio && tData) {
        const dIni = new Date(filtros.dataInicio);
        if (tData < dIni) return false;
      }
      if (filtros.dataFim && tData) {
        const dFim = new Date(filtros.dataFim);
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

    return JSON.stringify({
      saldo: ent - sai,
      entradas: ent,
      saidas: sai,
      historico: dadosFiltrados.reverse(),
      isSuperAdmin: isMaster // Retorna isMaster real para o Front-end
    });

  } catch (e) {
    console.error("Erro Admin Fin:", e);
    return JSON.stringify({ saldo: 0, entradas: 0, saidas: 0, historico: [], isSuperAdmin: false });
  }
}

/**
 * 🛠️ HELPER DE FORMULÁRIO BLINDADO
 * Se não for master, o Select de Academias para Registrar Contas só mostra a dele.
 */
function getDadosParaLancamento(loginSolicitante) {
  const auth = getPermissoesUsuario(loginSolicitante);
  const locaisRaw = lerTabelaDinamica("Locais_de_treino");

  let listaLocais = locaisRaw.map(l => String(l.nome_do_local).trim()).filter(n => n !== "");

  // 🛡️ Filtra locais
  if (!auth.isMaster) {
    listaLocais = listaLocais.filter(loc => auth.academias.some(myAcad => loc.toLowerCase().includes(myAcad)));
  }

  const alunosRaw = lerTabelaDinamica("cadastro_de_alunos");
  let listaAlunos = alunosRaw.filter(a => String(a.status).toLowerCase() === "ativo");

  // 🛡️ Filtra Alunos
  if (!auth.isMaster) {
    listaAlunos = listaAlunos.filter(a => auth.academias.some(myAcad => String(a.academia_vinculada).toLowerCase().includes(myAcad)));
  }

  listaAlunos = listaAlunos.map(a => ({ login: String(a.login).trim(), nome: String(a.nome_completo).trim() }))
    .sort((a, b) => a.nome.localeCompare(b.nome));

  return { locais: listaLocais, alunos: listaAlunos };
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
    let contatoResp = ""; // Fallback (Número do Mestre/Geral)
    const configApp = lerTabelaDinamica("Config_App");
    if (configApp && configApp.length > 0) {
      let linkLojaApp = String(configApp[0].link_loja || configApp[0].Link_Loja || "");
      if (linkLojaApp.includes("wa.me") || linkLojaApp.includes("api.whatsapp")) {
        contatoResp = linkLojaApp.replace(/\D/g, "");
      }
    }

    // Se o local tiver contato próprio, sobrepõe
    if (academiaAluno) {
      const localInfo = locais.find(l => String(l.nome_do_local).toLowerCase() === String(academiaAluno).toLowerCase());
      if (localInfo && localInfo.contato) {
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
 * 📦 GET PACOTES COM TRAVA DE FRANQUIA
 */
function getPacotesDisponiveis(login) {
  const auth = getPermissoesUsuario(login);
  const pacotes = lerTabelaDinamica("Fin_Pacotes");

  return pacotes.filter(p => {
    if (String(p.status_pacote).toLowerCase() !== "ativo") return false;
    const permitidas = String(p.academias_permitidas).toLowerCase();
    // 🛡️ TRAVA: Se o pacote é de outra franquia, esconde.
    if (permitidas.includes("todas") || auth.isMaster) return true;
    return auth.academias.some(myAcad => permitidas.includes(myAcad));
  });
}

/**
 * 📦 LISTAR PACOTES ADMIN (Com Trava de Franquia - Mostra Ativos e Inativos)
 */
function listarPacotesAdmin(login) {
  try {
    const auth = getPermissoesUsuario(login);
    const pacotes = lerTabelaDinamica("Fin_Pacotes");

    return pacotes.filter(p => {
      const permitidas = String(p.academias_permitidas || "Todas").toLowerCase();
      // O Master vê tudo. O franqueado vê os pacotes "Todas" (Matriz) e os específicos da unidade dele.
      return auth.isMaster || permitidas === "todas" || auth.academias.some(myAcad => permitidas.includes(myAcad));
    }).map(p => ({
      _linha: p._linha,
      nome_pacote: p.nome_pacote || "",
      valor_padrao: p.valor_padrao || 0,
      duracao_dias: p.duracao_dias || 30,
      academias_permitidas: p.academias_permitidas || "Todas",
      status_pacote: p.status_pacote || "Ativo",
      descricao: p.descricao || ""
    }));
  } catch (e) {
    registrarLogBlindado("ERRO", "listarPacotesAdmin", e.message);
    return [];
  }
}

/**
 * 💰 REGISTRAR MOVIMENTAÇÃO (Agora suporta Contas a Pagar/Receber Pendentes)
 */
function registrarMovimentacao(dados) {
  try {
    let dataParaSalvar = new Date();
    if (dados.data_registro && dados.data_registro !== "") {
      const p = String(dados.data_registro).split('-');
      if (p.length === 3) dataParaSalvar = new Date(p[0], p[1] - 1, p[2], 12, 0, 0);
    }

    const novaTransacao = {
      "ID_Transacao": "TRX-" + Date.now(),
      "Data_Registro": dataParaSalvar,
      "Tipo": dados.tipo,
      "Categoria": dados.categoria,
      "Descricao": dados.descricao,
      "Valor": dados.valor,
      "Forma_Pagto": dados.forma,
      "Responsavel": dados.responsavel,
      "Login_Aluno": dados.alunoLogin || "",
      "Academia_Ref": dados.academia || "Matriz",
      "Status": dados.status_pagamento || "Concluido", // 🛡️ LÊ DO NOVO CAMPO DO MODAL
      "Comprovante_Url": padronizarLinkDrive(dados.comprovante_url) || "",
      "Modalidade": dados.modalidade || "Geral"
    };

    salvarFinanceiroSeguro("Fin_Transacoes", novaTransacao);

    // 🛡️ REGRA DE NEGÓCIO: Só renova assinatura se o pagamento estiver CONCLUÍDO
    const statusLimpo = String(novaTransacao.Status).toLowerCase();
    if (statusLimpo.includes("conclu") && dados.tipo === "Receita" && (String(dados.categoria).toLowerCase().includes("mensalidade") || String(dados.categoria).toLowerCase().includes("pacote"))) {
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
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const normalize = (str) => String(str).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, "");

  const mapaColunas = {};
  headers.forEach((h, i) => { if (h) mapaColunas[normalize(h)] = i; });

  if (!idLinha || isNaN(idLinha)) {
    const novaLinha = new Array(headers.length).fill("");
    for (const [coluna, valor] of Object.entries(mapaDados)) {
      const idx = mapaColunas[normalize(coluna)];
      if (idx !== undefined) novaLinha[idx] = valor;
    }
    sheet.appendRow(novaLinha);
    return "✅ Registro criado com sucesso!";
  } else {
    const range = sheet.getRange(idLinha, 1, 1, headers.length);
    const dadosAtuais = range.getValues()[0];
    for (const [coluna, valor] of Object.entries(mapaDados)) {
      const idx = mapaColunas[normalize(coluna)];
      if (idx !== undefined) dadosAtuais[idx] = valor;
    }
    range.setValues([dadosAtuais]);
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
 * 🛡️ SISTEMA DE LOG MILITAR - REGISTRO DE ALTA PRECISÃO
 * @param {string} nível - INFO, SUCESSO, ALERTA, ERRO, CRÍTICO
 * @param {string} ação - O que foi feito (ex: "Acesso ao Admin", "Salvar Aluno")
 * @param {string} detalhes - Dados técnicos ou mensagem de erro
 */
function registrarLogBlindado(nivel, acao, detalhes) {
  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    let abaLogs = planilha.getSheetByName("Logs_Auditoria");

    // Auto-setup da aba de logs se não existir
    if (!abaLogs) {
      abaLogs = planilha.insertSheet("Logs_Auditoria");
      abaLogs.appendRow(["Data/Hora", "Usuário", "Nível", "Ação", "Detalhes", "Página"]);
    }

    const usuario = Session.getActiveUser().getEmail() || "Sistema/LocalStorage";
    const carimbo = new Date();

    abaLogs.appendRow([carimbo, usuario, nivel, acao, detalhes, "Backend"]);

    // Se o erro for Crítico, console.error para o Log do Google Cloud
    if (nivel === "CRÍTICO" || nivel === "ERRO") {
      console.error(`🚨 [${acao}] - ${detalhes}`);
    }
  } catch (e) {
    // Fallback silencioso apenas se a própria gravação de log falhar
    console.warn("Falha catastrófica no motor de log: " + e.message);
  }
}

/**
 * 🛡️ SISTEMA DE LOG MILITAR - REGISTRO DE ALTA PRECISÃO
 * @param {string} nível - INFO, SUCESSO, ALERTA, ERRO, CRÍTICO
 * @param {string} ação - O que foi feito (ex: "Acesso ao Admin", "Salvar Aluno")
 * @param {string} detalhes - Dados técnicos ou mensagem de erro
 */
function registrarLogAuditoria(nivel, acao, detalhes) {
  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    let abaLogs = planilha.getSheetByName("Logs_Auditoria");

    // Auto-setup da aba de logs se não existir
    if (!abaLogs) {
      abaLogs = planilha.insertSheet("Logs_Auditoria");
      abaLogs.appendRow(["Data/Hora", "Usuário", "Nível", "Ação", "Detalhes", "Página"]);
    }

    const usuario = Session.getActiveUser().getEmail() || "Sistema/LocalStorage";
    const carimbo = new Date();

    abaLogs.appendRow([carimbo, usuario, nivel, acao, detalhes, "Backend"]);

    // Se o erro for Crítico, console.error para o Log do Google Cloud
    if (nivel === "CRÍTICO" || nivel === "ERRO") {
      console.error(`🚨 [${acao}] - ${detalhes}`);
    }
  } catch (e) {
    // Fallback silencioso apenas se a própria gravação de log falhar
    console.warn("Falha catastrófica no motor de log: " + e.message);
  }
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

function doGet(e) {
  try {
    // 1. Auditoria de Boot: Verifica integridade das tabelas antes de servir a página
    verificarCriarAbasSistema();

    const page = e.parameter.page || 'login';
    const routes = {
      'agendar': 'Agendar', 'cursos': 'Cursos', 'dashboard': 'Dashboard',
      'locais': 'Locais_treino', 'instrutor': 'Instrutor', 'recuperar': 'RecuperarSenha',
      'login': 'Login', 'ajuda': 'Ajuda'
    };

    // Define o arquivo ou força Login se a rota for inválida
    let htmlFile = routes[page] || 'Login';

    // 2. Blindagem de Carregamento: Verifica se o arquivo físico existe no projeto
    let template;
    try {
      template = HtmlService.createTemplateFromFile(htmlFile);
    } catch (e) {
      registrarLogBlindado("ERRO", "Router", `Tentativa de acesso a arquivo inexistente: ${htmlFile}`);
      template = HtmlService.createTemplateFromFile('Login'); // Fallback de segurança
    }

    // 3. Injeção de Variáveis com Fallback Garantido (Padrão White-Label)
    template.scriptUrl = SCRIPT_URL;
    const config = getAppConfig(); // Puxa do Banco de Dados

    template.bgUrl = getDriveImageUrl(config.bgId) || DEFAULT_ASSETS.BACKGROUND_ID;
    template.logoUrl = getDriveImageUrl(config.iconId) || DEFAULT_ASSETS.ICON_ID;
    template.primaryColor = config.primaryColor || "#FFD700";
    template.secondaryColor = config.secondaryColor || "#1e1e1e";
    template.textColor = config.textColor || "#ffffff";
    template.btnTextColor = config.btnTextColor || "#000000";
    template.bgColor = config.bgColor || "#121212";

    template.linkLoja = config.linkLoja || DEFAULT_ASSETS.LINK_LOJA;
    template.linkInstagram = config.linkInstagram || DEFAULT_ASSETS.LINK_INSTA;
    template.linkYouTube = config.linkYouTube || DEFAULT_ASSETS.LINK_YT;
    template.linkCadastro = config.linkCadastro || DEFAULT_ASSETS.LINK_CAD;

    const tituloSaaS = config.nomeAcademia || "DojoManager SaaS";

    // 4. Log de Sucesso (Monitoramento Militar)
    sentinela("Router", "INFO", `Página servida: ${htmlFile} para o IP/Sessão atual.`);

    return template.evaluate()
      .setTitle(`${tituloSaaS} | Portal do Aluno`)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (err) {
    // 🚨 REGISTRO DE DESASTRE: Se o doGet falhar, o Log registra o motivo exato
    if (typeof registrarLogBlindado === 'function') {
      registrarLogBlindado("CRÍTICO", "Falha Fatal no Boot", err.message);
    }
    return HtmlService.createHtmlOutput(`
      <div style="font-family:sans-serif; text-align:center; padding:50px; background:#121212; color:white; height:100vh;">
        <h1 style="color:#ff5252;">⚠️ Sistema em Manutenção</h1>
        <p>Ocorreu um erro crítico na inicialização do portal.</p>
        <code style="background:#000; padding:10px; display:block; border:1px solid #333;">${err.message}</code>
        <br><button onclick="location.reload()" style="padding:10px 20px; cursor:pointer;">Tentar Novamente</button>
      </div>
    `);
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
/**
 * 🩺 LISTAR ALUNOS COM TRAVA DE FRANQUIA
 */
function listarAlunosAdmin(login) {
  try {
    const auth = getPermissoesUsuario(login);
    const alunos = lerTabelaDinamica("cadastro_de_alunos");
    const chamadas = lerTabelaDinamica("Registro_Chamada");

    // 🛡️ TRAVA: Filtra antes de processar
    const alunosFiltrados = alunos.filter(a => {
      const acadAluno = String(a.academia_vinculada).toLowerCase();
      return auth.isMaster || auth.academias.some(myAcad => acadAluno.includes(myAcad));
    });

    return alunosFiltrados.map(a => {
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
        turma: a.turma_vinculada || "Sem Turma",
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
 * 🛠️ BACKEND: BUSCA DADOS DE PAGAMENTO PARA O MODAL DO ALUNO
 * BLINDAGEM: Sanitizado para White-Label (Sem Hardcode de ZAP)
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

    // 🛡️ SISTEMA DE FALLBACK DINÂMICO DO WHATSAPP (SEM HARDCODE)
    let zapGlobal = "";
    let linkLojaApp = String(configApp.link_loja || configApp.Link_Loja || "");
    if (linkLojaApp.includes("wa.me") || linkLojaApp.includes("api.whatsapp")) {
      zapGlobal = linkLojaApp.replace(/\D/g, "");
    }

    const localResp = locais.find(l => String(l.nome_do_local).toLowerCase() === String(aluno.academia_vinculada).toLowerCase());

    // Se a academia tem telefone, usa ele. Se não, tenta o Zap Global. Se nenhum existir, retorna vazio.
    info.contatoWhatsApp = (localResp && localResp.contato) ? String(localResp.contato).replace(/\D/g, "") : zapGlobal;

    return info;
  } catch (e) {
    console.error("Erro fatal no servidor ao buscar Dados Pagamento: " + e.message);
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
    const idLinha = parseInt(form.aluno_id);
    const dadosParaSalvar = {
      "Nome Completo": form.aluno_nome,
      "CPF": form.aluno_cpf,
      "Data de Nascimento": form.aluno_nasc,
      "Telefone": form.aluno_tel,
      "Nome do Pai": form.aluno_pai,
      "Nome da Mãe": form.aluno_mae,
      "Endereço": form.aluno_end,
      "Academia Vinculada": form.aluno_acad,
      "Turma Vinculada": form.aluno_turma,
      "LOGIN": form.aluno_login,
      "Senha": form.aluno_senha,
      "Endereço de e-mail": form.aluno_email,
      "STATUS": form.aluno_status,
      "Nível do Praticante": form.aluno_nivel,
      "NivelAdministrativo": form.aluno_nivel,
      "GRADUACAO_ATUAL": form.aluno_grad,
      "PROX_GRADUACAO": form.aluno_prox_grad,
      "Data Próximo Exame": form.aluno_exame,
      "Data Ultima Carteirinha": form.aluno_data_cart,
      "Foto 3x4 (para a carteirinha)": form.aluno_foto,
      "Modalidade": form.aluno_modalidade
    };

    const res = salvarDadosSeguro(NOME_ABA_ALUNOS, dadosParaSalvar, isNaN(idLinha) ? null : idLinha);

    // Sincroniza Assinatura se houver pacote, PRESERVANDO a Data de Início Original
    if (form.aluno_pacote) {
      const assinaturas = lerTabelaDinamica("Fin_Assinaturas");
      const login = String(form.aluno_login).toLowerCase().trim();
      const aAtual = assinaturas.find(a => String(a.login_aluno || a["Login_Aluno"]).toLowerCase().trim() === login);

      const dadosAssinatura = {
        "Login_Aluno": form.aluno_login,
        "Pacote_Atual": form.aluno_pacote,
        "Data_Fim": form.aluno_vencimento,
        "Status_Assinatura": form.aluno_status_assinatura
      };

      // 🛡️ MÁGICA AQUI: Se a assinatura já existia, recupera a data inicial dela. Senão, põe a de hoje.
      if (aAtual && (aAtual.data_inicio || aAtual["Data_Inicio"])) {
        dadosAssinatura["Data_Inicio"] = formatDate(aAtual.data_inicio || aAtual["Data_Inicio"]);
      } else {
        dadosAssinatura["Data_Inicio"] = formatDate(new Date());
      }

      salvarDadosSeguro("Fin_Assinaturas", dadosAssinatura, aAtual ? aAtual._linha : null);
    }
    return res;
  } catch (e) {
    return "❌ Erro: " + e.message;
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

/**
 * 🌍 LISTAR LOCAIS/ACADEMIAS COM TRAVA DE FRANQUIA
 */
function listarLocaisAdmin(login) {
  try {
    const auth = getPermissoesUsuario(login);
    const data = lerTabelaDinamica("Locais_de_treino");

    // 🛡️ TRAVA
    const locaisFiltrados = data.filter(l => {
      const nomeLocal = String(l.nome_do_local).toLowerCase();
      return auth.isMaster || auth.academias.some(myAcad => nomeLocal.includes(myAcad));
    });

    return locaisFiltrados.map(row => ({
      id: row._linha, nome: row.nome_do_local || "", endereco: row['endereço'] || "",
      cidade: row['cidade/estado'] || "", contato: row.contato || "", linkMaps: row.link_google_maps || "",
      iframeHtml: row.html_mapa_off_lline || "", responsavel: row.responsavel || "", status: row.status || "Ativo",
      pixChave: row.pix_chave_local_de_treino || "", pixNome: row.pix_nome_local_de_treino || "",
      pixBanco: row.banco_local_de_treino || "", pixCidade: row.pix_cidade_local_de_treino || ""
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
      "Pix_Nome_Local_de_Treino": form.loc_pix_nome,   // <-- Mapeado Novo Cabeçalho
      "Banco_Local_de_Treino": form.loc_pix_banco,     // <-- Mapeado Novo Cabeçalho
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

/**
 * 🎬 SALVAR VÍDEO (Agora mapeia Modalidade)
 */
function salvarVideo(form) {
  try {
    let link = form.vid_link; if (link.includes("v=")) link = link.split("v=")[1].split("&")[0];
    const dados = {
      "Faixa": form.vid_faixa,
      "Modalidade": form.vid_mod || "Geral", // <-- INJETADO
      "Titulo": form.vid_titulo,
      "Descricao": form.vid_desc,
      "Youtube_Link": link,
      "Status": form.vid_status
    };
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

/**
 * 📦 SALVAR PACOTE (CRUD Completo e Blindado)
 */
function salvarPacote(form) {
  try {
    const dados = {
      "Nome_Pacote": form.pac_nome,
      "Valor_Padrao": String(form.pac_valor).replace(',', '.'),
      "Duracao_Dias": form.pac_dias,
      "Academias_Permitidas": form.pac_academias || "TODAS",
      "Status_Pacote": form.pac_status || "Ativo",
      "Descricao": form.pac_desc || ""
    };
    const idLinha = parseInt(form.linha_id); // Puxa do hidden input do AdmModais
    const res = salvarDadosSeguro("Fin_Pacotes", dados, isNaN(idLinha) ? null : idLinha);
    return res;
  } catch (e) {
    return "❌ Erro ao salvar pacote: " + e.message;
  }
}

/**
 * 📄 SALVAR PROGRAMA TÉCNICO (Agora mapeia Modalidade)
 */
function salvarProgramaTecnico(form) {
  try {
    let idArq = "";
    if (form.prog_link) { const m = form.prog_link.match(/id=([a-zA-Z0-9_-]+)/); if (m) idArq = m[1]; }
    const dados = {
      "Faixa": form.prog_faixa,
      "Modalidade": form.prog_mod || "Geral", // <-- INJETADO
      "ID_Arquivo": idArq,
      "Link_Original": padronizarLinkDrive(form.prog_link),
      "Descricao": form.prog_desc
    };
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
      "Pix_Chave_Global": form.cfg_pix_chave,       // <-- Mapeado Novo Cabeçalho
      "Pix_Nome": form.cfg_pix_nome,
      "Nome_Banco_PIX_Global": form.cfg_pix_banco,  // <-- Mapeado Novo Cabeçalho
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
    const sheet = getSheet("cadastro_de_alunos");
    const data = sheet.getDataRange().getDisplayValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());

    const col = {
      login: headers.indexOf("login"),
      senha: headers.indexOf("senha"),
      nome: headers.indexOf("nome completo"),
      nivelTatame: headers.indexOf("nível do praticante") !== -1 ? headers.indexOf("nível do praticante") : headers.indexOf("nivel do praticante"),
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

        const roleTatame = String((col.nivelTatame > -1 && row[col.nivelTatame]) ? row[col.nivelTatame] : gradInfo.nivel).toUpperCase();
        const roleOffice = String(col.nivelAdm > -1 ? row[col.nivelAdm] : "").toUpperCase();

        const isInstrutor = roleTatame.includes("INSTRUTOR") || roleTatame.includes("PROFESSOR") || roleTatame.includes("MESTRE");
        const isMestre = roleTatame.includes("MESTRE");
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
          LOGIN: String(row[col.login]).trim(),
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
          isAdmin: isAdmin
        };

        // ====================================================================
        // 🛑 CATRACA VIRTUAL: BLOQUEIO DE INADIMPLENTES (Apenas Alunos)
        // ====================================================================
        if (!isCheckOnly && !isInstrutor && !isMestre && !isAdmin) {
          try {
            const assinaturas = lerTabelaDinamica("Fin_Assinaturas");
            const assDoAluno = assinaturas.find(a => String(a.login_aluno).toLowerCase() === loginInput);

            // Se não tem assinatura ou não está ativo
            if (!assDoAluno || String(assDoAluno.status_assinatura).toLowerCase() !== "ativo") {
              if (!isCheckOnly) registrarLogLogin(loginInput, "BLOQUEIO_FINANCEIRO_S_PLANO");
              return { success: false, message: "🚫 Assinatura inativa ou não localizada. Procure a recepção." };
            }

            const dataVencimento = parseDataSegura(assDoAluno.data_fim);
            const hoje = new Date();
            hoje.setHours(0, 0, 0, 0);

            // Se a data de validade já passou
            if (dataVencimento && dataVencimento < hoje) {
              if (!isCheckOnly) registrarLogLogin(loginInput, "BLOQUEIO_FINANCEIRO_VENCIDO");

              const dataVencStr = ('0' + dataVencimento.getDate()).slice(-2) + '/' + ('0' + (dataVencimento.getMonth() + 1)).slice(-2) + '/' + dataVencimento.getFullYear();

              return {
                success: false,
                message: `🚫 ACESSO BLOQUEADO!\nSua mensalidade venceu dia ${dataVencStr}.\nProcure a recepção ou faça o pagamento para liberar seu acesso.`
              };
            }
          } catch (e) {
            // Falha silenciosa se houver erro na aba financeira (não bloqueia por erro sistêmico)
            console.error("Erro na leitura da Catraca Virtual:", e);
          }
        }
        // ====================================================================

        if (!isCheckOnly) registrarLogLogin(loginInput, "SUCESSO");
        return { success: true, user: userPayload, redirectUrl: SCRIPT_URL + "?page=dashboard" };
      }
    }
    return { success: false, message: "Usuário não encontrado." };
  } catch (e) { return { success: false, message: e.message }; }
}
/**
 * 🛡️ A FUNÇÃO QUE ESTAVA FALTANDO NO SEU ARQUIVO: getListaAcademias
 * Essencial para o Instrutor/Aluno escolher a unidade correta no dropdown.
 */
function getListaAcademias(login) {
  try {
    const auth = getPermissoesUsuario(login);
    const data = lerTabelaDinamica("Locais_de_treino");

    return data.filter(l => {
      const nomeLocal = String(l.nome_do_local).toLowerCase();
      if (String(l.status).toLowerCase() !== "ativo") return false;
      return auth.isMaster || auth.academias.some(myAcad => nomeLocal.includes(myAcad));
    }).map(l => l.nome_do_local).filter(n => n && n !== "");
  } catch (e) {
    return [];
  }
}

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
 * ============================================================================
 * 🛡️ MOTOR SRE: INICIAR AULA AO VIVO (Com Conteúdo Temporário)
 * ============================================================================
 */
function iniciarAulaDinamica(dados) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Aulas_Em_Andamento");
    if (!sheet) throw new Error("Aba Aulas_Em_Andamento não encontrada.");

    const idAula = "AULA-" + new Date().getTime();
    const pin = Math.floor(1000 + Math.random() * 9000).toString(); // Gera PIN de 4 dígitos

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const novaLinha = new Array(headers.length).fill("");

    const mapDados = {
      "ID_Aula": idAula,
      "Data_Aula": dados.data,
      "Hora_Inicio": dados.inicio,
      "Hora_Fim": dados.fim,
      "Academia": dados.local,
      "Turma": dados.turma,
      "Instrutor": dados.instrutor,
      "PIN": pin,
      "Status": "ABERTA",
      "Checkins_JSON": "[]",
      "Conteudo": dados.conteudo || "Conteúdo não especificado" // 🛡️ GUARDA O CONTEÚDO NA SALA DE ESPERA
    };

    headers.forEach((h, i) => {
      if (mapDados[h] !== undefined) novaLinha[i] = mapDados[h];
    });

    sheet.appendRow(novaLinha);
    return { success: true, idAula: idAula, pin: pin };
  } catch (e) {
    registrarLogBlindado("ERRO", "iniciarAulaDinamica", e.message);
    return { success: false, erro: e.message };
  }
}

/**
 * 📝 QA Note: Busca as aulas ativas para o aluno fazer check-in.
 * Mapeia dinamicamente os cabeçalhos para evitar quebra de índice.
 * 🛡️ FIX: Comparação de datas blindada (Dia, Mês e Ano).
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

    // Pegando Dia, Mês e Ano exatos de hoje no servidor
    const hojeAno = agora.getFullYear();
    const hojeMes = agora.getMonth();
    const hojeDia = agora.getDate();

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
        let dObj = parseDataSegura(diaPlanilha); // Usa nosso parser indestrutível

        // Se a data for válida, compara as partes essenciais
        if (dObj && dObj.getFullYear() === hojeAno && dObj.getMonth() === hojeMes && dObj.getDate() === hojeDia) {
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
  } catch (e) {
    registrarLogBlindado("ERRO", "buscarAulasDisponiveisAluno", e.message);
    return [];
  }
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

// 🛡️ MOTOR DE SALVAMENTO MANUAL (TAMBÉM BLINDADO)
function salvarChamada(dados) {
  try {
    const idChamada = "AULA-" + new Date().getTime();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro_Chamada");
    const nomes = dados.alunos.map(a => a.nome).join(", ");
    const ids = dados.alunos.map(a => a.login).join(", ");
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const novaLinha = new Array(headers.length).fill("");

    // 🛡️ Limpeza de Hora no modo Manual
    let horaLimpa = String(dados.horaTreino || "--:--");
    let match = horaLimpa.match(/\b([01]\d|2[0-3]):([0-5]\d)\b/);
    if (match) horaLimpa = match[0]; else horaLimpa = horaLimpa.substring(0, 5);

    const mapDados = {
      "ID_Chamada": idChamada, "Data_Registro": new Date(), "Data_Treino": dados.dataTreino, "Hora_Treino": horaLimpa,
      "Instrutor_Logado": dados.instrutor, "Local_Treino": dados.local, "Qtd_Presentes": dados.alunos.length, "Lista_Alunos_IDs": ids,
      "Lista_Nomes": nomes, "Observacoes": "Chamada Manual", "Conteudo": dados.conteudo || "Conteúdo não especificado"
    };

    headers.forEach((h, i) => { if (mapDados[h] !== undefined) novaLinha[i] = mapDados[h]; });
    sheet.appendRow(novaLinha);
    return { success: true, message: "Chamada salva com sucesso! Conteúdo registrado." };
  } catch (e) { throw new Error("Erro ao salvar chamada: " + e.message); }
}

/**
 * ============================================================================
 * 🛡️ MOTOR SRE: FINALIZAR AULA AO VIVO E SALVAR (Transferindo o Conteúdo)
 * ============================================================================
 */
// 🛡️ MOTOR DE SALVAMENTO DE AULA AO VIVO (COM BLINDAGEM CONTRA O "SAT D")
function finalizarAulaDinamica(idAula, listaAlunos) {
  try {
    const sheetAtivas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Aulas_Em_Andamento");
    const sheetRegistro = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro_Chamada");

    const dataAtivas = sheetAtivas.getDataRange().getValues();
    const headersAtivas = dataAtivas[0];
    const idxId = headersAtivas.indexOf("ID_Aula");
    const idxStatus = headersAtivas.indexOf("Status");
    const idxConteudo = headersAtivas.indexOf("Conteudo");
    const idxHora = headersAtivas.indexOf("Hora_Inicio"); // Pegando a hora original

    let rowIdx = -1; let aulaDado = null;

    for (let i = 1; i < dataAtivas.length; i++) {
      if (String(dataAtivas[i][idxId]) === String(idAula)) { rowIdx = i + 1; aulaDado = dataAtivas[i]; break; }
    }

    if (rowIdx === -1) throw new Error("Aula não encontrada.");

    sheetAtivas.getRange(rowIdx, idxStatus + 1).setValue("ENCERRADA");

    const confirmados = listaAlunos.filter(a => a.status === 'Confirmado');
    const nomes = confirmados.map(a => a.nome).join(", ");
    const ids = confirmados.map(a => a.login).join(", ");
    const localFormatado = aulaDado[headersAtivas.indexOf("Academia")] + " | " + aulaDado[headersAtivas.indexOf("Turma")];
    let conteudoSalvo = "Conteúdo não especificado";
    if (idxConteudo !== -1 && aulaDado[idxConteudo]) conteudoSalvo = aulaDado[idxConteudo];

    // 🛡️ A EXTERMINAÇÃO DO "SAT D": Traduzindo a data bruta do Google para Hora limpa
    let horaRaw = aulaDado[idxHora];
    let horaLimpa = "--:--";
    if (horaRaw) {
      if (horaRaw instanceof Date) {
        horaLimpa = Utilities.formatDate(horaRaw, Session.getScriptTimeZone(), "HH:mm");
      } else {
        let str = String(horaRaw);
        let match = str.match(/\b([01]\d|2[0-3]):([0-5]\d)\b/); // Busca só o padrão HH:MM
        if (match) horaLimpa = match[0];
        else horaLimpa = str.substring(0, 5);
      }
    }

    const headersReg = sheetRegistro.getRange(1, 1, 1, sheetRegistro.getLastColumn()).getValues()[0];
    const novaLinha = new Array(headersReg.length).fill("");

    const mapDados = {
      "ID_Chamada": idAula, "Data_Registro": new Date(), "Data_Treino": aulaDado[headersAtivas.indexOf("Data_Aula")],
      "Hora_Treino": horaLimpa, "Instrutor_Logado": aulaDado[headersAtivas.indexOf("Instrutor")],
      "Local_Treino": localFormatado, "Qtd_Presentes": confirmados.length, "Lista_Alunos_IDs": ids, "Lista_Nomes": nomes,
      "Observacoes": "Chamada Dinâmica App", "Conteudo": conteudoSalvo
    };

    headersReg.forEach((h, i) => { if (mapDados[h] !== undefined) novaLinha[i] = mapDados[h]; });
    sheetRegistro.appendRow(novaLinha);

    return { success: true, msg: "Aula encerrada com sucesso! (" + confirmados.length + " alunos presentes)" };
  } catch (e) { throw new Error("Falha ao finalizar: " + e.message); }
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
 * 🛠️ BACKEND: BUSCA DADOS PIX PARA GERADOR DE QR CODE
 * BLINDAGEM: Lê os cabeçalhos exatos da planilha e limpa espaços invisíveis
 * SANITIZADO: White-Label (Sem Hardcode de WhatsApp)
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
      chave: configApp.pix_chave_global ? String(configApp.pix_chave_global).trim() : "",
      beneficiario: configApp.pix_nome || "Dojo Manager",
      cidade: configApp.pix_cidade || "RECIFE",
      banco: configApp.nome_banco_pix_global || "Instituição Bancária",
      valor: pacote ? pacote.valor_padrao : 0,
      contatoWhatsApp: ""
    };

    // 2. Se o PIX Global estiver DESATIVADO, sobrepõe com a CONTA LOCAL
    if (String(configApp.pix_global_ativo).toLowerCase() !== "sim") {
      const local = locais.find(l => String(l.nome_do_local).toLowerCase() === String(aluno.academia_vinculada).toLowerCase());

      if (local && String(local.status).toLowerCase() === "ativo") {
        info.chave = local.pix_chave_local_de_treino ? String(local.pix_chave_local_de_treino).trim() : info.chave;
        info.beneficiario = local.pix_nome_local_de_treino || info.beneficiario;
        info.cidade = local.pix_cidade_local_de_treino || info.cidade;
        info.banco = local.banco_local_de_treino || info.banco;
      }
    }

    // 🛡️ SISTEMA DE FALLBACK DINÂMICO DO WHATSAPP (SEM HARDCODE)
    let zapGlobal = "";
    let linkLojaApp = String(configApp.link_loja || configApp.Link_Loja || "");
    if (linkLojaApp.includes("wa.me") || linkLojaApp.includes("api.whatsapp")) {
      zapGlobal = linkLojaApp.replace(/\D/g, "");
    }

    const localResp = locais.find(l => String(l.nome_do_local).toLowerCase() === String(aluno.academia_vinculada).toLowerCase());

    // Se a academia tem telefone, usa ele. Se não, tenta o Zap Global. Se nenhum existir, retorna vazio.
    info.contatoWhatsApp = (localResp && localResp.contato) ? String(localResp.contato).replace(/\D/g, "") : zapGlobal;

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
    template.cpf = alunoInfo["cpf"] || "NÃO INFORMADO";
    template.graduacao = alunoInfo["graduacao_atual"] || "NÃO INFORMADO";
    template.academia = alunoInfo["academia vinculada"] || "NÃO INFORMADO";

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

/**
 * 💳 CRUD: Lista todas as Formas de Pagamento para a tabela Administrativa
 * 🛡️ BLINDADA COM TRAVA DE FRANQUIA
 */
function listarFormasPagamentoAdmin(login) {
  try {
    const auth = getPermissoesUsuario(login);
    const formasPagamento = lerTabelaDinamica("Forma_Pagamento");

    return formasPagamento.filter(f => {
      const pagLocal = String(f.local || "Todas").toLowerCase();
      // O Master vê tudo. O franqueado vê as formas "Todas" e as específicas da unidade dele.
      return auth.isMaster || pagLocal === "todas" || auth.academias.some(myAcad => pagLocal.includes(myAcad));
    }).map(f => ({
      id: f._linha,
      nome: f.nome || "Sem Nome",
      local: f.local || "Todas",
      status: f.status || "Ativo"
    }));
  } catch (e) {
    registrarLogBlindado("ERRO", "listarFormasPagamentoAdmin", e.message);
    return [];
  }
}

/**
 * 💳 CRUD: Salva ou Edita Formas de Pagamento
 */
function salvarFormaPagamento(form) {
  try {
    const dados = {
      "Nome": form.pag_nome,
      "Local": form.pag_local,
      "Status": form.pag_status || "Ativo"
    };
    const idLinha = parseInt(form.pag_id);
    const res = salvarDadosSeguro("Forma_Pagamento", dados, isNaN(idLinha) ? null : idLinha);
    registrarLogBlindado("SUCESSO", "salvarFormaPagamento", form.pag_nome);
    return res;
  } catch (e) {
    registrarLogBlindado("ERRO", "salvarFormaPagamento", e.message);
    return "❌ Erro: " + e.message;
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
 * 🏫 LISTAR TURMAS COM TRAVA DE FRANQUIA
 */
function listarTurmasAdmin(login) {
  try {
    const auth = getPermissoesUsuario(login);
    const data = lerTabelaDinamica("Config_Turmas");

    // 🛡️ TRAVA
    const turmasFiltradas = data.filter(t => {
      const acadTurma = String(t.local_vinculado).toLowerCase();
      return auth.isMaster || auth.academias.some(myAcad => acadTurma.includes(myAcad));
    });

    return turmasFiltradas.map(t => ({
      idLinha: t._linha, idTurma: t.id_turma || "", nome: t.nome_da_turma || "Sem Nome",
      modalidade: t.modalidade || "Geral", faixaEtaria: t["faixa_etária"] || t.faixa_etaria || "Misto",
      local: t.local_vinculado || "Matriz", dias: t.dias_da_semana || "",
      inicio: t["horário_início"] || t.horario_inicio || "", fim: t["horário_fim"] || t.horario_fim || "",
      status: t.status || "Ativa", responsavel: t.responsavel || t['responsável'] || "", telefone: t.telefone || ""
    }));
  } catch (e) { return []; }
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

/**
 * 📊 DASHBOARD DE INTELIGÊNCIA (SUPER RELATÓRIO) - BLINDADO
 */
function getEstatisticasRelatorio(login, filtros) {
  try {
    const ws = SpreadsheetApp.getActiveSpreadsheet();
    const abaAlunos = ws.getSheetByName("cadastro_de_alunos");
    if (!abaAlunos) throw new Error("Aba 'cadastro_de_alunos' não encontrada.");

    const dadosAlunos = abaAlunos.getDataRange().getValues();
    if (dadosAlunos.length < 2) return { success: true, kpiAtivos: 0, kpiInadimplentes: 0, kpiReceitaPrevista: "0,00", kpiRecebido: "0,00", kpiPendente: "0,00", kpiDespesa: "0,00", listaCobranca: [], listaRanking: [] };

    const cabecalhoAlunosNorm = dadosAlunos[0].map(h => String(h).trim().toLowerCase().replace(/\s+/g, ""));
    const findColAluno = (name) => cabecalhoAlunosNorm.indexOf(String(name).trim().toLowerCase().replace(/\s+/g, ""));

    const idxLogin = findColAluno("LOGIN");
    const idxAcad = findColAluno("Academia Vinculada");
    const idxStatus = findColAluno("STATUS");
    const idxNome = findColAluno("Nome Completo");
    const idxTurma = findColAluno("Turma Vinculada");
    const idxTel = findColAluno("Telefone");

    let userAcads = [];
    let isMaster = false;
    let usuarioEncontrado = false;
    const loginBuscado = String(login).trim().toLowerCase();

    // VALIDAÇÃO CORTANTE DA ACADEMIA
    for (let i = 1; i < dadosAlunos.length; i++) {
      let loginPlanilha = idxLogin > -1 ? String(dadosAlunos[i][idxLogin]).trim().toLowerCase() : "";
      if (loginPlanilha === loginBuscado) {
        usuarioEncontrado = true;
        let academiasStr = idxAcad > -1 ? String(dadosAlunos[i][idxAcad]).toLowerCase() : "";
        userAcads = academiasStr.split(',').map(a => a.trim());

        // 🚨 REMOVIDA A TRAVA DE "ADMIN/DIRETOR". AGORA SÓ "TODAS" DÁ PODER DE DEUS!
        if (userAcads.includes("todas")) {
          isMaster = true;
        }
        break;
      }
    }

    if (!usuarioEncontrado || (userAcads.length === 0 && !isMaster)) {
      return { success: true, kpiAtivos: 0, kpiInadimplentes: 0, kpiReceitaPrevista: "0,00", kpiRecebido: "0,00", kpiPendente: "0,00", kpiDespesa: "0,00", listaCobranca: [], listaRanking: [] };
    }

    // 2. FILTROS DA TELA
    const targetAcad = (filtros && filtros.academia && filtros.academia !== "TODAS") ? String(filtros.academia).trim().toLowerCase() : "";
    const dataIniStr = (filtros && filtros.dataInicio) ? filtros.dataInicio : "";
    const dataFimStr = (filtros && filtros.dataFim) ? filtros.dataFim : "";
    const dIni = dataIniStr ? new Date(dataIniStr + "T00:00:00") : new Date(2000, 0, 1);
    const dFim = dataFimStr ? new Date(dataFimStr + "T23:59:59") : new Date(2100, 0, 1);
    const hoje = new Date(); hoje.setHours(0, 0, 0, 0);

    // 3. MAPEAR PACOTES E ASSINATURAS... (Permanece igual)
    const abaPacotes = ws.getSheetByName("Fin_Pacotes");
    const mapPacotes = {};
    if (abaPacotes) {
      const dadosPacotes = abaPacotes.getDataRange().getValues();
      if (dadosPacotes.length > 1) {
        const hPacotes = dadosPacotes[0].map(h => String(h).trim().toLowerCase().replace(/\s+/g, ""));
        const idxPacNome = hPacotes.indexOf("nome_pacote") > -1 ? hPacotes.indexOf("nome_pacote") : hPacotes.indexOf("nomedopacote");
        const idxPacValor = hPacotes.indexOf("valor_padrao") > -1 ? hPacotes.indexOf("valor_padrao") : hPacotes.indexOf("valorpadrão");
        if (idxPacNome > -1 && idxPacValor > -1) {
          for (let i = 1; i < dadosPacotes.length; i++) {
            let nome = String(dadosPacotes[i][idxPacNome]).trim().toLowerCase();
            let valor = parseFloat(String(dadosPacotes[i][idxPacValor]).replace(',', '.')) || 0;
            mapPacotes[nome] = valor;
          }
        }
      }
    }

    const abaAssinaturas = ws.getSheetByName("Fin_Assinaturas");
    const mapAssinaturas = {};
    if (abaAssinaturas) {
      const dadosAssinaturas = abaAssinaturas.getDataRange().getValues();
      if (dadosAssinaturas.length > 1) {
        const hAssinaturas = dadosAssinaturas[0].map(h => String(h).trim().toLowerCase().replace(/\s+/g, ""));
        const idxAssLogin = hAssinaturas.indexOf("login_aluno") > -1 ? hAssinaturas.indexOf("login_aluno") : hAssinaturas.indexOf("logindoaluno");
        const idxAssPacote = hAssinaturas.indexOf("pacote_atual") > -1 ? hAssinaturas.indexOf("pacote_atual") : hAssinaturas.indexOf("pacoteatual");
        const idxAssFim = hAssinaturas.indexOf("data_fim") > -1 ? hAssinaturas.indexOf("data_fim") : hAssinaturas.indexOf("datafim");
        const idxAssStatus = hAssinaturas.indexOf("status_assinatura") > -1 ? hAssinaturas.indexOf("status_assinatura") : hAssinaturas.indexOf("statusdaassinatura");

        if (idxAssLogin > -1 && idxAssPacote > -1 && idxAssFim > -1) {
          for (let i = 1; i < dadosAssinaturas.length; i++) {
            let l = String(dadosAssinaturas[i][idxAssLogin]).trim().toLowerCase();
            if (!l) continue;
            let v = dadosAssinaturas[i][idxAssFim];
            let dFimPacote = null;
            if (v instanceof Date) { dFimPacote = v; }
            else if (v && String(v).includes('/')) {
              let p = String(v).split('/');
              if (p.length === 3) dFimPacote = new Date(p[2], p[1] - 1, p[0]);
            } else if (v && String(v).includes('-')) {
              dFimPacote = new Date(String(v).split(' ')[0] + "T00:00:00");
            }
            mapAssinaturas[l] = {
              pacote: String(dadosAssinaturas[i][idxAssPacote]).trim().toLowerCase(),
              vencimento: dFimPacote,
              status: idxAssStatus > -1 ? String(dadosAssinaturas[i][idxAssStatus]).trim().toLowerCase() : "ativo"
            };
          }
        }
      }
    }

    let totalAtivos = 0; let totalInadimplentes = 0; let receitaPrevista = 0; let pagamentosPendentes = 0;
    let listaCobranca = []; let contagemTurmas = {};

    for (let i = 1; i < dadosAlunos.length; i++) {
      let statusAluno = idxStatus > -1 ? String(dadosAlunos[i][idxStatus]).trim().toLowerCase() : "ativo";
      if (statusAluno !== "ativo" && statusAluno !== "") continue;

      let aAcad = idxAcad > -1 ? String(dadosAlunos[i][idxAcad]).toLowerCase() : "";

      // 🛡️ A BARREIRA INTRANSPONÍVEL DA ACADEMIA
      if (!isMaster && !userAcads.some(myAcad => aAcad.includes(myAcad))) continue;
      if (targetAcad && !aAcad.includes(targetAcad)) continue;

      totalAtivos++;

      let l = idxLogin > -1 ? String(dadosAlunos[i][idxLogin]).trim().toLowerCase() : "";
      let nomeA = idxNome > -1 ? String(dadosAlunos[i][idxNome]).trim() : "Sem Nome";
      let turmaA = idxTurma > -1 ? String(dadosAlunos[i][idxTurma]).trim() : "Sem Turma";
      let acadA = idxAcad > -1 ? String(dadosAlunos[i][idxAcad]).trim() : "Sem Local";
      let telA = idxTel > -1 ? String(dadosAlunos[i][idxTel]).trim() : "";

      let chaveTurma = turmaA + "|||" + acadA;
      if (!contagemTurmas[chaveTurma]) contagemTurmas[chaveTurma] = { nome: turmaA, local: acadA, count: 0 };
      contagemTurmas[chaveTurma].count++;

      let ass = mapAssinaturas[l];
      let inadimplente = false; let valorPacote = 0; let pacoteNome = ""; let vencStr = "Sem Plano";

      if (ass) {
        pacoteNome = ass.pacote;
        valorPacote = mapPacotes[pacoteNome] || 0;
        if (ass.vencimento) {
          let v = ass.vencimento;
          vencStr = ('0' + v.getDate()).slice(-2) + '/' + ('0' + (v.getMonth() + 1)).slice(-2) + '/' + v.getFullYear();
          if (ass.status !== "ativo" || v < hoje) inadimplente = true;
        } else inadimplente = true;
      } else inadimplente = true;

      if (inadimplente) {
        if (acadA.toLowerCase() === "todas" || l.includes("master")) {
          totalAtivos--;
        } else {
          totalInadimplentes++; pagamentosPendentes += valorPacote;
          listaCobranca.push({ nome: nomeA, academia: acadA, turma: turmaA, telefone: telA, plano: pacoteNome, vencimento: vencStr });
        }
      } else {
        receitaPrevista += valorPacote;
      }
    }

    // PROCESSAR TRANSAÇÕES 
    let receitasRealizadas = 0; let despesasRealizadas = 0;
    const abaTransacoes = ws.getSheetByName("Fin_Transacoes");

    if (abaTransacoes) {
      const dadosTransacoes = abaTransacoes.getDataRange().getValues();
      if (dadosTransacoes.length > 1) {
        const hTrx = dadosTransacoes[0].map(h => String(h).trim().toLowerCase().replace(/\s+/g, ""));
        const idxTrxData = hTrx.indexOf("data_registro") > -1 ? hTrx.indexOf("data_registro") : hTrx.indexOf("dataderegistro");
        const idxTrxTipo = hTrx.indexOf("tipo"); const idxTrxValor = hTrx.indexOf("valor");
        const idxTrxAcad = hTrx.indexOf("academia_ref") > -1 ? hTrx.indexOf("academia_ref") : hTrx.indexOf("academiaref");
        const idxTrxStatus = hTrx.indexOf("status");

        if (idxTrxData > -1 && idxTrxValor > -1) {
          for (let i = 1; i < dadosTransacoes.length; i++) {
            let status = idxTrxStatus > -1 ? String(dadosTransacoes[i][idxTrxStatus]).trim().toLowerCase() : "concluido";
            if (status !== "concluido" && status !== "concluído") continue;

            let trxAcad = idxTrxAcad > -1 ? String(dadosTransacoes[i][idxTrxAcad]).toLowerCase() : "";

            // 🛡️ A BARREIRA NA TRANSAÇÃO
            if (!isMaster && !userAcads.some(myAcad => trxAcad.includes(myAcad))) continue;
            if (targetAcad && !trxAcad.includes(targetAcad)) continue;

            let dTrx = dadosTransacoes[i][idxTrxData];
            let dValid = null;
            if (dTrx instanceof Date) { dValid = dTrx; }
            else if (dTrx && String(dTrx).includes('/')) {
              let p = String(dTrx).split(' ')[0].split('/');
              if (p.length === 3) dValid = new Date(p[2], p[1] - 1, p[0]);
            } else if (dTrx && String(dTrx).includes('-')) {
              dValid = new Date(String(dTrx).split(' ')[0] + "T00:00:00");
            }

            if (dValid && dValid >= dIni && dValid <= dFim) {
              let tipo = idxTrxTipo > -1 ? String(dadosTransacoes[i][idxTrxTipo]).trim().toLowerCase() : "receita";
              let valor = parseFloat(String(dadosTransacoes[i][idxTrxValor]).replace(',', '.')) || 0;
              if (tipo === "receita") receitasRealizadas += valor;
              if (tipo === "despesa") despesasRealizadas += valor;
            }
          }
        }
      }
    }

    let turmasArr = Object.values(contagemTurmas).filter(t => t.local !== "todas" && t.local !== "" && t.local !== "Sem Local");
    turmasArr.sort((a, b) => b.count - a.count);

    return {
      success: true, kpiAtivos: totalAtivos, kpiInadimplentes: totalInadimplentes,
      kpiReceitaPrevista: receitaPrevista.toFixed(2).replace('.', ','),
      kpiRecebido: receitasRealizadas.toFixed(2).replace('.', ','),
      kpiPendente: pagamentosPendentes.toFixed(2).replace('.', ','),
      kpiDespesa: despesasRealizadas.toFixed(2).replace('.', ','),
      listaCobranca: listaCobranca, listaRanking: turmasArr,
      graficoFluxo: { labels: ["Período Filtrado"], receitas: [receitasRealizadas], despesas: [despesasRealizadas] },
      graficoOcupacao: { labels: turmasArr.slice(0, 5).map(t => t.nome), valores: turmasArr.slice(0, 5).map(t => t.count) }
    };
  } catch (e) { return { erro: e.message }; }
}

/**
 * Puxa os Locais Autorizados (Financeiro)
 */
function getLocaisParaFinanceiro(loginSolicitante) {
  const alunos = lerTabelaDinamica(NOME_ABA_ALUNOS);
  const user = alunos.find(a => String(a.login).toLowerCase() === String(loginSolicitante).toLowerCase());
  if (!user) return ["Matriz"];

  const acads = String(user.academia_vinculada || "Matriz").split(',').map(a => a.trim());
  if (acads.map(a => a.toLowerCase()).includes("todas")) {
    const locais = lerTabelaDinamica("Locais_de_treino");
    return locais.filter(l => String(l.status).toLowerCase() === 'ativo').map(l => String(l.nome_do_local).trim());
  }
  return acads;
}

function getDadosIniciaisFinanceiro() {
  try {
    const ws = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Busca Categorias
    const sCat = ws.getSheetByName("Categoria_financeira");
    if (!sCat) throw new Error("Aba 'Categoria_financeira' não encontrada.");
    const dCat = sCat.getDataRange().getValues();
    const mapCat = dCat[0].reduce((acc, col, i) => { acc[String(col).trim()] = i; return acc; }, {});
    let arrCat = [];
    for (let i = 1; i < dCat.length; i++) {
      let nomeCat = String(dCat[i][mapCat["Nome"]]).trim();
      if (!nomeCat) continue;
      arrCat.push({
        nome: nomeCat,
        tipo: String(dCat[i][mapCat["Tipo"]]).trim(),
        local: String(dCat[i][mapCat["Local"]] || "Todas").trim(),
        exibeAluno: String(dCat[i][mapCat["Exibe Aluno"]] || "Não").trim(),
        status: String(dCat[i][mapCat["Status"]] || "Ativo").trim()
      });
    }

    // 2. Busca Formas de Pagamento
    const sPag = ws.getSheetByName("Forma_Pagamento");
    if (!sPag) throw new Error("Aba 'Forma_Pagamento' não encontrada.");
    const dPag = sPag.getDataRange().getValues();
    const mapPag = dPag[0].reduce((acc, col, i) => { acc[String(col).trim()] = i; return acc; }, {});
    let arrPag = [];
    for (let i = 1; i < dPag.length; i++) {
      let nomePag = String(dPag[i][mapPag["Nome"]]).trim();
      if (!nomePag) continue;
      arrPag.push({
        nome: nomePag,
        local: String(dPag[i][mapPag["Local"]] || "Todas").trim(),
        status: String(dPag[i][mapPag["Status"]] || "Ativo").trim()
      });
    }

    // 3. Busca Locais de Treino (Unidades)
    const sLoc = ws.getSheetByName("Locais_de_treino");
    if (!sLoc) throw new Error("Aba 'Locais_de_treino' não encontrada.");
    const dLoc = sLoc.getDataRange().getValues();
    const idxLocNome = dLoc[0].indexOf("Nome do Local");
    const idxLocStatus = dLoc[0].indexOf("Status");
    let arrLocais = [];
    for (let i = 1; i < dLoc.length; i++) {
      let nomeLoc = String(dLoc[i][idxLocNome]).trim();
      let statusLoc = String(dLoc[i][idxLocStatus]).trim().toLowerCase();
      if (nomeLoc && statusLoc === "ativo") {
        arrLocais.push(nomeLoc);
      }
    }

    // 4. Busca Modalidades (Dinâmicas da aba GRADUACAO)
    const sGrad = ws.getSheetByName("GRADUACAO");
    if (!sGrad) throw new Error("Aba 'GRADUACAO' não encontrada.");
    const dGrad = sGrad.getDataRange().getValues();
    const idxGradMod = dGrad[0].indexOf("Modalidade");
    let setMods = new Set(); // Usamos Set para evitar repetições
    for (let i = 1; i < dGrad.length; i++) {
      let mod = String(dGrad[i][idxGradMod]).trim();
      if (mod) setMods.add(mod);
    }
    let arrMods = Array.from(setMods).sort(); // Transforma em Array e ordena em ordem alfabética

    // DEVOLVE TUDO PARA O FRONT-END DE UMA VEZ SÓ!
    return {
      categorias: arrCat,
      formasPagto: arrPag,
      locais: arrLocais,
      modalidades: arrMods
    };
  } catch (e) {
    return { erro: e.message };
  }
}

/**
 * ============================================================================
 * 📊 PONTE DE TRADUÇÃO DO DASHBOARD (FRONTEND <-> BACKEND)
 * Pega os dados brutos da getEstatisticasRelatorio e embala para o Chart.js e KPIs
 * ============================================================================
 */

function getDadosDashboardAdmin(filtros) {
  try {
    registrarLogBlindado("INFO", "DASH_PONTE", `Iniciando embalagem de dados para o Front...`);

    // 1. Puxa os dados da sua função mestre! Note que agora repassamos o login correto
    const stats = getEstatisticasRelatorio(filtros.usuario, filtros);

    if (stats.erro) throw new Error(stats.erro);

    registrarLogBlindado("SUCESSO", "DASH_PONTE", `Empacotando: ${stats.kpiAtivos} ativos, R$ ${stats.kpiReceitaPrevista}`);

    // 2. Formata e embala EXATAMENTE como o Front-end pediu, forçando o envio das chaves certas
    return {
      kpiAtivos: stats.kpiAtivos || 0,
      kpiInadimplentes: stats.kpiInadimplentes || 0,
      kpiReceitaPrevista: stats.kpiReceitaPrevista || "0,00",
      kpiRecebido: stats.kpiRecebido || "0,00",
      kpiPendente: stats.kpiPendente || "0,00",
      kpiDespesa: stats.kpiDespesa || "0,00",
      listaCobranca: stats.listaCobranca || [],
      listaRanking: stats.listaRanking || [],
      graficoFluxo: stats.graficoFluxo || { labels: ['Mês Filtrado'], receitas: [0], despesas: [0] },
      graficoOcupacao: stats.graficoOcupacao || { labels: [], valores: [] }
    };

  } catch (e) {
    registrarLogBlindado("ERRO", "DASH_PONTE", e.message);
    throw new Error("Falha ao embalar dados do Dashboard: " + e.message);
  }
}


/**
 * Busca categorias filtradas por Tipo e Local
 */
function getCategoriasFinanceirasCascata(tipo, local) {
  const categorias = lerTabelaDinamica("Categoria_financeira");
  return categorias.filter(c => {
    const matchTipo = String(c.tipo).toLowerCase() === String(tipo).toLowerCase();
    const matchStatus = String(c.status).toLowerCase() === 'ativo';
    const matchLocal = String(c.local).toLowerCase().includes('todas') ||
      String(c.local).toLowerCase().includes(String(local).toLowerCase());
    return matchTipo && matchStatus && matchLocal;
  });
}

/**
 * Busca Formas de Pagamento por Local
 */
function getFormasPagamentoCascata(local) {
  const formas = lerTabelaDinamica("Forma_Pagamento");
  return formas.filter(f => {
    const matchStatus = String(f.status).toLowerCase() === 'ativo';
    const matchLocal = String(f.local).toLowerCase().includes('todas') ||
      String(f.local).toLowerCase().includes(String(local).toLowerCase());
    return matchStatus && matchLocal;
  });
}

//CRUD financeiro

function listarCategoriasFinanceirasAdmin(login) {
  try {
    const auth = getPermissoesUsuario(login);
    const categorias = lerTabelaDinamica("Categoria_financeira");

    return categorias.filter(c => {
      const catLocal = String(c.local || "Todas").toLowerCase();
      // O Master vê tudo. O franqueado vê as categorias "Todas" (criadas pelo master) e as específicas da unidade dele.
      return auth.isMaster || catLocal === "todas" || auth.academias.some(myAcad => catLocal.includes(myAcad));
    }).map(c => ({
      id: c._linha, nome: c.nome || "", tipo: c.tipo || "Receita",
      local: c.local || "Todas", exibeAluno: c.exibe_aluno || "Não", status: c.status || "Ativo"
    }));
  } catch (e) {
    return [];
  }
}

function salvarCategoriaFinanceira(form) {
  const dados = {
    "Nome": form.cat_nome, "Tipo": form.cat_tipo, "Local": form.cat_local,
    "Exibe Aluno": form.cat_exibe, "Status": form.cat_status
  };
  return salvarDadosSeguro("Categoria_financeira", dados, parseInt(form.cat_id) || null);
}

/**
 * 🛡️ Função para receber logs do Frontend (Adm.html)
 */
function registrarLogFront(nivel, acao, detalhes) {
  registrarLogBlindado(nivel, acao, detalhes);
}

/**
 * 🛰️ PONTE DE SEGURANÇA (BACKEND)
 * Garante que chamadas feitas com nomes diferentes não quebrem o sistema.
 */
function sentinela(acao, nivel, detalhes) {
  registrarLogBlindado(nivel, acao, detalhes);
}

/**
 * ============================================================================
 * 🛡️ MOTOR SRE: HISTÓRICO DE CHAMADAS (MULTI-TENANCY)
 * Busca os registros de aula apenas para o instrutor logado.
 * ============================================================================
 */
function getHistoricoFiltrado(filtros) {
  try {
    const chamadas = lerTabelaDinamica("Registro_Chamada");
    if (!chamadas || chamadas.length === 0) return [];

    const hoje = new Date();

    // Tratamento seguro de datas
    let dIni = null; let dFim = null;
    if (filtros.inicio) dIni = new Date(filtros.inicio + "T00:00:00");
    if (filtros.fim) { dFim = new Date(filtros.fim + "T00:00:00"); dFim.setHours(23, 59, 59, 999); }

    const loginInstrutor = String(filtros.instrutor || "").trim().toLowerCase();

    // 🛡️ Obtém as credenciais de segurança
    const auth = getPermissoesUsuario(loginInstrutor);

    const historicoFiltrado = [];

    for (let i = 0; i < chamadas.length; i++) {
      const row = chamadas[i];
      if (!row.id_chamada) continue;

      const dataStr = String(row.data_treino || row.data_registro || "");
      let dataTreinoObj = parseDataSegura(dataStr);

      // Filtro de Data
      if (dIni && dataTreinoObj && dataTreinoObj < dIni) continue;
      if (dFim && dataTreinoObj && dataTreinoObj > dFim) continue;

      // Filtro de Local/Turma (Digitado na Tela)
      const localTreino = String(row.local_treino || "").toLowerCase();
      if (filtros.local && !localTreino.includes(String(filtros.local).toLowerCase())) continue;

      // 🛡️ Filtro de Franquia (O Cofre Multi-Tenancy)
      // Se não for Master, só vê aulas que contêm as academias dele OU aulas que ele mesmo deu.
      const donoDaAula = String(row.instrutor_logado || "").toLowerCase();

      let temPermissao = auth.isMaster;
      if (!temPermissao) {
        // O instrutor deu a aula?
        if (loginInstrutor && donoDaAula.includes(loginInstrutor)) temPermissao = true;
        // A aula ocorreu na franquia dele?
        else if (auth.academias.some(myAcad => localTreino.includes(myAcad))) temPermissao = true;
      }

      if (temPermissao) {
        historicoFiltrado.push({
          id: row.id_chamada,
          data: dataStr.includes('T') ? dataStr.split('T')[0] : dataStr,
          hora: String(row.hora_treino || ""),
          local: String(row.local_treino || "---"),
          instrutor: String(row.instrutor_logado || "---"),
          qtd: row.qtd_presentes || 0
        });
      }
    }

    return historicoFiltrado.reverse(); // Mais recentes primeiro

  } catch (e) {
    registrarLogBlindado("ERRO", "getHistoricoFiltrado", e.message);
    throw new Error("Falha no servidor ao buscar histórico: " + e.message);
  }
}

/**
 * 📄 BUSCAR DETALHES DE UMA CHAMADA ESPECÍFICA (Para gerar o PDF)
 */
function getDetalhesChamada(idChamada) {
  try {
    const chamadas = lerTabelaDinamica("Registro_Chamada");
    const reg = chamadas.find(c => String(c.id_chamada) === String(idChamada));

    if (!reg) return null;

    let dataFormatada = reg.data_treino || reg.data_registro;
    if (dataFormatada instanceof Date) {
      dataFormatada = Utilities.formatDate(dataFormatada, Session.getScriptTimeZone(), "dd/MM/yyyy");
    }

    return {
      id: reg.id_chamada,
      data: dataFormatada,
      hora: formatarHoraSimples(reg.hora_treino),
      local: reg.local_treino,
      instrutor: reg.instrutor_logado,
      listaNomes: reg.lista_nomes || "Ninguém marcou presença.",
      conteudo: reg.conteudo || reg.Conteudo || "Nenhum conteúdo especificado na aula." // 🛡️ CAMPO PUXADO DA PLANILHA AQUI
    };
  } catch (e) {
    registrarLogBlindado("ERRO", "getDetalhesChamada", e.message);
    return null;
  }
}


/**
 * Puxa os dados para alimentar as cascatas do modal financeiro
 */
function getDadosIniciaisCaixa() {
  try {
    const ws = SpreadsheetApp.getActiveSpreadsheet();

    // Busca Categorias (Tratando nomes de colunas com e sem underscore)
    const sCat = ws.getSheetByName("Categoria_financeira");
    const dCat = sCat ? sCat.getDataRange().getValues() : [];
    let arrCat = [];
    if (dCat.length > 1) {
      const hCat = dCat[0].map(h => String(h).trim().toLowerCase());
      const iNome = hCat.indexOf("nome"); const iTipo = hCat.indexOf("tipo");
      const iLoc = hCat.indexOf("local"); const iStat = hCat.indexOf("status");
      const iExibe = hCat.indexOf("exibe aluno") > -1 ? hCat.indexOf("exibe aluno") : hCat.indexOf("exibe_aluno");

      for (let i = 1; i < dCat.length; i++) {
        if (String(dCat[i][iNome]).trim() !== "") {
          arrCat.push({
            nome: String(dCat[i][iNome]).trim(),
            tipo: String(dCat[i][iTipo]).trim(),
            local: String(dCat[i][iLoc] || "Todas").trim(),
            exibeAluno: String(dCat[i][iExibe] || "Não").trim(),
            status: String(dCat[i][iStat] || "Ativo").trim()
          });
        }
      }
    }

    // Busca Formas de Pagamento
    const sPag = ws.getSheetByName("Forma_Pagamento");
    const dPag = sPag ? sPag.getDataRange().getValues() : [];
    let arrPag = [];
    if (dPag.length > 1) {
      const hPag = dPag[0].map(h => String(h).trim().toLowerCase());
      const iNomeP = hPag.indexOf("nome"); const iLocP = hPag.indexOf("local"); const iStatP = hPag.indexOf("status");

      for (let i = 1; i < dPag.length; i++) {
        if (String(dPag[i][iNomeP]).trim() !== "") {
          arrPag.push({
            nome: String(dPag[i][iNomeP]).trim(),
            local: String(dPag[i][iLocP] || "Todas").trim(),
            status: String(dPag[i][iStatP] || "Ativo").trim()
          });
        }
      }
    }

    // Busca Modalidades Dinâmicas (Aba GRADUACAO)
    const sGrad = ws.getSheetByName("GRADUACAO");
    let arrMods = ["Geral"]; // Default de segurança
    if (sGrad) {
      const dGrad = sGrad.getDataRange().getValues();
      const colMod = dGrad[0].map(h => String(h).trim().toLowerCase()).indexOf("modalidade");
      if (colMod > -1) {
        let setMods = new Set();
        for (let i = 1; i < dGrad.length; i++) {
          let mod = String(dGrad[i][colMod]).trim();
          if (mod) setMods.add(mod);
        }
        arrMods = Array.from(setMods).sort();
        if (!arrMods.includes("Geral")) arrMods.unshift("Geral");
      }
    }

    return { categorias: arrCat, formasPagto: arrPag, modalidades: arrMods };
  } catch (e) {
    registrarLogBlindado("ERRO", "getDadosIniciaisCaixa", e.message);
    return { categorias: [], formasPagto: [], modalidades: ["Geral"] };
  }
}

// ============================================================================
// 🌍 APIs PÚBLICAS (Para a Tela de Agendamento de Visitantes)
// ============================================================================

/**
 * Busca todas as Academias ATIVAS para a tela pública de agendamento (Sem trava de usuário)
 */
function getLocaisTreinoPublico() {
  try {
    const locais = lerTabelaDinamica("Locais_de_treino");
    return locais
      .filter(l => String(l.status).toLowerCase() === "ativo" && l.nome_do_local)
      .map(row => ({
        nome: row.nome_do_local,
        endereco: row['endereço'] || row.endereco || "Endereço sob consulta",
        cidade: row['cidade/estado'] || "",
        contato: row.contato || "",
        iframeHtml: row.html_mapa_off_lline || "",
        status: row.status || "Ativo"
      }));
  } catch (e) {
    registrarLogBlindado("ERRO", "API_PUBLICA_LOCAIS", e.message);
    return [];
  }
}

/**
 * Busca todas as Turmas ATIVAS para a tela pública de agendamento (Sem trava de usuário)
 */
function listarTurmasPublico() {
  try {
    const turmas = lerTabelaDinamica("Config_Turmas");
    return turmas
      .filter(t => String(t.status).toLowerCase() === "ativa" || String(t.status).toLowerCase() === "ativo")
      .map(t => ({
        nome: t.nome_da_turma || "Sem Nome",
        local: t.local_vinculado || "Matriz",
        dias: t.dias_da_semana || "",
        inicio: t["horário_início"] || t.horario_inicio || "",
        fim: t["horário_fim"] || t.horario_fim || "",
        status: t.status || "Ativa"
      }));
  } catch (e) {
    registrarLogBlindado("ERRO", "API_PUBLICA_TURMAS", e.message);
    return [];
  }
}

/**
 * 💰 MOTOR DE BI: RELATÓRIO ANALÍTICO FINANCEIRO AVANÇADO
 * Agora com cálculo de Previsto vs Realizado
 */
function getRelatorioAnaliticoFinanceiro(loginSolicitante, filtros) {
  try {
    const auth = getPermissoesUsuario(loginSolicitante);
    const transacoes = lerTabelaDinamica("Fin_Transacoes");
    const alunos = lerTabelaDinamica("cadastro_de_alunos");

    const mapAlunos = {};
    alunos.forEach(a => {
      const log = String(a.login || "").toLowerCase().trim();
      if (log) {
        mapAlunos[log] = { nome: a.nome_completo || a.nome || "", graduacao: a.graduacao_atual || "Iniciante", modalidade: a.modalidade || "Geral" };
      }
    });

    // 🛡️ NOVOS ACUMULADORES (Iguais aos Cards do Dashboard)
    let recebido = 0;
    let aReceber = 0;
    let despesasPagas = 0;
    let aPagar = 0;

    let filtroAcads = filtros.academias ? filtros.academias.map(a => String(a).toLowerCase().trim()) : [];
    let buscarTodas = filtroAcads.length === 0 || filtroAcads.includes("todas");

    let dadosFiltrados = transacoes.filter(t => {
      const tAcad = String(t.academia_ref || "").trim().toLowerCase();

      if (!auth.isMaster && !auth.academias.some(myAcad => tAcad.includes(myAcad))) return false;
      if (!buscarTodas && !filtroAcads.some(fAcad => tAcad.includes(fAcad))) return false;
      if (filtros.tipo && filtros.tipo !== "TODOS" && String(t.tipo).toLowerCase() !== String(filtros.tipo).toLowerCase()) return false;
      if (filtros.categoria && filtros.categoria !== "TODAS" && String(t.categoria).toLowerCase() !== String(filtros.categoria).toLowerCase()) return false;
      if (filtros.status && filtros.status !== "TODOS" && !String(t.status).toLowerCase().includes(String(filtros.status).toLowerCase())) return false;

      const tLogin = String(t.login_aluno || "").toLowerCase().trim();
      const alunoInfo = mapAlunos[tLogin] || { nome: "", graduacao: "", modalidade: "" };

      if (filtros.aluno && !alunoInfo.nome.toLowerCase().includes(String(filtros.aluno).toLowerCase())) return false;
      if (filtros.graduacao && filtros.graduacao !== "TODAS" && String(alunoInfo.graduacao).toLowerCase() !== String(filtros.graduacao).toLowerCase()) return false;
      if (filtros.modalidade && filtros.modalidade !== "TODAS" && String(alunoInfo.modalidade).toLowerCase() !== String(filtros.modalidade).toLowerCase()) return false;

      const tData = parseDataSegura(t.data_registro);
      if (filtros.dataInicio && tData) { if (tData < new Date(filtros.dataInicio + "T00:00:00")) return false; }
      if (filtros.dataFim && tData) { if (tData > new Date(filtros.dataFim + "T23:59:59")) return false; }

      t._alunoNome = alunoInfo.nome || "-";
      t._alunoGraduacao = alunoInfo.graduacao || "-";
      t._alunoModalidade = t.modalidade || alunoInfo.modalidade || "-";

      let vStr = String(t.valor || "0").replace("R$", "").trim();
      if (vStr.includes(",") && !vStr.includes(".")) vStr = vStr.replace(/\./g, "").replace(",", ".");
      const valNum = parseFloat(vStr) || 0;

      // 🛡️ SEPARAÇÃO INTELIGENTE (CAIXA REAL VS PREVISÃO)
      const isConcluido = String(t.status).toLowerCase().includes("conclu");

      if (String(t.tipo).toLowerCase() === "receita") {
        if (isConcluido) recebido += valNum; else aReceber += valNum;
      } else if (String(t.tipo).toLowerCase() === "despesa") {
        if (isConcluido) despesasPagas += valNum; else aPagar += valNum;
      }

      return true;
    });

    const resultado = dadosFiltrados.reverse().map(t => ({
      id: t.id_transacao, data: formatDate(t.data_registro), tipo: t.tipo || "---",
      categoria: t.categoria || "---", descricao: t.descricao || "---",
      valor: parseFloat(String(t.valor).replace(',', '.') || 0).toFixed(2).replace('.', ','),
      forma: t.forma_pagto || "---", status: t.status || "Concluido",
      academia: t.academia_ref || "---", aluno: t._alunoNome, graduacao: t._alunoGraduacao, modalidade: t._alunoModalidade
    }));

    return {
      success: true,
      data: resultado,
      resumo: {
        recebido: recebido,
        aReceber: aReceber,
        despesas: despesasPagas,
        aPagar: aPagar,
        saldo: recebido - despesasPagas
      }
    };
  } catch (e) { return { success: false, msg: e.message }; }
}

/**
 * ============================================================================
 * 💰 MOTOR DE BAIXA: TRANSFORMA PENDENTE EM CONCLUÍDO E RENOVA PLANO
 * ============================================================================
 */
function darBaixaTransacaoPendente(idTransacao, loginSolicitante) {
  try {
    const auth = getPermissoesUsuario(loginSolicitante);
    const sheet = getSheet("Fin_Transacoes");
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase().replace(/\s+/g, "_"));

    const colId = headers.indexOf("id_transacao");
    const colStatus = headers.indexOf("status");
    const colAcad = headers.indexOf("academia_ref");
    const colTipo = headers.indexOf("tipo");
    const colCat = headers.indexOf("categoria");
    const colAluno = headers.indexOf("login_aluno");

    let linhaAlvo = -1;
    let transacao = null;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][colId]) === idTransacao) {
        linhaAlvo = i + 1;
        transacao = data[i];

        const tAcad = String(data[i][colAcad]).toLowerCase();
        if (!auth.isMaster && !auth.academias.some(myAcad => tAcad.includes(myAcad))) {
          return { success: false, msg: "Acesso negado a esta unidade." };
        }
        break;
      }
    }

    if (linhaAlvo === -1) return { success: false, msg: "Transação não encontrada." };
    if (String(transacao[colStatus]).toLowerCase().includes("conclu")) return { success: false, msg: "Esta conta já está paga!" };

    // Atualiza status para Concluído
    sheet.getRange(linhaAlvo, colStatus + 1).setValue("Concluido");

    // Regra de Negócio: Se for mensalidade, renova o plano do aluno
    const tipo = String(transacao[colTipo]).toLowerCase();
    const cat = String(transacao[colCat]).toLowerCase();
    const alunoLogin = String(transacao[colAluno]);

    if (tipo === "receita" && (cat.includes("mensalidade") || cat.includes("pacote") || cat.includes("plano")) && alunoLogin) {
      const assinaturas = lerTabelaDinamica("Fin_Assinaturas");
      const assAtual = assinaturas.find(a => String(a.login_aluno).toLowerCase() === alunoLogin.toLowerCase());
      const pacoteDoAluno = assAtual ? assAtual.pacote_atual : "Mensalidade Padrão";
      processarRenovacaoAssinatura(alunoLogin, pacoteDoAluno);
    }

    registrarLogBlindado("SUCESSO", "DAR_BAIXA", `Transação ${idTransacao} concluída por ${loginSolicitante}`);
    return { success: true, msg: "Baixa realizada com sucesso! Saldo e Assinaturas atualizados." };

  } catch (e) {
    registrarLogBlindado("ERRO", "darBaixaTransacaoPendente", e.message);
    return { success: false, msg: "Erro Crítico: " + e.message };
  }
}

// ============================================================================
// 📁 DOSSIÊ DO ALUNO (MOTOR FINAL: LTV, ENGAJAMENTO, DATAS E FICHA)
// ============================================================================
function getDossierAlunoAdmin(login) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const obterDadosAba = (nomeAba) => {
      const sheet = ss.getSheetByName(nomeAba);
      if (!sheet) return [];
      const data = sheet.getDataRange().getValues();
      if (data.length < 2) return [];
      const headers = data[0];
      return data.slice(1).map(row => {
        let obj = {}; headers.forEach((h, i) => obj[h] = row[i]); return obj;
      });
    };

    const bdAlunos = obterDadosAba("cadastro_de_alunos");
    const bdTransacoes = obterDadosAba("Fin_Transacoes");
    const bdChamadas = obterDadosAba("Registro_Chamada");
    const bdTurmas = obterDadosAba("Config_Turmas");
    const bdAssinaturas = obterDadosAba("Fin_Assinaturas");

    const aluno = bdAlunos.find(a => String(a.LOGIN).trim().toLowerCase() === String(login).trim().toLowerCase());
    if (!aluno) throw new Error("Aluno não encontrado na base de dados.");

    // --- 1. CRUZAMENTO: Responsável, Assinaturas e Datas Físicas ---
    let respTurma = "Não Definido";
    if (aluno['Turma Vinculada']) {
      const turmaObj = bdTurmas.find(t => String(t['Nome da Turma']).trim().toLowerCase() === String(aluno['Turma Vinculada']).trim().toLowerCase());
      if (turmaObj) respTurma = turmaObj['Responsável'];
    }

    let planoAssociado = "Sem Plano";
    let dataVencimento = "N/A";
    const assinatura = bdAssinaturas.find(a => String(a.Login_Aluno).trim().toLowerCase() === String(login).trim().toLowerCase());
    if (assinatura) {
      planoAssociado = assinatura['Pacote_Atual'] || "Nenhum";
      let vData = assinatura['Data_Fim'];
      if (vData instanceof Date) dataVencimento = Utilities.formatDate(vData, Session.getScriptTimeZone(), "dd/MM/yyyy");
      else if (vData) dataVencimento = vData;
    }

    let nascFormatado = "N/A";
    let idade = "N/A";
    if (aluno['Data de Nascimento ']) {
      let dNasc = aluno['Data de Nascimento '];
      let dateNasc = (dNasc instanceof Date) ? dNasc : new Date(dNasc);
      if (dateNasc && !isNaN(dateNasc.getTime())) {
        nascFormatado = Utilities.formatDate(dateNasc, Session.getScriptTimeZone(), "dd/MM/yyyy");
        idade = Math.floor((Date.now() - dateNasc.getTime()) / (1000 * 60 * 60 * 24 * 365.25)) + " anos";
      }
    }

    // --- 2. CONSTRUÇÃO DA FICHA COMPLETA (PERFIL) ---
    const perfil = {
      cpf: aluno['CPF'] || "Não informado",
      foto: aluno['Foto 3x4 (para a carteirinha)'] || "",
      academia: aluno['Academia Vinculada'] || "Sem Registo",
      modalidade: aluno['Modalidade'] || "Geral",
      turma: aluno['Turma Vinculada'] || "Sem Turma",
      responsavelTurma: respTurma,
      plano: planoAssociado,
      vencimento: dataVencimento,
      graduacao: aluno['GRADUACAO_ATUAL'] || aluno['Graduação'] || "Branca",
      proxGraduacao: aluno['PROX_GRADUACAO'] || "N/A",
      idade: idade,
      nascimento: nascFormatado,
      peso: aluno['Peso'] ? aluno['Peso'] + " kg" : "---",
      altura: aluno['Altura'] ? aluno['Altura'] + " m" : "---",
      mae: aluno['Nome da Mãe'] || "",
      pai: aluno['Nome do Pai'] || "",
      endereco: aluno['Endereço'] || "Não informado",
      telefone: aluno['Telefone'] || "Sem telefone",
      email: aluno['E-mail'] || aluno['Endereço de e-mail'] || "Sem E-mail"
    };

    // --- 3. CÁLCULO DE TEMPO E MATRÍCULA ---
    let mesesAcademia = 0;
    let dataMatriculaStr = "N/A";
    if (aluno['Carimbo de data/hora']) {
      const dataReg = new Date(aluno['Carimbo de data/hora']);
      if (!isNaN(dataReg.getTime())) {
        const hoje = new Date();
        mesesAcademia = (hoje.getFullYear() - dataReg.getFullYear()) * 12 + (hoje.getMonth() - dataReg.getMonth());
        if (mesesAcademia < 0) mesesAcademia = 0;
        dataMatriculaStr = Utilities.formatDate(dataReg, Session.getScriptTimeZone(), "dd/MM/yyyy");
      }
    }

    // --- 4. LTV (Lifetime Value) ---
    let ltvTotal = 0;
    const faturas = [];
    bdTransacoes.forEach(t => {
      if (String(t.Login_Aluno).trim().toLowerCase() === String(login).trim().toLowerCase() &&
        String(t.Tipo).trim().toLowerCase() === "receita" && String(t.Status).trim().toLowerCase().includes("conclui")) {
        let val = parseFloat(String(t.Valor).replace(/\./g, '').replace(',', '.')) || 0;
        ltvTotal += val;
        let dataFormatada = t.Data_Registro instanceof Date ? Utilities.formatDate(t.Data_Registro, Session.getScriptTimeZone(), "dd/MM/yyyy") : t.Data_Registro;
        faturas.push({ data: dataFormatada, ref: t.Descricao || t.Categoria, valor: val });
      }
    });
    faturas.sort((a, b) => new Date(b.data.split('/').reverse().join('-')) - new Date(a.data.split('/').reverse().join('-')));

    // --- 5. AULAS (Datas da 1ª e Última) ---
    let totalAulas = 0;
    const aulasAssistidas = [];
    bdChamadas.forEach(c => {
      if (String(c.Lista_Alunos_IDs || "").toLowerCase().includes(String(login).trim().toLowerCase())) {
        totalAulas++;
        let dataFormatada = (c.Data_Treino || c.Data_Registro) instanceof Date ? Utilities.formatDate(c.Data_Treino || c.Data_Registro, Session.getScriptTimeZone(), "dd/MM/yyyy") : (c.Data_Treino || c.Data_Registro);
        aulasAssistidas.push({ data: dataFormatada, conteudo: c.Conteudo || "Treino Regular" });
      }
    });
    aulasAssistidas.sort((a, b) => new Date(b.data.split('/').reverse().join('-')) - new Date(a.data.split('/').reverse().join('-')));

    let primeiraAula = "N/A";
    let ultimaAula = "N/A";
    if (aulasAssistidas.length > 0) {
      ultimaAula = aulasAssistidas[0].data;
      primeiraAula = aulasAssistidas[aulasAssistidas.length - 1].data;
    }

    // --- 6. IA DE PARECER ---
    let parecer = "";
    if (ltvTotal >= 1000 && totalAulas >= 15) parecer = `⭐⭐⭐ Atleta VIP! Investiu R$ ${ltvTotal.toFixed(2).replace('.', ',')} e treinou ${totalAulas} vezes. Última aula em ${ultimaAula}.`;
    else if (totalAulas === 0 && mesesAcademia > 1) parecer = `⚠️ ALERTA: Em evasão! Zero aulas, mas matriculado desde ${dataMatriculaStr}.`;
    else if (ltvTotal === 0) parecer = `ℹ️ Sem registo financeiro. Verifique se tem bolsa ou dívidas.`;
    else parecer = `✅ Atleta Engajado. Tempo: ${mesesAcademia} meses | Aulas: ${totalAulas} | LTV: R$ ${ltvTotal.toFixed(2).replace('.', ',')}`;

    return {
      success: true,
      perfil: perfil,
      meses: mesesAcademia,
      dataMatricula: dataMatriculaStr,
      ltv: ltvTotal,
      totalAulas: totalAulas,
      primeiraAula: primeiraAula,
      ultimaAula: ultimaAula,
      parecer: parecer,
      faturas: faturas,
      aulas: aulasAssistidas
    };

  } catch (e) {
    return { success: false, msg: e.message };
  }
}

// ============================================================================
// 🥋 MOTOR MULTICANAL DE CONTEÚDOS (CHAMADA DINÂMICA)
// ============================================================================
function getConteudoMulticanal() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let conteudos = [];

    // 1. VIDEOTECA (Filtrando Ativos)
    const sheetVid = ss.getSheetByName("Config_Videoteca");
    if (sheetVid) {
      const dataVid = sheetVid.getDataRange().getValues();
      if (dataVid.length > 1) {
        dataVid.slice(1).forEach(row => {
          if (String(row[5]).trim().toLowerCase() === 'ativo' && row[1]) {
            // Entrega titulo e modalidade (coluna 4)
            conteudos.push({ titulo: "🎥 Vídeo: " + String(row[1]).trim(), modalidade: String(row[4] || "") });
          }
        });
      }
    }

    // 2. CURSOS (Filtrando Eventos Ativos)
    const sheetCur = ss.getSheetByName("Cursos");
    if (sheetCur) {
      const dataCur = sheetCur.getDataRange().getValues();
      if (dataCur.length > 1) {
        dataCur.slice(1).forEach(row => {
          // Status na Coluna F (5), Nome na Coluna A (0)
          if (String(row[5]).trim().toLowerCase() === 'ativo' && row[0]) {
            conteudos.push({ titulo: "🏆 Curso: " + String(row[0]).trim(), modalidade: "Evento Especial" });
          }
        });
      }
    }

    // 3. PROGRAMAS TÉCNICOS (Filtrando PDFs válidos)
    const sheetProg = ss.getSheetByName("Config_Programas");
    if (sheetProg) {
      const dataProg = sheetProg.getDataRange().getValues();
      if (dataProg.length > 1) {
        dataProg.slice(1).forEach(row => {
          const idArq = String(row[1]).trim().toLowerCase();
          if (idArq !== 'indisponivel' && idArq !== '' && row[0]) {
            const desc = row[4] ? " - " + String(row[4]).trim() : "";
            // Modalidade na Coluna D (3)
            conteudos.push({ titulo: "📄 Prog: " + String(row[0]).trim() + desc, modalidade: String(row[3] || "") });
          }
        });
      }
    }

    // Ordena alfabeticamente para agrupar os Emojis no Select
    conteudos.sort((a, b) => a.titulo.localeCompare(b.titulo));
    return conteudos;

  } catch (e) {
    console.error("Erro no getConteudoMulticanal:", e);
    return [];
  }
}