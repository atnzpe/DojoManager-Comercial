# 🗺️ Roadmap de Desenvolvimento - DojoManager

Acompanhe a evolução, as metas e as próximas entregas do nosso sistema.

## ✅ Fase 1: Fundação Serverless (Concluído)
- [x] Arquitetura de Banco de Dados via Google Sheets.
- [x] Sistema de Login e Sessão (localStorage).
- [x] Dashboard Responsivo em Dark Mode.
- [x] CRUD Administrativo completo (Alunos, Turmas, Locais).

## 🧹 Fase 1.5: Auditoria, UX e White-Label (Concluído)
- [x] Implementar a personalização de imagens de logo, background e 5 cores.
- [x] Ajustar CSS para que todas as telas tenham o mesmo estilo (Padronização UI).
- [x] Remover Hardcode do botão loja, link financeiro e mídias sociais.
- [x] **Ajuda.html:** Reescrever manual neutro e modularizar.
- [x] **Social.html:** Transformar botões em links dinâmicos.
- [x] **Código.js:** Alterar `.setTitle()` para buscar o nome da academia dinamicamente.
- [x] **CardTemplate.html:** Recriar a carteirinha usando apenas HTML e CSS puro (Adeus imagens estáticas de fundo).
- [x] **Motor de Graduação:** Remover Hardcode. Permitir cadastro de ID (ordem), cor e modalidade.
- [x] Permitir o Instrutor/Mestre acompanhar o desenvolvimento do Aluno (aulas, tempo de tatame).
- [x] Exibir no Perfil do Aluno e no CRUD de Alunos: Peso, Idade (anos, meses e dias), Modalidade, Altura. 
- [x] Criar o "Super Relatório" administrativo.
- [x] 🚨 **Auditoria de Integridade de CRUDs:** Aplicada regra de *Hydration V5* na edição de alunos para não apagar colunas ocultas no frontend (Fim do Hard Delete Acidental).

## 🚧 Fase 2: Gestão Financeira e Turmas (Concluído)
- [x] Dashboard Financeiro (Inadimplência, Faturamento).
- [x] **Fluxo de Pagamento Híbrido:** Exibir QR Code/Chave PIX da academia e botão dinâmico no painel do aluno.
- [x] **Gestão de Turmas Administrativa:** Criar CRUD de Turmas associando Modalidade, Dia e Horário.
- [x] **Separação de Turmas por Nível e Idade:** (Infantil, Iniciante, Avançado) com trava de segurança inteligente no momento do Check-in.
- [x] Modularizar a tela de Dashboard (Escolher quem vê qual card via níveis de acesso RBAC).
- [x] Aluno Multimodalidade (Filtrar vídeos e programas baseados nas modalidades cruzadas).

## 🎯 Fase 2.5: Máquina de Vendas e Automação de Base (Concluído)
- [x] **Agendamento Dinâmico (Leads):** Criar página externa para Visitantes agendarem Aulas Experimentais com Chips de Turmas.
- [x] **Integração WhatsApp (Leads):** O agendamento da aula experimental dispara uma mensagem pré-formatada para o WhatsApp da unidade.
- [x] **Função Auto-Setup e Verificação de Banco de Dados:** Criar uma função para validar e corrigir todas as abas e colunas.
- [x] **Rodapé White-Label e Direitos:** Incluir em todas as páginas o atalho para a página do desenvolvedor.
- [x] **[🔥 ALTO ROI] Blindagem Ultra Militar:** Registrar cada passo do código e erros para que nenhum procedimento seja silencioso (Trilha de Auditoria).
- [x] **[🔥 ALTO ROI] Implementar Controle MultiAcademia e Mono Academia (Estilo Franquia):** Filtro de visibilidade de dados baseado na coluna `academia_vinculada` do usuário.
- [x] **[🔥 ALTO ROI] Crud Formas de Pagamento:** Cada academia gere as suas formas de recebimento. Fim do Hardcode.
- [x] **[🔥 ALTO ROI] Crud Categoria Financeira:** Regras para despesas/receitas vinculadas a unidades específicas.
- [x] Incluir no campo senha a opção do usuário visualizar a senha digitada antes de confirmar (Olhinho da Senha).

## 💰 Fase 3: Inteligência Financeira e Relatórios (Em Andamento)
- [x] **Nova Aba "Relatórios" no Adm.html:** Central de Inteligência com relatórios de Cadastros, Inadimplência e Lotação de Turmas.
- [x] **Contas a Pagar e Receber:** Lançamento de despesas (aluguel, repasse) e receitas cruzadas.
- [x] **Dashboard de Previsões (Financeiro.html):** Visão de Faturamento Previsto vs. Realizado (Saldo Final).
- [ ] Gerar Relatório Específico de Inadimplentes (Focado na Régua de Cobrança).
- [ ] Gerar Relatório de Alunos "Fantasmas" (Alunos sem Turma vinculada e sem Assinatura de Pacote).

## 🔮 Fase 4: Automação, CRM e Compliance (FOCO ATUAL)
- [x] **Expandir Hydration:** Aplicar blindagem anti-apagão nos CRUDs menores (Cursos, Locais, Vídeos, etc).
- [x] **Histórico do Aluno (Dossiê/CRM Avançado):** Visão analítica e sintética (LTV total investido, tempo na turma, aulas assistidas, progressão de peso/graduação e pareceres para bolsas/descontos).
- [x] **Conteúdo de Aula Dinâmico:** Permitir que, além da Videoteca, possamos listar e selecionar "Cursos" e "Programas Técnicos" como Chips no momento de abertura da aula/chamada.
- [x] **Compliance/Canal Seguro:** Incluir os tópicos "Denúncia", "Doping" e "Assédio" no menu de envio de contato para suporte ao atleta.
- [x] **Sistema de Alerta Crítico (Compliance):** Disparo imediato de notificação (Email/WhatsApp API) para a diretoria quando houver envio de relatório sigiloso (Denúncia, Doping ou Assédio).
- [ ] Check-in Antecipado: Permitir que o aluno agende a sua presença na grade de aulas com antecedência.
- [ ] **Compliance:** Controle de Anuidade (Bloqueio em caso de inadimplência).
- [ ] **Compliance:** Trava de inatividade por falta de presença no Congresso Oficial.
- [ ] **Compliance:** Controle de Assiduidade em Treinos de Instrutores.
- [ ] Relatórios Avançados (Gráficos de evasão, horários de pico).
- [ ] Validar a renderização das imagens (Base64) na carteirinha em PDF em diferentes dispositivos.
- [ ] Validar se a planilha Mestra funciona perfeitamente para clonagem rápida.
- [ ] Cobrança Automática: Disparo de WhatsApp via API para alunos atrasados.
- [ ] Alertas Automáticos de Aniversário via API.
- [ ] Criar alertas de limite de escalabilidade (Avisar QA/Dev quando o banco Sheets chegar a 50% da capacidade).

## 🌐 Projetos Especiais (Customizações Institucionais)
- [ ] Criar uma apresentação institucional do App para vender a outras academias.
- [ ] Criar uma Landing Page comercial e direcionar o tráfego diretamente para a tela de Agendamento do App.
- [ ] Adaptação do sistema base para a **Confederação Brasileira de Sambo (Pauta Março/2026):**
  - [ ] **Módulo de Gestão de Eventos/Competições:** Coleta obrigatória de peso, tamanho de camisa para credenciamento.
  - [ ] **Gestão de Comissões Oficiais:** Módulo de indicação de membros com controle de prazos para as federações estaduais.
  - [ ] **Padronização Sistêmica de Exames de Faixa:** Workflow de aprovação pelas federações (Trava de aprovação por nível).
  - [ ] **Módulo de Compliance e Documentação (GED):** Upload e gestão de estatutos, alvarás e atestados médicos com semáforo de status.
  - [ ] **Motor de Ranking Nacional Dinâmico:** Lançamento de resultados de eventos e cálculo automático de ranking pontuado.
  - [ ] **Trilha de Carreira para Árbitros e Técnicos:** Cadastro de níveis de formação (ex: Nível 1, 2, 3) e travas de inscrição.
  - [ ] **Centro de Custos Específico para Projetos (Leis de Incentivo):** Tags de projetos nas transações financeiras para exportação rápida de prestação de contas.

## 🚀 Fase 6: Escalabilidade e Nuvem
- [ ] Implementar uma base de dados estruturada Firebase (Google).
- [ ] Implementar login OAuth 2.0 (Google Login).