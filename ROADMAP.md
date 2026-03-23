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

## 🎯 Fase 2.5: Máquina de Vendas e Automação de Base (FOCO ATUAL)
- [x] **Agendamento Dinâmico (Leads):** Criar página externa para Visitantes agendarem Aulas Experimentais com Chips de Turmas.
- [x] **Integração WhatsApp (Leads):** O agendamento da aula experimental dispara uma mensagem pré-formatada para o WhatsApp da unidade.
- [x] **Função Auto-Setup e Verificação de Banco de Dados:** Criar uma função para validar e corrigir todas as abas e colunas.
- [x] **Rodapé White-Label e Direitos:** Incluir em todas as páginas o atalho para a página do desenvolvedor. *
- [ ] **[🔥 ALTO ROI]Blindagem Ultra Militar para registea cada passo do código e erros por menor que sejam** Nenhum erro ou porcedimento pode ser silenciosos. A pagina Adm.html deve ter blindagem miltar desde de seu acesso a sua saida. 
- [ ] **[🔥 ALTO ROI]Implementar Controle MultiAcademia e mono Academia para ADm (Estilo Franquia):** Um usuario ADM só podera acessar a academia que estiver vinclulada ano seu cadastro dentro da coluna academia_vinclulada. Assimum franqueador podera ver todoas as acedemias se no seu cadastro estiver escrito 'todas' e seus franqueados so poderam ver duas ou mais acadeoias se estiver escrito o nome dela na coluna 'academias_vinclualdas'. (revisar Código.gs, Adm.html e GestãoAlunos.html)
- [ ] **[🔥 ALTO ROI] Crud Formas de pagamento:** Cada acedemia deve ter as suas forma de pagamento cadastrada . Nenhuma forma de pagamento pode ser hardcode (revisar Código.gs, Adm.html e Finaiceiro.html)
- [ ] **[🔥 ALTO ROI] Crud Categoria Finaiceira:** Tipo (Se mensalidade, aluguel, pacote procional, curso, despesa propoganda, etc) e dizer se é um despsa ou receita, de qual academia esta vinculda, se uma ou todas, a data de vencimento ou recimento  e o status se esta paga /  recebida ou pendente. Criar um relatori analitivo fiultardo por data, academia, tipo e status. Revisar (Código,gs, Financeiro.html, Adm.html)


## 💰 Fase 3: Inteligência Financeira e Relatórios (Próximos Passos - Foco Emilly)
- [ ]Nova Aba "Relatórios" no Adm.html:** Central de Inteligência com relatórios de Cadastros, Inadimplência e Lotação de Turmas.
- [ ] Contas a Pagar e Receber:** Lançamento de despesas (aluguel, repasse de professores) e receitas cruzadas.
- [ ] Dashboard de Previsões (Financeiro.html):** Visão de Faturamento Previsto vs. Realizado.

## 🔮 Fase 4: Automação, CRM e Compliance
- [ ] Expandir Hydration:** Aplicar blindagem anti-apagão nos CRUDs menores (Alunos, financeiros, Cursos, Locais, Vídeos, etc).
- [ ] Check-in Antecipado: Permitir que o aluno agende a sua presença na grade de aulas com antecedência.
- [ ] Histórico do Aluno (tempo na turma, aulas assistidas, conteúdo visualizado).
- [ ] **Compliance:** Controle de Anuidade (Bloqueio em caso de inadimplência).
- [ ] **Compliance:** Trava de inatividade por falta de presença no Congresso Oficial.
- [ ] **Compliance:** Controle de Assiduidade em Treinos de Instrutores.
- [ ] Relatórios Avançados (Gráficos de evasão, horários de pico).
- [ ] Validar a renderização das imagens (Base64) na carteirinha em PDF em diferentes dispositivos.
- [ ] Validar se a planilha Mestra funciona perfeitamente para clonagem rápida.
- [ ] Cobrança Automática: Disparo de WhatsApp via API para alunos atrasados.
- [ ] Alertas Automáticos de Aniversário via API.
- [ ] Criar alertas de limite de escalabilidade (Avisar QA/Dev quando o banco Sheets chegar a 50% da capacidade).

## 🌐 Projetos Especiais (Customizações)
- [ ] Criar uma apresentação institucional do App para vender a outras academias.
- [ ] Adaptação do sistema base para a **Confederação Pernambucana de Sambo**. 
- [ ] Criar uma Landing Page comercial e direcionar o tráfego diretamente para a tela de Agendamento do App.