# 🗺️ Roadmap de Desenvolvimento - DojoManager

Acompanhe a evolução, as metas e as próximas entregas do nosso sistema.

## ✅ Fase 1: Fundação Serverless (Concluído)
- [x] Arquitetura de Banco de Dados via Google Sheets.
- [x] Sistema de Login e Sessão (localStorage).
- [x] Dashboard Responsivo em Dark Mode.
- [x] CRUD Administrativo completo (Alunos, Turmas, Locais).

## 🧹 Fase 1.5: Auditoria, UX e White-Label (Quase Concluído)
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
- [ ] Replicar blindagem Anti-Apagão (*Hydration*) para os demais CRUDs menores (Cursos, Locais, Vídeos).
- [ ] Validar a renderização das imagens (Base64) na carteirinha em PDF em diferentes dispositivos.
- [ ] Validar se a planilha Mestra funciona perfeitamente para clonagem rápida.
- [ ] Criar alertas de limite de escalabilidade (Avisar QA/Dev quando o banco Sheets chegar a 50% da capacidade).

## 🚧 Fase 2: Gestão Financeira e Turmas (Em Andamento - FOCO ATUAL 🎯)
- [x] Dashboard Financeiro (Inadimplência, Faturamento).
- [x] **[🚀 QUICK WIN] Fluxo de Pagamento Híbrido:** Exibir QR Code/Chave PIX da academia e botão dinâmico no painel do aluno.
- [ ] **[🔥 ALTO ROI] Gestão de Turmas Administrativa:** Criar CRUD de Turmas associando Modalidade, Dia e Horário. *(Próxima Tarefa)*
- [ ] **[🔥 ALTO ROI] Separação de Turmas por Nível e Idade:** (Infantil, Iniciante, Avançado) para resolver a dor do professor multitarefa.
- [ ] Check-in Antecipado (Agendamento de grade de aulas pelo aluno).

## 🔮 Fase 3: Automação, CRM e Compliance (Próximos Passos)
- [x] Aluno Multimodalidade (Filtrar vídeos e programas baseados nas modalidades cruzadas).
- [ ] Cobrança Automática: Disparo de WhatsApp (API) para atrasados.
- [ ] Alertas Automáticos de Aniversário.
- [ ] Relatórios Avançados (Gráficos de evasão, horários de pico).
- [ ] Histórico do Aluno (tempo na turma, aulas assistidas, conteúdo visualizado).
- [ ] Modularizar a tela de Dashboard (Escolher quem vê qual card via níveis de acesso).
- [ ] **Compliance:** Controle de Anuidade (Bloqueio em caso de inadimplência).
- [ ] **Compliance:** Trava de inatividade por falta de presença no Congresso Oficial.
- [ ] **Compliance:** Controle de Assiduidade em Treinos de Instrutores.
- [ ] Criar uma apresentação institucional do App.
- [ ] Atualizar o manual (ajuda.html) a cada nova implementação.
- [ ] Incluir em todas as páginas o atalho para a página do desenvolvedor (Rodapé White-Label).

## 🌐 Projetos Especiais (Customizações)
- [ ] Adaptação do sistema base para a **Confederação Pernambucana de Sambo**. Criar uma Landpage e da landpage direcinar para o App