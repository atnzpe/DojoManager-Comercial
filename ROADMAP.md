# 🗺️ Roadmap de Desenvolvimento - DojoManager

Acompanhe a evolução do nosso sistema.

## ✅ Fase 1: Fundação Serverless (Concluído)

- [x] Arquitetura de Banco de Dados via Google Sheets.
- [x] Sistema de Login e Sessão (localStorage).
- [x] Dashboard Responsivo em Dark Mode.
- [x] CRUD Administrativo completo (Alunos, Turmas, Locais).

## 🧹 Fase 1.5: Auditoria e White-Label (Prioridade Atual)

- [x] Implementar a personalização de imagens de logo, background e 5 cores.
- [x] Ajustar CSS para que todas as telas tenham o mesmo estilo CSS.
- [x] Remover Hardcode do botão loja.
- [x] Remover Hardcode do link financeiro.
- [x] Remover Hardcode da tela de mídias sociais.
- [x] **Ajuda.html:** Reescrever manual neutro e modularizar.
- [x] **Social.html:** Transformar botões em links dinâmicos.
- [x] **Código.js:** Alterar `.setTitle()` para buscar o nome da academia dinamicamente.
- [x] **CardTemplate.html:** Recriar a carteirinha usando apenas HTML e CSS puro (Adeus imagens estáticas de fundo).
- [ ] Validar a renderização das imagens (Base64) na carteirinha em PDF em diferentes dispositivos antes de implementar na Confederação Pernambucana de Sambo.
- [ ] **O CHEFÃO:** Remover Hardcode de graduação. Permitir cadastro de ID (ordem), cor e modalidade (Karate, Muay Thai, etc).
- [ ] Gestão de Turmas.
- [ ] Gestão de cadastro de Modalidades, associados aos locais de treino.
- [ ] Permitir o Instrutor/Mestre acompanhar o desenvolvimento do Aluno (aulas, tempo).
- [ ] Validar todos os arquivos do projeto (Pente fino).
- [ ] Validar se a planilha Mestra funciona perfeitamente para clonagem.
- [ ] Criar alertas de limite de escalabilidade (Avisar QA/Dev quando o banco Sheets chegar a 50% da capacidade).

## 🚧 Fase 2: Gestão Financeira e Turmas (Em Andamento)

- [x] Dashboard Financeiro (Inadimplência, Faturamento).
- [ ] Separação de Turmas por Nível e Idade (Infantil, Iniciante, Avançado).
- [ ] Check-in Antecipado (Agendamento de grade de aulas).

## 🔮 Fase 3: Automação, CRM e Compliance (Próximos Passos)

- [ ] Cobrança Automática: Disparo de WhatsApp (API) para atrasados.
- [ ] Alertas Automáticos de Aniversário.
- [ ] Relatórios Avançados (Gráficos de evasão, horários de pico).
- [ ] Histórico do Aluno (tempo na turma, aulas assistidas, conteúdo visualizado).
- [ ] Aluno Multimodalidade (Filtrar vídeos e programas baseados nas modalidades).
- [ ] Modularizar a tela de Dashboard (Escolher quem vê qual card).
- [ ] **Compliance:** Controle de Anuidade (Bloqueio em caso de inadimplência).
- [ ] **Compliance:** Trava de inatividade por falta de presença no Congresso Oficial.
- [ ] **Compliance:** Controle de Assiduidade em Treinos de Instrutores.
- [ ] Validar se ao fazer a cópia da planilha, o SaaS continua a funcionar de forma autônoma.

## 🌐 Projetos Especiais

- [ ] Adaptação do sistema base para a **Confederação Pernambucana de Sambo**.
