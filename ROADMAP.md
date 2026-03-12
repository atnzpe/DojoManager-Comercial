# 🗺️ Roadmap de Desenvolvimento - DojoManager

Acompanhe a evolução do nosso sistema.

## ✅ Fase 1: Fundação Serverless (Concluído)
- [x] Arquitetura de Banco de Dados via Google Sheets.
- [x] Sistema de Login e Sessão (localStorage).
- [x] Dashboard Responsivo em Dark Mode.
- [x] CRUD Administrativo completo (Alunos, Turmas, Locais).

## 🧹 Fase 1.5: Auditoria e White-Label (Prioridade Atual)
- [x] Implementar a personalização de imagens de logo, background e 5 cores personalizadas.
- [x] Ajustar CSS para que todas as telas tenham o mesmo estilo CSS com cores personalizadas (White-Label Integrado).
- [ ] Remover Hardcode do botão loja. Deixar o usuário escolher o link que deseja incluir. Deve ser ajustado através do painel Admin.
- [ ] Remover Hardcode do link financeiro (despesas, receitas, planos, etc). Deixar o usuário escolher o link. Deve ser ajustado através do painel Admin.
- [ ] Remover Hardcode da tela de mídias sociais (YouTube/Instagram) para que o usuário possa incluir redes e links através do painel Admin.
- [ ] Remover o Hardcode que controla o nível de graduação. Permitir cadastro de ID (ordem), cor e modalidade (Karate, Muay Thai, etc).
- [ ] Gestão de Turmas.
- [ ] Gestão de cadastro de Modalidades, associados aos locais de treino (Local + Modalidade + Instrutor).
- [ ] Permitir o Instrutor, Mestre ou Admin acompanhar o desenvolvimento do Aluno (aulas assistidas, tempo de inscrição) para avaliar elegibilidade para exames.
- [ ] Validar todos os arquivos do projeto (Pente fino em todos os arquivos).
- [ ] Validar se a planilha Mestra funciona.
- [ ] Remover todo o hardcode. Tudo deve ser gerenciado pelo painel Admin.
- [ ] **Código.js:** Alterar `.setTitle("FBKMK - Leão do Norte")` para buscar o nome da academia dinamicamente.
- [ ] **Ajuda.html:** Reescrever o manual com texto neutro de suporte técnico (remover menções específicas).
- [ ] **Social.html:** Transformar botões de Youtube e Instagram em links gerenciáveis/dinâmicos.
- [ ] **CardTemplate.html:** Recriar a carteirinha usando apenas HTML e CSS puro, para que a logo e as cores do cliente apareçam dinamicamente.

## 🚧 Fase 2: Gestão Financeira e Turmas (Em Andamento)
- [x] Dashboard Financeiro (Inadimplência, Faturamento).
- [ ] Separação de Turmas por Nível e Idade (Infantil, Iniciante, Avançado).
- [ ] Check-in Antecipado (Agendamento de grade de aulas).

## 🔮 Fase 3: Automação e CRM (Próximos Passos)
- [ ] Cobrança Automática Inteligente: Disparo de mensagens via WhatsApp (API) para atrasados.
- [ ] Alertas Automáticos de Aniversário.
- [ ] Relatórios Avançados (Gráficos de evasão, horários de pico).
- [ ] Histórico do Aluno (tempo na turma, aulas assistidas, conteúdo visualizado).
- [ ] Aluno Multimodalidade (Filtrar vídeos e programas técnicos baseados nas modalidades que o aluno pratica).
- [ ] Modularizar a tela de Dashboard (Botões essenciais, permissões de visualização por card).
- [ ] Validar se ao fazer a cópia da planilha (para venda), o SaaS continua a funcionar de forma autônoma.

## 🌐 Projetos Especiais
- [ ] Adaptação do sistema base para a **Confederação Pernambucana de Sambo**.