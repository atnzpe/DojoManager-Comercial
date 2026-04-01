function gerarApresentacaoComercialVIP() {
  // 1. Cria o arquivo no Google Drive
  const apresentacao = SlidesApp.create("Pitch VIP - DojoManager SaaS");
  
  // 🎨 PALETA B2B PREMIUM (Psicologia de Vendas Tech)
  const corFundo = "#0B1120";         // Fundo super escuro (Deep Space - Elegância)
  const corCabecalho = "#1E293B";     // Fundo do título (Slate - Modernidade)
  const corTextoTitulo = "#F59E0B";   // Dourado/Âmbar (Chama a atenção, Poder)
  const corTextoCorpo = "#F8FAFC";    // Branco Gelo (Leitura limpa e confortável)

  // 3. O Roteiro Estratégico de Alta Conversão
  const roteiro = [
    {
      titulo: "DojoManager SaaS\nA Evolução do Seu Tatame",
      corpo: "Gestão, Retenção e Vendas na palma da mão do seu aluno.\n\nTransforme a sua academia com tecnologia de ponta."
    },
    {
      titulo: "O Problema (A Dor Invisível)",
      corpo: "• Inadimplência invisível e constrangedora de cobrar.\n• Alunos desmotivados sem clareza da sua evolução.\n• Planilhas confusas e trabalho manual na secretaria.\n• Falta de padrão e profissionalismo na gestão de filiais."
    },
    {
      titulo: "A Solução (Ecossistema Integrado)",
      corpo: "• Um aplicativo com a SUA marca (White-Label).\n• Muito mais que gestão: é um portal de ensino e comunidade.\n• Financeiro, CRM, E-learning e Catraca Virtual 100% conectados."
    },
    {
      titulo: "O App do Aluno (Foco em Retenção)",
      corpo: "• Carteirinha Digital Oficial com QR Code.\n• Videoteca inteligente: filtrada pela faixa e modalidade do aluno.\n• Programas Técnicos (Manuais e PDFs de exame).\n• Emissão automática de Certificados e Diplomas."
    },
    {
      titulo: "Inteligência Financeira (O Fim do Atraso)",
      corpo: "• Pagamentos via PIX 'Copia e Cola' direto no App do aluno.\n• Renovação automática de planos no caixa.\n• Painel de Cobrança: Disparo rápido para WhatsApp com 1 clique.\n• Dashboard de LTV (Lifetime Value) e Previsão de Receita."
    },
    {
      titulo: "Gestão de Tatame (Para os Mestres)",
      corpo: "• Diário de aula dinâmico: associe vídeos e técnicas ao treino.\n• Check-in por PIN com trava de idade de segurança.\n• Dossiê Analítico do Atleta: Raio-X completo de faltas e engajamento."
    },
    {
      titulo: "Compliance e Ouvidoria Segura",
      corpo: "• Proteção total dos dados (LGPD).\n• Canal sigiloso para denúncias, assédio e doping.\n• Modo de denúncia 100% anônima para a segurança do aluno.\n• Roteamento criptografado exclusivo para a diretoria."
    },
    {
      titulo: "A Máquina de Vendas (Leads)",
      corpo: "• Captação de novos alunos 24h por dia.\n• Link de 'Aula Experimental' integrado ao Instagram da academia.\n• O visitante agenda sozinho e o WhatsApp da secretaria é notificado na hora."
    },
    {
      titulo: "Escalabilidade Multi-Franquias",
      corpo: "• O controle total da sua rede num único lugar.\n• O Master vê tudo. Os franqueados gerenciam apenas a própria unidade.\n• Filtros avançados no Super Relatório Administrativo."
    },
    {
      titulo: "Leve a sua academia para o próximo nível!",
      corpo: "Próximos Passos:\n\n• Agende uma demonstração prática do sistema.\n• Conheça as nossas condições especiais de implantação.\n\nDojoManager SaaS - O braço digital da sua academia."
    }
  ];

  // 4. Pega o primeiro slide
  const slideCapa = apresentacao.getSlides()[0];
  
  // 5. Motor de Criação
  roteiro.forEach((slideData, index) => {
    let slideCorrente;
    
    // Layout de Capa para o primeiro, Título+Corpo para os restantes
    if (index === 0) {
      slideCorrente = slideCapa;
    } else {
      slideCorrente = apresentacao.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
    }
    
    // Pinta o Fundo de Deep Space
    slideCorrente.getBackground().setSolidFill(corFundo);

    // Formata a Caixa de TÍTULO
    const placeholderTitulo = slideCorrente.getPlaceholder(index === 0 ? SlidesApp.PlaceholderType.CENTERED_TITLE : SlidesApp.PlaceholderType.TITLE);
    if (placeholderTitulo) {
      // 🛡️ MÁGICA AQUI: Converte o Placeholder para Shape
      const shapeTitulo = placeholderTitulo.asShape();
      
      // Pinta a caixa do título
      shapeTitulo.getFill().setSolidFill(corCabecalho);
      
      const textoTitulo = shapeTitulo.getText();
      textoTitulo.setText(slideData.titulo);
      
      // Estilo do Título
      const estiloTitulo = textoTitulo.getTextStyle();
      estiloTitulo.setFontFamily("Montserrat"); 
      estiloTitulo.setForegroundColor(corTextoTitulo);
      estiloTitulo.setBold(true);
    }

    // Formata a Caixa de CORPO
    const placeholderCorpo = slideCorrente.getPlaceholder(index === 0 ? SlidesApp.PlaceholderType.SUBTITLE : SlidesApp.PlaceholderType.BODY);
    if (placeholderCorpo) {
      // 🛡️ MÁGICA AQUI: Converte o Placeholder para Shape
      const shapeCorpo = placeholderCorpo.asShape();
      
      const textoCorpo = shapeCorpo.getText();
      textoCorpo.setText(slideData.corpo);
      
      // Estilo do Corpo
      const estiloCorpo = textoCorpo.getTextStyle();
      estiloCorpo.setFontFamily("Roboto");
      estiloCorpo.setForegroundColor(corTextoCorpo);
      
      // Ajusta o parágrafo para ficar mais espaçado (pula linha se não for a capa)
      if (index > 0) {
          textoCorpo.getParagraphStyle().setLineSpacing(130);
      }
    }
  });

  console.log("✅ APRESENTAÇÃO VIP GERADA!");
  console.log("🔗 Link: " + apresentacao.getUrl());
}