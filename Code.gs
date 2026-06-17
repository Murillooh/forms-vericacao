/**
 * TI · Inventário Corporativo
 * Projeto: Cadastro de Periféricos (Locagora)
 * 
 * Este arquivo contém as funções do servidor executadas no ambiente do Google Apps Script.
 */

/**
 * Serve o formulário HTML quando a URL da Web App for acessada ou retorna opções de equipamentos em formato JSON.
 */
function doGet(e) {
  // Se a requisição vier do GitHub Pages ou Localhost para obter as opções de equipamentos
  if (e && e.parameter && e.parameter.action === "getOptions") {
    var opcoes = getOpcoesEquipamentos();
    return ContentService.createTextOutput(JSON.stringify(opcoes))
        .setMimeType(ContentService.MimeType.JSON);
  }

  var cacheBuster = new Date().getTime();
  var url = "https://raw.githubusercontent.com/Murillooh/forms-vericacao/main/index.html?cb=" + cacheBuster;
  var response = UrlFetchApp.fetch(url);
  var html = response.getContentText();
  
  return HtmlService.createHtmlOutput(html)
      .setTitle('Cadastro de Periféricos')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Ponto de entrada para receber requisições POST externas do GitHub Pages ou Localhost.
 */
function doPost(e) {
  try {
    var dados;
    // Se o payload vier como JSON stringificado em text/plain
    if (e && e.postData && e.postData.contents) {
      dados = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      dados = e.parameter;
    } else {
      throw new Error("Nenhum dado recebido no payload da requisição.");
    }
    
    // Processa o cadastro
    var resposta = processarCadastro(dados);
    
    return ContentService.createTextOutput(JSON.stringify(resposta))
        .setMimeType(ContentService.MimeType.JSON);
        
  } catch (erro) {
    var respostaErro = {
      sucesso: false,
      mensagem: "Erro ao processar cadastro (CORS/API): " + erro.toString()
    };
    return ContentService.createTextOutput(JSON.stringify(respostaErro))
        .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Retorna as opções de equipamentos para popular as seleções de Notebook e Celular.
 * Tenta ler de uma aba chamada "Opcoes" na planilha ativa. Se a aba ou planilha
 * não existir, retorna opções padrão/fallback.
 * 
 * @return {Object} Objeto contendo listas de notebooks e celulares
 */
function getOpcoesEquipamentos() {
  try {
    var ss;
    try {
      ss = SpreadsheetApp.openById("14QIAJQARsNntOpgsxSy6qFzfTPY1c74UPXSgCgOmHAA");
    } catch(e) {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    }
    if (ss) {
      var sheet = ss.getSheetByName('Opcoes');
      if (sheet) {
        var data = sheet.getDataRange().getValues();
        var notebooks = [];
        var celulares = [];
        
        // Assume que a primeira linha (índice 0) é o cabeçalho: "Notebooks", "Celulares"
        for (var i = 1; i < data.length; i++) {
          var row = data[i];
          if (row[0] && row[0].toString().trim() !== "") {
            notebooks.push(row[0].toString().trim());
          }
          if (row[1] && row[1].toString().trim() !== "") {
            celulares.push(row[1].toString().trim());
          }
        }
        
        // Se encontramos dados, retorna eles
        if (notebooks.length > 0 || celulares.length > 0) {
          return {
            notebooks: notebooks,
            celulares: celulares
          };
        }
      }
    }
  } catch (e) {
    Logger.log("Aviso: Não foi possível carregar as opções da planilha. Usando valores padrão. Detalhes: " + e.toString());
  }
  
  // Opções fallback padrão caso a planilha ou aba não estejam configuradas
  return {
    notebooks: [
      "Notebook Lenovo preto",
      "Notebook Lenovo Prata"
    ],
    celulares: [
      "Moto G24",
      "Poco C65",
      "Redmi 13C"
    ]
  };
}

/**
 * Gera um recibo PDF estilizado com as informações cadastradas e salva no Google Drive.
 * 
 * @param {Object} dados Dados recebidos do formulário
 * @return {String} URL do arquivo PDF gerado
 */
function gerarPDF(dados) {
  var dataFormatada = new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
  
  var htmlContent = 
    "<!DOCTYPE html><html><head><meta charset='utf-8'>" +
    "<style>" +
      "body { font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; color: #2d3748; line-height: 1.6; margin: 0; padding: 40px; background-color: #f7fafc; }" +
      ".container { max-width: 650px; margin: 0 auto; background: #ffffff; padding: 40px; border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); border-top: 6px solid #c9a449; }" +
      ".header { text-align: center; margin-bottom: 30px; border-bottom: 1px solid #e2e8f0; padding-bottom: 20px; }" +
      ".header h1 { font-size: 24px; color: #1a202c; margin: 0 0 8px; font-weight: 700; }" +
      ".header p { font-size: 13px; color: #718096; margin: 0; }" +
      ".section-title { font-size: 15px; text-transform: uppercase; letter-spacing: 0.05em; color: #c9a449; font-weight: 700; margin: 25px 0 12px; border-bottom: 1px solid #edf2f7; padding-bottom: 6px; }" +
      ".info-grid { display: table; width: 100%; margin-bottom: 20px; }" +
      ".info-row { display: table-row; }" +
      ".info-label { display: table-cell; font-weight: bold; width: 35%; padding: 8px 0; color: #4a5568; font-size: 14px; border-bottom: 1px solid #edf2f7; }" +
      ".info-value { display: table-cell; padding: 8px 0; color: #2d3748; font-size: 14px; border-bottom: 1px solid #edf2f7; }" +
      ".footer { text-align: center; margin-top: 40px; font-size: 11px; color: #a0aec0; border-top: 1px solid #e2e8f0; padding-top: 20px; }" +
    "</style>" +
    "</head><body>" +
      "<div class='container'>" +
        "<div class='header'>" +
          "<h1>Recibo de Atribuição de Ativos</h1>" +
          "<p>TI · Inventário Locagora Periféricos</p>" +
        "</div>" +
        
        "<div class='section-title'>Dados Pessoais</div>" +
        "<div class='info-grid'>" +
          "<div class='info-row'><div class='info-label'>Nome Completo:</div><div class='info-value'>" + dados.nome + "</div></div>" +
          "<div class='info-row'><div class='info-label'>CPF:</div><div class='info-value'>" + formatarCPF(dados.cpf) + "</div></div>" +
          "<div class='info-row'><div class='info-label'>E-mail Corporativo:</div><div class='info-value'>" + dados.email + "</div></div>" +
          "<div class='info-row'><div class='info-label'>Cargo:</div><div class='info-value'>" + dados.cargo + "</div></div>" +
          "<div class='info-row'><div class='info-label'>Unidade de Atuação:</div><div class='info-value'>" + dados.unidade + "</div></div>" +
        "</div>" +
        
        "<div class='section-title'>Notebook Atribuído</div>" +
        "<div class='info-grid'>" +
          "<div class='info-row'><div class='info-label'>Modelo do Notebook:</div><div class='info-value'>" + dados.modeloNotebook + "</div></div>" +
          "<div class='info-row'><div class='info-label'>Número de Patrimônio:</div><div class='info-value'>" + dados.patrimonioNotebook + "</div></div>" +
        "</div>" +
        
        "<div class='section-title'>Celular Atribuído</div>" +
        "<div class='info-grid'>" +
          "<div class='info-row'><div class='info-label'>Modelo do Celular:</div><div class='info-value'>" + dados.modeloCelular + "</div></div>" +
          "<div class='info-row'><div class='info-label'>IMEI do Celular:</div><div class='info-value'>" + dados.imeiCelular + "</div></div>" +
        "</div>" +
        
        "<div class='footer'>" +
          "<p>Documento de controle interno registrado em " + dataFormatada + "</p>" +
          "<p>Este documento serve como comprovante digital de termo de responsabilidade de uso de ativos.</p>" +
        "</div>" +
      "</div>" +
    "</body></html>";

  var blob = HtmlService.createHtmlOutput(htmlContent).getAs('application/pdf');
  blob.setName("Recibo_Ativos_" + dados.nome.replace(/\s+/g, "_") + "_" + new Date().getTime() + ".pdf");

  var folder = obterOuCriarPasta();
  var file = folder.createFile(blob);
  
  // Define permissão para qualquer pessoa com o link visualizar
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  return file.getUrl();
}

/**
 * Obtém ou cria a pasta de recebidos no Google Drive.
 */
function obterOuCriarPasta() {
  var folderName = "Recibos_Cadastro_Locagora";
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(folderName);
}

/**
 * Processa a submissão de cadastro e grava em uma linha na planilha ativa,
 * em uma aba chamada "Respostas". Cria e formata a aba automaticamente se não existir.
 * 
 * @param {Object} dados Dados recebidos do formulário
 * @return {Object} Status do processamento
 */
function processarCadastro(dados) {
  try {
    var ss;
    try {
      ss = SpreadsheetApp.openById("14QIAJQARsNntOpgsxSy6qFzfTPY1c74UPXSgCgOmHAA");
    } catch(e) {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    }
    if (!ss) {
      return {
        sucesso: false,
        mensagem: "Erro no Servidor: Não foi possível abrir a Planilha do Google."
      };
    }
    
    // Tenta obter a aba "Respostas". Se não existir, cria e formata
    var sheet = ss.getSheetByName('Respostas');
    if (!sheet) {
      sheet = ss.insertSheet('Respostas');
      
      // Cabeçalho da tabela
      var cabecalhos = [
        "Data/Hora de Registro", 
        "Nome Completo", 
        "CPF",
        "E-mail Corporativo", 
        "Cargo", 
        "Unidade", 
        "Modelo do Notebook", 
        "Patrimônio do Notebook", 
        "Modelo do Celular", 
        "IMEI do Celular",
        "Link do PDF de Recibo"
      ];
      
      sheet.appendRow(cabecalhos);
      sheet.setFrozenRows(1);
    } else {
      // Garante que o cabeçalho tem o CPF
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var hasCpf = headers.indexOf("CPF") !== -1;
      if (!hasCpf) {
        // Insere a coluna CPF na posição 3 (depois de Nome Completo)
        sheet.insertColumnBefore(3);
        sheet.getRange(1, 3).setValue("CPF");
      }
      
      // Garante que tem a coluna de Recibo PDF no final
      if (sheet.getLastColumn() < 11) {
        var headersAtualizados = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        if (headersAtualizados.indexOf("Link do PDF de Recibo") === -1) {
          sheet.getRange(1, sheet.getLastColumn() + 1).setValue("Link do PDF de Recibo");
        }
      }
    }

    // Aplica a formatação do cabeçalho (Fundo Azul e Letras Douradas)
    var cabecalhoRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    cabecalhoRange.setFontWeight("bold");
    cabecalhoRange.setBackground("#0f4c81"); // Azul Corporativo (Classic Blue)
    cabecalhoRange.setFontColor("#dfb247"); // Letras Douradas
    cabecalhoRange.setHorizontalAlignment("center");
    
    // Validações de segurança no lado do servidor
    if (!dados.nome || !dados.nome.trim() ||
        !dados.cpf || !dados.cpf.trim() ||
        !dados.email || !dados.email.trim() ||
        !dados.cargo || !dados.cargo.trim() ||
        !dados.unidade || !dados.unidade.trim() ||
        !dados.modeloNotebook ||
        !dados.patrimonioNotebook || !dados.patrimonioNotebook.trim() ||
        !dados.modeloCelular ||
        !dados.imeiCelular || !dados.imeiCelular.trim()) {
      return {
        sucesso: false,
        mensagem: "Todos os campos do formulário são de preenchimento obrigatório."
      };
    }
    
    // Validação de CPF no servidor
    var cpfLimpo = dados.cpf.replace(/\D/g, '');
    if (cpfLimpo.length !== 11) {
      return {
        sucesso: false,
        mensagem: "O CPF deve possuir exatamente 11 dígitos numéricos."
      };
    }
    
    // Validação de e-mail simplificada no servidor
    var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(dados.email.trim())) {
      return {
        sucesso: false,
        mensagem: "O e-mail corporativo fornecido é inválido."
      };
    }
    
    // Validação de IMEI (deve conter exatamente 15 dígitos numéricos)
    var imeiLimpo = dados.imeiCelular.replace(/\D/g, '');
    if (imeiLimpo.length !== 15) {
      return {
        sucesso: false,
        mensagem: "O IMEI celular deve possuir exatamente 15 dígitos numéricos."
      };
    }
    
    // Gera o recibo PDF e obtém o link do arquivo
    var pdfUrl = "";
    try {
      pdfUrl = gerarPDF(dados);
    } catch (ePdf) {
      Logger.log("Erro ao gerar PDF: " + ePdf.toString());
    }
    
    // Insere o registro de cadastro com timestamp
    sheet.appendRow([
      new Date(),
      dados.nome.trim(),
      formatarCPF(dados.cpf),
      dados.email.trim(),
      dados.cargo.trim(),
      dados.unidade.trim(),
      dados.modeloNotebook,
      dados.patrimonioNotebook.trim(),
      dados.modeloCelular,
      imeiLimpo,
      pdfUrl
    ]);
    
    // Aplica a formatação das linhas de dados (Fundo Preto e Letras Douradas)
    if (sheet.getLastRow() > 1) {
      var dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
      dataRange.setBackground("#111111"); // Fundo Preto (Suave/Elegante)
      dataRange.setFontColor("#ffd97d"); // Letras Douradas Brilhantes
      dataRange.setHorizontalAlignment("left");
      
      // Centraliza coluna de Data/Hora, CPF e link do PDF
      var dataHoraRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1);
      dataHoraRange.setHorizontalAlignment("center");
      
      var cpfCellRange = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1);
      cpfCellRange.setHorizontalAlignment("center");
      
      var linkPdfRange = sheet.getRange(2, 11, sheet.getLastRow() - 1, 1);
      linkPdfRange.setHorizontalAlignment("center");
    }
    
    // Auto-ajusta as colunas após inserção para legibilidade
    try {
      sheet.autoResizeColumns(1, 11);
    } catch(e) {}
    
    return {
      sucesso: true,
      mensagem: "Cadastro de periféricos efetuado e registrado com sucesso!"
    };
    
  } catch (erro) {
    return {
      sucesso: false,
      mensagem: "Erro crítico ao gravar dados na planilha: " + erro.toString()
    };
  }
}

/**
 * Formata uma string de CPF para o padrão 000.000.000-00.
 */
function formatarCPF(cpf) {
  var cpfLimpo = cpf.replace(/\D/g, '');
  if (cpfLimpo.length === 11) {
    return cpfLimpo.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, "$1.$2.$3-$4");
  }
  return cpf;
}