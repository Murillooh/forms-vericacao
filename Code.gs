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

  // Se a requisição for para obter os dados do usuário logado (SSO)
  if (e && e.parameter && e.parameter.action === "getUserData") {
    var dadosUsuario = obterDadosUsuarioLogado();
    return ContentService.createTextOutput(JSON.stringify(dadosUsuario))
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
      "Notebook Lenovo Prata",
      "Notebook HP"
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
  
  var acessoriosHtml = "";
  if (dados.acessorios) {
    var listaAcessorios = Array.isArray(dados.acessorios) ? dados.acessorios : dados.acessorios.toString().split(",");
    var listaLimpa = listaAcessorios.map(function(item) { return item.trim(); }).filter(function(item) { return item !== ""; });
    if (listaLimpa.length > 0) {
      acessoriosHtml = "<div class='section-title'>Acessórios Adicionais</div>" +
                       "<div class='info-grid'>";
      listaLimpa.forEach(function(acess) {
        acessoriosHtml += "<div class='info-row'><div class='info-label'>• " + acess + "</div><div class='info-value'>Entregue e atribuído</div></div>";
      });
      acessoriosHtml += "</div>";
    }
  }

  var itensFaltaHtml = "";
  if (dados.itensFalta && dados.itensFalta.trim() !== "") {
    itensFaltaHtml = "<div class='section-title'>Itens em Falta / Observações</div>" +
                     "<div style='font-size: 13px; color: #2d3748; background: #fffaf0; border: 1px solid #fbd38d; padding: 10px; border-radius: 6px; margin-bottom: 15px;'>" +
                       dados.itensFalta.trim() +
                     "</div>";
  }

  var assinaturaHtml = "";
  if (dados.assinatura) {
    var termoTexto = "Declaro ter recebido os equipamentos descritos neste termo em perfeitas condições de uso, assumindo a responsabilidade pela guarda, conservação e devolução dos mesmos ao término do vínculo de trabalho.";
    assinaturaHtml = "<div class='section-title'>Termo de Responsabilidade & Assinatura</div>" +
                     "<p style='font-size: 11px; color: #4a5568; margin-bottom: 20px;'>" + termoTexto + "</p>";
                     
    if (dados.foto) {
      assinaturaHtml += 
        "<div style='width: 100%; display: block; margin-top: 15px; border-top: 1px dashed #cbd5e0; padding-top: 15px;'>" +
          "<div style='display: inline-block; width: 45%; vertical-align: top; text-align: center; border-right: 1px solid #edf2f7; padding-right: 15px;'>" +
            "<div style='font-size: 11px; font-weight: bold; color: #4a5568; margin-bottom: 8px;'>Foto de Validação</div>" +
            "<img src='" + dados.foto + "' style='max-height: 110px; max-width: 100%; border-radius: 6px; border: 1.5px solid #cbd5e0;' />" +
          "</div>" +
          "<div style='display: inline-block; width: 45%; vertical-align: top; text-align: center; padding-left: 15px;'>" +
            "<div style='font-size: 11px; font-weight: bold; color: #4a5568; margin-bottom: 8px;'>Assinatura Digital</div>" +
            "<img src='" + dados.assinatura + "' style='max-height: 60px; max-width: 100%;' /><br/>" +
            "<span style='font-size: 11px; font-weight: bold; color: #2d3748;'>" + dados.nome + "</span><br/>" +
            "<span style='font-size: 9px; color: #718096;'>CPF: " + formatarCPF(dados.cpf) + "</span>" +
          "</div>" +
        "</div>";
    } else {
      assinaturaHtml += 
        "<div style='text-align: center; margin-top: 20px; border-top: 1px dashed #cbd5e0; padding-top: 10px;'>" +
          "<img src='" + dados.assinatura + "' style='max-height: 70px; max-width: 260px; display: block; margin: 0 auto 5px;' /><br/>" +
          "<span style='font-size: 12px; font-weight: bold; color: #2d3748;'>" + dados.nome + "</span><br/>" +
          "<span style='font-size: 10px; color: #718096;'>CPF: " + formatarCPF(dados.cpf) + "</span>" +
        "</div>";
    }
  }

  var htmlContent = 
    "<!DOCTYPE html><html><head><meta charset='utf-8'>" +
    "<style>" +
      "body { font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; color: #2d3748; line-height: 1.5; margin: 0; padding: 20px; background-color: #f7fafc; }" +
      ".container { max-width: 650px; margin: 0 auto; background: #ffffff; padding: 30px; border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); border-top: 6px solid #c9a449; }" +
      ".header { text-align: center; margin-bottom: 20px; border-bottom: 1px solid #e2e8f0; padding-bottom: 15px; }" +
      ".header h1 { font-size: 22px; color: #1a202c; margin: 0 0 6px; font-weight: 700; }" +
      ".header p { font-size: 12px; color: #718096; margin: 0; }" +
      ".section-title { font-size: 13px; text-transform: uppercase; letter-spacing: 0.05em; color: #c9a449; font-weight: 700; margin: 20px 0 10px; border-bottom: 1px solid #edf2f7; padding-bottom: 4px; }" +
      ".info-grid { display: table; width: 100%; margin-bottom: 15px; }" +
      ".info-row { display: table-row; }" +
      ".info-label { display: table-cell; font-weight: bold; width: 35%; padding: 6px 0; color: #4a5568; font-size: 13px; border-bottom: 1px solid #edf2f7; }" +
      ".info-value { display: table-cell; padding: 6px 0; color: #2d3748; font-size: 13px; border-bottom: 1px solid #edf2f7; }" +
      ".footer { text-align: center; margin-top: 30px; font-size: 10px; color: #a0aec0; border-top: 1px solid #e2e8f0; padding-top: 15px; }" +
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
          "<div class='info-row'><div class='info-label'>E-mail:</div><div class='info-value'>" + dados.email + "</div></div>" +
          "<div class='info-row'><div class='info-label'>Cargo:</div><div class='info-value'>" + dados.cargo + "</div></div>" +
          "<div class='info-row'><div class='info-label'>Unidade de Atuação:</div><div class='info-value'>" + dados.unidade + "</div></div>" +
        "</div>" +
        
        "<div class='section-title'>Notebook Atribuído</div>" +
        "<div class='info-grid'>" +
          "<div class='info-row'><div class='info-label'>Modelo do Notebook:</div><div class='info-value'>" + dados.modeloNotebook + "</div></div>" +
          "<div class='info-row'><div class='info-label'>Número de Patrimônio:</div><div class='info-value'>" + dados.patrimonioNotebook + "</div></div>" +
          "<div class='info-row'><div class='info-label'>Acompanha Carregador?</div><div class='info-value'>" + (dados.notebookCarregador || "Sim") + "</div></div>" +
        "</div>" +
        
        "<div class='section-title'>Celular Atribuído</div>" +
        "<div class='info-grid'>" +
          "<div class='info-row'><div class='info-label'>Modelo do Celular:</div><div class='info-value'>" + dados.modeloCelular + "</div></div>" +
          "<div class='info-row'><div class='info-label'>IMEI do Celular:</div><div class='info-value'>" + dados.imeiCelular + "</div></div>" +
          "<div class='info-row'><div class='info-label'>Acompanha Carregador?</div><div class='info-value'>" + (dados.celularCarregador || "Sim") + "</div></div>" +
        "</div>" +
        
        acessoriosHtml +
        
        itensFaltaHtml +
        
        assinaturaHtml +
        
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
  
  // Tenta alterar a permissão para qualquer pessoa com o link.
  // Se as políticas de segurança corporativas do domínio bloquearem, ignora o erro para garantir que a URL seja gerada e salva.
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (eShare) {
    try {
      // Tenta compartilhar apenas dentro da organização (Locagora) como fallback
      file.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW);
    } catch (eDomain) {
      Logger.log("Aviso: Falha ao definir permissão de compartilhamento. O arquivo permanecerá privado. Detalhes: " + eDomain.toString());
    }
  }
  
  return {
    url: file.getUrl(),
    blob: blob
  };
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
        "Carregador Notebook",
        "Modelo do Celular", 
        "IMEI do Celular",
        "Carregador Celular",
        "Acessórios",
        "Itens em Falta / Observações",
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
      
      // Recarrega headers e confere a coluna "Carregador Notebook"
      headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (headers.indexOf("Carregador Notebook") === -1) {
        var idxPatr = headers.indexOf("Patrimônio do Notebook");
        if (idxPatr !== -1) {
          sheet.insertColumnAfter(idxPatr + 1);
          sheet.getRange(1, idxPatr + 2).setValue("Carregador Notebook");
        } else {
          sheet.insertColumnBefore(sheet.getLastColumn());
          sheet.getRange(1, sheet.getLastColumn() - 1).setValue("Carregador Notebook");
        }
      }
      
      // Recarrega headers e confere a coluna "Carregador Celular"
      headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (headers.indexOf("Carregador Celular") === -1) {
        var idxImei = headers.indexOf("IMEI do Celular");
        if (idxImei !== -1) {
          sheet.insertColumnAfter(idxImei + 1);
          sheet.getRange(1, idxImei + 2).setValue("Carregador Celular");
        } else {
          sheet.insertColumnBefore(sheet.getLastColumn());
          sheet.getRange(1, sheet.getLastColumn() - 1).setValue("Carregador Celular");
        }
      }
      
      // Recarrega headers e confere a coluna Acessórios
      headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var hasAcessorios = headers.indexOf("Acessórios") !== -1;
      if (!hasAcessorios) {
        // Insere a coluna Acessórios antes da coluna de PDF (que costuma ser a última)
        var lastCol = sheet.getLastColumn();
        sheet.insertColumnBefore(lastCol);
        sheet.getRange(1, lastCol).setValue("Acessórios");
      }
      
      // Recarrega headers e confere a coluna "Itens em Falta / Observações"
      headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (headers.indexOf("Itens em Falta / Observações") === -1) {
        var idxLink = headers.indexOf("Link do PDF de Recibo");
        if (idxLink !== -1) {
          sheet.insertColumnBefore(idxLink + 1);
          sheet.getRange(1, idxLink + 1).setValue("Itens em Falta / Observações");
        } else {
          sheet.insertColumnBefore(sheet.getLastColumn() + 1);
          sheet.getRange(1, sheet.getLastColumn()).setValue("Itens em Falta / Observações");
        }
      }
      

      
      // Garante que tem a coluna de Recibo PDF no final
      headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (headers.indexOf("Link do PDF de Recibo") === -1) {
        sheet.getRange(1, sheet.getLastColumn() + 1).setValue("Link do PDF de Recibo");
      }
    }

    // Aplica a formatação do cabeçalho (Fundo Azul e Letras Douradas)
    var cabecalhoRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    cabecalhoRange.setFontWeight("bold");
    cabecalhoRange.setBackground("#0f4c81"); // Azul Corporativo
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
    
    // Validação matemática de CPF no servidor
    if (!validarCPF(dados.cpf)) {
      return {
        sucesso: false,
        mensagem: "O CPF fornecido é inválido matematicamente."
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
    
    // Validação de Patrimônio do Notebook (deve conter exatamente 8 caracteres)
    if (dados.patrimonioNotebook.trim().length !== 8) {
      return {
        sucesso: false,
        mensagem: "O patrimônio do notebook deve conter exatamente 8 caracteres."
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
    
    // Prevenção de duplicidades por Patrimônio de Notebook ou IMEI de Celular
    var patrimonioLimpo = dados.patrimonioNotebook.trim().toLowerCase();
    if (sheet.getLastRow() > 1) {
      var headersVerif = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var idxPatrimonio = headersVerif.indexOf("Patrimônio do Notebook");
      var idxImei = headersVerif.indexOf("IMEI do Celular");
      var valuesVerif = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
      
      for (var i = 0; i < valuesVerif.length; i++) {
        var row = valuesVerif[i];
        if (idxPatrimonio !== -1 && row[idxPatrimonio]) {
          if (row[idxPatrimonio].toString().trim().toLowerCase() === patrimonioLimpo) {
            return {
              sucesso: false,
              mensagem: "Erro: O Notebook com o Patrimônio '" + dados.patrimonioNotebook + "' já está cadastrado no sistema."
            };
          }
        }
        if (idxImei !== -1 && row[idxImei]) {
          if (row[idxImei].toString().replace(/\D/g, '') === imeiLimpo) {
            return {
              sucesso: false,
              mensagem: "Erro: O Celular com o IMEI '" + dados.imeiCelular + "' já está cadastrado no sistema."
            };
          }
        }
      }
    }
    


    // Gera o recibo PDF e obtém o link do arquivo
    var pdfUrl = "";
    var pdfBlob = null;
    try {
      var resultadoPdf = gerarPDF(dados);
      pdfUrl = resultadoPdf.url;
      pdfBlob = resultadoPdf.blob;
    } catch (ePdf) {
      Logger.log("Erro ao gerar PDF: " + ePdf.toString());
    }
    
    // Recarrega os headers definitivos para mapeamento dinâmico de inserção
    var headersFinais = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var novaLinha = [];
    
    for (var k = 0; k < headersFinais.length; k++) {
      var headerName = headersFinais[k];
      switch (headerName) {
        case "Data/Hora de Registro":
          novaLinha.push(new Date());
          break;
        case "Nome Completo":
          novaLinha.push(dados.nome.trim());
          break;
        case "CPF":
          novaLinha.push(formatarCPF(dados.cpf));
          break;
        case "E-mail Corporativo":
          novaLinha.push(dados.email.trim());
          break;
        case "Cargo":
          novaLinha.push(dados.cargo.trim());
          break;
        case "Unidade":
          novaLinha.push(dados.unidade.trim());
          break;
        case "Modelo do Notebook":
          novaLinha.push(dados.modeloNotebook);
          break;
        case "Patrimônio do Notebook":
          novaLinha.push(dados.patrimonioNotebook.trim());
          break;
        case "Carregador Notebook":
          novaLinha.push(dados.notebookCarregador || "Sim");
          break;
        case "Modelo do Celular":
          novaLinha.push(dados.modeloCelular);
          break;
        case "IMEI do Celular":
          novaLinha.push(imeiLimpo);
          break;
        case "Carregador Celular":
          novaLinha.push(dados.celularCarregador || "Sim");
          break;
        case "Acessórios":
          var acessString = "";
          if (dados.acessorios) {
            acessString = Array.isArray(dados.acessorios) ? dados.acessorios.join(", ") : dados.acessorios.toString();
          }
          novaLinha.push(acessString);
          break;
        case "Itens em Falta / Observações":
          novaLinha.push(dados.itensFalta ? dados.itensFalta.trim() : "");
          break;

        case "Link do PDF de Recibo":
          novaLinha.push(pdfUrl);
          break;
        default:
          novaLinha.push("");
          break;
      }
    }
    
    // Insere os dados
    sheet.appendRow(novaLinha);
    
    // Envio automático do termo por e-mail usando GmailApp se o PDF foi gerado
    if (pdfBlob) {
      try {
        GmailApp.sendEmail(
          dados.email.trim(),
          "Recibo de Atribuição de Ativos - " + dados.nome.trim(),
          "Olá " + dados.nome.trim() + ",\n\n" +
          "Confirmamos a recepção e o registro dos seus equipamentos periféricos no inventário da empresa.\n\n" +
          "Anexo a este e-mail está o PDF do seu Recibo de Atribuição de Ativos / Termo de Responsabilidade assinado digitalmente.\n\n" +
          "Atenciosamente,\n" +
          "TI · Locagora Periféricos",
          {
            attachments: [pdfBlob]
          }
        );
      } catch (eEmail) {
        Logger.log("Erro ao enviar e-mail automático: " + eEmail.toString());
      }
    }
    
    // Aplica a formatação das linhas de dados (Fundo Preto e Letras Douradas)
    if (sheet.getLastRow() > 1) {
      var dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
      dataRange.setBackground("#111111"); // Fundo Preto
      dataRange.setFontColor("#ffd97d"); // Letras Douradas
      dataRange.setHorizontalAlignment("left");
      
      // Centraliza colunas específicas para alinhamento profissional
      for (var colIdx = 0; colIdx < headersFinais.length; colIdx++) {
        var hName = headersFinais[colIdx];
        if (hName === "Data/Hora de Registro" || hName === "CPF" || hName === "Link do PDF de Recibo" || hName === "IMEI do Celular") {
          sheet.getRange(2, colIdx + 1, sheet.getLastRow() - 1, 1).setHorizontalAlignment("center");
        }
      }
    }
    
    // Auto-ajusta as colunas
    try {
      sheet.autoResizeColumns(1, sheet.getLastColumn());
    } catch(e) {}
    
    return {
      sucesso: true,
      mensagem: "Cadastro de periféricos efetuado e termo enviado por e-mail!"
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

/**
 * Valida matematicamente um número de CPF.
 */
function validarCPF(cpf) {
  var cpfLimpo = cpf.replace(/\D/g, '');
  if (cpfLimpo.length !== 11) return false;
  
  // Impede CPFs conhecidos de dígitos repetidos
  if (/^(\d)\1{10}$/.test(cpfLimpo)) return false;
  
  var soma = 0;
  var resto;
  
  for (var i = 1; i <= 9; i++) {
    soma += parseInt(cpfLimpo.substring(i - 1, i)) * (11 - i);
  }
  resto = (soma * 10) % 11;
  if (resto === 10 || resto === 11) resto = 0;
  if (resto !== parseInt(cpfLimpo.substring(9, 10))) return false;
  
  soma = 0;
  for (var i = 1; i <= 10; i++) {
    soma += parseInt(cpfLimpo.substring(i - 1, i)) * (12 - i);
  }
  resto = (soma * 10) % 11;
  if (resto === 10 || resto === 11) resto = 0;
  if (resto !== parseInt(cpfLimpo.substring(10, 11))) return false;
  
  return true;
}

/**
 * Retorna os dados prováveis do usuário logado (SSO) com base na sessão ativa do domínio.
 */
function obterDadosUsuarioLogado() {
  try {
    var email = Session.getActiveUser().getEmail();
    var nome = "";
    if (email) {
      var username = email.split('@')[0];
      // Substitui pontos e hífens por espaço e capitaliza cada palavra
      var partes = username.split(/[._-]/);
      nome = partes.map(function(parte) {
        if (!parte) return "";
        return parte.charAt(0).toUpperCase() + parte.slice(1).toLowerCase();
      }).filter(Boolean).join(' ');
    }
    return {
      sucesso: true,
      email: email || "",
      nome: nome || ""
    };
  } catch (e) {
    return {
      sucesso: false,
      mensagem: "Falha na sessão do Workspace (SSO): " + e.toString()
    };
  }
}