// Esta constante é escrita na coluna C para linhas nas quais um e-mail
// foi enviado com sucesso.
function unknown() {
    const ss = SpreadsheetApp.getActive();// Obtém a planilha ativa do google sheets
    const sh = ss.getSheetByName("cadastros");// Obtém a folha de cálculo com o nome  "cadastros"
    var rg = sh.getDataRange();// Obtém a faixa de dados na folha de trabalho
    var vs = rg.getValues();  // Obtém os valores da faixa de dados
    const esh = ss.getSheetByName("email");  // Obtém a folha de cálculo com o nome "email"
    var message = esh.getRange("A4").getValue();  // Obtém o conteúdo da célula A4 na planilha "email"
    var t = esh.getRange("A2").getValue();  // Obtém o conteúdo da célula A2 na planilha "email"
  
    for (var i = 1; i < vs.length; i++) {  // Itera sobre as linhas da faixa de dados a partir da segunda linha (índice 1)
      var row = vs[i];    // Obtém os valores da linha(row) atual
      var name = row[0];
      var emailDestination = row[1]; 
      var emailSent = row[2]; 
      if (emailDestination && emailSent != EMAIL_SENT) { 
        var subjectTemplate = t;//Reinicializar o template antes cada troca (Replace)
        var subject = subjectTemplate.replace("<nome>", name);
        var message2 = message.replace("<nome>", name); // Replaces the <nome> placeholder with the name from the sheet
        sendEmailWithLogo(emailDestination, subject, message2, logoUrl= "https://media.licdn.com/dms/image/C4E12AQHwckJ-hv-0rg/article-cover_image-shrink_600_2000/0/1632191739115?e=2147483647&v=beta&t=CbEz6-6qwHRuHy65RGXc2BOzy7Fix_eRlh_3dZFut2Q"); // The email is sent along with the logo (logoUrl)
        ss.getSheetByName("cadastros").getRange(1 + i, 3).setValue(EMAIL_SENT); // The column with status (emailsent) is marked as EMAIL_SENT, avoiding duplicates
        Logger.log("Email sent to " + emailDestination + " successfully!") // Logged in console that the email was sent successfully
        SpreadsheetApp.flush();
      }
      else {
        Logger.log("No new email to send!")
      }
    }
  }
  var EMAIL_SENT = 'EMAIL_SENT';
  
  // Função principal para enviar e-mails
  function sendEmail() {
    // URL do logo a ser incluído no e-mail (Lembrar de usar uma logo em .png)
    var logoUrl = "https://media.licdn.com/dms/image/C4E12AQHwckJ-hv-0rg/article-cover_image-shrink_600_2000/0/1632191739115?e=2147483647&v=beta&t=CbEz6-6qwHRuHy65RGXc2BOzy7Fix_eRlh_3dZFut2Q";
    var ss = SpreadsheetApp.getActiveSpreadsheet();
  
    // Obter o endereço do email
    var dataRange = ss.getSheetByName("cadastros").getDataRange();
  
    // Faz com que seja obtido o assunto e a mensagem do email (A2/A4)
    var subjectTemplate = ss.getSheetByName("email").getRange("A2").getValue(); // Refere-se à parte de Título do assunto, nos sheets posicionado em A2
    var message = ss.getSheetByName("email").getRange("A4").getValue(); // Refere-se à parte de e-mail, nos sheets posicionado em A4
  
    // Obtem valores para cada linha no alcance determinado
    var data = dataRange.getValues();
  
    // Loop através das linhas da planilha "cadastros"
    for (var i = 1; i < data.length; ++i) {
      var row = data[i];
      var name = row[0]; // Primeira coluna do planilha (Nome)
      var emailDestination = row[1]; // Segunda coluna da planilha (Email)
      var emailSent = row[2]; // Terceira Coluna da planilha (status)
  
      // Se o email estiver vazio, dá um break
      if (emailDestination == "") {
        break;
      }
  
      // Verifica se o e-mail ainda não foi enviado
      if (emailSent != EMAIL_SENT) {
        
        var subject = subjectTemplate.replace("<nome>", name);
        var message2 = message.replace("<nome>", name);// Troca a marcação <nome> pelo nome que está escrito na planilha
  
  
        // Envie o e-mail com o logotipo
        sendEmailWithLogo(emailDestination, subject, message2, logoUrl);
  
        // Marque a coluna 'Status' como 'EMAIL_SENT', evitando duplicatas
        ss.getSheetByName("cadastros").getRange(1 + i, 3).setValue(EMAIL_SENT);
  
        // Registre que o e-mail foi enviado com sucesso no log/console
        Logger.log("Email enviado para " + emailDestination + " com sucesso!");
  
        // Garanta que a célula seja atualizada imediatamente, caso o script seja interrompido
        SpreadsheetApp.flush();
      } else {
        Logger.log("Nenhum novo email para enviar!");
      }
    }
  }
  
  // Função para enviar e-mail com logotipo
  function sendEmailWithLogo(emailDestination, subject, message, logoUrl) {
    // Busca o Blob do logotipo
    var logoBlob = UrlFetchApp.fetch(logoUrl).getBlob().setName("logoBlob");
    // Envia o e-mail
    MailApp.sendEmail({
      to: emailDestination,
      subject: subject,
      htmlBody: message + "<br><br><img src='cid:logo' align='middle' style='width:100px; height:100px;'>",
      inlineImages: {
        logo: logoBlob,
      },
    });
  }
  
  // Função para criar o acionador
  function createTrigger() {
    // Esta função cria um acionador para executar a função sendEmail a cada semana, às 15 horas
    ScriptApp.newTrigger('sendEmail')
      .timeBased()
      .atHour(15)
      .everyDays(7)
      .create();
  }