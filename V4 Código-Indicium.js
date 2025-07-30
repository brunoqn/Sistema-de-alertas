// VERSÃO FINAL COMPLETA - Ciclo de Vida do Acompanhamento (0 a 90 dias)

function executarAlerta() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = planilha.getDataRange();
  const dados = range.getValues();
  
  const fusoHorario = Session.getScriptTimeZone();
  const hojeString = Utilities.formatDate(new Date(), fusoHorario, "dd/MM/yyyy");
  
  const seuEmail = "fabriciapinheiro24@gmail.com";

  Logger.log("--- Iniciando execução do Ciclo de Acompanhamento Completo ---");
  Logger.log("Data de hoje para comparação: " + hojeString);

  const cabecalho = dados[0];
  // --- DEFINIÇÃO DAS COLUNAS ---
  const COL_NOME_AFILHADO = 0;   // Coluna A
  const COL_EMAIL_AFILHADO = 1;    // Coluna B
  const COL_NOME_PADRINHO = 2;     // Coluna C <-- AJUSTE REALIZADO AQUI
  const COL_DATA_ADMISSAO = 6;     // Coluna G
  const COL_EMAIL_PADRINHO = 24;   // Coluna Y
  const COL_EMAIL_LIDERANCA = 25;  // Coluna Z
  
  const COL_30_DIAS = encontrarIndiceColuna("Av. Exp.\n30 dias", cabecalho);
  const COL_38_DIAS = encontrarIndiceColuna("Alerta liderança\n38 dias", cabecalho);
  const COL_75_DIAS = encontrarIndiceColuna("Av. Exp.\n75 dias", cabecalho);
  const COL_83_DIAS = encontrarIndiceColuna("Alerta liderança\n83 dias", cabecalho);
  const COL_STATUS = encontrarIndiceColuna("Status do E-mail", cabecalho);
  
  if (COL_STATUS === -1) {
    Logger.log("ERRO CRÍTICO: A coluna 'Status do E-mail' não foi encontrada.");
    return;
  }
  
  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];
    const statusAtual = linha[COL_STATUS].toString();
    
    if (!linha[COL_EMAIL_AFILHADO]) continue;

    const nomeAfilhado = linha[COL_NOME_AFILHADO];
    const emailAfilhado = linha[COL_EMAIL_AFILHADO];
    const dataAdmissao = linha[COL_DATA_ADMISSAO];
    const nomePadrinho = linha[COL_NOME_PADRINHO];
    const emailPadrinho = linha[COL_EMAIL_PADRINHO];
    const emailLideranca = linha[COL_EMAIL_LIDERANCA];
    
    const destinatariosAlerta = [emailPadrinho, emailLideranca].filter(e => e && e.toString().trim() !== "");
    const opcoesBcc = { bcc: destinatariosAlerta.join(',') };
    
    // --- LÓGICA DE PERIODICIDADE COMPLETA ---
    
    if (datasIguais(dataAdmissao, hojeString, fusoHorario) && !statusAtual.includes("sent_admission")) {
      enviarEmail(emailAfilhado, "Boas-vindas ao Programa de Apadrinhamento Indicium!", criarMensagemBoasVindasAfilhado(nomeAfilhado));
      enviarEmail(emailPadrinho, "Início da Jornada: Seu/Sua Afilhado(a) Começou!", criarMensagemBoasVindasPadrinho(nomePadrinho, nomeAfilhado));
      planilha.getRange(i + 1, COL_STATUS + 1).setValue(statusAtual + "sent_admission; ");
      Logger.log('>> E-mails de boas-vindas enviados para a jornada de ' + nomeAfilhado);

    } else if (datasIguais(linha[COL_30_DIAS], hojeString, fusoHorario) && !statusAtual.includes("sent_30")) {
      enviarEmail(seuEmail, '📌 Alerta – 30 dias de experiência: ' + nomeAfilhado, criarMensagem30Dias(nomeAfilhado), opcoesBcc);
      planilha.getRange(i + 1, COL_STATUS + 1).setValue(statusAtual + "sent_30; ");
      Logger.log('>> Alerta de 30 dias enviado para ' + nomeAfilhado);

    } else if (datasIguais(linha[COL_38_DIAS], hojeString, fusoHorario) && !statusAtual.includes("sent_38")) {
      enviarEmail(seuEmail, '⚠️ Alerta – 38 dias de experiência: ' + nomeAfilhado, criarMensagem38Dias(nomeAfilhado), opcoesBcc);
      enviarEmail(emailAfilhado, "Acompanhamento - 38 dias", "<p>Olá! Chegamos ao marco de 38 dias do seu acompanhamento.</p>");
      planilha.getRange(i + 1, COL_STATUS + 1).setValue(statusAtual + "sent_38; ");
      Logger.log('>> Alerta de 38 dias enviado para ' + nomeAfilhado);

    } else if (datasIguais(linha[COL_75_DIAS], hojeString, fusoHorario) && !statusAtual.includes("sent_75")) {
      enviarEmail(seuEmail, '📌 Alerta – 75 dias de experiência: ' + nomeAfilhado, criarMensagem75Dias(nomeAfilhado), opcoesBcc);
      planilha.getRange(i + 1, COL_STATUS + 1).setValue(statusAtual + "sent_75; ");
      Logger.log('>> Alerta de 75 dias enviado para ' + nomeAfilhado);

    } else if (datasIguais(linha[COL_83_DIAS], hojeString, fusoHorario) && !statusAtual.includes("sent_83")) {
      enviarEmail(seuEmail, '⚠️ Alerta – 83 dias de experiência: ' + nomeAfilhado, criarMensagem83Dias(nomeAfilhado), opcoesBcc);
      enviarEmail(emailAfilhado, "Acompanhamento - 83 dias", "<p>Olá! Chegamos ao marco de 83 dias do seu acompanhamento.</p>");
      planilha.getRange(i + 1, COL_STATUS + 1).setValue(statusAtual + "sent_83; ");
      Logger.log('>> Alerta de 83 dias enviado para ' + nomeAfilhado);

    } else if (dataAdmissao instanceof Date && !statusAtual.includes("sent_90_day_feedback")) {
      let data90dias = new Date(dataAdmissao);
      data90dias.setDate(data90dias.getDate() + 90);
      if (datasIguais(data90dias, hojeString, fusoHorario)) {
        enviarEmail(emailAfilhado, "Avalie sua Experiência no Programa de Apadrinhamento", criarMensagemFeedback90Dias(nomeAfilhado));
        planilha.getRange(i + 1, COL_STATUS + 1).setValue(statusAtual + "sent_90_day_feedback; ");
        Logger.log('>> E-mail de avaliação de 90 dias enviado para ' + nomeAfilhado);
      }
    }
  }
  Logger.log("--- Fim da execução ---");
}

// --- NOVOS MODELOS DE E-MAIL ---
function criarMensagemBoasVindasAfilhado(nome) { return '<p>Olá <b>' + nome + '</b>,</p><p>Desejamos as boas-vindas oficialmente na Indicium e ao Programa de Apadrinhamento.</p>'; }
function criarMensagemBoasVindasPadrinho(nomePadrinho, nomeAfilhado) { return '<p>Olá <b>' + nomePadrinho + '</b>,</p><p>Passando para lembrá-lo que o seu/sua afilhado/a <b>' + nomeAfilhado + '</b> iniciou hoje na Indicium.</p>'; }
function criarMensagemFeedback90Dias(nome) { return '<p>Olá <b>' + nome + '</b>,</p><p>Chegou o momento de avaliar o seu padrinho ou madrinha do Programa de Apadrinhamento.</p><p>Clique no link e envie a sua avaliação.</p><p><a href="https://forms.gle/HooXxTAyrHMj4K8C8">Link para Avaliação</a></p>'; }

// --- MODELOS DE E-MAIL EXISTENTES ---
function criarMensagem30Dias(nome) { return '<p>📌 Alerta – 30 dias de experiência</p><p>O(a) afilhado(a) <b>' + nome + '</b> está completando 30 dias na Indicium!</p><p>Este é um momento importante para acompanhar a adaptação e os primeiros resultados da jornada.<br>Você tem 8 dias para preencher a 1ª avaliação de experiência na HiBob.</p><p>Em caso de dúvidas sobre o preenchimento, estou à disposição.</p>'; }
function criarMensagem38Dias(nome) { return '<p>⚠️ Alerta – 38 dias de experiência</p><p>Hoje é o último dia para preencher a 1ª avaliação de experiência na HiBob para o(a) afilhado(a) <b>' + nome + '</b>.</p><p>Essa etapa é essencial para registrar percepções iniciais sobre o desenvolvimento do(a) afilhado(a).</p><p>Caso já tenha finalizado, lembre-se de realizar o feedback de 30 dias com ele(a) antes de completar 45 dias.</p>'; }
function criarMensagem75Dias(nome) { return '<p>📌 Alerta – 75 dias de experiência</p><p>O(a) afilhado(a) <b>' + nome + '</b> está completando 75 dias na casa!</p><p>Este é o momento de acompanhar a consolidação da performance e o alinhamento com o time e entregas.<br>Você tem 8 dias para preencher a 2ª avaliação de experiência na HiBob.</p><p>Se precisar de apoio, estou por aqui.</p>'; }
function criarMensagem83Dias(nome) { return '<p>⚠️ Alerta – 83 dias de experiência</p><p>Hoje é o último dia para preencher a 2ª avaliação de experiência na HiBob para o(a) afilhado(a) <b>' + nome + '</b>.</p><p>Esse registro é essencial para fechar o ciclo de experiência com visão de desempenho, evolução e integração.</p><p>Caso já tenha finalizado, lembre-se de realizar o feedback final com o afilhado(a) antes de completar 90 dias.</p>'; }

// --- FUNÇÕES AUXILIARES ---
function datasIguais(dataDaPlanilha, hojeString, fusoHorario) { if (!dataDaPlanilha || !(dataDaPlanilha instanceof Date)) return false; const dataPlanilhaString = Utilities.formatDate(dataDaPlanilha, fusoHorario, 'dd/MM/yyyy'); return dataPlanilhaString === hojeString; }
function enviarEmail(destinatario, assunto, corpoHtml, opcoes) { if (!destinatario || destinatario.length === 0) return; const opcoesDeEnvio = opcoes || {}; opcoesDeEnvio.htmlBody = corpoHtml; opcoesDeEnvio.subject = '=?UTF-8?B?' + Utilities.base64Encode(assunto, Utilities.Charset.UTF_8) + '?='; opcoesDeEnvio.charset = 'UTF-8'; GmailApp.sendEmail(destinatario, null, null, opcoesDeEnvio); }
function encontrarIndiceColuna(nomeColuna, cabecalho) { const indice = cabecalho.findIndex(function(celula) { return celula.toString().trim() === nomeColuna.trim(); }); if (indice === -1) { Logger.log("AVISO: A coluna '" + nomeColuna.replace('\n', ' ') + "' não foi encontrada."); } return indice; }