// Inclui if/else para condição de datas (Testar)

function executarAlerta() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = planilha.getDataRange();
  const dados = range.getValues();
  
  const fusoHorario = Session.getScriptTimeZone();
  const hojeString = Utilities.formatDate(new Date(), fusoHorario, "dd/MM/yyyy");
  
  const seuEmail = "fabriciapinheiro24@gmail.com";

  Logger.log("--- Iniciando execução com Lógica de Periodicidade ---");
  Logger.log("Data de hoje para comparação: " + hojeString);

  const cabecalho = dados[0];
  const COL_NOME_AFILHADO = 0;
  const COL_EMAIL_AFILHADO = 1;
  const COL_EMAIL_PADRINHO = 24;
  const COL_EMAIL_LIDERANCA = 25;
  
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
    const statusAtual = linha[COL_STATUS];
    
    // Pula a linha se não houver e-mail ou se já tiver um status para o dia de hoje
    if (!linha[COL_EMAIL_AFILHADO] || statusAtual.toString().includes(hojeString)) {
      continue;
    }

    const nomeAfilhado = linha[COL_NOME_AFILHADO];
    const emailAfilhado = linha[COL_EMAIL_AFILHADO];
    const emailPadrinho = linha[COL_EMAIL_PADRINHO];
    const emailLideranca = linha[COL_EMAIL_LIDERANCA];
    
    const destinatariosAlerta = [emailPadrinho, emailLideranca].filter(e => e && e.toString().trim() !== "");
    const opcoesBcc = { bcc: destinatariosAlerta.join(',') };

    // --- LÓGICA DE PERIODICIDADE ---
    // O script agora só enviará o PRIMEIRO alerta que encontrar para o dia.
    
    let emailEnviado = false;

    if (datasIguais(linha[COL_30_DIAS], hojeString, fusoHorario)) {
      const assunto = `📌 Alerta – 30 dias de experiência: ${nomeAfilhado}`;
      const corpoHtml = criarMensagem30Dias(nomeAfilhado);
      enviarEmail(seuEmail, assunto, corpoHtml, opcoesBcc);
      emailEnviado = true;
    } else if (datasIguais(linha[COL_38_DIAS], hojeString, fusoHorario)) {
      const assunto = `📌 Alerta – 38 dias de experiência: ${nomeAfilhado}`;
      const corpoHtml = criarMensagem38Dias(nomeAfilhado);
      enviarEmail(seuEmail, assunto, corpoHtml, opcoesBcc);
      enviarEmail(emailAfilhado, "Acompanhamento - 38 dias", "<p>Olá! Chegamos ao marco de 38 dias do seu acompanhamento.</p>");
      emailEnviado = true;
    } else if (datasIguais(linha[COL_75_DIAS], hojeString, fusoHorario)) {
      const assunto = `📌 Alerta – 75 dias de experiência: ${nomeAfilhado}`;
      const corpoHtml = criarMensagem75Dias(nomeAfilhado);
      enviarEmail(seuEmail, assunto, corpoHtml, opcoesBcc);
      emailEnviado = true;
    } else if (datasIguais(linha[COL_83_DIAS], hojeString, fusoHorario)) {
      const assunto = `📌 Alerta – 83 dias de experiência: ${nomeAfilhado}`;
      const corpoHtml = criarMensagem83Dias(nomeAfilhado);
      enviarEmail(seuEmail, assunto, corpoHtml, opcoesBcc);
      enviarEmail(emailAfilhado, "Acompanhamento - 83 dias", "<p>Olá! Chegamos ao marco de 83 dias do seu acompanhamento.</p>");
      emailEnviado = true;
    }

    // Se qualquer um dos e-mails foi enviado, marca a linha com o status
    if (emailEnviado) {
      const linhaDaPlanilha = i + 1; 
      planilha.getRange(linhaDaPlanilha, COL_STATUS + 1).setValue("Enviado - " + hojeString);
      Logger.log(`  >> Linha ${linhaDaPlanilha} (${nomeAfilhado}) marcada como 'Enviado'.`);
    }
  }
  Logger.log("--- Fim da execução ---");
}

// --- MODELOS DE E-MAIL (HTML) ---
function criarMensagem30Dias(nome) {
  return `<p>⚠️ Alerta – 30 dias de experiência</p><p>O(a) afilhado(a) <b>${nome}</b> está completando 30 dias na Indicium!</p><p>Este é um momento importante para acompanhar a adaptação e os primeiros resultados da jornada.<br>Você tem 8 dias para preencher a 1ª avaliação de experiência na HiBob.</p><p>Em caso de dúvidas sobre o preenchimento, estou à disposição.</p>`;
}
function criarMensagem38Dias(nome) {
    return `<p>⚠️ Alerta – 38 dias de experiência</p><p>Hoje é o último dia para preencher a 1ª avaliação de experiência na HiBob para o(a) afilhado(a) <b>${nome}</b>.</p><p>Essa etapa é essencial para registrar percepções iniciais sobre o desenvolvimento do(a) afilhado(a).</p><p>Caso já tenha finalizado, lembre-se de realizar o feedback de 30 dias com ele(a) antes de completar 45 dias.</p>`;
}
function criarMensagem75Dias(nome) {
  return `<p>⚠️ Alerta – 75 dias de experiência</p><p>O(a) afilhado(a) <b>${nome}</b> está completando 75 dias na casa!</p><p>Este é o momento de acompanhar a consolidação da performance e o alinhamento com o time e entregas.<br>Você tem 8 dias para preencher a 2ª avaliação de experiência na HiBob.</p><p>Se precisar de apoio, estou por aqui.</p>`;
}
function criarMensagem83Dias(nome) {
    return `<p>⚠️ Alerta – 83 dias de experiência</p><p>Hoje é o último dia para preencher a 2ª avaliação de experiência na HiBob para o(a) afilhado(a) <b>${nome}</b>.</p><p>Esse registro é essencial para fechar o ciclo de experiência com visão de desempenho, evolução e integração.</p><p>Caso já tenha finalizado, lembre-se de realizar o feedback final com o afilhado(a) antes de completar 90 dias.</p>`;
}

// --- FUNÇÕES AUXILIARES ---
function datasIguais(dataDaPlanilha, hojeString, fusoHorario) {
  if (!dataDaPlanilha || !(dataDaPlanilha instanceof Date)) return false;
  const dataPlanilhaString = Utilities.formatDate(dataDaPlanilha, fusoHorario, "dd/MM/yyyy");
  return dataPlanilhaString === hojeString;
}

function enviarEmail(destinatario, assunto, corpoHtml, opcoes) {
  if (!destinatario || destinatario.length === 0) return;
  const opcoesDeEnvio = opcoes || {};
  opcoesDeEnvio.htmlBody = corpoHtml;
  opcoesDeEnvio.subject = `=?UTF-8?B?${Utilities.base64Encode(assunto, Utilities.Charset.UTF_8)}?=`;
  opcoesDeEnvio.charset = 'UTF-8';
  GmailApp.sendEmail(destinatario, null, null, opcoesDeEnvio);
}

function encontrarIndiceColuna(nomeColuna, cabecalho) {
  const indice = cabecalho.findIndex(celula => celula.toString().trim() === nomeColuna.trim());
  if (indice === -1) {
    Logger.log("AVISO: A coluna '" + nomeColuna.replace("\n", " ") + "' não foi encontrada.");
  }
  return indice;
}
