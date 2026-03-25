
function enviarPendentes() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Envios');
  if (!sheet) throw new Error("Aba 'Envios' não encontrada.");

  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length &lt; 2) return;

  const headers = values[0].map(h => String(h).trim());
  const idx = (name) => headers.indexOf(name);

  const iPrestador = idx('Prestador');
  const iTo = idx('Email_To');
  const iCc = idx('Email_Cc');
  const iAssunto = idx('Assunto');
  const iMsg = idx('Mensagem');
  const iAnexoUrl = idx('Anexo_URL');
  const iStatus = idx('Status');
  const iEnviadoEm = idx('Enviado_em');
  const iMsgErro = idx('Msg_erro');

  const required = [iPrestador, iTo, iAssunto, iMsg, iStatus, iEnviadoEm, iMsgErro];
  if (required.some(i => i === -1)) {
    throw new Error("Cabeçalhos obrigatórios faltando. Verifique: Prestador, Email_To, Assunto, Mensagem, Status, Enviado_em, Msg_erro (e Email_Cc/Anexo_URL se usar).");
  }

  for (let r = 1; r &lt; values.length; r++) {
    const row = values[r];
    const status = String(row[iStatus] || '').trim().toUpperCase();

    if (status !== 'PENDENTE') continue;

    const prestador = String(row[iPrestador] || '').trim();
    const toRaw = String(row[iTo] || '').trim();
    const ccRaw = iCc !== -1 ? String(row[iCc] || '').trim() : '';
    const subject = String(row[iAssunto] || '').trim();
    const message = String(row[iMsg] || '').trim();
    const anexoUrl = iAnexoUrl !== -1 ? String(row[iAnexoUrl] || '').trim() : '';

    try {
      if (!toRaw) throw new Error('Email_To vazio.');
      if (!subject) throw new Error('Assunto vazio.');
      if (!message) throw new Error('Mensagem vazia.');

      const to = parseEmailsSemicolon(toRaw);
      const cc = ccRaw ? parseEmailsSemicolon(ccRaw) : [];

      // Corpo HTML simples (você pode sofisticar)
      const htmlBody = `
        <p>Prezado(a) ${escapeHtml(prestador || 'prestador')},</p>
        <p>${escapeHtml(message).replace(/\n/g, '<br>')}</p>
        <p>Atenciosamente,<br>Backoffice</p>
      `;

      const options = {
        htmlBody,
      };

      if (cc.length) options.cc = cc.join(',');

      // Anexo via Drive (se link foi informado)
      if (anexoUrl) {
        const fileId = extractDriveFileId(anexoUrl);
        if (!fileId) throw new Error('Não consegui extrair o fileId do Anexo_URL.');
        const file = DriveApp.getFileById(fileId);
        options.attachments = [file.getBlob()];
      }

      GmailApp.sendEmail(to.join(','), subject, stripHtml(htmlBody), options);

      // Marca como enviado
      sheet.getRange(r + 1, iStatus + 1).setValue('ENVIADO');
      sheet.getRange(r + 1, iEnviadoEm + 1).setValue(new Date());
      sheet.getRange(r + 1, iMsgErro + 1).setValue('');

    } catch (err) {
      sheet.getRange(r + 1, iStatus + 1).setValue('ERRO');
      sheet.getRange(r + 1, iEnviadoEm + 1).setValue(new Date());
      sheet.getRange(r + 1, iMsgErro + 1).setValue(String(err && err.message ? err.message : err));
    }
  }
}

function parseEmailsSemicolon(raw) {
  return String(raw)
    .split(';')
    .map(s => s.trim())
    .filter(s => s.includes('@'))
    .filter((v, i, a) => a.indexOf(v) === i);
}

function extractDriveFileId(url) {
  // Funciona para links tipo:
  // https://drive.google.com/file/d/FILE_ID/view?...
  // https://drive.google.com/open?id=FILE_ID
  // https://docs.google.com/spreadsheets/d/FILE_ID/edit...
  const s = String(url);
  let m = s.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (m && m[1]) return m[1];
  m = s.match(/[?&]id=([a-zA-Z0-9-_]+)/);
  if (m && m[1]) return m[1];
  return null;
}

function escapeHtml(text) {
  return String(text)
    .replaceAll('&', '&amp;')
    .replaceAll('&lt;', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function stripHtml(html) {
  return String(html).replace(/&lt;[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
}
