/**
 * AMARRAÇÃO DA PLATAFORMA COM A PLANILHA
 * -----------------------------------------------------------------
 * Este script faz a página "ENVIO DE CARGAS - CRIAÇÃO DE CTE.html"
 * puxar os dados direto da sua planilha do Google Sheets, sempre que
 * ela for aberta (ou quando você clicar no botão de atualizar).
 *
 * COMO INSTALAR:
 * 1. Abra a planilha no Google Sheets (a mesma onde você sobe o
 *    export do BI a cada 3 dias, aba "Dados").
 * 2. Extensões > Apps Script.
 * 3. Apague o conteúdo do Code.gs e cole este arquivo inteiro.
 * 4. Salve.
 * 5. Clique em "Implantar" (Deploy) > "Nova implantação".
 *    - Tipo: "App da Web" (Web app)
 *    - Executar como: Eu (seu e-mail)
 *    - Quem pode acessar: "Qualquer pessoa" (Anyone) — precisa ser
 *      "Anyone", senão a página não consegue ler os dados de fora.
 * 6. Clique em "Implantar". Vai aparecer uma URL terminando em /exec.
 * 7. Copie essa URL.
 * 8. Abra o arquivo "ENVIO DE CARGAS - CRIAÇÃO DE CTE.html" num editor
 *    de texto, procure a linha:
 *        const SCRIPT_URL = "";
 *    e cole a URL entre as aspas. Salve o HTML.
 * 9. Pronto — toda vez que abrir o HTML, ele vai buscar os dados
 *    atuais da aba "Dados" da planilha.
 *
 * Sempre que você reimplantar o script (nova implantação), a URL
 * muda — nesse caso, atualize o SCRIPT_URL no HTML de novo. Para
 * evitar isso, ao reimplantar escolha "Gerenciar implantações" >
 * editar a implantação existente em vez de criar uma nova.
 * -----------------------------------------------------------------
 */

const ABA_DADOS = "Dados";
const ABA_EMAILS = "Emails_Transportadoras";
const ABA_ENVIOS = "Envios";

function getIdx(cab) {
  return {
    filial: cab.indexOf("FILIAL"),
    dtEmissao: cab.indexOf("DT_EMISSAO_NF"),
    romaneio: cab.indexOf("NR_ROMANEIO"),
    cdTransp: cab.indexOf("CD_TRANSP"),
    dsTransp: cab.indexOf("DS_TRANSP"),
    motorista: cab.indexOf("DS_MOTORISTA"),
    placa: cab.indexOf("PLACA"),
    peso: cab.indexOf("PESO"),
    tarifa: cab.indexOf("TARIFA"),
    vlrFrete: cab.indexOf("VLR_FRETE"),
    valorNf: cab.indexOf("VALOR_NF"),
    cidade: cab.indexOf("CIDADE"),
    uf: cab.indexOf("UF"),
    loja: cab.indexOf("LOJA"),
    nrNf: cab.indexOf("NR_NF"),
    chave: cab.indexOf("CHAVENF"),
    ocComplemento: cab.indexOf("OC_COMPLEMENTO"),
    nrOcorrencia: cab.indexOf("NR_OCORRENCIA"),
    dataOc: cab.indexOf("DATA_OC"),
    ocorrencia: cab.indexOf("OCORRENCIA"),
    descrSubOc: cab.indexOf("DESCR_SUB_OC"),
    motOcor: cab.indexOf("MOT_OCOR"),
    cte: cab.indexOf("CTE")
  };
}

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(ABA_DADOS);
  const dados = aba.getDataRange().getValues();
  const cab = dados[0];
  const idx = getIdx(cab);

  const romaneios = {};
  const ordem = [];

  for (let i = 1; i < dados.length; i++) {
    const r = dados[i];
    if (!r[idx.filial]) continue;

    const key = r[idx.filial] + "|" + r[idx.romaneio] + "|" + r[idx.cdTransp];

    if (!romaneios[key]) {
      const dt = r[idx.dtEmissao];
      romaneios[key] = {
        filial: String(r[idx.filial]),
        romaneio: String(r[idx.romaneio]),
        cdTransp: String(r[idx.cdTransp]),
        transportadora: limpar(r[idx.dsTransp]),
        motorista: limpar(r[idx.motorista]),
        placa: limpar(r[idx.placa]),
        peso: parseNumeroBR(r[idx.peso]),
        tarifa: parseNumeroBR(r[idx.tarifa]),
        frete: parseNumeroBR(r[idx.vlrFrete]),
        valorCarga: parseNumeroBR(r[idx.valorNf]),
        dataEmissao: dt instanceof Date ? Utilities.formatDate(dt, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(dt || ""),
        cte: limpar(r[idx.cte]),
        nfs: [],
        cidades: [],
        ocorrencias025: []
      };
      ordem.push(key);
    }

    const rom = romaneios[key];
    const cidadeUf = limpar(r[idx.cidade]) + "/" + limpar(r[idx.uf]);
    if (rom.cidades.indexOf(cidadeUf) === -1) rom.cidades.push(cidadeUf);

    rom.nfs.push({
      nrNf: String(r[idx.nrNf]),
      loja: limpar(r[idx.loja]),
      cidade: limpar(r[idx.cidade]),
      uf: limpar(r[idx.uf]),
      chave: String(r[idx.chave])
    });

    const descr = limpar(r[idx.descrSubOc]);
    const mot = limpar(r[idx.motOcor]);
    if (descr.indexOf("025") === 0 || mot.indexOf("025") === 0) {
      const tipos = [];
      if (mot.indexOf("025") === 0) tipos.push("Complemento de frete");
      if (descr.indexOf("025") === 0) tipos.push("Diária");
      const dataOc = r[idx.dataOc];
      const desc = r[idx.ocorrencia];
      rom.ocorrencias025.push({
        nrNf: String(r[idx.nrNf]),
        nrOcorrencia: limpar(r[idx.nrOcorrencia]),
        dataOc: dataOc instanceof Date ? Utilities.formatDate(dataOc, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(dataOc || ""),
        tipos: tipos,
        valor: parseNumeroBR(r[idx.ocComplemento]),
        descricao: (typeof desc === "string" && desc.trim() !== "" && desc.trim() !== "-") ? limpar(desc) : null
      });
    }
  }

  const resultado = ordem.map(k => romaneios[k]);

  // Marca quais romaneios já tiveram e-mail enviado (aba "Envios")
  const envios = lerEnvios(ss);
  resultado.forEach(function (r) {
    const chave = r.filial + "|" + r.romaneio + "|" + r.cdTransp;
    const info = envios[chave];
    r.emailEnviado = !!info;
    r.dataEnvioEmail = info ? info.data : null;
  });

  return ContentService
    .createTextOutput(JSON.stringify(resultado))
    .setMimeType(ContentService.MimeType.JSON);
}

function lerEnvios(ss) {
  const aba = ss.getSheetByName(ABA_ENVIOS);
  const mapa = {};
  if (!aba) return mapa;
  const valores = aba.getDataRange().getValues();
  for (let i = 1; i < valores.length; i++) {
    const chave = valores[i][0];
    const data = valores[i][4];
    if (!chave) continue;
    mapa[chave] = {
      data: data instanceof Date ? data.toISOString() : String(data || "")
    };
  }
  return mapa;
}

/**
 * Recebe requisições POST da página (botão "Enviar e-mail").
 * Corpo esperado (JSON): { action: "enviarEmail", filial, romaneio, cdTransp, subject, greeting }
 */
function doPost(e) {
  let payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ success: false, error: "Requisição inválida." });
  }

  if (payload.action === "enviarEmail") {
    return enviarEmailRomaneio(payload);
  }

  return jsonResponse({ success: false, error: "Ação desconhecida: " + payload.action });
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function enviarEmailRomaneio(payload) {
  const filial = String(payload.filial || "");
  const romaneio = String(payload.romaneio || "");
  const cdTransp = String(payload.cdTransp || "");
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1) E-mail da transportadora
  const abaEmails = ss.getSheetByName(ABA_EMAILS);
  if (!abaEmails) return jsonResponse({ success: false, error: 'Aba "Emails_Transportadoras" não encontrada na planilha.' });

  const emailsValores = abaEmails.getDataRange().getValues();
  const emailsCab = emailsValores[0];
  const idxCd = emailsCab.indexOf("CD_TRANSP");
  const idxDest = emailsCab.findIndex(h => String(h).indexOf("EMAIL_DESTINATARIO") === 0);
  const idxCc = emailsCab.findIndex(h => String(h).indexOf("EMAIL_CC_GTF") === 0);

  let destinatario = "", cc = "";
  for (let i = 1; i < emailsValores.length; i++) {
    if (normalizarCodigo(emailsValores[i][idxCd]) === normalizarCodigo(cdTransp)) {
      destinatario = limpar(emailsValores[i][idxDest]);
      cc = limpar(emailsValores[i][idxCc]);
      break;
    }
  }
  if (!destinatario) {
    return jsonResponse({ success: false, error: "Essa transportadora ainda não tem e-mail cadastrado na aba Emails_Transportadoras (código " + cdTransp + ")." });
  }

  // 2) Linhas do romaneio na aba Dados (fonte da verdade, não confia no que veio do navegador)
  const abaDados = ss.getSheetByName(ABA_DADOS);
  const dados = abaDados.getDataRange().getValues();
  const cab = dados[0];
  const idx = getIdx(cab);

  const linhas = [];
  for (let i = 1; i < dados.length; i++) {
    const r = dados[i];
    if (String(r[idx.filial]) === filial && String(r[idx.romaneio]) === romaneio && String(r[idx.cdTransp]) === cdTransp) {
      linhas.push(r);
    }
  }
  if (linhas.length === 0) {
    return jsonResponse({ success: false, error: "Não encontrei esse romaneio na aba Dados." });
  }

  // 3) Montar tabela igual ao layout do relatório (copiável, não é imagem)
  const primeira = linhas[0];
  const dt = primeira[idx.dtEmissao];
  const dtStr = dt instanceof Date ? Utilities.formatDate(dt, Session.getScriptTimeZone(), "dd/MM/yyyy") : String(dt || "");
  const transportadora = limpar(primeira[idx.dsTransp]);
  const motorista = limpar(primeira[idx.motorista]);
  const placa = limpar(primeira[idx.placa]);
  const peso = parseNumeroBR(primeira[idx.peso]);
  const frete = parseNumeroBR(primeira[idx.vlrFrete]);
  const valorCarga = parseNumeroBR(primeira[idx.valorNf]);

  const fmtMoeda = v => v == null ? "" : "R$ " + v.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const fmtNum = v => v == null ? "" : v.toLocaleString("pt-BR");

  let linhasHtml = "";
  linhas.forEach(function (r, i) {
    const nrNf = limpar(r[idx.nrNf]);
    const chave = limpar(r[idx.chave]);
    if (i === 0) {
      linhasHtml += "<tr>"
        + "<td>" + dtStr + "</td>"
        + "<td>" + romaneio + "</td>"
        + "<td>" + transportadora + "</td>"
        + "<td>" + motorista + "</td>"
        + "<td>" + placa + "</td>"
        + "<td>" + fmtNum(peso) + "</td>"
        + "<td>" + fmtMoeda(frete) + "</td>"
        + "<td>" + fmtMoeda(valorCarga) + "</td>"
        + "<td>" + nrNf + "</td>"
        + "<td style='font-size:11px'>" + chave + "</td>"
        + "</tr>";
    } else {
      linhasHtml += "<tr>"
        + "<td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td>"
        + "<td>" + nrNf + "</td>"
        + "<td style='font-size:11px'>" + chave + "</td>"
        + "</tr>";
    }
  });

  const tabela =
    "<table border='1' cellpadding='6' cellspacing='0' style='border-collapse:collapse;font-family:Arial,sans-serif;font-size:12px;width:100%;margin-top:12px'>"
    + "<tr style='background:#4472C4;color:#ffffff'>"
    + "<th>DT_EMISSAO_NF</th><th>NR_ROMANEIO</th><th>DS_TRANSP</th><th>DS_MOTORISTA</th><th>PLACA</th>"
    + "<th>PESO</th><th>VLR_FRETE</th><th>VALOR_NF</th><th>NR_NF</th><th>CHAVENF</th>"
    + "</tr>"
    + linhasHtml
    + "</table>";

  // Ocorrências 025 (complemento de frete / diária) desse romaneio - aparecem
  // acima da tabela, uma linha por ocorrência: "Ocorrência: 000 - descrição VALOR: X,XX"
  const linhasOcorrenciaHtml = [];
  const linhasOcorrenciaTexto = [];
  linhas.forEach(function (r) {
    const descrSub = limpar(r[idx.descrSubOc]);
    const mot = limpar(r[idx.motOcor]);
    if (descrSub.indexOf("025") === 0 || mot.indexOf("025") === 0) {
      const nrOc = limpar(r[idx.nrOcorrencia]);
      const desc = limpar(r[idx.ocorrencia]);
      const valorOc = parseNumeroBR(r[idx.ocComplemento]);
      let linha = "Ocorrência: " + nrOc;
      if (desc) linha += " - " + desc;
      if (valorOc != null) linha += " VALOR: " + valorOc.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      linhasOcorrenciaHtml.push("<p style='margin:2px 0'>" + linha + "</p>");
      linhasOcorrenciaTexto.push(linha);
    }
  });
  const ocorrenciasHtml = linhasOcorrenciaHtml.join("");
  const ocorrenciasTexto = linhasOcorrenciaTexto.length ? linhasOcorrenciaTexto.join("\n") + "\n\n" : "";

  const saudacao = payload.greeting || "Bom dia";
  const corpoHtml =
    "<div style=\"font-family:Arial,sans-serif;font-size:13px;color:#1a1a1a;line-height:1.5\">"
    + "<p>" + saudacao + " a todos, Favor enviar o cte nesse mesmo e-mail, não tire ninguém da cópia</p>"
    + "<p><b>OBS:- FAVOR SEMPRE CONFERIR SE AS NOTAS ANEXA ESTÃO CONFERINDO COM O RELATÓRIO ABAIXO</b></p>"
    + "<p>Segue anexo as notas e abaixo a formação para emissão do cte, lembrando que após a emissão devem enviar nesse mesmo e-mail cte para lançamento e posterior pagamento.</p>"
    + "<p>Comprovantes de despesas enviar anexo no e-mail junto com o cte</p>"
    + ocorrenciasHtml
    + tabela
    + "</div>";

  let corpoTexto =
    saudacao + " a todos, Favor enviar o cte nesse mesmo e-mail, não tire ninguém da cópia\n\n"
    + "OBS:- FAVOR SEMPRE CONFERIR SE AS NOTAS ANEXA ESTÃO CONFERINDO COM O RELATÓRIO ABAIXO\n\n"
    + "Segue anexo as notas e abaixo a formação para emissão do cte, lembrando que após a emissão devem enviar nesse mesmo e-mail cte para lançamento e posterior pagamento.\n\n"
    + "Comprovantes de despesas enviar anexo no e-mail junto com o cte\n\n"
    + ocorrenciasTexto
    + "DT_EMISSAO_NF\tNR_ROMANEIO\tDS_TRANSP\tDS_MOTORISTA\tPLACA\tPESO\tVLR_FRETE\tVALOR_NF\tNR_NF\tCHAVENF\n";

  linhas.forEach(function (r, i) {
    const nrNf = limpar(r[idx.nrNf]);
    const chave = limpar(r[idx.chave]);
    if (i === 0) {
      corpoTexto += [dtStr, romaneio, transportadora, motorista, placa, fmtNum(peso), fmtMoeda(frete), fmtMoeda(valorCarga), nrNf, chave].join("\t") + "\n";
    } else {
      corpoTexto += ["", "", "", "", "", "", "", "", nrNf, chave].join("\t") + "\n";
    }
  });

  const assunto = payload.subject ||
    ("EMISSÃO CTE - CARGA: " + romaneio + " - TRANSPORTADORA: " + transportadora + " - DATA: " + dtStr);

  marcarPreparado(ss, filial, romaneio, cdTransp, destinatario, cc);

  return jsonResponse({
    success: true,
    destinatario: destinatario,
    cc: cc,
    subject: assunto,
    htmlBody: corpoHtml,
    plainBody: corpoTexto
  });
}

function marcarPreparado(ss, filial, romaneio, cdTransp, destinatario, cc) {
  let aba = ss.getSheetByName(ABA_ENVIOS);
  if (!aba) {
    aba = ss.insertSheet(ABA_ENVIOS);
    aba.appendRow(["CHAVE", "FILIAL", "ROMANEIO", "CD_TRANSP", "DATA_PREPARO", "DESTINATARIO", "CC"]);
  }
  const chave = filial + "|" + romaneio + "|" + cdTransp;
  aba.appendRow([chave, filial, romaneio, cdTransp, new Date(), destinatario, cc]);
}

function limpar(v) {
  if (v === null || v === undefined) return "";
  return String(v).replace(/\s+/g, " ").trim();
}

/**
 * Normaliza códigos de transportadora pra comparação - remove zeros à
 * esquerda, já que a mesma transportadora pode vir como "2073" na aba
 * Dados e "002073" na aba Emails_Transportadoras.
 */
function normalizarCodigo(v) {
  return String(v == null ? "" : v).trim().replace(/^0+(?=\d)/, "");
}

/**
 * Converte valores numéricos que podem vir como número puro OU como texto
 * formatado (ex: "R$ 168.337,38", "7.448,00", "5117") - comum quando a
 * coluna da planilha está formatada como moeda/texto em vez de número.
 */
function parseNumeroBR(v) {
  if (typeof v === "number") return v;
  if (v === null || v === undefined) return null;
  var s = String(v).trim();
  if (s === "" || s === "-") return null;
  s = s.replace(/R\$/gi, "").replace(/\s/g, "");
  if (s.indexOf(",") > -1) {
    s = s.replace(/\./g, "").replace(",", ".");
  }
  var n = parseFloat(s);
  return isNaN(n) ? null : n;
}
