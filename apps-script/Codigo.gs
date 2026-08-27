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

/**
 * Preenche com zero à esquerda até o tamanho padrão do campo, porque a
 * planilha guarda esses códigos como número puro e perde os zeros
 * (ex: FILIAL "01026" vira 1026, NR_ROMANEIO "08111701" vira 8111701).
 * Se o valor já tiver esse tamanho ou mais, não mexe nele.
 */
function zpad(v, largura) {
  return String(v == null ? "" : v).trim().padStart(largura, "0");
}
const LARGURA_FILIAL = 5;
const LARGURA_ROMANEIO = 8;
const LARGURA_NF = 9;
const LARGURA_TRANSP = 6;

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

    const filialPad = zpad(r[idx.filial], LARGURA_FILIAL);
    const romaneioPad = zpad(r[idx.romaneio], LARGURA_ROMANEIO);
    const cdTranspPad = zpad(r[idx.cdTransp], LARGURA_TRANSP);
    const nrNfPad = zpad(r[idx.nrNf], LARGURA_NF);

    const key = filialPad + "|" + romaneioPad + "|" + cdTranspPad;

    if (!romaneios[key]) {
      const dt = r[idx.dtEmissao];
      romaneios[key] = {
        filial: filialPad,
        romaneio: romaneioPad,
        cdTransp: cdTranspPad,
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
      nrNf: nrNfPad,
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
        nrNf: nrNfPad,
        nrOcorrencia: limpar(r[idx.nrOcorrencia]),
        dataOc: dataOc instanceof Date ? Utilities.formatDate(dataOc, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(dataOc || ""),
        tipos: tipos,
        valor: parseNumeroBR(r[idx.ocComplemento]),
        descricao: (typeof desc === "string" && desc.trim() !== "" && desc.trim() !== "-") ? limpar(desc) : null
      });
    }
  }

  const resultado = ordem.map(k => romaneios[k]);

  // Marca o status de cada romaneio a partir da aba "Envios":
  // - "nunca": nunca foi preparado
  // - "enviado": preparado e travado (bolinha verde)
  // - "liberado": foi liberado por senha depois de já ter sido enviado (plaquinha de atenção, pode reenviar)
  const eventos = lerUltimoEvento(ss);
  resultado.forEach(function (r) {
    const chave = r.filial + "|" + r.romaneio + "|" + r.cdTransp;
    const info = eventos[chave];
    if (!info) {
      r.emailStatus = "nunca";
      r.dataEnvioEmail = null;
    } else if (info.tipo === "liberacao") {
      r.emailStatus = "liberado";
      r.dataEnvioEmail = info.data;
    } else {
      r.emailStatus = "enviado";
      r.dataEnvioEmail = info.data;
    }
  });

  return ContentService
    .createTextOutput(JSON.stringify(resultado))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Lê a aba "Envios" e devolve, para cada romaneio (chave), o evento mais
 * recente (seja um preparo de e-mail ou uma liberação pra reenvio).
 */
function lerUltimoEvento(ss) {
  const aba = ss.getSheetByName(ABA_ENVIOS);
  const mapa = {};
  if (!aba) return mapa;
  const valores = aba.getDataRange().getValues();
  // colunas: CHAVE, FILIAL, ROMANEIO, CD_TRANSP, TIPO, DATA, DESTINATARIO, CC, OBS
  for (let i = 1; i < valores.length; i++) {
    const linha = valores[i];
    const chave = linha[0];
    if (!chave) continue;
    const tipo = linha[4] || "preparo";
    const data = linha[5];
    const dataIso = data instanceof Date ? data.toISOString() : String(data || "");
    const atual = mapa[chave];
    if (!atual || dataIso > atual.data) {
      mapa[chave] = { tipo: tipo, data: dataIso };
    }
  }
  return mapa;
}

/**
 * Recebe requisições POST da página.
 * Ações suportadas:
 * - enviarEmail: { action, filial, romaneio, cdTransp, subject, greeting }
 * - solicitarLiberacao: { action, filial, romaneio, cdTransp, nomeSolicitante }
 * - liberarRomaneio: { action, filial, romaneio, cdTransp, senha }
 */
function doPost(e) {
  let payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ success: false, error: "Requisição inválida." });
  }

  if (payload.action === "enviarEmail") return enviarEmailRomaneio(payload);
  if (payload.action === "solicitarLiberacao") return solicitarLiberacao(payload);
  if (payload.action === "liberarRomaneio") return liberarRomaneio(payload);

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
    if (zpad(r[idx.filial], LARGURA_FILIAL) === filial && zpad(r[idx.romaneio], LARGURA_ROMANEIO) === romaneio && zpad(r[idx.cdTransp], LARGURA_TRANSP) === cdTransp) {
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
    const nrNf = zpad(r[idx.nrNf], LARGURA_NF);
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
  // acima da tabela, uma linha por ocorrência: "➢ Ocorrência: 000 - descrição VALOR: X,XX"
  const linhasOcorrenciaHtml = [];
  const linhasOcorrenciaTexto = [];
  linhas.forEach(function (r) {
    const descrSub = limpar(r[idx.descrSubOc]);
    const mot = limpar(r[idx.motOcor]);
    if (descrSub.indexOf("025") === 0 || mot.indexOf("025") === 0) {
      const nrOc = limpar(r[idx.nrOcorrencia]);
      // remove um "VALOR: ..." que às vezes já vem dentro do próprio texto da ocorrência,
      // pra não duplicar quando a gente adiciona o nosso VALOR formatado no final
      const desc = limpar(r[idx.ocorrencia]).replace(/\s*VALOR\s*:?\s*[\d.,]+\s*$/i, "").trim();
      const valorOc = parseNumeroBR(r[idx.ocComplemento]);
      let linha = "Ocorrência: " + nrOc;
      if (desc) linha += " - " + desc;
      if (valorOc != null) linha += " VALOR: " + valorOc.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      linhasOcorrenciaHtml.push("<p style='margin:2px 0'>➢ " + linha + "</p>");
      linhasOcorrenciaTexto.push("➢ " + linha);
    }
  });
  const ocorrenciasHtml = linhasOcorrenciaHtml.join("");
  const ocorrenciasTexto = linhasOcorrenciaTexto.length ? linhasOcorrenciaTexto.join("\n") + "\n\n" : "";

  const saudacao = payload.greeting || "Bom dia";
  const corpoHtml =
    "<div style=\"font-family:Arial,sans-serif;font-size:13px;color:#1a1a1a;line-height:1.5\">"
    + "<p>" + saudacao + " a todos, Favor enviar o cte nesse mesmo e-mail, não tire ninguém da cópia.</p>"
    + "<p><b>OBS:- FAVOR SEMPRE CONFERIR SE AS NOTAS ANEXA ESTÃO CONFERINDO COM O RELATÓRIO ABAIXO.</b></p>"
    + "<p>Segue anexo as notas e abaixo a formação para emissão do cte, lembrando que após a emissão devem enviar nesse mesmo e-mail cte para lançamento e posterior pagamento.</p>"
    + "<p>Comprovantes de despesas enviar anexo no e-mail junto com o cte:</p>"
    + ocorrenciasHtml
    + tabela
    + "</div>";

  let corpoTexto =
    saudacao + " a todos, Favor enviar o cte nesse mesmo e-mail, não tire ninguém da cópia.\n\n"
    + "OBS:- FAVOR SEMPRE CONFERIR SE AS NOTAS ANEXA ESTÃO CONFERINDO COM O RELATÓRIO ABAIXO.\n\n"
    + "Segue anexo as notas e abaixo a formação para emissão do cte, lembrando que após a emissão devem enviar nesse mesmo e-mail cte para lançamento e posterior pagamento.\n\n"
    + "Comprovantes de despesas enviar anexo no e-mail junto com o cte:\n\n"
    + ocorrenciasTexto
    + "DT_EMISSAO_NF\tNR_ROMANEIO\tDS_TRANSP\tDS_MOTORISTA\tPLACA\tPESO\tVLR_FRETE\tVALOR_NF\tNR_NF\tCHAVENF\n";

  linhas.forEach(function (r, i) {
    const nrNf = zpad(r[idx.nrNf], LARGURA_NF);
    const chave = limpar(r[idx.chave]);
    if (i === 0) {
      corpoTexto += [dtStr, romaneio, transportadora, motorista, placa, fmtNum(peso), fmtMoeda(frete), fmtMoeda(valorCarga), nrNf, chave].join("\t") + "\n";
    } else {
      corpoTexto += ["", "", "", "", "", "", "", "", nrNf, chave].join("\t") + "\n";
    }
  });

  const assunto = payload.subject ||
    ("EMISSÃO CTE - CARGA: " + romaneio + " - TRANSPORTADORA: " + transportadora + " - DATA EMISSÃO NF: " + dtStr);

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
  registrarEvento(ss, filial, romaneio, cdTransp, "preparo", destinatario, cc, "");
}

/**
 * Garante que a aba "Envios" existe com o cabeçalho certo (colunas novas
 * inclusas), mesmo se ela já existia no formato antigo.
 */
function getAbaEnvios(ss) {
  let aba = ss.getSheetByName(ABA_ENVIOS);
  if (!aba) {
    aba = ss.insertSheet(ABA_ENVIOS);
    aba.appendRow(["CHAVE", "FILIAL", "ROMANEIO", "CD_TRANSP", "TIPO", "DATA", "DESTINATARIO", "CC", "OBS"]);
  } else {
    const cab = aba.getRange(1, 1, 1, Math.max(aba.getLastColumn(), 9)).getValues()[0];
    if (cab[4] !== "TIPO") {
      // Planilha ainda no formato antigo (sem coluna TIPO) - insere e ajusta
      aba.insertColumnBefore(5);
      aba.getRange(1, 5).setValue("TIPO");
      const nLinhas = aba.getLastRow();
      if (nLinhas > 1) {
        aba.getRange(2, 5, nLinhas - 1, 1).setValue("preparo");
      }
    }
  }
  return aba;
}

function registrarEvento(ss, filial, romaneio, cdTransp, tipo, destinatario, cc, obs) {
  const aba = getAbaEnvios(ss);
  const chave = filial + "|" + romaneio + "|" + cdTransp;
  aba.appendRow([chave, filial, romaneio, cdTransp, tipo, new Date(), destinatario || "", cc || "", obs || ""]);
}

/**
 * E-mails autorizados a liberar um romaneio pra reenvio (recebem o pedido
 * de liberação e são os únicos "esperados" a saber a senha).
 */
const EMAILS_LIBERACAO = [
  "diego.lopes@gtf.com.br",
  "geovana.canevarolli1@gtf.com.br",
  "raphael.sauer@gtf.com.br",
  "gabrieli.lima@gtf.com.br"
];

/**
 * Senha compartilhada pra liberar reenvio de um romaneio já enviado.
 * TROQUE ESSE VALOR e não compartilhe o código-fonte com quem não deveria ter a senha.
 */
const SENHA_LIBERACAO = "gtf2026";

/**
 * Manda um e-mail pros responsáveis pedindo pra liberarem um romaneio
 * específico pra reenvio dentro da plataforma.
 */
function solicitarLiberacao(payload) {
  const filial = String(payload.filial || "");
  const romaneio = String(payload.romaneio || "");
  const cdTransp = String(payload.cdTransp || "");
  const transportadora = limpar(payload.transportadora || "");
  const nome = limpar(payload.nomeSolicitante || "Alguém da equipe");

  const assunto = "Liberação de reenvio - Romaneio " + romaneio;
  const corpo =
    "<p>" + nome + " solicitou a liberação do romaneio <b>" + romaneio + "</b>"
    + (transportadora ? " (" + transportadora + ")" : "") + " para reenvio de e-mail dentro da plataforma.</p>"
    + "<p>Esse romaneio já tinha sido enviado antes e está travado. Se estiver de acordo, entre na plataforma e libere com a senha.</p>";

  try {
    GmailApp.sendEmail(EMAILS_LIBERACAO.join(","), assunto, "", {
      htmlBody: corpo,
      name: "Envio de Cargas - Criação de CTe"
    });
  } catch (err) {
    return jsonResponse({ success: false, error: "Erro ao enviar o pedido de liberação: " + err });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  registrarEvento(ss, filial, romaneio, cdTransp, "pedido_liberacao", EMAILS_LIBERACAO.join(","), "", "Solicitado por: " + nome);

  return jsonResponse({ success: true });
}

/**
 * Confere a senha e, se estiver certa, libera o romaneio pra reenvio
 * (o próximo "enviarEmail" vai passar a funcionar de novo).
 */
function liberarRomaneio(payload) {
  const senha = String(payload.senha || "");
  if (senha !== SENHA_LIBERACAO) {
    return jsonResponse({ success: false, error: "Senha incorreta." });
  }

  const filial = String(payload.filial || "");
  const romaneio = String(payload.romaneio || "");
  const cdTransp = String(payload.cdTransp || "");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  registrarEvento(ss, filial, romaneio, cdTransp, "liberacao", "", "", "Liberado manualmente na plataforma");

  return jsonResponse({ success: true });
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
