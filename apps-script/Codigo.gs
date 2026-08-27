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

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(ABA_DADOS);
  const dados = aba.getDataRange().getValues();
  const cab = dados[0];

  const idx = {
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

  return ContentService
    .createTextOutput(JSON.stringify(resultado))
    .setMimeType(ContentService.MimeType.JSON);
}

function limpar(v) {
  if (v === null || v === undefined) return "";
  return String(v).replace(/\s+/g, " ").trim();
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
