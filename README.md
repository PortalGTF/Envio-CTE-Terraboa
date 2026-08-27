# Envio de Cargas — Criação de CTe

Painel para acompanhar romaneios, notas fiscais e ocorrências (código 025) puxando os dados direto de uma planilha do Google Sheets.

## Estrutura

```
index.html          → o painel (abre no navegador ou publica no GitHub Pages)
apps-script/Codigo.gs → script que conecta a planilha ao painel
```

## Passo 1 — Preparar a planilha

Crie (ou use) uma planilha no Google Sheets com uma aba chamada exatamente **`Dados`**, contendo as colunas do export do BI:

```
FILIAL, DT_EMISSAO_NF, NR_ROMANEIO, CD_TRANSP, DS_TRANSP, DS_MOTORISTA, PLACA,
PESO, TARIFA, VLR_FRETE, VALOR_NF, CIDADE, UF, LOJA, NR_NF, CHAVENF,
OC_COMPLEMENTO, NR_OCORRENCIA, DATA_OC, OCORRENCIA, DESCR_SUB_OC, MOT_OCOR, CTE
```

Essa é a aba que você atualiza a cada 3 dias com o export do BI (só a sua filial).

## Passo 2 — Publicar o script que alimenta o painel

1. Na planilha: **Extensões → Apps Script**.
2. Apague o conteúdo do `Code.gs` e cole o conteúdo de `apps-script/Codigo.gs` deste repositório.
3. Salve (Ctrl+S).
4. **Implantar → Nova implantação**:
   - Tipo: **App da Web**
   - Executar como: **Eu**
   - Quem pode acessar: **Qualquer pessoa**
5. Implante e copie a URL gerada (termina em `/exec`).

> Se você já tem uma implantação antiga (de uma versão anterior do script), edite ela em **Implantar → Gerenciar implantações → editar → Nova versão**, em vez de criar uma nova — assim a URL não muda.

## Passo 3 — Conectar o painel à planilha

Abra `index.html` em um editor de texto, encontre a linha:

```js
const SCRIPT_URL = "";
```

e cole a URL do passo 2 entre as aspas.

## Passo 4 — Publicar no GitHub Pages (opcional)

1. Crie um repositório novo no GitHub e suba os arquivos deste projeto (`index.html` e a pasta `apps-script/`).
2. Vá em **Settings → Pages** do repositório.
3. Em "Source", selecione a branch `main` e a pasta `/root`.
4. Salve. Em alguns minutos o painel estará disponível em `https://SEU-USUARIO.github.io/NOME-DO-REPO/`.

Sem GitHub Pages, o `index.html` também funciona normalmente aberto direto no navegador (duplo clique).

## Como funciona

- `index.html` busca os dados publicados pelo `doGet` do Apps Script sempre que é aberto (ou quando você clica no botão ↻ no topo).
- Se `SCRIPT_URL` estiver vazio ou a conexão falhar, o painel cai automaticamente em modo demonstração, usando um instantâneo de dados salvo dentro do próprio arquivo — assim ele nunca fica em branco.
- Ocorrências mostradas são filtradas apenas para o código **025** (complemento de frete e diária), conforme os campos `MOT_OCOR` e `DESCR_SUB_OC` da planilha.
- Peso, frete e valor da carga são lidos uma única vez por romaneio (esses campos vêm repetidos em cada linha do export do BI).
