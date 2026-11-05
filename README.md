# Contraprova · NF-e destinadas à União (jan–out/2026)

Análise de 1.311.245 notas fiscais eletrônicas e 4.885.929 itens publicados no Portal da Transparência,
com foco em notas divergentes e na produção de cada estado. O resultado é uma apresentação em HTML
(com mapa do Brasil interativo) e uma versão em PowerPoint.

**Contraprova é uma organização fictícia**, criada para este estudo acadêmico (pós-graduação em Gestão de TI).

- Apresentação HTML: [`index.html`](index.html) (publicável no GitHub Pages)
- Apresentação PowerPoint: [`apresentacao/contraprova_nfe.pptx`](apresentacao/contraprova_nfe.pptx)

## Premissas

| | |
|---|---|
| **Escopo** | NF-e de mercadorias (modelo 55) emitidas por fornecedores para órgãos do governo federal, quase todos do Poder Executivo, de jan a out/2026. |
| **Fora da base** | Compras de estados e municípios, e serviços (limpeza, vigilância, consultoria, obras), que usam nota fiscal de serviço municipal. |
| **Nota não é pagamento** | A nota mostra o que o fornecedor faturou. Não confirma empenho, entrega nem pagamento. |
| **Uma linha, uma nota** | No arquivo de notas, cada linha é uma nota fiscal (após remover duplicadas). |
| **Item é linha de produto** | Os itens são linhas de produto dentro das notas, não a quantidade comprada. |
| **Valor × soma dos itens** | Valor da nota é o total do cabeçalho; soma dos itens é a soma do valor total de cada linha de produto. |
| **Nota sinalizada** | Nota que caiu em pelo menos uma checagem de divergência. Indica que precisa ser conferida, não que é irregular. |

## Estrutura

```
contraprova-nfe/
├── index.html                      apresentação HTML (arquivo único)
├── etl/
│   └── tratar_nfs.py               ETL: leitura, limpeza, checagens e agregações
├── apresentacao/
│   ├── template.html               modelo da apresentação HTML
│   ├── montar_apresentacao.py      preenche o modelo com os números do ETL
│   ├── mapa_svg.py                 gera os contornos dos estados a partir de um GeoJSON
│   ├── montar_pptx.js              gera o PowerPoint (pptxgenjs)
│   ├── aplicar_tema.js             grava as cores do tema no PowerPoint
│   └── contraprova_nfe.pptx        apresentação PowerPoint
├── dados/
│   ├── resumo.json                 números agregados usados nas apresentações
│   ├── producao_por_uf.csv         valor e notas por UF (emitente e destinatário)
│   ├── R9_precos_atipicos.csv      itens com preço ≥ 10× a mediana de itens similares
│   ├── notas_divergentes.zip       todas as notas sinalizadas, com a checagem que acionou
│   └── mapa.json                   contornos simplificados das 27 UFs
└── dados_brutos/                   CSVs do portal (não versionados, ver LEIAME.md)
```

## ETL

**Extração.** 10 arquivos mensais do Portal da Transparência, cada um com três CSVs (notas, itens e
eventos), cerca de 3,9 GB descompactados. Leitura com pandas: separador `;`, encoding latin-1,
tudo como texto e só as colunas usadas.

**Transformação** (`etl/tratar_nfs.py`)
- Valores no formato brasileiro (`1.234,56`) e datas (`dd/mm/aaaa`) convertidos.
- Remoção de 40.527 linhas de nota e 157.674 linhas de item publicadas em duplicidade.
- Agregação dos itens por chave de acesso e junção com o cabeçalho da nota.
- Checagens de divergência, todas vetorizadas: valor da nota × soma dos itens, quantidade × valor
  unitário, decomposição da chave de acesso (UF, ano/mês, CNPJ, série, número, DV módulo 11),
  duplicidade de série/número, CFOP × destino, nota zerada ou sem itens, nota sem autorização,
  datas inconsistentes e preço unitário atípico (mediana por NCM + unidade + primeira palavra da descrição).
- Agregações por UF, fluxo entre estados, mês, órgão e fornecedor. CPFs de emitentes ficam mascarados.

**Carga.** Saídas em JSON e CSV (`;`, vírgula decimal, UTF-8) prontas para Excel e para as apresentações.

Bibliotecas: **pandas** e **NumPy** (Python); **pptxgenjs** e **JSZip** (Node.js) para o PowerPoint.
O restante é biblioteca padrão.

## Como reproduzir

```bash
# 1. Python
pip install -r requirements.txt
python etl/tratar_nfs.py dados_brutos saida          # ~4 minutos para 10 meses

# 2. Apresentação HTML
python apresentacao/montar_apresentacao.py apresentacao/template.html saida/resumo.json dados/mapa.json saida/apresentacao.html
#    gera saida/apresentacao.html e saida/apresentacao_standalone.html (esta é a versão do index.html)

# 3. PowerPoint
npm install
node apresentacao/montar_pptx.js saida/resumo.json dados/mapa.json apresentacao/contraprova_nfe.pptx
```

O `dados/mapa.json` já vem pronto. Para regerar, baixe o GeoJSON de estados do projeto
[click_that_hood](https://github.com/codeforgermany/click_that_hood/blob/main/public/data/brazil-states.geojson) e rode
`python apresentacao/mapa_svg.py brazil-states.geojson dados/mapa.json`.

## Publicar no GitHub Pages

Settings → Pages → *Deploy from a branch* → branch `main`, pasta `/ (root)`. A apresentação fica em
`https://<seu-usuario>.github.io/contraprova-nfe/`.

## Fonte e avisos

- Dados: Portal da Transparência, Notas Fiscais Eletrônicas, competências 202601 a 202610.
- Notas sinalizadas são indícios estatísticos e de consistência, sujeitos a verificação com contrato e empenho.
  Citar um fornecedor nesta análise não significa que houve irregularidade.
- Descontos legais em medicamentos: Convênio ICMS 87/2002 e Resolução CTE-CMED nº 1/2026 (CAP).
