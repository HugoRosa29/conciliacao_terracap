# Conciliação Financeira — Terracap

Conciliação diária do extrato do BRB com a francesinha Bolepix e
o arquivo de retorno da cobrança (DCB), com a identidade visual
*Eixos & Curvas* da Terracap.

Reproduz automaticamente o que antes era feito à mão na planilha
*Conciliação Boleto c. movimento* — e, além disso, aponta **nome
por nome** quais Bolepix não baixaram no arquivo de retorno, que
é justamente a parte que a planilha não resolvia.

A aplicação roda **inteiramente dentro do navegador**. Não há
servidor, upload nem armazenamento: os arquivos financeiros nunca
saem da máquina de quem usa.

## O que ela calcula

| Bloco | Origem | Resultado |
| --- | --- | --- |
| Bolepix | `CRED PIX QR CODE DINAMICO` no extrato | total recebido por PIX |
| Bolebarras | `CREDITO COBRANCA BRB` no extrato | total recebido por código de barras |
| Total de pagamentos | soma dos dois | *Total pagamentos QR Code + C.Barras* |
| Arquivo de retorno | títulos pagos no DCB (CNAB 400) | *Total arquivo de retorno* |
| Divergência | pagamentos − retorno | Bolepix não baixado, item a item |

A francesinha Bolepix ainda serve para **identificar o pagador**
de cada crédito PIX do extrato, preenchendo sozinha as colunas
*Cliente / CNPJ / Data-Hora* que antes eram digitadas à mão.

## Como usar

### No navegador (GitHub Pages)

Abra a página publicada, arraste os três arquivos do dia e clique
em **Conciliar**:

- **Extrato BRB** — o TXT do Internet Banking
- **Francesinha Bolepix** — o relatório *Extrato/Devolução* (XLS)
- **DCB** — o arquivo de retorno da cobrança (TXT, CNAB 400)

A francesinha Bolebarras é opcional: serve para conferir também o
total da cobrança por código de barras.

O botão **Baixar Excel** gera uma planilha com cinco abas
(Resumo, PIX detalhado, Bolepix não baixado, DCB liquidados e
Extrato créditos).

### Na linha de comando

```bash
pip install -r requirements.txt

python main.py --pasta ../PROMPT
python main.py --pasta ../PROMPT --excel conciliacao.xlsx

python main.py --extrato "28 09 2026 m.txt" \
               --francesinha "Francesinha bolepix.xls" \
               --dcb "DCB_1219001012.txt"
```

### Servidor local (opcional)

Só para quem preferir rodar na intranet em vez do GitHub Pages:

```bash
python -m uvicorn app:app --reload
```

## Publicar no GitHub Pages

A pasta `docs/` já é o site pronto.

1. `git push` para o branch `main`.
2. Em **Settings → Pages**, escolha **GitHub Actions** como origem
   (o workflow `.github/workflows/pages.yml` publica `docs/`).

   Alternativa sem workflow: escolha **Deploy from a branch**,
   branch `main`, pasta `/docs`.

O workflow confere, antes de publicar, se `docs/py/` está em dia
com o núcleo.

> Depois de alterar qualquer módulo do núcleo, rode `python build.py`
> para atualizar `docs/py/`.

## Como está organizado

```
conciliacao.py   núcleo: leitura dos arquivos e conciliação
planilhas.py     leitura de .xls/.xlsx sem pandas
exportar.py      geração do Excel (xlsxwriter)
ponte.py         ligação entre o navegador e o núcleo
main.py          interface de linha de comando
app.py           servidor local opcional (FastAPI)
validar.py       validação contra os arquivos reais
build.py         copia o núcleo para docs/py/
docs/            o site publicado no GitHub Pages
```

O mesmo núcleo em Python roda nos três lugares — navegador, linha
de comando e servidor. Ele depende apenas de `xlrd`, `openpyxl` e
`XlsxWriter`, que são bibliotecas Python puras; é por isso que o
Pyodide consegue carregá-lo no navegador sem baixar pandas.

## Validação

`validar.py` confere o resultado contra os arquivos reais de
28/09/2026 e contra os números da planilha feita à mão:

```bash
python validar.py ../PROMPT
```

```
EXTRATO BRB
  OK   Total PIX (Bolepix)              R$ 687.438,67
  OK   Total cobrança (Bolebarras)      R$ 3.645.052,09
  OK   Total pagamentos                 R$ 4.332.490,76
DCB - ARQUIVO DE RETORNO
  OK   Total arquivo de retorno         R$ 4.319.046,36
DIVERGÊNCIA
  OK   Divergência                      R$ 13.444,40
  OK   Itens Bolepix não baixados       5

14/14 verificações passaram
```

## Sobre o layout do DCB

O DCB é um retorno de cobrança CNAB 400 do BRB (registros de 400
caracteres). As posições usadas estão em `CAMPOS_DCB`, em
[`conciliacao.py`](conciliacao.py), e foram confirmadas contra o
arquivo real e contra a conciliação manual.

Um ponto importante: um título é considerado **liquidado** quando
o *valor pago* (posições 258–270) é maior que zero — e não pelo
código de ocorrência. As ocorrências `005` e `105` representam
liquidação, e somar apenas a `005` deixaria dois títulos
(R$ 5.740,60) de fora do total do arquivo.

## Identidade visual

A interface segue o conceito *Eixos & Curvas* da Terracap: a
estrutura reta de Lucio Costa (eixos, coordenadas, carimbos de
status, tabelas monoespaçadas) suavizada pelas curvas de Niemeyer
(molduras em arco, nunca ângulo reto puro).

A cor comunica o status antes do texto: **verde** para o que
fecha, **dourado** para o que merece atenção, **vermelho** para
divergência. O fundo azul profundo ("modo satélite") aparece uma
única vez na página, reservado para o bloco da divergência.

Tipografia: Space Grotesk (títulos), Public Sans (texto) e
JetBrains Mono (valores, códigos e coordenadas).

## Aviso

Os arquivos de extrato, francesinha e DCB contêm dados
financeiros e pessoais. O `.gitignore` já impede que sejam
versionados por engano — confira antes de cada commit.
