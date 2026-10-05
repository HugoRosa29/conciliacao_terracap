# Conciliação Financeira — Terracap

Conciliação diária da cobrança do BRB: a francesinha do dia
contra o arquivo de retorno (DCB) e contra os dois relatórios do
GAI, com a identidade visual *Eixos & Curvas* da Terracap.

Reproduz automaticamente o que antes era feito à mão na planilha
*Conciliação Boleto c. movimento* — e vai além dela: aponta,
**nome por nome**, qual cobrança não baixou e por quê. Na
conciliação de 30/09/2026 a planilha manual parava numa
divergência de R$ 4.213,39 com a anotação *"Estimativa GAI, é
preciso especificar"*. O programa diz de quem é: boleto 834036,
alienação 101071, que o GAI leu e registrou com valor 0,00.

A aplicação roda **inteiramente dentro do navegador**. Não há
servidor, upload nem armazenamento: os arquivos financeiros nunca
saem da máquina de quem usa.

## As três conferências

### Bolebarra × DCB — título por título

A francesinha Bolebarra traz o **nosso número** de cada cobrança,
e o arquivo de retorno também. A conferência casa os dois por esse
número, não pelo valor, e separa:

| Situação | O que significa |
| --- | --- |
| Conferida | está liquidada no retorno, pelo mesmo valor |
| Valor diferente | liquidou no retorno por outro valor |
| Sem baixa | está na francesinha e não liquidou no retorno |
| Fora da francesinha | liquidou no retorno e não está na Bolebarra — em geral são os Bolepix do dia |

### Bolebarra × GAI — o que o GAI fez com cada cobrança

O nosso número do BRB tem 12 dígitos e carrega, nos seis do meio,
o número do boleto que o GAI imprime nos relatórios dele:

```
6 8 3 4 3 4 2 0 7 0 1 5
|  \____ ____/  \_/ \_/
|       |        |   dígitos verificadores
|       |        código do banco (070 = BRB)
|       boleto — o "Nosso Número" do GAI
prefixo da carteira
```

Por esse número, cada cobrança da francesinha encontra o seu par
na *Relação de Parcelas Lidas*. Um boleto pode pagar várias
parcelas, e o relatório imprime uma linha por parcela: somadas por
boleto, elas fecham com o valor da francesinha. Daí saem quatro
situações:

| Situação | O que significa |
| --- | --- |
| Baixado no GAI | o GAI baixou, e o valor fecha |
| Lido com valor 0,00 | o GAI leu o pagamento e não aplicou a baixa |
| Recusado | o GAI recusou, e o relatório de *Baixas Não Efetivadas* diz o motivo |
| Fora dos relatórios do GAI | não está em nenhum dos dois — em geral é cobrança de outra gerência (GIR, GGR, GOP) |

### Bolepix × DCB — pelo valor

O relatório de PIX do banco não traz nosso número, então aqui a
conferência é pelo valor: aponta, por nome e valor, os
recebimentos que não têm baixa no arquivo de retorno.

### E, se o extrato vier, o fechamento do dia

| Bloco | Origem | Resultado |
| --- | --- | --- |
| Bolepix | `CRED PIX QR CODE DINAMICO` no extrato | total recebido por PIX |
| Bolebarras | `CREDITO COBRANCA BRB` no extrato | total recebido por código de barras |
| Total de pagamentos | soma dos dois | *Total pagamentos QR Code + C.Barras* |
| Arquivo de retorno | títulos pagos no DCB (CNAB 400) | *Total arquivo de retorno* |
| Divergência | pagamentos − retorno | o que o dia não fechou |

**O extrato é opcional.** Ele fecha o total do dia; as três
conferências acima são título por título e não dependem dele.

## Como usar

### No navegador (GitHub Pages)

Abra a página publicada, arraste os arquivos do dia e clique em
**Conciliar**. Cada conferência aparece quando os arquivos dela
chegam:

- **DCB** — o arquivo de retorno da cobrança (TXT, CNAB 400)
- **Francesinha Bolebarra** — o relatório *Francesinha movimento* (XLS)
- **GAI / Parcelas lidas** — *Relação de Parcelas Lidas* (PDF)
- **GAI / Não baixadas** — *Baixas de Pagamentos Não Efetivadas* (PDF)
- **Francesinha Bolepix** — o relatório *Extrato/Devolução* (XLS)
- **Extrato BRB** — o TXT do Internet Banking (opcional)

Os dois PDFs do GAI se identificam pelo título impresso dentro
deles, então trocá-los de lugar não muda o resultado.

O botão **Baixar Excel** gera a planilha, com uma aba por bloco —
entre elas *Bolebarra nao baixada*, que é só o que não fechou.

### Na linha de comando

```bash
pip install -r requirements.txt

python main.py --pasta ../PROMPT/"TESTE 2"
python main.py --pasta ../PROMPT/"TESTE 2" --excel conciliacao.xlsx

python main.py --dcb "DCB_1219001012_30092026_031844.txt"                --bolebarras "Francesinha movimento Bolebarra 30 09 2026.xlt"                --gai-lidas "VALORES BAIXADOS NO GAI.pdf"                --gai-nao-baixadas "VALORES NAO BAIXADOS NO GAI.pdf"
```

Com `--pasta`, os arquivos são reconhecidos pelo nome, e os `.pdf`
da pasta vão para os campos do GAI na ordem em que aparecem.

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
gai.py           leitura dos relatórios do GAI em PDF
exportar.py      geração do Excel (xlsxwriter)
ponte.py         ligação entre o navegador e o núcleo
main.py          interface de linha de comando
app.py           servidor local opcional (FastAPI)
validar.py       validação contra os arquivos reais
build.py         copia o núcleo para docs/py/
docs/            o site publicado no GitHub Pages
```

O mesmo núcleo em Python roda nos três lugares — navegador, linha
de comando e servidor. Ele depende apenas de `xlrd`, `openpyxl`,
`XlsxWriter` e `pypdf`, que são bibliotecas Python puras; é por
isso que o Pyodide consegue carregá-lo no navegador sem baixar
pandas. Os wheels ficam em `docs/vendor/`, servidos da própria
página.

## Validação

`validar.py` confere o resultado contra os arquivos reais de dois
dias e contra os números das planilhas feitas à mão:

```bash
python validar.py ../PROMPT
```

28/09/2026 fecha o dia pelo extrato. 30/09/2026 não tem extrato: é
o dia em que a conferência passa pelos relatórios do GAI.

```
28/09/2026 — fechamento pelo extrato
  OK   Total pagamentos QR Code + C.Barras    R$ 4.332.490,76
  OK   Total arquivo de retorno               R$ 4.319.046,36
  OK   Divergência                                R$ 13.444,40

30/09/2026 — conferência com o GAI
  OK   Cobranças casadas pelo nosso número               270
  OK   Total conferido                        R$ 9.740.777,55
  OK   Total baixado no GAI                  R$ 10.013.491,52
  OK   Cobranças NÃO baixadas no GAI                      29
  OK   Lidas pelo GAI com valor 0,00               R$ 4.213,39
  OK   Boleto da cobrança lida sem baixa              834036

43/43 verificações passaram
```

Os totais lidos dos PDFs são conferidos contra o rodapé impresso
pelo próprio GAI (*TOTAL GERAL RECEBIDO*, *Total de Valores Não
Baixados*): se a leitura escorregar, o resultado sai com aviso em
vez de sair errado.

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

## Sobre os relatórios do GAI

Os dois relatórios são texto posicionado: cada campo é desenhado
numa coordenada fixa da página. `gai.py` agrupa os pedaços de
texto por linha e decide a coluna de cada um pela coordenada
horizontal — as faixas estão em `COLUNAS_LIDAS` e
`COLUNAS_NAO_EFETIVADAS`, conferidas contra todas as páginas dos
relatórios reais.

Na *Relação de Parcelas Lidas*, uma linha com valor 0,00 não é um
pagamento de R$ 0,00: é um boleto que o GAI leu do arquivo e não
aplicou, e o número dele sai impresso na coluna *Imóvel*. É assim
que a cobrança de R$ 4.213,39 de 30/09/2026 aparece — lida, nunca
baixada.

## Identidade visual

A interface segue o conceito *Eixos & Curvas* da Terracap: a
estrutura reta de Lucio Costa (eixos, coordenadas, carimbos de
status, tabelas monoespaçadas) suavizada pelas curvas de Niemeyer
(molduras em arco, nunca ângulo reto puro).

A cor comunica o status antes do texto: **verde** para o que
fecha, **dourado** para o que merece atenção, **vermelho** para
divergência. O fundo azul profundo ("modo satélite") aparece uma
única vez na página, reservado para a prestação de contas — o que
não fechou.

Tipografia: Space Grotesk (títulos), Public Sans (texto) e
JetBrains Mono (valores, códigos e coordenadas).

## Aviso

O extrato, as francesinhas, o DCB e os PDFs do GAI contêm dados
financeiros e pessoais — nome, CPF/CNPJ, alienação e valor de cada
comprador. O `.gitignore` já impede que sejam versionados por
engano — confira antes de cada commit.
