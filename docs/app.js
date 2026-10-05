/*
 * Interface da conciliação financeira.
 *
 * Carrega o Pyodide, instala as bibliotecas de planilha
 * (Python puro, servidas da própria pasta vendor/) e roda o
 * mesmo núcleo em Python usado pelo CLI e pelo servidor.
 *
 * Nenhum arquivo é enviado para a rede: o conteúdo é lido como
 * bytes e passado direto para o Python dentro do navegador.
 */

"use strict";

const MODULOS = [
  "planilhas.py",
  "gai.py",
  "conciliacao.py",
  "exportar.py",
  "ponte.py",
];

const WHEELS = [
  "vendor/xlrd-2.0.2-py2.py3-none-any.whl",
  "vendor/et_xmlfile-2.0.0-py3-none-any.whl",
  "vendor/openpyxl-3.1.5-py2.py3-none-any.whl",
  "vendor/xlsxwriter-3.2.9-py3-none-any.whl",
  "vendor/pypdf-6.19.0-py3-none-any.whl",
];

const CAMPOS = [
  "dcb",
  "bolebarras",
  "gai_lidas",
  "gai_nao_baixadas",
  "francesinha",
  "extrato",
];

/* Como cada situação da conferência com o GAI é escrita na
   tela, e a cor que ela merece. */
const SITUACAO_GAI = {
  baixado: { texto: "Baixado no GAI", classe: "ok" },
  lido_sem_baixa: {
    texto: "Lido com valor 0,00",
    classe: "encerrado",
  },
  recusado: { texto: "Recusado pelo GAI", classe: "encerrado" },
  ausente: { texto: "Fora dos relatórios do GAI", classe: "estudo" },
};

const arquivos = {};

let pyodide = null;
let ponte = null;
let resultadoAtual = null;

const el = (selector) => document.querySelector(selector);
const todos = (selector) => Array.from(document.querySelectorAll(selector));

const moeda = new Intl.NumberFormat("pt-BR", {
  style: "currency",
  currency: "BRL",
});

const inteiro = new Intl.NumberFormat("pt-BR");

const formatarMoeda = (valor) => moeda.format(Number(valor) || 0);

/* ============================================================
 * Carregamento do motor Python
 * ============================================================ */

/* Carimbo de status: a cor diz o estado antes do texto —
   verde em andamento, dourado em estudo, vermelho encerrado. */
function estado(texto, classe) {
  el("#estado-texto").textContent = texto;
  el("#estado-runtime").className = `carimbo carimbo-${classe}`;
}

async function iniciarPython() {
  try {
    estado("Carregando o motor de cálculo", "estudo");

    pyodide = await loadPyodide({
      stdout: () => {},
      stderr: (linha) => console.warn("[python]", linha),
    });

    estado("Instalando bibliotecas de planilha", "estudo");

    await pyodide.loadPackage("micropip");

    const micropip = pyodide.pyimport("micropip");
    const base = new URL(".", window.location.href).href;

    // Lista convertida para Python: os wheels são de Python
    // puro e já incluem as dependências, por isso deps=False.
    const lista = pyodide.toPy(WHEELS.map((caminho) => base + caminho));

    try {
      await micropip.install(lista, false, false);
    } finally {
      lista.destroy();
    }

    estado("Carregando a conciliação", "estudo");

    for (const modulo of MODULOS) {
      const resposta = await fetch(`py/${modulo}`, { cache: "no-cache" });

      if (!resposta.ok) {
        throw new Error(`Não foi possível carregar py/${modulo}`);
      }

      pyodide.FS.writeFile(`/home/pyodide/${modulo}`, await resposta.text());
    }

    ponte = pyodide.pyimport("ponte");

    estado("Pronto · nada sai deste computador", "andamento");

    atualizarBotao();
  } catch (erro) {
    console.error(erro);
    estado("Falha ao carregar", "encerrado");

    mostrarMensagem(
      "Não foi possível iniciar o motor de cálculo. Verifique a " +
        "conexão e recarregue a página.",
      true,
    );
  }
}

/* ============================================================
 * Seleção de arquivos
 * ============================================================ */

function registrarCaixa(caixa) {
  const campo = caixa.dataset.campo;
  const entrada = caixa.querySelector("input[type=file]");

  entrada.addEventListener("change", () => {
    if (entrada.files.length) {
      definirArquivo(campo, entrada.files[0]);
    }
  });

  ["dragenter", "dragover"].forEach((evento) =>
    caixa.addEventListener(evento, (e) => {
      e.preventDefault();
      caixa.classList.add("arrastando");
    }),
  );

  ["dragleave", "drop"].forEach((evento) =>
    caixa.addEventListener(evento, (e) => {
      e.preventDefault();
      caixa.classList.remove("arrastando");
    }),
  );

  caixa.addEventListener("drop", (e) => {
    const arquivo = e.dataTransfer?.files?.[0];

    if (arquivo) {
      definirArquivo(campo, arquivo);
    }
  });
}

function definirArquivo(campo, arquivo) {
  arquivos[campo] = arquivo;

  const caixa = el(`.caixa-arquivo[data-campo="${campo}"]`);
  const tamanho = (arquivo.size / 1024).toFixed(0);

  caixa.querySelector(".nome").textContent = `${arquivo.name} · ${tamanho} KB`;
  caixa.classList.add("preenchida");

  mostrarMensagem("");
  atualizarBotao();
}

function atualizarBotao() {
  const temArquivo = CAMPOS.some((campo) => arquivos[campo]);

  el("#botao-conciliar").disabled = !(temArquivo && ponte);
}

function limpar() {
  CAMPOS.forEach((campo) => {
    delete arquivos[campo];

    const caixa = el(`.caixa-arquivo[data-campo="${campo}"]`);

    caixa.classList.remove("preenchida");
    caixa.querySelector(".nome").textContent = "Clique ou arraste o arquivo";
    caixa.querySelector("input[type=file]").value = "";
  });

  resultadoAtual = null;
  el("#resultado").hidden = true;
  mostrarMensagem("");
  atualizarBotao();
}

function mostrarMensagem(texto, erro = false) {
  const alvo = el("#mensagem");

  alvo.textContent = texto;
  alvo.className = erro ? "mensagem erro" : "mensagem";
}

/* ============================================================
 * Execução
 * ============================================================ */

async function conciliar() {
  if (!ponte) {
    return;
  }

  const botao = el("#botao-conciliar");

  botao.disabled = true;
  mostrarMensagem("Processando os arquivos…");

  try {
    const entrada = {};

    for (const campo of CAMPOS) {
      const arquivo = arquivos[campo];

      entrada[campo] = arquivo
        ? {
            nome: arquivo.name,
            dados: new Uint8Array(await arquivo.arrayBuffer()),
          }
        : null;
    }

    const mapa = pyodide.toPy(entrada);

    let bruto;

    try {
      bruto = ponte.processar(mapa);
    } finally {
      mapa.destroy();
    }

    const resultado = JSON.parse(bruto);

    if (resultado.erro) {
      mostrarMensagem(resultado.erro, true);
      return;
    }

    resultadoAtual = resultado;

    desenhar(resultado);
    mostrarMensagem("Conciliação concluída.");

    el("#resultado").scrollIntoView({ behavior: "smooth", block: "start" });
  } catch (erro) {
    console.error(erro);
    mostrarMensagem(`Erro ao processar: ${erro.message || erro}`, true);
  } finally {
    botao.disabled = false;
    atualizarBotao();
  }
}

async function exportar() {
  if (!resultadoAtual || !ponte) {
    return;
  }

  const botao = el("#botao-exportar");

  botao.disabled = true;

  try {
    // bytes do Python chega como PyProxy; versões que já
    // convertem sozinhas devolvem um Uint8Array direto.
    const bytes = ponte.exportar();

    let dados = bytes;

    if (typeof bytes?.toJs === "function") {
      dados = bytes.toJs({ create_proxies: false });
      bytes.destroy();
    }

    const blob = new Blob([dados], {
      type:
        "application/vnd.openxmlformats-officedocument" +
        ".spreadsheetml.sheet",
    });

    const agora = new Date()
      .toISOString()
      .slice(0, 16)
      .replace("T", "_")
      .replace(":", "");

    const link = document.createElement("a");

    link.href = URL.createObjectURL(blob);
    link.download = `conciliacao_${agora}.xlsx`;
    link.click();

    URL.revokeObjectURL(link.href);
  } catch (erro) {
    console.error(erro);
    mostrarMensagem(`Erro ao gerar o Excel: ${erro.message || erro}`, true);
  } finally {
    botao.disabled = false;
  }
}

/* ============================================================
 * Desenho do resultado
 * ============================================================ */

function desenhar(resultado) {
  el("#resultado").hidden = false;

  desenharAvisos(resultado.avisos || []);
  desenharIndicadores(resultado);
  desenharPrestacaoDeContas(resultado);

  desenharBolebarraGai(resultado.bolebarras_gai);
  desenharBolebarraDcb(resultado.bolebarras_dcb);
  desenharNaoBaixados(resultado.bolepix_dcb);
  desenharMotivosGai(resultado.gai_nao_baixadas);
  desenharPix(resultado.pix_detalhado);
  desenharCobranca(resultado.extrato);
  desenharOcorrencias(resultado.dcb);
}

function desenharAvisos(avisos) {
  el("#avisos").innerHTML = avisos
    .map((aviso) => `<div class="aviso">${escapar(aviso)}</div>`)
    .join("");
}

function indicador({ titulo, valor, detalhe = "", classe = "" }) {
  return `
    <div class="indicador ${classe}">
      <div class="titulo">${escapar(titulo)}</div>
      <div class="valor">${escapar(valor)}</div>
      ${detalhe ? `<div class="detalhe">${escapar(detalhe)}</div>` : ""}
    </div>`;
}

function desenharIndicadores(resultado) {
  const { extrato, francesinha, bolebarras, dcb, divergencia } = resultado;

  const cartoes = [];

  if (extrato) {
    cartoes.push(
      indicador({
        titulo: "PIX recebido (Bolepix)",
        valor: formatarMoeda(extrato.total_pix),
        detalhe: `${inteiro.format(extrato.quantidade_pix)} créditos no extrato`,
      }),
      indicador({
        titulo: "Cobrança (Bolebarras)",
        valor: formatarMoeda(extrato.total_cobranca),
        detalhe: `${inteiro.format(
          extrato.quantidade_cobranca,
        )} créditos no extrato`,
      }),
      indicador({
        titulo: "Total de pagamentos",
        valor: formatarMoeda(extrato.total_pagamentos),
        detalhe: "QR Code + código de barras",
      }),
    );
  }

  if (francesinha) {
    const conferencia = francesinha.conferencia_extrato;

    cartoes.push(
      indicador({
        titulo: "Francesinha Bolepix",
        valor: formatarMoeda(francesinha.total),
        detalhe: conferencia
          ? conferencia.confere
            ? "Fecha com o PIX do extrato"
            : `Diferença de ${formatarMoeda(conferencia.diferenca)}`
          : `${inteiro.format(francesinha.quantidade)} recebimentos`,
        classe: conferencia ? (conferencia.confere ? "ok" : "alerta") : "",
      }),
    );
  }

  if (bolebarras) {
    const conferencia = bolebarras.conferencia_extrato;

    cartoes.push(
      indicador({
        titulo: "Francesinha Bolebarra",
        valor: formatarMoeda(bolebarras.total),
        detalhe: conferencia
          ? conferencia.confere
            ? "Fecha com a cobrança do extrato"
            : `Diferença de ${formatarMoeda(conferencia.diferenca)}`
          : `${inteiro.format(bolebarras.quantidade)} cobranças`,
        classe: conferencia ? (conferencia.confere ? "ok" : "alerta") : "",
      }),
    );
  }

  if (dcb) {
    cartoes.push(
      indicador({
        titulo: "Arquivo de retorno (DCB)",
        valor: formatarMoeda(dcb.total_liquidado),
        detalhe: `${inteiro.format(
          dcb.quantidade_liquidados,
        )} títulos liquidados de ${inteiro.format(dcb.quantidade_registros)}`,
      }),
    );
  }

  const comDcb = resultado.bolebarras_dcb;

  if (comDcb) {
    cartoes.push(
      indicador({
        titulo: "Bolebarra x DCB",
        valor: formatarMoeda(comDcb.total_conferido),
        detalhe: comDcb.confere
          ? `${inteiro.format(
              comDcb.quantidade_casados,
            )} cobranças liquidadas, todas pelo valor certo`
          : `${inteiro.format(
              comDcb.quantidade_sem_baixa + comDcb.quantidade_valor_divergente,
            )} cobranças não conferem`,
        classe: comDcb.confere ? "ok" : "alerta",
      }),
    );
  }

  const lidas = resultado.gai_lidas;

  if (lidas) {
    cartoes.push(
      indicador({
        titulo: "Baixado no GAI",
        valor: formatarMoeda(lidas.total_baixado),
        detalhe: `${inteiro.format(lidas.quantidade_boletos)} boletos na Relação de Parcelas Lidas`,
      }),
    );
  }

  const comGai = resultado.bolebarras_gai;

  if (comGai) {
    cartoes.push(
      indicador({
        titulo: "Bolebarra não baixada no GAI",
        valor: formatarMoeda(comGai.total_nao_baixado),
        detalhe: comGai.confere
          ? "Toda a cobrança foi baixada no GAI"
          : `${inteiro.format(comGai.quantidade_nao_baixados)} cobranças`,
        classe: comGai.confere ? "ok" : "alerta",
      }),
    );
  }

  const bolepix = resultado.bolepix_dcb;

  if (bolepix) {
    cartoes.push(
      indicador({
        titulo: "Bolepix não baixado",
        valor: formatarMoeda(bolepix.total_nao_baixados),
        detalhe: bolepix.confere
          ? "Todo o PIX recebido tem baixa no retorno"
          : `${inteiro.format(
              bolepix.quantidade_nao_baixados,
            )} recebimentos sem baixa`,
        classe: bolepix.confere ? "ok" : "alerta",
      }),
    );
  }

  if (divergencia) {
    const zerada = Math.abs(Number(divergencia.divergencia)) < 0.005;

    cartoes.push(
      indicador({
        titulo: "Divergência do dia",
        valor: formatarMoeda(divergencia.divergencia),
        detalhe: zerada
          ? "Nada pendente de baixa"
          : "Pagamentos do extrato menos o arquivo de retorno",
        classe: zerada ? "ok" : "alerta",
      }),
    );
  }

  el("#indicadores").innerHTML =
    cartoes.join("") ||
    '<p class="vazio">Envie os arquivos para ver os totais.</p>';
}

function linhaConta(rotulo, valor, classe = "") {
  return `
    <div class="linha-conta ${classe}">
      <span>${escapar(rotulo)}</span>
      <span>${valor}</span>
    </div>`;
}

/* A prestação de contas do dia: a conferência da Bolebarra com
   o GAI, que não depende do extrato, e — quando o extrato vem —
   a divergência entre o que entrou na conta e o que o arquivo
   de retorno baixou. */
function desenharPrestacaoDeContas(resultado) {
  const cartao = el("#cartao-divergencia");

  const comGai = resultado.bolebarras_gai;
  const divergencia = resultado.divergencia;

  if (!comGai && !divergencia) {
    cartao.hidden = true;
    return;
  }

  cartao.hidden = false;

  const blocos = [];

  if (comGai) {
    const contas = [
      linhaConta(
        "Francesinha Bolebarra",
        formatarMoeda(comGai.total_francesinha),
      ),
      linhaConta(
        `Baixado no GAI (${inteiro.format(comGai.quantidade_baixados)} cobranças)`,
        `&minus; ${formatarMoeda(comGai.total_baixado)}`,
      ),
      linhaConta(
        "Bolebarra não baixada no GAI",
        formatarMoeda(comGai.total_nao_baixado),
        "total",
      ),
    ];

    /* A explicação de cada situação já diz o que aconteceu, por
       isso o carimbo não se repete aqui. */
    const detalhe = (comGai.situacoes || [])
      .map(
        (situacao) => `<li>
            <strong>${formatarMoeda(situacao.total)}</strong>
            em ${inteiro.format(situacao.quantidade)}
            ${situacao.quantidade === 1 ? "cobrança" : "cobranças"}.
            ${escapar(situacao.explicacao || "")}
          </li>`,
      )
      .join("");

    blocos.push(
      contas.join("") +
        (comGai.confere
          ? `<div class="selo ok">
               Toda a cobrança por código de barras foi baixada no GAI.
             </div>`
          : `<div class="selo"><ul class="motivos">${detalhe}</ul></div>`),
    );
  }

  if (divergencia) {
    const explicada = divergencia.explicada;

    blocos.push(
      [
        linhaConta(
          "Total de pagamentos (extrato)",
          formatarMoeda(divergencia.total_pagamentos),
        ),
        linhaConta(
          "Total do arquivo de retorno (DCB)",
          `&minus; ${formatarMoeda(divergencia.total_retorno)}`,
        ),
        linhaConta(
          "Divergência do dia",
          formatarMoeda(divergencia.divergencia),
          "total",
        ),
      ].join("") +
        (explicada === null
          ? ""
          : explicada
            ? `<div class="selo ok">
                 A divergência está totalmente explicada pelos
                 ${inteiro.format(divergencia.quantidade_nao_baixados)}
                 Bolepix sem baixa no retorno.
               </div>`
            : `<div class="selo">
                 Os Bolepix não baixados somam
                 ${formatarMoeda(divergencia.total_nao_baixados)} e não
                 explicam toda a divergência. Veja a conferência da
                 Bolebarra com o DCB.
               </div>`),
    );
  }

  el("#divergencia").innerHTML = blocos
    .map((bloco) => `<div class="conta">${bloco}</div>`)
    .join("");
}

function tabela(colunas, linhas, vazio) {
  if (!linhas.length) {
    return `<p class="vazio">${escapar(vazio)}</p>`;
  }

  const cabecalho = colunas
    .map(
      (coluna) =>
        `<th class="${coluna.numero ? "numero" : ""}">${escapar(
          coluna.titulo,
        )}</th>`,
    )
    .join("");

  const corpo = linhas
    .map(
      (linha) =>
        `<tr>${colunas
          .map(
            (coluna) =>
              `<td class="${coluna.numero ? "numero" : ""}">${
                coluna.html ? linha[coluna.chave] : escapar(linha[coluna.chave])
              }</td>`,
          )
          .join("")}</tr>`,
    )
    .join("");

  return `
    <div class="rolagem">
      <table>
        <thead><tr>${cabecalho}</tr></thead>
        <tbody>${corpo}</tbody>
      </table>
    </div>`;
}

function carimbo(situacao) {
  const { texto, classe } = SITUACAO_GAI[situacao] || {
    texto: situacao,
    classe: "estudo",
  };

  return `<span class="carimbo carimbo-${classe}">${escapar(texto)}</span>`;
}

/* A tabela central: cobrança por cobrança, o que o GAI fez com
   ela. Traz também as que o GAI baixou por outro valor. */
function desenharBolebarraGai(conferencia) {
  const linhas = [
    ...(conferencia?.nao_baixados || []),
    ...(conferencia?.valor_divergente || []),
  ].map((item, indice) => ({
    numero: indice + 1,
    situacao: carimbo(item.situacao),
    nome: item.nome || "—",
    boleto: item.boleto || "—",
    alienacao: item.alienacao || "—",
    valor: formatarMoeda(item.valor),
    // Zero aqui é ausência de baixa, não um valor baixado.
    valor_gai: Number(item.valor_gai)
      ? formatarMoeda(item.valor_gai)
      : "—",
    motivo: item.motivo || "—",
  }));

  el('[data-painel="bolebarra-gai"]').innerHTML = tabela(
    [
      { titulo: "#", chave: "numero", numero: true },
      { titulo: "Situação", chave: "situacao", html: true },
      { titulo: "Sacado", chave: "nome" },
      { titulo: "Boleto", chave: "boleto" },
      { titulo: "Alienação", chave: "alienacao" },
      { titulo: "Valor recebido", chave: "valor", numero: true },
      { titulo: "Baixado no GAI", chave: "valor_gai", numero: true },
      { titulo: "Motivo informado pelo GAI", chave: "motivo" },
    ],
    linhas,
    conferencia
      ? "Toda a cobrança da francesinha Bolebarra foi baixada no GAI."
      : "Envie a francesinha Bolebarra e os PDFs do GAI para ver este bloco.",
  );
}

/* Só os desencontros: a cobrança sem baixa no retorno e a que
   liquidou por outro valor. O que confere fica no Excel. */
function desenharBolebarraDcb(conferencia) {
  const linhas = [];

  (conferencia?.valor_divergente || []).forEach((item) =>
    linhas.push({
      situacao: '<span class="carimbo carimbo-encerrado">Valor diferente</span>',
      nosso_numero: item.nosso_numero,
      nome: item.nome || "—",
      valor: formatarMoeda(item.valor),
      valor_dcb: formatarMoeda(item.valor_dcb),
      detalhe: `Diferença de ${formatarMoeda(item.diferenca)}`,
    }),
  );

  (conferencia?.sem_baixa || []).forEach((item) =>
    linhas.push({
      situacao: '<span class="carimbo carimbo-encerrado">Sem baixa</span>',
      nosso_numero: item.nosso_numero,
      nome: item.nome || "—",
      valor: formatarMoeda(item.valor),
      valor_dcb: "—",
      detalhe: item.no_arquivo
        ? `No arquivo, ocorrências ${item.ocorrencias}`
        : "Não está no arquivo de retorno",
    }),
  );

  el('[data-painel="bolebarra-dcb"]').innerHTML =
    tabela(
      [
        { titulo: "Situação", chave: "situacao", html: true },
        { titulo: "Nosso número", chave: "nosso_numero" },
        { titulo: "Sacado", chave: "nome" },
        { titulo: "Valor na francesinha", chave: "valor", numero: true },
        { titulo: "Valor pago no DCB", chave: "valor_dcb", numero: true },
        { titulo: "Observação", chave: "detalhe" },
      ],
      linhas,
      conferencia
        ? `As ${inteiro.format(
            conferencia.quantidade_casados,
          )} cobranças da francesinha Bolebarra estão liquidadas no arquivo de retorno, todas pelo valor certo.`
        : "Envie a francesinha Bolebarra e o DCB para ver este bloco.",
    ) +
    (conferencia
      ? `<p class="nota-tabela">
           Além delas, ${inteiro.format(
             conferencia.quantidade_fora_da_francesinha,
           )} título(s) liquidado(s) no retorno
           (${formatarMoeda(conferencia.total_fora_da_francesinha)})
           não estão na francesinha Bolebarra: são os recebimentos
           do Bolepix, que não têm francesinha de cobrança própria.
         </p>`
      : "");
}

function desenharNaoBaixados(bolepix) {
  const linhas = (bolepix?.nao_baixados || []).map((registro, indice) => ({
    numero: indice + 1,
    data: registro.data || "—",
    nome: registro.nome || "—",
    valor: formatarMoeda(registro.valor),
  }));

  el('[data-painel="nao-baixados"]').innerHTML = tabela(
    [
      { titulo: "#", chave: "numero", numero: true },
      { titulo: "Data/Hora do pagamento", chave: "data" },
      { titulo: "Cliente / Contraparte", chave: "nome" },
      { titulo: "Valor", chave: "valor", numero: true },
    ],
    linhas,
    bolepix
      ? "Todos os Bolepix recebidos têm baixa no arquivo de retorno."
      : "Envie a francesinha Bolepix e o DCB para ver este bloco.",
  );
}

/* O retrato do relatório de Baixas Não Efetivadas: por que o
   GAI recusou cada pagamento. */
function desenharMotivosGai(relatorio) {
  const linhas = (relatorio?.motivos || []).map((motivo) => ({
    motivo: motivo.motivo,
    quantidade: inteiro.format(motivo.quantidade),
    total: formatarMoeda(motivo.total),
  }));

  el('[data-painel="gai-motivos"]').innerHTML = tabela(
    [
      { titulo: "Motivo informado pelo GAI", chave: "motivo" },
      { titulo: "Registros", chave: "quantidade", numero: true },
      { titulo: "Total recusado", chave: "total", numero: true },
    ],
    linhas,
    relatorio
      ? "O relatório de Baixas Não Efetivadas do GAI veio sem registros."
      : "Envie o PDF de Baixas Não Efetivadas do GAI para ver este bloco.",
  );
}

function desenharPix(detalhado) {
  const linhas = (detalhado?.linhas || []).map((linha) => ({
    data: linha.data,
    valor: formatarMoeda(linha.valor),
    nome: linha.identificado
      ? escapar(linha.nome)
      : '<span class="pendente">não identificado</span>',
    documento: linha.documento || "—",
    pagamento: linha.data_pagamento || "—",
  }));

  el('[data-painel="pix"]').innerHTML = tabela(
    [
      { titulo: "Data do crédito", chave: "data" },
      { titulo: "Valor", chave: "valor", numero: true },
      { titulo: "Cliente / Contraparte", chave: "nome", html: true },
      { titulo: "CNPJ / CPF", chave: "documento" },
      { titulo: "Data/Hora do pagamento", chave: "pagamento" },
    ],
    linhas,
    "Envie o extrato e a francesinha Bolepix para identificar os pagadores.",
  );
}

function desenharCobranca(extrato) {
  const linhas = (extrato?.cobranca || []).map((registro) => ({
    data: registro.data,
    descricao: registro.descricao,
    doc: registro.doc,
    valor: formatarMoeda(registro.valor),
  }));

  el('[data-painel="cobranca"]').innerHTML = tabela(
    [
      { titulo: "Data", chave: "data" },
      { titulo: "Descrição", chave: "descricao" },
      { titulo: "DOC", chave: "doc" },
      { titulo: "Valor", chave: "valor", numero: true },
    ],
    linhas,
    "Envie o extrato para ver os créditos de cobrança.",
  );
}

function desenharOcorrencias(dcb) {
  const linhas = (dcb?.ocorrencias || []).map((ocorrencia) => ({
    codigo: ocorrencia.codigo,
    descricao: ocorrencia.descricao,
    quantidade: inteiro.format(ocorrencia.quantidade),
  }));

  el('[data-painel="ocorrencias"]').innerHTML = tabela(
    [
      { titulo: "Código", chave: "codigo" },
      { titulo: "Descrição", chave: "descricao" },
      { titulo: "Títulos", chave: "quantidade", numero: true },
    ],
    linhas,
    "Envie o DCB para ver as ocorrências do arquivo de retorno.",
  );
}

function escapar(valor) {
  return String(valor ?? "").replace(
    /[&<>"']/g,
    (caractere) =>
      ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#39;",
      })[caractere],
  );
}

/* ============================================================
 * Abas
 * ============================================================ */

function registrarAbas() {
  todos(".aba").forEach((aba) => {
    aba.addEventListener("click", () => {
      todos(".aba").forEach((outra) =>
        outra.classList.toggle("ativa", outra === aba),
      );

      todos(".painel").forEach((painel) => {
        painel.hidden = painel.dataset.painel !== aba.dataset.aba;
      });
    });
  });
}

/* ============================================================
 * Tema claro / escuro
 * ============================================================ */

const CHAVE_TEMA = "conciliacao-tema";

function lerTemaSalvo() {
  try {
    const salvo = localStorage.getItem(CHAVE_TEMA);
    return salvo === "dark" || salvo === "light" ? salvo : null;
  } catch {
    // Janela anônima ou armazenamento bloqueado: segue no
    // tema do sistema, sem quebrar a página.
    return null;
  }
}

function temaEmUso() {
  const explicito = document.documentElement.dataset.theme;

  if (explicito === "dark" || explicito === "light") {
    return explicito;
  }

  return window.matchMedia("(prefers-color-scheme: dark)").matches
    ? "dark"
    : "light";
}

function aplicarTema(tema) {
  document.documentElement.dataset.theme = tema;

  el("#tema-texto").textContent =
    tema === "dark" ? "Tema claro" : "Tema escuro";

  el("#botao-tema").setAttribute(
    "aria-label",
    tema === "dark" ? "Mudar para o tema claro" : "Mudar para o tema escuro",
  );

  try {
    localStorage.setItem(CHAVE_TEMA, tema);
  } catch {
    /* sem persistência, o tema vale só para esta visita */
  }
}

function registrarTema() {
  const salvo = lerTemaSalvo();

  if (salvo) {
    document.documentElement.dataset.theme = salvo;
  }

  // Sincroniza o rótulo com o tema realmente em vigor.
  el("#tema-texto").textContent =
    temaEmUso() === "dark" ? "Tema claro" : "Tema escuro";

  el("#botao-tema").addEventListener("click", () => {
    aplicarTema(temaEmUso() === "dark" ? "light" : "dark");
  });

  // Enquanto o usuário não escolher, acompanha o sistema.
  window
    .matchMedia("(prefers-color-scheme: dark)")
    .addEventListener("change", () => {
      if (!lerTemaSalvo()) {
        el("#tema-texto").textContent =
          temaEmUso() === "dark" ? "Tema claro" : "Tema escuro";
      }
    });
}

/* ============================================================
 * Revelar ao rolar
 * ============================================================ */

function registrarRevelacao() {
  const alvos = todos(".revelar");

  if (!("IntersectionObserver" in window)) {
    alvos.forEach((alvo) => alvo.classList.add("visivel"));
    return;
  }

  const observador = new IntersectionObserver(
    (entradas) => {
      entradas.forEach((entrada) => {
        if (entrada.isIntersecting) {
          entrada.target.classList.add("visivel");
          observador.unobserve(entrada.target);
        }
      });
    },
    { threshold: 0.08, rootMargin: "0px 0px -40px 0px" },
  );

  alvos.forEach((alvo) => observador.observe(alvo));
}

/* ============================================================
 * Início
 * ============================================================ */

todos(".caixa-arquivo").forEach(registrarCaixa);
registrarAbas();
registrarRevelacao();
registrarTema();

el("#botao-conciliar").addEventListener("click", conciliar);
el("#botao-limpar").addEventListener("click", limpar);
el("#botao-exportar").addEventListener("click", exportar);

iniciarPython();
