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
  "sistemas.py",
  "benner.py",
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
  "gir",
  "ggr",
  "gop",
  "benner",
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

  desenharBolebarraGai(resultado.bolebarras_sistemas);
  desenharBolebarraDcb(resultado.bolebarras_dcb);
  desenharNaoBaixados(resultado.bolepix_dcb);
  desenharMotivosGai(resultado.gai_nao_baixadas);
  desenharSistemas(resultado);
  desenharBenner(resultado);
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

  for (const relatorio of resultado.sistemas || []) {
    cartoes.push(indicador({
      titulo: `Pagamentos no ${relatorio.sistema}`,
      valor: formatarMoeda(relatorio.total),
      detalhe: `${inteiro.format(relatorio.quantidade)} pagamentos com valor positivo`,
    }));
  }

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

  const sistemas = resultado.bolebarras_sistemas;
  if (sistemas) {
    cartoes.push(indicador({
      titulo: "Bolebarra sem baixa nos sistemas",
      valor: formatarMoeda(sistemas.total_ausente),
      detalhe: `${inteiro.format(sistemas.quantidade_ausentes)} sem baixa localizada; ${inteiro.format(sistemas.divergentes.length)} com valor divergente`,
      classe: sistemas.quantidade_pendentes ? "alerta" : "ok",
    }));
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
  const sistemas = resultado.bolebarras_sistemas;
  const divergencia = resultado.divergencia;
  if (!sistemas && !divergencia) {
    cartao.hidden = true;
    return;
  }
  cartao.hidden = false;
  const blocos = [];
  if (sistemas) {
    const contas = [
      linhaConta("Francesinha Bolebarra", formatarMoeda(sistemas.total_francesinha)),
      linhaConta(`Com baixa localizada nos sistemas (${inteiro.format(sistemas.quantidade_localizados)} cobranças)`,
        `&minus; ${formatarMoeda(sistemas.total_localizado)}`),
      linhaConta("Bolebarra sem baixa nos sistemas", formatarMoeda(sistemas.total_ausente), "total"),
    ];
    const notas = `<p>Relatórios considerados: ${escapar(sistemas.sistemas_enviados.join(", "))}.
      O valor com baixa localizada usa o valor da Bolebarra, contado uma vez por cobrança.
      Diferenças de valor estão discriminadas abaixo.</p>`;
    const divergentes = sistemas.divergentes.map((r) => ({...r,
      valor: formatarMoeda(r.valor), valor_sistemas: formatarMoeda(r.valor_sistemas),
      diferenca: formatarMoeda(Number(r.valor) - Number(r.valor_sistemas)),
    }));
    let detalhes = "";
    if (divergentes.length) {
      detalhes += '<h3>Diferenças de valor por sistema</h3>' + tabela([
        {titulo: "Sistema", chave: "sistema"},
        {titulo: "Boleto", chave: "boleto"},
        {titulo: "Sacado", chave: "nome"},
        {titulo: "Data pagamento", chave: "data_pagamento"},
        {titulo: "Bolebarra", chave: "valor", numero: true},
        {titulo: "Valor no sistema", chave: "valor_sistemas", numero: true},
        {titulo: "Diferença (Bolebarra - sistema)", chave: "diferenca", numero: true},
      ], divergentes, "");
    }
    if (sistemas.ambiguos.length) {
      detalhes += '<h3>Identificações que exigem revisão</h3>' + tabela([
        {titulo: "Boleto", chave: "boleto"}, {titulo: "Sacado", chave: "nome"},
        {titulo: "Sistemas e valores informados", chave: "valores"},
        {titulo: "Situação", chave: "situacao"},
      ], sistemas.ambiguos.map((r) => ({...r, valores: Object.entries(r.valores_por_sistema || {})
        .map(([s, v]) => `${s}: ${formatarMoeda(v)}`).join("; ")})), "");
    }
    if (!sistemas.quantidade_pendentes) {
      detalhes += '<div class="selo ok">Toda a Bolebarra confere com os sistemas enviados.</div>';
    }
    blocos.push(contas.join("") + notas + detalhes);
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

/* Cobranças sem baixa localizada no conjunto de relatórios enviados. */
function desenharBolebarraGai(conferencia) {
  const ausentes = conferencia?.ausentes || [];
  const consultados = conferencia?.sistemas_enviados || [];
  const naoEnviados = ["GAI", "GIR", "GGR", "GOP"].filter((s) => !consultados.includes(s));
  el('[data-painel="bolebarra-gai"]').innerHTML =
    (conferencia ? `<p class="nota-tabela">${inteiro.format(conferencia.quantidade_ausentes)} cobrança(s) sem baixa localizada,
      total de ${formatarMoeda(conferencia.total_ausente)}.
      ${naoEnviados.length ? `Relatórios não enviados: ${escapar(naoEnviados.join(", "))}.` : "Todos os quatro sistemas foram consultados."}
      Uma baixa em qualquer sistema enviado retira a cobrança desta lista.</p>` : "") + tabela(
    [
      { titulo: "Boleto", chave: "boleto" },
      { titulo: "Sacado", chave: "nome" },
      { titulo: "Data pagamento", chave: "data_pagamento" },
      { titulo: "Valor recebido", chave: "valor", numero: true },
      { titulo: "Baixa não localizada em", chave: "nao_localizada_em" },
      { titulo: "Situação", chave: "situacao" },
    ],
    ausentes.map((r) => ({...r, valor: formatarMoeda(r.valor),
      nao_localizada_em: r.situacao === "Sem nosso número válido"
        ? "Identificação insuficiente para consultar" : consultados.join(", "),
    })),
    conferencia
      ? "Nenhuma cobrança sem baixa nos sistemas enviados."
      : "Envie a Bolebarra e os relatórios de pagamentos para conferir.",
  );
}

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

/* O relatório de Baixas Não Efetivadas, recusa por recusa: o
   motivo, a alienação e — quando a francesinha Bolebarra sabe
   dizer — o nome de quem pagou. O relatório do GAI não traz
   nome; ele vem da francesinha, pelo boleto ou pela alienação. */
function desenharMotivosGai(relatorio) {
  const recusas = (relatorio?.registros || []).filter(
    (registro) => Number(registro.total_pago) > 0,
  );
  const porMotivo = new Map();
  for (const registro of recusas) {
    const motivo = registro.motivo || "—";
    const grupo = porMotivo.get(motivo) || { motivo, quantidade: 0, centavos: 0 };
    grupo.quantidade += 1;
    grupo.centavos += Math.round(Number(registro.total_pago) * 100);
    porMotivo.set(motivo, grupo);
  }
  const resumo = Array.from(porMotivo.values()).map((motivo) => ({
    motivo: motivo.motivo,
    quantidade: inteiro.format(motivo.quantidade),
    total: formatarMoeda(motivo.centavos / 100),
  }));

  const linhas = recusas.map((registro) => ({
    motivo: registro.motivo || "—",
    alienacao: registro.alienacao || "—",
    boleto: registro.boleto || "—",
    parcela: registro.parcela || "—",
    nome: registro.nome
      ? escapar(registro.nome)
      : '<span class="pendente">não identificado</span>',
    data: registro.data_pagamento || "—",
    total: formatarMoeda(registro.total_pago),
  }));

  el('[data-painel="gai-motivos"]').innerHTML =
    tabela(
      [
        { titulo: "Motivo informado pelo GAI", chave: "motivo" },
        { titulo: "Recusas com valor", chave: "quantidade", numero: true },
        { titulo: "Total recusado", chave: "total", numero: true },
      ],
      resumo,
      relatorio
        ? "Nenhuma recusa com valor positivo neste relatório."
        : "Envie o PDF de Baixas Não Efetivadas do GAI para ver este bloco.",
    ) +
    (linhas.length
      ? `<p class="nota-tabela">
           Exibindo ${inteiro.format(linhas.length)} recusa(s) com valor
           positivo. No relatório completo,
           ${inteiro.format(relatorio.quantidade_identificados)} de
           ${inteiro.format(relatorio.quantidade)} registros têm sacado
           identificado na francesinha Bolebarra. A ausência de nome
           indica que não foi possível identificar o sacado com os
           arquivos enviados. Todos os registros estão no Excel.
         </p>` +
        tabela(
          [
            { titulo: "Alienação", chave: "alienacao" },
            { titulo: "Sacado", chave: "nome", html: true },
            { titulo: "Boleto", chave: "boleto" },
            { titulo: "Parcela", chave: "parcela", numero: true },
            { titulo: "Data pagamento", chave: "data" },
            { titulo: "Valor recusado", chave: "total", numero: true },
            { titulo: "Motivo", chave: "motivo" },
          ],
          linhas,
          "",
        )
      : relatorio
        ? '<p class="nota-tabela">Nenhuma recusa com valor positivo neste relatório. Todos os registros estão no Excel.</p>'
        : "");
}

function desenharBenner(resultado) {
  const painel = el('[data-painel="gai-benner"]');
  const conferencia = resultado.gai_benner;
  if (!conferencia || conferencia.status === "aguardando_gai") {
    painel.innerHTML = '<p class="vazio">Envie a integração BENNER e o PDF de Parcelas Lidas do GAI do mesmo período para comparar.</p>';
    return;
  }
  if (conferencia.status === "datas_incompativeis") {
    painel.innerHTML = '<p class="vazio">GAI e BENNER não têm datas de pagamento em comum. Envie arquivos do mesmo período.</p>';
    return;
  }
  const colunas = [
    { titulo: "Alienação", chave: "alienacao" },
    { titulo: "Imóvel", chave: "imovel" },
    { titulo: "Parcela", chave: "parcela" },
    { titulo: "Boleto GAI", chave: "boleto" },
    { titulo: "Data pagamento", chave: "data_pagamento" },
    { titulo: "Valor GAI", chave: "valor", numero: true },
  ];
  const linhas = (itens) => itens.map((r) => ({...r, valor: formatarMoeda(r.valor),
    valor_benner: r.valor_benner == null ? "—" : formatarMoeda(r.valor_benner)}));
  painel.innerHTML = `<p class="nota-tabela">${inteiro.format(conferencia.quantidade_ausentes)} parcela(s)
    não localizada(s) no BENNER, total de ${formatarMoeda(conferencia.total_ausente)}.
    ${inteiro.format(conferencia.quantidade_encontrados)} parcela(s) conferida(s).
    A identificação usa alienação, imóvel, parcela e data de pagamento.
    Somente parcelas com valor positivo são comparadas.</p>` +
    tabela(colunas, linhas(conferencia.ausentes), "Nenhuma parcela ausente entre as que puderam ser comparadas.") +
    (conferencia.divergentes.length ? '<h3>Encontradas com valor diferente</h3>' +
      tabela([...colunas, {titulo: "Valor BENNER", chave: "valor_benner", numero: true},
        {titulo: "Linha BENNER", chave: "linha_benner"}], linhas(conferencia.divergentes), "") : "") +
    (conferencia.nao_comparados.length ? '<h3>Parcelas não comparadas</h3>' +
      tabela(colunas, linhas(conferencia.nao_comparados), "") : "") +
    (resultado.benner?.quantidade_nao_identificados ?
      `<p class="nota-tabela">${inteiro.format(resultado.benner.quantidade_nao_identificados)} crédito(s)
      do BENNER sem identificação completa. Revise a aba BENNER revisar do Excel antes de concluir sobre as ausências.</p>` : "");
}

function desenharSistemas(resultado) {
  const relatorios = resultado.sistemas || [];
  const linhas = relatorios.flatMap((relatorio) => relatorio.linhas.map((r) => ({
    ...r,
    nome: r.nome || "Não informado no PDF",
    valor: formatarMoeda(r.valor),
    valor_dcb: r.valor_dcb == null ? "—" : formatarMoeda(r.valor_dcb),
    valor_bolebarra: r.valor_bolebarra == null ? "—" : formatarMoeda(r.valor_bolebarra),
  })));
  el('[data-painel="sistemas"]').innerHTML =
    '<p class="nota-tabela">Pagamentos positivos das tabelas de GIR, GGR e GOP. ' +
    'O cruzamento usa boleto e data de pagamento. Não encontrado na Bolebarra ' +
    'não significa ausência de pagamento: o recebimento pode ser Bolepix.</p>' + tabela([
      { titulo: "Sistema", chave: "sistema" },
      { titulo: "Documento / Processo", chave: "documento" },
      { titulo: "Contrato", chave: "contrato" },
      { titulo: "Boleto", chave: "boleto" },
      { titulo: "Cliente / Ocupante", chave: "nome" },
      { titulo: "Data pagamento", chave: "data_pagamento" },
      { titulo: "Valor pago", chave: "valor", numero: true },
      { titulo: "DCB", chave: "situacao_dcb" },
      { titulo: "Valor DCB", chave: "valor_dcb", numero: true },
      { titulo: "Bolebarra", chave: "situacao_bolebarra" },
      { titulo: "Valor Bolebarra", chave: "valor_bolebarra", numero: true },
    ], linhas, relatorios.length ? "Nenhum pagamento com valor positivo." : "Envie um PDF do GIR, GGR ou GOP.");


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
