import * as XLSX from "xlsx";

const NAMESPACE_CB =
  "http://www.tce.sp.gov.br/audesp/xml/conciliacoes";

const NAMESPACE_GEN =
  "http://www.tce.sp.gov.br/audesp/xml/generico";

/**
 * Retorna o nome local de uma tag XML.
 *
 * Exemplo:
 *
 * {http://www.tce.sp.gov.br/audesp/xml/generico}Banco
 *
 * retorna:
 *
 * Banco
 */
function getLocalName(element) {
  return element.localName || element.tagName.split(":").pop();
}

/**
 * Retorna o texto de um elemento.
 */
function getElementText(element) {
  if (!element) {
    return "";
  }

  return (element.textContent || "").trim();
}

/**
 * Lê os campos do Descritor.
 */
function lerDescritor(xmlDoc) {
  const descritores = xmlDoc.getElementsByTagNameNS(
    NAMESPACE_CB,
    "Descritor"
  );

  if (descritores.length === 0) {
    throw new Error(
      "A tag <cb:Descritor> não foi encontrada no XML."
    );
  }

  const descritor = descritores[0];

  const dados = {};

  Array.from(descritor.children).forEach((elemento) => {
    const nome = getLocalName(elemento);

    dados[nome] = getElementText(elemento);
  });

  return dados;
}

/**
 * Lê todos os DomicilioBancario existentes
 * dentro das Conciliacoes.
 */
function lerDomiciliosBancarios(xmlDoc) {
  const conciliacoes = xmlDoc.getElementsByTagNameNS(
    NAMESPACE_CB,
    "Conciliacao"
  );

  if (conciliacoes.length === 0) {
    throw new Error(
      "Nenhuma tag <cb:Conciliacao> foi encontrada no XML."
    );
  }

  const domicilios = [];

  Array.from(conciliacoes).forEach((conciliacao) => {
    const domiciliosDaConciliacao =
      conciliacao.getElementsByTagNameNS(
        NAMESPACE_CB,
        "DomicilioBancario"
      );

    Array.from(domiciliosDaConciliacao).forEach(
      (domicilio) => {
        const dados = {};

        Array.from(domicilio.children).forEach(
          (elemento) => {
            const nome = getLocalName(elemento);

            dados[nome] = getElementText(elemento);
          }
        );

        domicilios.push(dados);
      }
    );
  });

  if (domicilios.length === 0) {
    throw new Error(
      "Nenhuma tag <cb:DomicilioBancario> foi encontrada."
    );
  }

  return domicilios;
}

/**
 * Monta as colunas do Excel.
 *
 * Primeiro entram os campos do Descritor.
 * Depois os campos do DomicilioBancario.
 */
function montarColunas(descritor, domicilios) {
  const colunas = [];

  Object.keys(descritor).forEach((campo) => {
    if (!colunas.includes(campo)) {
      colunas.push(campo);
    }
  });

  domicilios.forEach((domicilio) => {
    Object.keys(domicilio).forEach((campo) => {
      if (!colunas.includes(campo)) {
        colunas.push(campo);
      }
    });
  });

  return colunas;
}

/**
 * Converte o XML em dados tabulares.
 */
function montarDados(descritor, domicilios) {
  const colunas = montarColunas(
    descritor,
    domicilios
  );

  const dados = [];

  // Cabeçalho
  dados.push(colunas);

  // Linhas
  domicilios.forEach((domicilio) => {
    const linha = colunas.map((coluna) => {
      if (
        Object.prototype.hasOwnProperty.call(
          descritor,
          coluna
        )
      ) {
        return descritor[coluna];
      }

      return domicilio[coluna] ?? "";
    });

    dados.push(linha);
  });

  return {
    colunas,
    dados
  };
}

/**
 * Cria e baixa o arquivo XLSX.
 */
function gerarExcel(dados, nomeArquivo) {
  const worksheet = XLSX.utils.aoa_to_sheet(
    dados
  );

  const workbook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    "DomiciliosBancarios"
  );

  // Congelar cabeçalho
  worksheet["!freeze"] = {
    xSplit: 0,
    ySplit: 1
  };

  // Filtro automático
  const range = XLSX.utils.decode_range(
    worksheet["!ref"]
  );

  worksheet["!autofilter"] = {
    ref: XLSX.utils.encode_range(range)
  };

  // Ajuste automático das colunas
  const larguras = dados[0].map(
    (_, colunaIndex) => {
      let maior = 0;

      dados.forEach((linha) => {
        const valor = linha[colunaIndex];

        if (valor !== undefined && valor !== null) {
          maior = Math.max(
            maior,
            String(valor).length
          );
        }
      });

      return {
        wch: Math.min(
          Math.max(maior + 2, 10),
          50
        )
      };
    }
  );

  worksheet["!cols"] = larguras;

  // Gera o arquivo e inicia o download
  XLSX.writeFile(
    workbook,
    nomeArquivo,
    {
      bookType: "xlsx"
    }
  );
}

/**
 * Função principal.
 *
 * Recebe um File selecionado pelo usuário.
 */
export async function converterXmlParaExcel(file) {
  if (!file) {
    throw new Error(
      "Nenhum arquivo foi selecionado."
    );
  }

  if (
    !file.name
      .toLowerCase()
      .endsWith(".xml")
  ) {
    throw new Error(
      "Selecione um arquivo com extensão .xml."
    );
  }

  const conteudo = await file.text();

  const parser = new DOMParser();

  const xmlDoc = parser.parseFromString(
    conteudo,
    "application/xml"
  );

  // Detecta erro de parsing
  const parserError =
    xmlDoc.getElementsByTagName(
      "parsererror"
    );

  if (parserError.length > 0) {
    throw new Error(
      "O arquivo XML possui uma estrutura inválida."
    );
  }

  // Verifica se é um XML AUDESP de Conciliação
  const conciliacoes =
    xmlDoc.getElementsByTagNameNS(
      NAMESPACE_CB,
      "ConciliacoesBancarias"
    );

  if (conciliacoes.length === 0) {
    throw new Error(
      "O XML não parece ser um arquivo de Conciliações Bancárias AUDESP."
    );
  }

  const descritor = lerDescritor(xmlDoc);

  const domicilios =
    lerDomiciliosBancarios(xmlDoc);

  const resultado = montarDados(
    descritor,
    domicilios
  );

  /*
   * Remove a extensão original.
   *
   * Exemplo:
   *
   * 201_01_Conc_Bancaria.xml
   *
   * vira:
   *
   * 201_01_Conc_Bancaria.xlsx
   */
  const nomeArquivo =
    file.name.replace(
      /\.xml$/i,
      ".xlsx"
    );

  gerarExcel(
    resultado.dados,
    nomeArquivo
  );

  return {
    nomeArquivo,
    quantidadeRegistros:
      domicilios.length,
    quantidadeColunas:
      resultado.colunas.length,
    colunas: resultado.colunas
  };
}