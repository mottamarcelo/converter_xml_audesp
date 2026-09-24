import { useState } from "react";
import {
  converterXmlParaExcel
} from "./services/xmlToExcel.js";

function App() {
  const [arquivo, setArquivo] = useState(null);
  const [arrastando, setArrastando] =
    useState(false);

  const [processando, setProcessando] =
    useState(false);

  const [resultado, setResultado] =
    useState(null);

  const [erro, setErro] = useState("");

  function selecionarArquivo(file) {
    if (!file) {
      return;
    }

    setErro("");
    setResultado(null);
    setArquivo(file);
  }

  function handleInputChange(event) {
    const file =
      event.target.files?.[0];

    selecionarArquivo(file);
  }

  function handleDragOver(event) {
    event.preventDefault();

    setArrastando(true);
  }

  function handleDragLeave(event) {
    event.preventDefault();

    setArrastando(false);
  }

  function handleDrop(event) {
    event.preventDefault();

    setArrastando(false);

    const file =
      event.dataTransfer.files?.[0];

    selecionarArquivo(file);
  }

  async function converter() {
    if (!arquivo) {
      setErro(
        "Selecione um arquivo XML antes de converter."
      );

      return;
    }

    try {
      setProcessando(true);
      setErro("");
      setResultado(null);

      const resultado =
        await converterXmlParaExcel(
          arquivo
        );

      setResultado(resultado);
    } catch (error) {
      console.error(error);

      setErro(
        error.message ||
          "Ocorreu um erro durante a conversão."
      );
    } finally {
      setProcessando(false);
    }
  }

  function limpar() {
    setArquivo(null);
    setResultado(null);
    setErro("");
  }

  return (
    <div className="app">
      <main className="container">
        <header className="header">
          <div className="logo">
            XML
          </div>

          <div>
            <h1>
              Conversor XML → Excel
            </h1>

            <p>
              Conciliações Bancárias AUDESP
            </p>
          </div>
        </header>

        <section className="card">
          <h2>
            Selecione o arquivo XML
          </h2>

          <p className="description">
            Selecione um arquivo de
            Conciliações Bancárias AUDESP
            para gerar automaticamente um
            arquivo Excel.
          </p>

          <label
            className={`drop-zone ${
              arrastando
                ? "dragging"
                : ""
            }`}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
          >
            <input
              type="file"
              accept=".xml,text/xml"
              onChange={
                handleInputChange
              }
            />

            <div className="upload-icon">
              ↑
            </div>

            <strong>
              Arraste o XML aqui
            </strong>

            <span>
              ou clique para selecionar
            </span>
          </label>

          {arquivo && (
            <div className="file-selected">
              <div>
                <span className="file-icon">
                  XML
                </span>

                <div>
                  <strong>
                    {arquivo.name}
                  </strong>

                  <small>
                    {formatarTamanho(
                      arquivo.size
                    )}
                  </small>
                </div>
              </div>

              <button
                className="button-secondary"
                onClick={limpar}
                disabled={processando}
              >
                Remover
              </button>
            </div>
          )}

          {erro && (
            <div className="error">
              <strong>
                Erro
              </strong>

              <span>
                {erro}
              </span>
            </div>
          )}

          <button
            className="button-primary"
            onClick={converter}
            disabled={
              !arquivo ||
              processando
            }
          >
            {processando
              ? "Convertendo..."
              : "Converter para Excel"}
          </button>

          {resultado && (
            <div className="success">
              <div className="success-title">
                ✓ Arquivo convertido
                com sucesso!
              </div>

              <div className="success-info">
                <div>
                  <strong>
                    Arquivo
                  </strong>

                  <span>
                    {resultado.nomeArquivo}
                  </span>
                </div>

                <div>
                  <strong>
                    Registros
                  </strong>

                  <span>
                    {
                      resultado.quantidadeRegistros
                    }
                  </span>
                </div>

                <div>
                  <strong>
                    Colunas
                  </strong>

                  <span>
                    {
                      resultado.quantidadeColunas
                    }
                  </span>
                </div>
              </div>

              <p>
                O download do arquivo
                Excel foi iniciado.
              </p>
            </div>
          )}
        </section>

        <footer>
          <p>
            Conversão realizada
            diretamente no navegador.
          </p>

          <p>
            O arquivo XML não é enviado
            para nenhum servidor.
          </p>
        </footer>
      </main>
    </div>
  );
}

function formatarTamanho(bytes) {
  if (bytes < 1024) {
    return `${bytes} B`;
  }

  if (bytes < 1024 * 1024) {
    return `${(
      bytes / 1024
    ).toFixed(1)} KB`;
  }

  return `${(
    bytes /
    (1024 * 1024)
  ).toFixed(1)} MB`;
}

export default App;