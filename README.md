# Clientdash — Dashboard de Clientes

Dashboard web para importar, analisar e exportar uma tabela de clientes com as colunas **Id**, **nome**, **produto**, **valor** e **data**. A aplicação foi criada em português do Brasil e processa os dados localmente no navegador.

## Recursos

O dashboard oferece uma visão analítica dos clientes importados, com:

- indicadores de quantidade de clientes, valor total, preço médio e intervalo de datas;
- estatísticas de preço com mínimo, máximo e mediana estimada;
- gráfico de evolução dos valores por data;
- gráfico de valor acumulado por produto;
- distribuição temporal dos registros;
- tabela responsiva com busca por nome ou produto;
- filtros por valor mínimo, valor máximo, data inicial e data final;
- importação de arquivos `.xlsx`, `.xls` e `.csv`;
- exportação de dados em CSV e de relatórios em PDF e Word;
- amostra inicial com 15 clientes para demonstração;
- botão para restaurar a amostra original.

## Requisitos

- Node.js compatível com Vite;
- pnpm 11.25.0 ou versão compatível;
- navegador moderno com suporte a JavaScript, Canvas e downloads locais.

## Instalação e execução

Na raiz do projeto, instale as dependências:

```bash
pnpm install
```

Inicie o servidor de desenvolvimento:

```bash
pnpm dev
```

O dashboard ficará disponível em `http://localhost:3000`.

Para gerar a versão de produção:

```bash
pnpm build
```

Para visualizar localmente o build gerado:

```bash
pnpm preview
```

## Formato da planilha

A primeira linha do arquivo deve conter os cabeçalhos abaixo. A aplicação também reconhece variações comuns de maiúsculas, acentos e separadores:

| Coluna | Descrição | Exemplos |
| --- | --- | --- |
| `Id` | Identificador do cliente ou registro | `01`, `102` |
| `nome` | Nome do cliente | `Ana Souza` |
| `produto` | Produto associado ao registro | `Notebook` |
| `valor` | Valor monetário do registro | `3499,90`, `R$ 3.499,90`, `3499.90` |
| `data` | Data do registro | `01/10/2026`, `2026-10-01` |

O valor é convertido para número e a data é normalizada internamente para o formato ISO. Datas armazenadas como números seriais do Excel também são tratadas durante a importação.

## Como usar

1. Abra o dashboard.
2. Clique em **Importar Excel** ou arraste um arquivo para a área de importação.
3. Confira os indicadores e os gráficos atualizados.
4. Use o campo de busca para localizar um nome ou produto.
5. Abra **Filtros** para definir limites de valor e intervalo de datas.
6. Abra **Exportar** para gerar um arquivo CSV, PDF ou Word.
7. Use **Restaurar amostra** para retornar aos 15 registros de demonstração.

Os filtros afetam os indicadores, os gráficos, a tabela e os relatórios exportados. A importação substitui o conjunto de dados atualmente carregado.

## Exportações

- **CSV:** exporta os registros atualmente filtrados em formato tabular.
- **PDF:** gera um relatório com título, resumo dos indicadores e tabela de registros.
- **Word:** gera um documento editável com resumo e tabela dos registros.

Os arquivos são gerados no navegador e baixados para o dispositivo do usuário. Nenhum dado é enviado para um servidor pela aplicação frontend.

## Estrutura principal

```text
clientdash/
├── index.html              # Documento HTML de entrada
├── src/
│   ├── main.js              # Interface, importação, filtros, gráficos e exportações
│   └── styles.css           # Layout responsivo e identidade visual
├── public/
│   └── manus-routes.json    # Rotas reconhecidas pelo ambiente Webdev
├── app.config.ts            # Metadata do projeto
├── package.json             # Scripts e dependências
├── vite.config.js           # Configuração do servidor Vite
└── README.md                # Esta documentação
```

## Dependências principais

- [Vite](https://vitejs.dev/) para desenvolvimento e build frontend;
- [SheetJS / xlsx](https://sheetjs.com/) para leitura e geração de planilhas;
- [Chart.js](https://www.chartjs.org/) para gráficos;
- [jsPDF](https://github.com/parallax/jsPDF) e `jspdf-autotable` para PDF;
- [docx](https://docx.js.org/) para relatórios Word.

## Privacidade e limitações

O processamento ocorre localmente no navegador. A aplicação não implementa autenticação, banco de dados, sincronização entre usuários ou armazenamento remoto dos arquivos importados. Ao atualizar ou fechar a página, os dados importados podem ser perdidos; use uma das opções de exportação para guardar uma cópia.

A análise depende da qualidade dos dados fornecidos. Registros sem valor ou data válidos podem ser ignorados na importação ou não aparecer corretamente nos indicadores correspondentes.
