# Reference Equity DocSync

Reference Equity DocSync e uma ferramenta portavel para auditar paridade entre documentacao tecnica e referencias de implementacao. Ela compara um modelo de referencia usado como fonte de verdade com uma documentacao publicada e gera evidencias de divergencia: campos ausentes, datasets nao documentados, candidatos selecionados e rejeitados, linhagem de origem e dashboards prontos para revisao.

O caso de uso original nasceu em governanca de documentacao de modelos de dados. Para o portfolio, a versao publica usa um exemplo seguro baseado na tabela oficial `ZX_LINES`, documentada publicamente pela Oracle, sem expor dados internos.

## O que o projeto resolve

Documentacoes e referencias de implementacao costumam sair de sincronia. Este projeto torna essa diferenca visivel respondendo a tres perguntas:

- Quais campos existem no modelo tecnico, mas ainda nao aparecem na documentacao publicada?
- Qual arquivo de referencia foi selecionado quando ha multiplas versoes, backups, testes ou copias?
- Qual tabela e coluna de origem explicam cada campo ausente?

## Funcionalidades

- Le pacotes `.xdmz` ou pastas com multiplos arquivos `.xdmz`.
- Extrai texto de documentos `.docx` diretamente do XML interno do Word.
- Compara documentacao e referencia usando escopo por ancoras ou correspondencia por nome.
- Gera saidas em JSON, Markdown, CSV de diagnostico e dashboard HTML.
- Roda sem servicos externos por padrao.
- Permite enriquecimento por arquivos de mapeamento quando existe um catalogo de referencia disponivel.

## Estrutura do repositorio

```text
validator_py/         Motor Python de comparacao e pos-processamento
scripts/              Entradas PowerShell portaveis
samples/              Modelo e documentacao de demonstracao com referencia publica ZX_LINES
assets/               Exemplos publicos de mapeamento para documentacao oficial
tests/                Testes de linhagem SQL e leitura de documentos
docs/                 Proposta e posicionamento para portfolio
```

## Como executar

```powershell
python -m unittest discover -s tests

powershell -ExecutionPolicy Bypass -File .\scripts\run_reference_equity.ps1 -IncludePass
```

As saidas sao gravadas em `output/demo`:

- `comparison_report.json`
- `comparison_report.md`
- `comparison_dashboard.html`
- `by_extractor/*.csv`

## Demonstracao com ZX_LINES

A demo usa uma referencia publica real da Oracle Financials: `ZX_LINES`, tabela que armazena linhas detalhadas de impostos de transacoes. O documento de exemplo deixa alguns campos fora da documentacao publicada, como `TAX_RATE_CODE` e `TAX_AMT_FUNCL_CURR`, para demonstrar como o dashboard identifica lacunas e aponta para a documentacao oficial de enriquecimento.

Referencia oficial usada na demo: https://docs.oracle.com/en/cloud/saas/financials/25c/oedmf/zxlines-13072.html

## Uso direto com Python

```powershell
python .\validator_py\docsync_pipeline.py `
  --doc .\samples\docs\reference_catalog.docx `
  --xdmz-dir .\samples\models `
  --out .\output\demo `
  --include-pass-in-dashboard `
  --skip-web-post
```

## Leitura para portfolio

A ideia central e equidade entre documentacao e referencia: documentacao e implementacao devem ter autoridade equilibrada. Quando um lado muda, o outro deve ser auditado com evidencia rastreavel, nao apenas com revisao manual por amostragem.

Este projeto demonstra automacao em Python, leitura de XML/ZIP, heuristicas de linhagem SQL, geracao de dashboard e uma abordagem pragmatica para governanca de qualidade documental.