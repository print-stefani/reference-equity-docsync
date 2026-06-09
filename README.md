# Reference Equity DocSync

Reference Equity DocSync é uma ferramenta portável para auditar paridade entre documentação e referências técnicas. Ela compara um modelo de referência usado como fonte de verdade com uma documentação publicada e gera evidências de divergência: campos ausentes, datasets não documentados, candidatos selecionados e rejeitados, linhagem de origem e dashboards prontos para revisão.

O caso de uso original nasceu em governança de documentação de modelos de dados, mas esta versão de portfólio usa exemplos sintéticos e linguagem genérica para poder ser compartilhada publicamente.

## O que o projeto resolve

Documentações e referências de implementação costumam sair de sincronia. Este projeto torna essa diferença visível respondendo a três perguntas:

- Quais campos existem no modelo técnico, mas ainda não aparecem na documentação publicada?
- Qual arquivo de referência foi selecionado quando há múltiplas versões, backups, testes ou cópias?
- Qual tabela e coluna de origem explicam cada campo ausente?

## Funcionalidades

- Lê pacotes `.xdmz` ou pastas com múltiplos arquivos `.xdmz`.
- Extrai texto de documentos `.docx` diretamente do XML interno do Word.
- Compara documentação e referência usando escopo por âncoras ou correspondência por nome.
- Gera saídas em JSON, Markdown, CSV de diagnóstico e dashboard HTML.
- Roda sem serviços externos por padrão.
- Permite enriquecimento opcional por arquivos de mapeamento quando existe um catálogo de referência disponível.

## Estrutura do repositório

```text
validator_py/         Motor Python de comparação e pós-processamento
scripts/              Entradas PowerShell portáveis
samples/              Modelo e documentação sintéticos para demonstração
assets/               Exemplos públicos e genéricos de mapeamento
tests/                Testes de linhagem SQL e leitura de documentos
docs/                 Proposta e posicionamento para portfólio
```

## Como executar

```powershell
python -m unittest discover -s tests

powershell -ExecutionPolicy Bypass -File .\scripts\run_reference_equity.ps1 -IncludePass
```

As saídas são gravadas em `output/demo`:

- `comparison_report.json`
- `comparison_report.md`
- `comparison_dashboard.html`
- `by_extractor/*.csv`

## Uso direto com Python

```powershell
python .\validator_py\docsync_pipeline.py `
  --doc .\samples\docs\reference_catalog.docx `
  --xdmz-dir .\samples\models `
  --out .\output\demo `
  --include-pass-in-dashboard `
  --skip-web-post
```

## Leitura para portfólio

A ideia central é equidade entre documentação e referência: a documentação e a implementação devem ter autoridade equilibrada. Quando um lado muda, o outro deve ser auditado com evidência rastreável, não apenas com revisão manual por amostragem.

Este projeto demonstra automação em Python, leitura de XML/ZIP, heurísticas de linhagem SQL, geração de dashboard e uma abordagem pragmática para governança de qualidade documental.