# Proposta de Portfolio: Equidade Entre Documentacao e Referencias

## Posicionamento

Reference Equity DocSync e uma ferramenta de governanca para times que mantem documentacao tecnica junto de referencias de implementacao. O projeto trata documentacao e artefatos de referencia como duas fontes que precisam permanecer equivalentes, rastreaveis e revisaveis.

## Problema

Quando um modelo de dados, contrato de API, definicao de relatorio ou referencia de integracao muda, a documentacao frequentemente fica para tras. A revisao manual e lenta e tende a focar nas mudancas mais obvias, deixando divergencias menores, campo a campo, invisiveis ate que usuarios encontrem inconsistencias.

## Solucao

A ferramenta le pacotes de referencia, delimita as secoes correspondentes na documentacao e gera um conjunto de evidencias:

- campos ausentes por dataset
- tabela e coluna de origem quando a linhagem esta disponivel
- auditoria de selecao do modelo quando ha multiplos arquivos candidatos
- JSON para automacao
- Markdown para notas de revisao
- dashboard HTML para publicos tecnicos e funcionais

## Por que isso importa

A proposta nao e apenas comparar arquivos. E criar equidade documental: implementacao e documentacao escrita devem ter autoridade balanceada, e as lacunas precisam ficar visiveis antes da publicacao.

## Narrativa da demonstracao

A demo publica usa a tabela oficial `ZX_LINES` como exemplo de referencia real. Campos como `TAX_RATE_CODE` e `TAX_AMT_FUNCL_CURR` existem no modelo de origem, mas foram omitidos da documentacao de exemplo para que o dashboard mostre lacunas reais de enriquecimento, incluindo link para a documentacao oficial.