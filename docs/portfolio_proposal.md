# Proposta de Portfólio: Equidade Entre Documentação e Referências

## Posicionamento

Reference Equity DocSync é uma ferramenta de governança para times que mantêm documentação técnica junto de referências de implementação. O projeto trata documentação e artefatos de referência como duas fontes que precisam permanecer equivalentes, rastreáveis e revisáveis.

## Problema

Quando um modelo de dados, contrato de API, definição de relatório ou referência de integração muda, a documentação frequentemente fica para trás. A revisão manual é lenta e tende a focar nas mudanças mais óbvias, deixando divergências menores, campo a campo, invisíveis até que usuários encontrem inconsistências.

## Solução

A ferramenta lê pacotes de referência, delimita as seções correspondentes na documentação e gera um conjunto de evidências:

- campos ausentes por dataset
- tabela e coluna de origem quando a linhagem está disponível
- auditoria de seleção do modelo quando há múltiplos arquivos candidatos
- JSON para automação
- Markdown para notas de revisão
- dashboard HTML para públicos técnicos e funcionais

## Por que isso importa

A proposta não é apenas comparar arquivos. É criar equidade documental: implementação e documentação escrita devem ter autoridade balanceada, e as lacunas precisam ficar visíveis antes da publicação.

## Narrativa da demonstração

A demo sintética inclui um modelo de referência de clientes em que `ACCOUNT_STATUS` existe no modelo de origem, mas está ausente na documentação publicada. O dashboard destaca essa lacuna e mostra a referência de origem que deve ser revisada.