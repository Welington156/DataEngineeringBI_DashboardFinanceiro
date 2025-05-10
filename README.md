Teste Técnico – Data Engineering & BI - Dashboard Financeiro

Este repositório contém uma solução completa para automação e visualização em Excel (com VBA) e Power BI, dividido em níveis de complexidade.

Nível 1 – Fórmulas e Dashboards Excel

Objetivo: em até 90 minutos, preencher colunas de datas e valores na aba BASE através de fórmulas e criar dashboards que apresentem:

Valor total de entradas faturadas (por mês & ano)

Valor total de saídas faturadas (por mês & ano)

Valor total de estornos faturados (por mês & ano)

Saldo final faturado (por mês)

Ranking de clientes (entradas) – Top 5

Ranking de credores (saídas) – Top 5

Solução:

Uso de IFERROR para tratar erros e VLOOKUP / XLOOKUP para prazos de pagamento.

Fórmulas de agregação em SUMIFS com colunas auxiliares de MÊS e ANO.

Dashboards na aba DASHBOARD usando gráficos de colunas, linhas e barras.

Nível 2 – Formulário VBA para Inserções

Objetivo: permitir inserção/alteração de registros via formulário VBA, validando emissores e natureza de operação.

Solução:

UserForm em VBA que insere novas linhas no topo da tabela.

Validação contra listas na aba CADASTROS.

Proteção da aba BASE para impedir edição direta.

Nível 3 – Macro de Envio de E‑mail e Importação de XML

Objetivo: botão que envia e‑mail com resumo de KPIs e importa apenas registros existentes de um XML.

Solução:

Macro usando Outlook.Application para envio de e‑mail.

Parsing de XML via MSXML2.DOMDocument e inserção seletiva em BASE.


Nível 4 – Observações e Evoluções

Power BI: implementação adicional para dashboards avançados.

Testes Automatizados: cobertura com VBA unit testing frameworks.

Melhoria na Usabilidade: integração web para formulário e envio via Power Automate.

Documentação e Vídeo: apresentação gravada das funcionalidades.

Como Executar

Abra o arquivo Excel correspondente ao nível desejado.

Habilite macros quando solicitado.


