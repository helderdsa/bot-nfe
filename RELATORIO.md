# Relatório de Automação — Bot Emissor de NF-e

Implementei um bot de automação para emissão de Notas Fiscais de Serviço Eletrônicas (NFS-e), usando **Python**, **Selenium WebDriver** e a API interna do Advbox para coleta das transações financeiras, integrados a uma interface gráfica construída com **CustomTkinter**.

O bot consome as transações do período informado via API, filtra as que requerem emissão de nota, e preenche automaticamente todos os formulários no portal [nfse.gov.br](https://www.nfse.gov.br), incluindo data de competência, dados do tomador, código de tributação, descrição do serviço e valores — com tratamento de erros, retentativas automáticas e geração de relatório Excel ao final.

O resultado foi a **eliminação de um processo manual repetitivo** que ocupava a equipe do financeiro por mais de uma semana a cada fechamento: a emissão, que antes era feita nota por nota no portal, agora é concluída em horas pelo bot, com rastreabilidade completa via planilha de status (SUCESSO / ERRO / IGNORADO) e suporte a retomada a partir de execuções anteriores.
