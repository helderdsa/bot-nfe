# Bot NFS-e — Emissor Automático de Notas Fiscais

Automação para emissão de NFS-e (Nota Fiscal de Serviços Eletrônica) no [Emissor Nacional](https://www.nfse.gov.br/EmissorNacional) com base nas transações financeiras obtidas da API do **AdvBox**.

---

## Como funciona

1. Consulta a API do AdvBox buscando transações de receita (`income`) em um período configurado.
2. Abre o Google Chrome via Selenium e acessa o Emissor Nacional de NFS-e.
3. Realiza login com certificado digital.
4. Para cada transação encontrada, preenche e emite uma NFS-e automaticamente com:
   - Data de competência
   - CPF/CNPJ do tomador (`identification`)
   - Descrição do serviço (`description`)
   - Valor (`amount`)
   - Código de tributação nacional: `171401`
   - Município: Acari
   - Configurações de tributação federal (PIS/COFINS)
5. Em caso de falha, tenta novamente até **5 vezes** por transação antes de pular para a próxima.

---

## Requisitos

- Python 3.8+
- Google Chrome instalado
- ChromeDriver compatível com a versão do Chrome instalada
- Certificado digital configurado no navegador
- Conta com acesso à API do AdvBox

### Dependências Python

Instale as dependências com:

```bash
pip install selenium requests
```

---

## Configuração

Antes de executar, edite as seguintes variáveis no arquivo [scriptFin.py](scriptFin.py):

| Variável / Campo | Localização no código | Descrição |
|---|---|---|
| `api_url` | linha com `api_url = ...` | URL da API com o período desejado (`date_payment_start` e `date_payment_end`) |
| `Authorization` | header `Bearer ...` | Token de autenticação da API do AdvBox |
| `"01/01/2026"` | `send_keys("01/01/2026")` | Data de competência das notas fiscais |

---

## Execução

```bash
python scriptFin.py
```

O script irá:
- Exibir o status de cada transação processada no terminal.
- Abrir o Chrome automaticamente para realizar as emissões.
- Retornar ao Dashboard do Emissor Nacional após cada NFS-e emitida.

---

## Observações

- O login no Emissor Nacional é feito via **certificado digital** (clique no ícone de certificado na tela de login).
- O script **não fecha o navegador** ao final, permitindo inspeção manual caso necessário.
- Transações que falharem após todas as tentativas serão **puladas**, com o erro registrado no terminal.
- O token de autenticação Bearer presente no código deve ser mantido em segredo e **não deve ser versionado publicamente**.
