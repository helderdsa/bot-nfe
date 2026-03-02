import os
import time
import requests
import openpyxl
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ===== REQUISIÇÃO À API =====
print("Fazendo requisição à API...")

options = Options()
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--disable-gpu')
options.add_argument("--start-maximized")
options.add_experimental_option(
    "prefs",
    {
        "profile.default_content_setting_values.automatic_downloads": 1
    }
)
options.add_argument("--ignore-certificate-errors")

driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 20)

# Listas de rastreamento
transactions = []
ids_emitidas = []
ids_erro = []
ids_puladas = []

try:
    # URL base da API (sem offset/limit)
    api_base_url = "https://app.advbox.com.br/api/v1/transactions?date_payment_start=2026-02-01&date_payment_end=2026-02-28"
    # api_base_url = "https://app.advbox.com.br/api/v1/transactions?date_payment_start=2026-01-01&date_payment_end=2026-01-31&entry_type=income"
    
    # Criar sessão para manter cookies
    session = requests.Session()
    
    # Header de autenticação
    headers = {
        "Authorization": "Bearer 2gAucrPVxikEyTeHNj0QIhLQci2NE9u2hZTndQPV2D1E96J2RBfEaaVfG2Xh",
        "Content-Type": "application/json",
        "Accept": "application/json",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
        "Accept-Encoding": "gzip, deflate, br",
        "Connection": "keep-alive",
        "Upgrade-Insecure-Requests": "1"
    }
    
    LIMIT = 1000
    transactions = []
    offset = 0
    total_count = None

    # Buscar todos os registros usando paginação por offset
    while True:
        api_url = f"{api_base_url}&offset={offset}&limit={LIMIT}"
        response = session.get(api_url, headers=headers)

        print(f"Status Code: {response.status_code} (offset={offset})")

        if response.status_code != 200:
            print(f"✗ Erro ao fazer requisição à API. Status code: {response.status_code}")
            print(f"✗ Resposta da API: {response.text}")
            break

        data_json = response.json()

        if total_count is None:
            total_count = data_json.get('totalCount', 0)
            print(f"✓ Total geral (totalCount): {total_count}")

        page_data = data_json.get('data', [])
        transactions.extend(page_data)
        print(f"✓ Página carregada: {len(page_data)} registros (total acumulado: {len(transactions)})")

        if len(transactions) >= total_count or len(page_data) == 0:
            break

        offset += LIMIT

    print(f"✓ Todos os dados obtidos com sucesso!")
    print(f"✓ Total de transações encontradas: {len(transactions)}")

    # Verificar se a requisição foi bem-sucedida
    if len(transactions) > 0:
        
        driver.get("https://www.nfse.gov.br/EmissorNacional/Login?ReturnUrl=%2fEmissorNacional")
        wait.until(EC.presence_of_element_located((By.CLASS_NAME, "img-certificado")))
        time.sleep(0.5)
        driver.find_element(By.CLASS_NAME, "img-certificado").click()
        
        # Iterar sobre cada transação com tratamento de erro
        print("\n===== Processando transações =====")
        index = 0
        max_tentativas = 5  # Número máximo de tentativas por transação
        
        while index < len(transactions):
            transaction = transactions[index]

            if transaction.get('entry_type') == "expense":
                print(f"\n--- Transação {index + 1} de {len(transactions)} ---")
                print(f"⏭ Pulando transação ID {transaction.get('id')} - categoria 'TAXAS BANCÁRIAS' não gera NF.")
                ids_puladas.append(transaction.get('id'))
                index += 1
                continue

            tentativas = 0
            sucesso = False
            
            while tentativas < max_tentativas and not sucesso:
                try:
                    if tentativas > 0:
                        print(f"\n⚠ Tentativa {tentativas + 1} de {max_tentativas} para a transação {index + 1}")
                    
                    print(f"\n--- Transação {index + 1} de {len(transactions)} ---")
                    print(f"ID: {transaction.get('id')}")
                    print(f"Tipo de Entrada: {transaction.get('entry_type')}")
                    print(f"Data Vencimento: {transaction.get('date_due')}")
                    print(f"Data Pagamento: {transaction.get('date_payment')}")
                    print(f"Competência: {transaction.get('competence')}")
                    print(f"Valor: {transaction.get('amount')}")
                    print(f"Descrição: {transaction.get('description')}")
                    print(f"Responsável: {transaction.get('responsible')}")
                    print(f"Categoria: {transaction.get('category')}")
                    print(f"ID Processo: {transaction.get('lawsuit_id')}")
                    print(f"Número Processo: {transaction.get('process_number')}")
                    print(f"Número Protocolo: {transaction.get('protocol_number')}")
                    print(f"Nome: {transaction.get('name')}")
                    print(f"Identificação: {transaction.get('identification')}")
                    print(f"Banco Débito: {transaction.get('debit_bank')}")
                    print(f"Banco Crédito: {transaction.get('credit_bank')}")
                    print(f"Centro de Custo: {transaction.get('cost_center')}")
                    

                    
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".btnAcesso")))
                    time.sleep(0.5)
                    driver.find_element(By.CSS_SELECTOR, ".btnAcesso").click()

                    wait.until(EC.presence_of_element_located((By.ID, "DataCompetencia")))
                    time.sleep(0.5)
                    driver.find_element(By.ID, "DataCompetencia").send_keys("01/01/2026") #data de pagamento aqui

                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "body")))
                    time.sleep(0.5)
                    driver.find_element(By.CSS_SELECTOR, "body").click()

                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".radiobutton")))
                    time.sleep(0.5)
                    radio_options = driver.find_elements(By.CSS_SELECTOR, ".radiobutton")
                    radio_options[4].find_element(By.CSS_SELECTOR, "label").click()

                    wait.until(EC.presence_of_element_located((By.ID, "Tomador_Inscricao")))
                    time.sleep(0.5)
                    driver.find_element(By.ID, "Tomador_Inscricao").click()

                    time.sleep(0.5)
                    driver.find_element(By.ID, "Tomador_Inscricao").send_keys(str(transaction.get('identification', ''))) #CPF aqui

                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "body")))
                    time.sleep(0.5)
                    driver.find_element(By.CSS_SELECTOR, "body").click()

                    wait.until(EC.presence_of_element_located((By.ID, "btnAvancar")))
                    time.sleep(0.5)
                    driver.find_element(By.ID, "btnAvancar").click()

                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".select2-selection")))
                    time.sleep(0.5)
                    driver.find_element(By.CSS_SELECTOR, ".select2-selection").click()

                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input.select2-search__field")))
                    time.sleep(0.5)
                    driver.find_element(By.CSS_SELECTOR, "input.select2-search__field").send_keys("Acari")

                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".select2-results__option--highlighted")))
                    time.sleep(0.5)
                    driver.find_element(By.CSS_SELECTOR, ".select2-results__option--highlighted").click()

                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "span[aria-labelledby=select2-ServicoPrestado_CodigoTributacaoNacional-container]")))
                    time.sleep(0.5)
                    driver.find_element(By.CSS_SELECTOR, "span[aria-labelledby=select2-ServicoPrestado_CodigoTributacaoNacional-container]").click()

                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input.select2-search__field")))
                    time.sleep(0.5)
                    driver.find_element(By.CSS_SELECTOR, "input.select2-search__field").send_keys("171401")

                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".select2-results__option--highlighted")))
                    time.sleep(0.5)
                    driver.find_element(By.CSS_SELECTOR, ".select2-results__option--highlighted").click()

                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.radiobutton > label")))
                    time.sleep(0.5)
                    driver.find_element(By.CSS_SELECTOR, "div.radiobutton > label").click()

                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#ServicoPrestado_Descricao")))
                    time.sleep(0.5)
                    driver.find_element(By.CSS_SELECTOR, "#ServicoPrestado_Descricao").click()

                    time.sleep(0.5)
                    if transaction.get('category') == "IMPLANTAÇÕES":
                        driver.find_element(By.CSS_SELECTOR, "#ServicoPrestado_Descricao").send_keys("Pagamento referente a implantação de letras, parcela "+str(transaction.get('description', ''))) #descrição aqui
                    elif transaction.get('category') == "PRECATÓRIOS" or transaction.get('category') == "ALVARÁS" or transaction.get('category') == "HONORÁRIOS DE SUCUMBÊNCIA":
                        driver.find_element(By.CSS_SELECTOR, "#ServicoPrestado_Descricao").send_keys("Pagamento referente aos honorários contratuais, nº do processo: "+str(transaction.get('process_number', ''))) #descrição aqui
                    else :
                        driver.find_element(By.CSS_SELECTOR, "#ServicoPrestado_Descricao").send_keys("Pagamento referente a "+str(transaction.get('category', ''))) #descrição aqui
                    wait.until(EC.presence_of_element_located((By.CLASS_NAME, "btn-primary")))
                    time.sleep(0.5)
                    driver.find_element(By.CLASS_NAME, "btn-primary").click()

                    wait.until(EC.presence_of_element_located((By.ID, "Valores_ValorServico")))
                    time.sleep(0.5)
                    # Formatar valor com 2 casas decimais e vírgula (padrão BR)
                    valor = transaction.get('amount', '0')
                    valor_formatado = f"{float(valor):.2f}".replace('.', ',')
                    driver.find_element(By.ID, "Valores_ValorServico").send_keys(valor_formatado) #valor aqui

                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".radiobutton")))
                    time.sleep(0.5)
                    radio_options = driver.find_elements(By.CSS_SELECTOR, ".radiobutton")
                    radio_options[0].find_element(BY.CSS_SELECTOR, "label").click()

                    radio_options[2].find_element(BY.CSS_SELECTOR, "label").click()

                    radio_options[6].find_element(BY.CSS_SELECTOR, "label").click()

                    wait.until(EC.presence_of_element_located((By.ID, "TributacaoFederal_PISCofins_SituacaoTributaria_chosen")))
                    driver.find_elements(By.ID, "TributacaoFederal_PISCofins_SituacaoTributaria_chosen")[0].click()

                    driver.find_elements(By.ID, "TributacaoFederal_PISCofins_SituacaoTributaria_chosen")[0].find_element(By.CSS_SELECTOR, "div > div > input").send_keys("00")

                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div > div > ul > li")))
                    driver.find_elements(By.ID, "TributacaoFederal_PISCofins_SituacaoTributaria_chosen")[0].find_element(By.CSS_SELECTOR, "div > div > ul > li").click()
                    
                    wait.until(EC.presence_of_element_located((By.ID, "TributacaoFederal_PISCofins_TipoRetencao_chosen")))
                    driver.find_elements(By.ID, "TributacaoFederal_PISCofins_TipoRetencao_chosen")[0].click()

                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div > ul > li[data-option-array-index='1']")))
                    driver.find_elements(By.ID, "TributacaoFederal_PISCofins_TipoRetencao_chosen")[0].find_element(By.CSS_SELECTOR, "div > ul > li[data-option-array-index='1']").click()
                    time.sleep(0.5)

                    wait.until(EC.presence_of_element_located((By.CLASS_NAME, "btn-primary")))
                    time.sleep(0.5)
                    driver.find_element(By.CLASS_NAME, "btn-primary").click()


                    wait.until(EC.presence_of_element_located((By.ID, "btnProsseguir")))
                    time.sleep(0.5)
                    driver.find_element(By.ID, "btnProsseguir").click()

                    wait.until(EC.presence_of_element_located((By.ID, "btnDownloadDANFSE")))
                    time.sleep(0.5)
                    driver.get("https://www.nfse.gov.br/EmissorNacional/Dashboard")
                    
                    # FINALIZAR ID: btnDownloadDANFSE

                    # Se chegou aqui sem erro, marca como sucesso
                    sucesso = True
                    ids_emitidas.append(transaction.get('id'))
                    print(f"✓ Transação {index + 1} processada com sucesso!")
                    
                except Exception as e:
                    tentativas += 1
                    print(f"\n✗ Erro ao processar transação {index + 1}: {e}")
                    print(f"✗ Tipo do erro: {type(e).__name__}")
                    
                    # Retorna para o Dashboard para tentar novamente
                    try:
                        driver.get("https://www.nfse.gov.br/EmissorNacional/Dashboard")
                        time.sleep(10)  # Aguarda um pouco antes de tentar novamente
                    except:
                        pass
                    
                    if tentativas >= max_tentativas:
                        print(f"✗ Número máximo de tentativas atingido para a transação {index + 1}")
                        print(f"✗ Pulando para a próxima transação...")
                        ids_erro.append(transaction.get('id'))
                        sucesso = True  # Marca como "sucesso" para pular para a próxima
            
            # Avança para a próxima transação
            index += 1
        
except requests.exceptions.RequestException as e:
    print(f"✗ Erro na requisição à API: {e}")
    print(f"✗ Tipo do erro: {type(e).__name__}")
except Exception as e:
    print(f"✗ Erro inesperado: {e}")
    print(f"✗ Tipo do erro: {type(e).__name__}")
finally:
    # Calcular IDs que faltaram (não chegaram a ser processadas)
    ids_processados = set(ids_emitidas) | set(ids_erro) | set(ids_puladas)
    ids_faltaram = [t.get('id') for t in transactions if t.get('id') not in ids_processados]

    # Gerar planilha de relatório
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Relatório NF"
    ws.append(["Emitidas", "Puladas por Erro", "Faltaram"])

    max_len = max(len(ids_emitidas), len(ids_erro), len(ids_faltaram), 1)
    for i in range(max_len):
        ws.append([
            ids_emitidas[i] if i < len(ids_emitidas) else "",
            ids_erro[i] if i < len(ids_erro) else "",
            ids_faltaram[i] if i < len(ids_faltaram) else ""
        ])

    script_dir = os.path.dirname(os.path.abspath(__file__))
    filename = os.path.join(script_dir, f"relatorio_nf_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    wb.save(filename)
    print(f"\n===== RELATÓRIO FINAL =====")
    print(f"✓ Planilha salva: {filename}")
    print(f"  - Emitidas:        {len(ids_emitidas)}")
    print(f"  - Puladas por erro: {len(ids_erro)}")
    print(f"  - Faltaram:        {len(ids_faltaram)}")

# ===== FIM DA REQUISIÇÃO À API =====#