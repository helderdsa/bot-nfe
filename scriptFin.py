import time
import requests
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

try:
    # URL da API
    api_url = "https://app.advbox.com.br/api/v1/transactions?date_payment_start=2026-01-01&date_payment_end=2026-01-31&entry_type=income"
    
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
    
    # Fazer requisição à API
    response = session.get(api_url, headers=headers)
    
    # Debug: Mostrar status da resposta
    print(f"Status Code: {response.status_code}")
    
    # Verificar se a requisição foi bem-sucedida
    if response.status_code == 200:
        data_json = response.json()
        transactions = data_json.get('data', [])
        
        print(f"✓ Dados obtidos com sucesso!")
        print(f"✓ Total de transações encontradas: {len(transactions)}")
        print(f"✓ Total geral (totalCount): {data_json.get('totalCount', 'N/A')}")
        
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
                    driver.find_element(By.CSS_SELECTOR, "#ServicoPrestado_Descricao").send_keys(str(transaction.get('description', ''))) #descrição aqui

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
                    radio_options[0].find_element(By.CSS_SELECTOR, "label").click()

                    radio_options[2].find_element(By.CSS_SELECTOR, "label").click()

                    radio_options[6].find_element(By.CSS_SELECTOR, "label").click()

                    wait.until(EC.presence_of_element_located((By.ID, "TributacaoFederal_PISCofins_SituacaoTributaria_chosen")))
                    time.sleep(0.5)
                    driver.find_elements(By.ID, "TributacaoFederal_PISCofins_SituacaoTributaria_chosen")[0].click()

                    time.sleep(0.5)
                    driver.find_elements(By.ID, "TributacaoFederal_PISCofins_SituacaoTributaria_chosen")[0].find_element(By.CSS_SELECTOR, "div > div > input").send_keys("00")

                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div > div > ul > li")))
                    time.sleep(0.5)
                    driver.find_elements(By.ID, "TributacaoFederal_PISCofins_SituacaoTributaria_chosen")[0].find_element(By.CSS_SELECTOR, "div > div > ul > li").click()
                    
                    wait.until(EC.presence_of_element_located((By.ID, "TributacaoFederal_PISCofins_TipoRetencao_chosen")))
                    time.sleep(0.5)
                    driver.find_elements(By.ID, "TributacaoFederal_PISCofins_TipoRetencao_chosen")[0].click()

                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div > ul > li[data-option-array-index='1']")))
                    time.sleep(0.5)
                    driver.find_elements(By.ID, "TributacaoFederal_PISCofins_TipoRetencao_chosen")[0].find_element(By.CSS_SELECTOR, "div > ul > li[data-option-array-index='1']").click()

                    driver.get("https://www.nfse.gov.br/EmissorNacional/Dashboard")
                    
                    # Se chegou aqui sem erro, marca como sucesso
                    sucesso = True
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
                        sucesso = True  # Marca como "sucesso" para pular para a próxima
            
            # Avança para a próxima transação
            index += 1
    else:
        print(f"✗ Erro ao fazer requisição à API. Status code: {response.status_code}")
        print(f"✗ Resposta da API: {response.text}")
        print(f"✗ Headers enviados: {headers}")
        transactions = []
        
except requests.exceptions.RequestException as e:
    print(f"✗ Erro na requisição à API: {e}")
    print(f"✗ Tipo do erro: {type(e).__name__}")
    transactions = []
except Exception as e:
    print(f"✗ Erro inesperado: {e}")
    print(f"✗ Tipo do erro: {type(e).__name__}")
    transactions = []

# ===== FIM DA REQUISIÇÃO À API =====



