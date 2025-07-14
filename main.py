from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import time
import json
import os


CHROME_DRIVER_PATH = r"D:\chromedriver-win64\chromedriver.exe"
TARGET_URL = "https://www.gov.br/ana/pt-br/assuntos/regulacao-e-fiscalizacao/normativos-e-resolucoes/resolucoes"
CAMINHO_JSON = r"D:\cod\Arsae\Automatiza_arsae\verificacao_de_resolucoes_ANA\resolucoes.json"


def aceitar_cookies(driver):
    """Tenta clicar no botão de aceitação de cookies, se disponível."""
    try:
        cookie_button = driver.find_element(By.XPATH, "/html/body/div[5]/div/div/div/div/div[2]/button[3]")
        cookie_button.click()
        print("Cookies aceitos.")
    except (NoSuchElementException, TimeoutException):
        print("Botão de cookies não encontrado — seguindo mesmo assim.")



def extrair_dados(driver):
    try:
        wait = WebDriverWait(driver, 20)

        # Espera o iframe carregar e troca para ele
        iframe = wait.until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
        driver.switch_to.frame(iframe)

        # Agora espera as resoluções
        wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "#conteudo_listagem .resolucao")))
        resolucoes_divs = driver.find_elements(By.CSS_SELECTOR, "#conteudo_listagem .resolucao")

        if not resolucoes_divs:
            print("Nenhuma resolução encontrada no iframe.")
            return []

        resolucoes_dados = []

        for div in resolucoes_divs:
            try:
                id_div = div.get_attribute('id')
                conteudo_b = div.find_element(By.CSS_SELECTOR, ".titulo_resolucao a b").text
                conteudos_i = [i.text for i in div.find_elements(By.TAG_NAME, "i")]
                onclick_content = div.find_element(By.CSS_SELECTOR, "a.ico_pdf").get_attribute('onclick')

                resolucao_data = {
                    'id_div': id_div,
                    'conteudo_b': conteudo_b,
                    'conteudos_i': conteudos_i,
                    'onclick_content': onclick_content
                }

                resolucoes_dados.append(resolucao_data)
                print(f"Resolução extraída: {conteudo_b}")

            except Exception as e:
                print(f"Erro ao extrair dados da resolução: {e}")
                continue

        return resolucoes_dados

    except Exception as e:
        print(f"Erro ao acessar as resoluções: {e}")
        return []

def salvar_json(dados, caminho_arquivo):
    """
    Salva os dados extraídos em um arquivo JSON
    
    Args:
        dados (list): Lista com os dados das resoluções
        caminho_arquivo (str): Caminho completo do arquivo JSON
    """
    try:
        # Criar o diretório se não existir
        os.makedirs(os.path.dirname(caminho_arquivo), exist_ok=True)
        
        # Salvar os dados em JSON
        with open(caminho_arquivo, 'w', encoding='utf-8') as arquivo:
            json.dump(dados, arquivo, indent=2, ensure_ascii=False)
        
        print(f"Dados salvos com sucesso em: {caminho_arquivo}")
        print(f"Total de resoluções extraídas: {len(dados)}")
        
    except Exception as e:
        print(f"Erro ao salvar arquivo JSON: {e}")

def main():
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--disable-notifications")
    # chrome_options.add_argument("--headless")  # opcional

    service = Service(CHROME_DRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=chrome_options)
    wait = WebDriverWait(driver, 15)

    try:
        print(f"Acessando a URL: {TARGET_URL}")
        driver.get(TARGET_URL)

        aceitar_cookies(driver)

        dados_resolucoes = extrair_dados(driver)

        if dados_resolucoes:
            # Salvar dados em JSON
            salvar_json(dados_resolucoes, CAMINHO_JSON)
        else:
            print("Nenhuma resolução encontrada.")

        input()
    except Exception as e:
        print(f"Erro inesperado: {e}")
    finally:
        driver.quit()
        print("Driver finalizado.")


if __name__ == "__main__":
    main()
