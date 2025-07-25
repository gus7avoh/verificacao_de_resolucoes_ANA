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
import win32com.client as win32

CHROME_DRIVER_PATH = r"D:\chromedriver-win64\chromedriver.exe"
TARGET_URL = "https://www.gov.br/ana/pt-br/assuntos/regulacao-e-fiscalizacao/normativos-e-resolucoes/resolucoes"
CAMINHO_JSON_ATUAL= r"D:\cod\Arsae\Automatiza_arsae\verificacao_de_resolucoes_ANA\atual.json"
CAMINHO_JSON_ANTIGO= r"D:\cod\Arsae\Automatiza_arsae\verificacao_de_resolucoes_ANA\antigo.json"
CAMINHO_JSON_ALTERACOES= r"D:\cod\Arsae\Automatiza_arsae\verificacao_de_resolucoes_ANA\alteracoes.json"
CAMINHO_JSON_CONTAGEM= r"D:\cod\Arsae\Automatiza_arsae\verificacao_de_resolucoes_ANA\Contagem.json"
EMAIL = "gustavoh.a.m1409@gmail.com"



def aceitar_cookies(driver):
    """Tenta clicar no botão de aceitação de cookies, se disponível."""
    try:
        cookie_button = driver.find_element(By.XPATH, "/html/body/div[5]/div/div/div/div/div[2]/button[3]")
        cookie_button.click()
        print("Cookies aceitos.")
    except (NoSuchElementException, TimeoutException):
        print("Botão de cookies não encontrado — seguindo mesmo assim.")

def Tratar_url(texto):
    """
    Extrai e retorna o número da resolução a partir do parâmetro 'onclick',
    como em: abreArquivo('URL_AQUI')
    """
    if not texto.startswith("abreArquivo("):
        return None 

    try:
        # Remove o prefixo "abreArquivo('"
        url = texto.split("abreArquivo('")[1].split("')")[0]
        return url
    except IndexError:
        return None

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
                url = Tratar_url(onclick_content)

                resolucao_data = {
                    'id_div': id_div,
                    'Titulo': conteudo_b,
                    'Subtitulo': conteudos_i,
                    'url': url
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

def Verificar_alteracao():
    modificacao = []
    try:
        with open(CAMINHO_JSON_ANTIGO, 'r', encoding='utf-8') as antigo:
            dados_antigo = json.load(antigo)

        with open(CAMINHO_JSON_ATUAL, 'r', encoding='utf-8') as atual:
            dados_atual = json.load(atual)


        ids_antigos = set(item['id_div'] for item in dados_antigo)
        for item in dados_atual:
            if item['id_div'] not in ids_antigos:
                item["estado"] = "adicionado"
                modificacao.append(item)


        ids_atual = set(item['id_div'] for item in dados_atual)
        for item in dados_antigo:
            if item['id_div'] not in ids_atual:
                item["estado"] = "removido"
                modificacao.append(item)


    except Exception as erro:
        print(f"Erro: {erro}")

    return modificacao


def atualisar_json_antigo():
    try:
        with open(CAMINHO_JSON_ATUAL, 'r', encoding='utf-8') as atual:
            dados_atual = json.load(atual)

        salvar_json(dados_atual, CAMINHO_JSON_ANTIGO)

    except Exception as erro:
        print(f"erro: {erro}")

def limpar_json_alteracoes():
    try:
        with open(CAMINHO_JSON_ALTERACOES, 'w', encoding='utf-8') as arquivo:
            json.dump([], arquivo, indent=2, ensure_ascii=False)

        print("JSON de alterações limpo com sucesso.")
    except Exception as erro:
        print(f"Erro ao limpar JSON de alterações: {erro}")


def Enviar_email_alteracoes_outlook(alteracoes):
    # Garante que o arquivo de contagem exista e seja lido
    try:
        if os.path.exists(CAMINHO_JSON_CONTAGEM):
            with open(CAMINHO_JSON_CONTAGEM, 'r', encoding='utf-8') as arquivo:
                Contagem = json.load(arquivo)
        else:
            Contagem = {"Contagem": 0}
    except Exception as e:
        print(f"Erro ao ler JSON de contagem: {e}")
        Contagem = {"Contagem": 0}

    # CASO NÃO TENHA ALTERAÇÕES
    if not alteracoes:
        Contagem["Contagem"] += 1
        print("Nenhuma alteração encontrada.")

        if Contagem["Contagem"] == 7:
            try:
                corpo_email = "Nenhuma atualização encontrada durante a semana.\n\n"
                outlook = win32.Dispatch('outlook.application')
                mail = outlook.CreateItem(0)
                mail.To = "; ".join(EMAIL)
                mail.Subject = "Sem alterações detectadas em um período de 7 dias"
                mail.Body = corpo_email
                mail.Send()
                print("E-mail enviado com sucesso via Outlook (sem alterações).")
                Contagem["Contagem"] = 0  # zera após envio
            except Exception as erro:
                print(f"Erro ao enviar e-mail via Outlook: {erro}")

        salvar_json(Contagem, CAMINHO_JSON_CONTAGEM)
        return

    # CASO TENHA ALTERAÇÕES
    try:
        corpo_email = "Foram detectadas as seguintes alterações:\n\n"
        for item in alteracoes:
            corpo_email += f"Estado: {item.get('estado', 'N/A')}\n"
            corpo_email += f"Título: {item['Titulo']}\n"
            corpo_email += f"Subtítulo: {' - '.join(item['Subtitulo'])}\n"
            corpo_email += f"Link: {item['url']}\n"
            corpo_email += f"ID: {item['id_div']}\n"
            corpo_email += "-" * 40 + "\n\n"

        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = "; ".join(EMAIL)
        mail.Subject = "Alterações Detectadas no Sistema"
        mail.Body = corpo_email
        mail.Send()
        print("E-mail enviado com sucesso via Outlook (com alterações).")

        Contagem["Contagem"] = 0  # zera após envio
        salvar_json(Contagem, CAMINHO_JSON_CONTAGEM)

    except Exception as erro:
        print(f"Erro ao enviar e-mail via Outlook: {erro}")


def main():
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--disable-notifications")
    #chrome_options.add_argument("--headless") 

    service = Service(CHROME_DRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=chrome_options)
    wait = WebDriverWait(driver, 15)

    try:
        limpar_json_alteracoes()

        teste = True
        if teste  == False:
            print(f"Acessando a URL: {TARGET_URL}")
            driver.get(TARGET_URL)

            aceitar_cookies(driver)

            dados_resolucoes = extrair_dados(driver)

            if dados_resolucoes:
                salvar_json(dados_resolucoes, CAMINHO_JSON_ATUAL)
            else:
                print("Nenhuma resolução encontrada.")
        
        alteracoes = Verificar_alteracao()
        if alteracoes:
            salvar_json(alteracoes, CAMINHO_JSON_ALTERACOES)
        else:
            print("Nenhuma alteração encontrada.")

        with open(CAMINHO_JSON_ALTERACOES, 'r', encoding='utf-8') as arquivo:
            dados = json.load(arquivo)

        Enviar_email_alteracoes_outlook(dados)
        atualisar_json_antigo()

    except Exception as e:
        print(f"Erro inesperado: {e}")
    finally:
        driver.quit()
        print("Driver finalizado.")


if __name__ == "__main__":
    main()
