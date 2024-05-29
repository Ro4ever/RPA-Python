import re
import os
import xlwings as xw
from datetime import datetime
from playwright.sync_api import Playwright, sync_playwright, expect

file_path = r"Caminho do Arquivo"
wb = xw.Book(file_path)
sheet = wb.sheets['Planilha1']
data = sheet.range('A1').expand('table').value
serie = sheet.range('B1').expand('table').value
data_nota = sheet.range('C1').expand('table').value
wb.close()

# Função para formatar a data
def format_date(date_value):
    if isinstance(date_value, datetime):
        return date_value.strftime('%d/%m/%Y')
    return date_value

def run(playwright: Playwright) -> None:
    download_path = os.path.join(os.getcwd(), "downloads")
    if not os.path.exists(download_path):
        os.makedirs(download_path)

def run(playwright: Playwright) -> None:
    download_path = r"Caminho do Download"
    if not os.path.exists(download_path):
        os.makedirs(download_path)

    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context(accept_downloads=True)
    page = context.new_page()
    page.goto("Link")
    page.get_by_label("Usuário").fill("login")
    page.get_by_label("Senha").fill("senha")
    page.get_by_role("button", name="Entrar").click()

    for row in data:
        print(f"Processando linha: {row}")
        
        numero_documento = str(row[0])  
        serie = str(row[1])
        data_emissao = format_date(row[2])

        print(f"Preenchendo formulário com Nº Documento: {numero_documento} e Série: {serie} e Data: {data_emissao}")
        
        page.get_by_label("Data Final").fill(data_emissao)
        page.get_by_label("Data Inicial").fill(data_emissao)
        page.get_by_label("Nº Documento").fill(str(numero_documento))
        page.get_by_label("Série").fill(serie)
        page.get_by_text("Pesquisar").click()
              
        try:
            page.locator("input[name=\"checkedRecords\"]").check()
            
            page.get_by_title("Fazer download do arquivo XML").locator("div").nth(1).click()
            
            page.get_by_text("select").nth(4).click()
            page.locator("div").filter(has_text=re.compile(r"^Documento Originalselect$")).first.press("ArrowDown")
            page.locator("div").filter(has_text=re.compile(r"^Documento enviado para SEFAZselect$")).first.press("ArrowDown")
            page.locator("div").filter(has_text=re.compile(r"^Documento processado \(Proc/ProcEvento\)select$")).first.press("ArrowDown")
            page.locator("div").filter(has_text=re.compile(r"^Documento PDFselect$")).first.press("Enter")
            
            with page.expect_download() as download_info:
                page.get_by_role("button", name="Download").click()
            download = download_info.value

            download_target_path = os.path.join(download_path, download.suggested_filename)
            download.save_as(download_target_path)
            
            print(f"Arquivo baixado para: {download_target_path}")
            
        except Exception as e:
            print(f"Erro ao processar linha {row}: {e}")
        
       
        page.wait_for_timeout(2000)

    context.close()
    browser.close()

with sync_playwright() as playwright:
    run(playwright)
