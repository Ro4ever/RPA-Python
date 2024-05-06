import tkinter as tk
import re
import pyautogui
import time
from datetime import datetime
from tkinter import messagebox
from selenium import webdriver
from selenium import webdriver 
from selenium.webdriver.common.by import By 
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support import expected_conditions 
from selenium.webdriver.common.keys import Keys
from playwright.sync_api import sync_playwright
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

class RPAInterface:
    def __init__(self, root):
        self.root = root
        self.root.title("Comunicação Geral De Ocorrências")

        # Labels
        self.label_date_time = tk.Label(root, text="Data (dd/MM/yyyy):")
        self.label_date_time.grid(row=0, column=0, padx=10, pady=5, sticky="w")

        self.label_start_time = tk.Label(root, text="Hora Inicial (HH:MM):")
        self.label_start_time.grid(row=1, column=0, padx=10, pady=5, sticky="w")

        self.label_end_time = tk.Label(root, text="Hora Final (HH:MM):")
        self.label_end_time.grid(row=2, column=0, padx=10, pady=5, sticky="w")

        self.label_info = tk.Label(root, text="Informações sobre a ocorrência:")
        self.label_info.grid(row=3, column=0, padx=10, pady=5, sticky="w")

        # Caixas para informar os horários de início e final da ocorrência
        self.entry_start_time = tk.Entry(root)
        self.entry_start_time.grid(row=1, column=0, padx=10, pady=0)

        self.entry_end_time = tk.Entry(root)
        self.entry_end_time.grid(row=2, column=0, padx=10, pady=5)

        self.entry_date_time = tk.Entry(root)
        self.entry_date_time.grid(row=0, column=0, padx=10, pady=5)

        # Caixa de texto para inserir informações sobre a ocorrência
        self.text_info = tk.Text(root, width=50, height=5)
        self.text_info.grid(row=4, column=0, padx=10, pady=5)

        # Checkboxes com scrollbar
        self.options_frame = tk.Frame(root)
        self.options_frame.grid(row=5, column=0, columnspan=2, padx=10, pady=5, sticky="w")

        self.options_canvas = tk.Canvas(self.options_frame)
        self.options_canvas.pack(side="left", fill="both", expand=True)

        self.options_scrollbar = tk.Scrollbar(self.options_frame, orient="vertical", command=self.options_canvas.yview)
        self.options_scrollbar.pack(side="right", fill="y")

        self.options_canvas.configure(yscrollcommand=self.options_scrollbar.set)

        self.options_inner = tk.Frame(self.options_canvas, bg="white")
        self.options_canvas.configure(bg="white")
        self.options_canvas.create_window((0, 0), window=self.options_inner, anchor="nw")

        options = ["Opções CDDs"]  # Adicione suas opções aqui
        options.sort()
        self.checkbox_options = {}
        for idx, option in enumerate(options):
            var = tk.BooleanVar()
            checkbox = tk.Checkbutton(self.options_inner, text=option, variable=var, bg="white")
            checkbox.grid(row=idx, column=0, padx=5, pady=5, sticky="w")
            self.checkbox_options[option] = var

        self.options_inner.bind("<Configure>", self.onFrameConfigure)
        self.options_canvas.bind_all("<MouseWheel>", self.onMouseWheel)

        # Botões para selecionar/deselecionar todas as opções
        self.select_all_button = tk.Button(root, text="Selecionar Todas", command=self.select_all_options, width=20)
        self.select_all_button.grid(row=6, column=0, padx=10, pady=5)

        self.deselect_all_button = tk.Button(root, text="Desmarcar Todas", command=self.deselect_all_options, width=20)
        self.deselect_all_button.grid(row=7, column=0, padx=10, pady=5)

        # Botão para iniciar o processo
        self.start_button = tk.Button(root, text="Iniciar", command=self.start_process, width=20)
        self.start_button.grid(row=8, columnspan=2, padx=10, pady=10)


    def select_all_options(self):
        for var in self.checkbox_options.values():
            var.set(True)

    def deselect_all_options(self):
        for var in self.checkbox_options.values():
            var.set(False)

    def validate_time(self, time_str):
        # Regex para verificar o formato HH:MM
        time_regex = re.compile(r'^([01]\d|2[0-3]):([0-5]\d)$')
        return bool(time_regex.match(time_str))
    
    def validate_date(self, date_str):
        formated_date = re.compile(r'^(0[1-9]|[12][0-9]|3[01])/(0[1-9]|1[0-2])/(19|20)\d\d$')
        return bool(formated_date.match(date_str))

    def start_process(self):
        # Coletando os valores inseridos pelo usuário
        global additional_info
        global start_time
        global end_time
        global date_time
        start_time = self.entry_start_time.get()
        end_time = self.entry_end_time.get()
        date_time = self.entry_date_time.get()
        additional_info = self.text_info.get("1.0", tk.END)  # Obter texto da caixa de texto

        # Validando a data inserida
        if not self.validate_date(date_time):
            messagebox.showerror("Erro", "Data inválida. Por favor, insira no formato dd/MM/yyyy")
            return

        # Validando os horários inseridos
        if not self.validate_time(start_time):
            messagebox.showerror("Erro", "Hora inicial inválida. Por favor, insira no formato HH:MM.")
            return

        if not self.validate_time(end_time):
            messagebox.showerror("Erro", "Hora final inválida. Por favor, insira no formato HH:MM.")
            return

        # Validando se a caixa de texto possui alguma informação
        if additional_info.replace("\n", "") == "":
            messagebox.showerror("Erro", "Insira algum texto na caixa")
            return
        
        global selected_options
        selected_options = [option for option, var in self.checkbox_options.items() if var.get()]

        if selected_options == []:
            messagebox.showerror("Erro", "Selecione pelo menos um CDD")
            return

        #messagebox.showinfo("Informações", f"Hora Inicial: {start_time}\nHora Final: {end_time}\nInformações Adicionais:\n{additional_info}\nOpções Selecionadas: {selected_options}")
        
        root.destroy()

    def onFrameConfigure(self, event):
        '''Reset the scroll region to encompass the inner frame'''
        self.options_canvas.configure(scrollregion=self.options_canvas.bbox("all"))

    def onMouseWheel(self, event):
        self.options_canvas.yview_scroll(int(-1*(event.delta/120)), "units")

root = tk.Tk()
app = RPAInterface(root)
root.mainloop()

# Abre o navegador no link setado
navegador = webdriver.Chrome()
navegador.get('link')
navegador.maximize_window()

# Realizar login
WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class='select2-choice']"))).click()
navegador.find_element(By.XPATH, "//*[@class='select2-choice']").send_keys(Keys.ARROW_DOWN)
time.sleep(1)
navegador.find_element(By.XPATH, "//*[@class='select2-choice']").send_keys(Keys.ENTER)

# Obtem e insere o Usuário
WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class='select2-choice']")))
navegador.find_element(By.XPATH, "//*[@class='form-control mv-input-radius p-5 username-input-login-service']").send_keys("login")

# Obtem e insere a Senha
WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By. XPATH, "//*[@class='form-control mv-input-radius p-5 password-input-login-service']")))
navegador.find_element(By.XPATH, "//*[@class='form-control mv-input-radius p-5 password-input-login-service']").send_keys("senha")

# Obtem e clica no botão Entrar
WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By. XPATH, "//*[@id='btnSubmit']"))).click()
try:
    # Obtem janela de erro, quando já tem usuário logado
    WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class = 'btn-mv btn-mv-confirm md-confirm-action']"))).click()
except:
    print("Não precisou")

# Processo de Envio de Ocorrências de CDDs
try:
    for options in selected_options:
        print(options)
        print(type(options))

        # Obtem e clica no Botão de novo Ticket
        WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class='button-more-icon-container']"))).click()

        # Obtem e clica no Botão de novo Ticket 2
        WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "(//*[@class='dropdown-item-text-movidesk'])[1]"))).click()

        # Obtem e clica no botão solicitante
        WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class='select2-container mv-input-select2-container clients']"))).click()
        time.sleep(3)
        
        # Insere o CDD
        WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class='select2-choice']"))).send_keys(options)
        #pyautogui.typewrite(options, interval=0.3)
        time.sleep(5)
        navegador.find_element(By.XPATH, "//*[@class='select2-choice']").send_keys(Keys.ENTER)
        time.sleep(3)

        # Seleciona o tipo de chamado
        WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class='md-select-treeview-container md-single-select md-small']"))).click()
        navegador.find_element(By.XPATH, "//*[@class='md-select-treeview-search']").send_keys("Servico de Monitoramento Indisponivel")
        navegador.find_element(By.XPATH, "//*[@class='md-select-treeview-search']").click()
        time.sleep(3)
        navegador.find_element(By.XPATH, "//*[@class='md-select-treeview-search']").send_keys(Keys.ENTER)
        time.sleep(3)
        WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class='jqx-tree-item-li']"))).click()
        time.sleep(3)

        # Seleciona o tipo de ocorrência
        WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class='select2-container form-control category']"))).click()
        time.sleep(3)
        navegador.find_element(By.XPATH, "//*[@id='s2id_autogen18_search']").send_keys("Ocorrencias")
        navegador.find_element(By.XPATH, "//*[@id='s2id_autogen18_search']").click()
        time.sleep(3)
        WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class='select2-match']"))).click()
        time.sleep(3)

        # Seleciona o tipo de urgência
        WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class='select2-container form-control urgency']"))).click()
        navegador.find_element(By.XPATH, "//*[@id='s2id_autogen19_search']").send_keys("Urgente")
        navegador.find_element(By.XPATH, "//*[@id='s2id_autogen19_search']").click()
        time.sleep(3)
        WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class='select2-results-dept-0 select2-result select2-result-selectable']"))).click()
        time.sleep(3)

        # Seleciona e insere a data na caixa
        WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class='input-mv-new custom-fields hasDatepicker']"))).click()
        navegador.find_element(By.XPATH, "//*[@class='input-mv-new custom-fields hasDatepicker']").send_keys(date_time)
        navegador.find_element(By.XPATH, "//*[@class='input-mv-new custom-fields hasDatepicker']").click()
        time.sleep(3)
        navegador.find_element(By.XPATH, "//*[@class='input-mv-new custom-fields hasDatepicker']").send_keys(Keys.ENTER)
        time.sleep(3)
        

        # Seleciona e insere a hora 
        navegador.find_element(By.XPATH, "//*[@class='input-mv-new custom-fields hasDatepicker']").send_keys(Keys.TAB)
        time.sleep(3)
        pyautogui.typewrite(start_time, interval=0.3)
        time.sleep(3)

        # Seleciona e aplica a macro
        navegador.find_element(By.XPATH, "//*[@class='input-mv-new custom-fields hasDatepicker']").send_keys(Keys.ALT, 'm')
        #WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@id='s2id_autogen20']"))).click()
        time.sleep(3)
        navegador.find_element(By.XPATH, "//*[@id='s2id_autogen21_search']").send_keys(Keys.BACK_SPACE, "Ocorrencias")
        navegador.find_element(By.XPATH, "//*[@id='s2id_autogen21_search']").click()
        time.sleep(3)
        WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class='select2-results-dept-0 select2-result select2-result-selectable select2-highlighted']"))).click()
        time.sleep(3)
        navegador.find_element(By.XPATH, "//*[@class='input-mv-new custom-fields hasDatepicker']").send_keys(Keys.ALT, 'm')

        # Define a data final do preenchimento
        navegador.find_element(By.XPATH, "//*[@class='input-mv-new custom-fields hasDatepicker']").send_keys(Keys.ESCAPE)
        time.sleep(5)
        hora_atual = datetime.now().strftime("%H:%M")
        print(hora_atual)
        WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@name='Appointment[0].PeriodEnd']"))).click()
        time.sleep(5)
        pyautogui.typewrite(hora_atual, interval=0.3)

        #Seleciona e insere descrição
        framePagina = navegador.find_element(By.XPATH, "/html/body/section[1]/div[3]/div/div[1]/section/form/div[3]/div[2]/div[2]/div[1]/div[8]/section/div[3]/div[2]/div/div/div[2]/iframe")
        navegador.switch_to.frame(framePagina)
        time.sleep(3)
        WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@id='tinymce']/table/tbody/tr[5]/td/p[3]/span"))).click()
        time.sleep(3)
        navegador.find_element(By.XPATH, "//*[@id='tinymce']/table/tbody/tr[5]/td/p[3]/span").send_keys(additional_info)
        navegador.find_element(By.XPATH, "//*[@id='tinymce']/table/tbody/tr[5]/td/p[3]/span").click()
        time.sleep(3)
        navegador.switch_to.default_content()

        # Obter e clicar no botão salvar
        WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class='right-group']"))).click()
        time.sleep(10)

        # Obter e clicar no botão continuar editando
        WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[text()='Continuar editando']"))).click()
        time.sleep(5)

        # Obter e clicar no botão de editar
        WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class='action-item-option action-item-edit action-item-menu-subitem']"))).click()
        time.sleep(5)

        # Obter e clicar no botão OK
        WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class='btn-mv btn-mv-confirm md-confirm-action']"))).click()
        time.sleep(5)

        # Obter e clicar na caixa de Resolvido
        WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class='select2-container finish-ticket-status-list']"))).click()
        time.sleep(5)
        pyautogui.typewrite("Resolvido", interval=0.3)
        time.sleep(5)
        WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class='select2-result-label']"))).click()
        time.sleep(5)

        # Obter e clicar em finalizar edição
        WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[text()='Finalizar edição']"))).click()
        time.sleep(5)

        # Fechar aba de chamado
        pyautogui.hotkey('alt', 'w')
        time.sleep(5)
        pyautogui.hotkey('f5')

# Dispara um email caso o processo encontre um problema
except:
    def conexaoSMTP(email, senha):
        servidor_email = smtplib.SMTP('smtp.office365.com', 587)
        servidor_email.starttls()
        servidor_email.login(email, senha)
        return servidor_email

    email = "login email"
    senha = "senha email"
    servidor_email = conexaoSMTP(email, senha)

    # Cabeçalho do E-mail
    mensagem = MIMEMultipart()
    mensagem['From'] = email #Remetente
    mensagem['To'] = ", ".join(["emails para enviar"]) # Destinatários 
    mensagem['Subject'] = f"RPA - Comunicação Geral de Ocorrências {datetime.now().date().strftime("%d/%m/%Y")}" # Assunto

    corpo_email = f"""</body></html><p>Olá, ocorreu uma falha ao abrir chamados no processo Comunicação Geral de Ocorrências</p>"""

    mensagem.attach(MIMEText(corpo_email, 'html'))
    try:
        servidor_email.sendmail(email, mensagem['To'].split(", "), mensagem.as_string())
    except (smtplib.SMTPServerDisconnected, smtplib.SMTPException) as e:
        print(f"Erro ao enviar e-mail: {e}")
        # Reconectar ao servidor SMTP
        servidor_email = conexaoSMTP(email, senha)
        servidor_email.sendmail(email, mensagem['To'].split(", "), mensagem.as_string())
    servidor_email.quit()

# Dispara um email quando o processo for concluído
def conexaoSMTP(email, senha):
    servidor_email = smtplib.SMTP('smtp.office365.com', 587)
    servidor_email.starttls()
    servidor_email.login(email, senha)
    return servidor_email

email = "login email"
senha = "senha email"
servidor_email = conexaoSMTP(email, senha)

# Cabeçalho do E-mail
mensagem = MIMEMultipart()
mensagem['From'] = email #Remetente
mensagem['To'] = ", ".join(["emails para enviar"]) # Destinatários 
mensagem['Subject'] = f"RPA - Comunicação Geral de Ocorrências {datetime.now().date().strftime("%d/%m/%Y")}" # Assunto

corpo_email = f"""</body></html><p>Olá, Processo Comunicação Geral de Ocorrências finalizado com sucesso!!</p>"""

mensagem.attach(MIMEText(corpo_email, 'html'))
try:
    servidor_email.sendmail(email, mensagem['To'].split(", "), mensagem.as_string())
except (smtplib.SMTPServerDisconnected, smtplib.SMTPException) as e:
    print(f"Erro ao enviar e-mail: {e}")
    # Reconectar ao servidor SMTP
    servidor_email = conexaoSMTP(email, senha)
    servidor_email.sendmail(email, mensagem['To'].split(", "), mensagem.as_string())
servidor_email.quit()

time.sleep(30)