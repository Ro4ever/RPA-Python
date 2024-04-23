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

class RPAInterface:
    def __init__(self, root):
        self.root = root
        self.root.title("Comunicação Geral De Ocorrências")

        # Labels
        self.label_start_time = tk.Label(root, text="Hora Inicial (HH:MM):")
        self.label_start_time.grid(row=0, column=0, padx=10, pady=5, sticky="w")

        self.label_end_time = tk.Label(root, text="Hora Final (HH:MM):")
        self.label_end_time.grid(row=1, column=0, padx=10, pady=5, sticky="w")

        self.label_info = tk.Label(root, text="Informações sobre a ocorrência:")
        self.label_info.grid(row=2, column=0, padx=10, pady=5, sticky="w")

        # Caixas para informar os horários de início e final da ocorrência
        self.entry_start_time = tk.Entry(root)
        self.entry_start_time.grid(row=0, column=0, padx=10, pady=0)

        self.entry_end_time = tk.Entry(root)
        self.entry_end_time.grid(row=1, column=0, padx=10, pady=5)

        # Caixa de texto para inserir informações sobre a ocorrência
        self.text_info = tk.Text(root, width=50, height=5)
        self.text_info.grid(row=3, column=0, padx=10, pady=5)

        # Checkboxes com scrollbar
        self.options_frame = tk.Frame(root)
        self.options_frame.grid(row=4, column=0, columnspan=2, padx=10, pady=5, sticky="w")

        self.options_canvas = tk.Canvas(self.options_frame)
        self.options_canvas.pack(side="left", fill="both", expand=True)

        self.options_scrollbar = tk.Scrollbar(self.options_frame, orient="vertical", command=self.options_canvas.yview)
        self.options_scrollbar.pack(side="right", fill="y")

        self.options_canvas.configure(yscrollcommand=self.options_scrollbar.set)

        self.options_inner = tk.Frame(self.options_canvas, bg="white")
        self.options_canvas.configure(bg="white")
        self.options_canvas.create_window((0, 0), window=self.options_inner, anchor="nw")

        options = ["CDD Brasilia", "CDD Cuiabá", "CDD Gama", "CDD Goiânia", "CDD Manaus", "CDD Rondonópolis", "CDD AS Minas", "CDD Alfenas", "CDD Contagem ", "CDD Ipatinga ", "CDD Poços de Caldas", "CDD Pouso Alegre", "CDD Santa Luzia ", "CDD Sete Lagoas ", "CDD Uberaba", "CDD Uberlândia", "CDD Varginha", "CDD Aracaju ", "CDD Barreiras", "CDD Bom Jesus da Lapa", "CDD Cabo Santo Agostinho", "CDD Campina Grande", "CDD Caruaru ", "CDD Feira de Santana", "CDD Guanambi", "CDD Ilhéus", "CDD João Pessoa", "CDD Maceió", "CDD Olinda", "CDD Salvador", "CDD Vitória da Conquista", "CDD Belém ", "CDD Fortaleza", "CDD Imperatriz", "CDD Marabá", "CDD Natal", "CDD Cachoeiro", "CDD Campos", "CDD Itaperuna", "CDD Jacarepaguá", "CDD Niterói", "CDD Nova Friburgo", "CDD Nova Iguaçu", "CDD Pavuna", "CDD Petrópolis", "CDD São Cristóvão", "CDD Vitória", "CDD Volta Redonda", "CDD Agudos ", "CDD Araçatuba", "CDD Araraquara", "CDD Barueri", "CDD Campinas", "CDD Cubatão", "CDD Diadema", "CDD Embu das Artes", "CDD Guarulhos", "CDD Jaú", "CDD Jundiaí", "CDD Mauá", "CDD Mogi Mirim", "CDD Norte SP", "CDD Osasco", "CDD Piracicaba", "CDD Ribeirão Preto", "CDD São José dos Campos ", "CDD Suzano", "CDD Taubate", "CDL Salto", "CVV Itaquera", "PA Atibaia", "CDD Blumenau", "CDD Camboriú", "CDD Cascavel ", "CDD Caxias do Sul", "CDD Curitiba", "CDD Florianópolis", "CDD Foz do Iguaçu", "CDD Francisco Beltrão", "CDD Joinville ", "CDD Londrina ", "CDD Paranaguá", "CDD Pelotas", "CDD Porto Alegre - Eldorado do Sul", "CDD Santa Cruz do Sul", "CDD Santa Maria", "CDD São José dos Pinhais", "CDD Sapucaia do Sul", "GEO Curitiba", "PA Londrina"]  # Adicione suas opções aqui
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
        self.select_all_button.grid(row=5, column=0, padx=10, pady=5)

        self.deselect_all_button = tk.Button(root, text="Desmarcar Todas", command=self.deselect_all_options, width=20)
        self.deselect_all_button.grid(row=6, column=0, padx=10, pady=5)

        # Botão para iniciar o processo
        self.start_button = tk.Button(root, text="Iniciar", command=self.start_process, width=20)
        self.start_button.grid(row=7, columnspan=2, padx=10, pady=10)


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

    def start_process(self):
        # Coletando os valores inseridos pelo usuário
        global additional_info
        global start_time
        start_time = self.entry_start_time.get()
        end_time = self.entry_end_time.get()
        additional_info = self.text_info.get("1.0", tk.END)  # Obter texto da caixa de texto

        # Validando os horários inseridos
        if not self.validate_time(start_time):
            messagebox.showerror("Erro", "Hora inicial inválida. Por favor, insira no formato HH:MM.")
            return

        if not self.validate_time(end_time):
            messagebox.showerror("Erro", "Hora final inválida. Por favor, insira no formato HH:MM.")
            return

        global selected_options
        selected_options = [option for option, var in self.checkbox_options.items() if var.get()]


        # Aqui você pode realizar as operações do RPA com os dados coletados
        # Por enquanto, apenas exibiremos as informações

        #messagebox.showinfo("Informações", f"Hora Inicial: {start_time}\nHora Final: {end_time}\nInformações Adicionais:\n{additional_info}\nOpções Selecionadas: {selected_options}")
        #print(start_time)
        #print(end_time)
        #print(additional_info)
        #print(selected_options)
        
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
navegador.get('https://atendimento.governarti.com.br/Account/Login?ReturnUrl=%2f')
navegador.maximize_window()

# Realizar login
WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class='select2-choice']"))).click()
navegador.find_element(By.XPATH, "//*[@class='select2-choice']").send_keys(Keys.ARROW_DOWN)
time.sleep(1)
navegador.find_element(By.XPATH, "//*[@class='select2-choice']").send_keys(Keys.ENTER)

# Obtem e insere o Usuário
WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class='select2-choice']")))
navegador.find_element(By.XPATH, "//*[@class='form-control mv-input-radius p-5 username-input-login-service']").send_keys("fernanda.theiss")

# Obtem e insere a Senha
WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By. XPATH, "//*[@class='form-control mv-input-radius p-5 password-input-login-service']")))
navegador.find_element(By.XPATH, "//*[@class='form-control mv-input-radius p-5 password-input-login-service']").send_keys("teste@1")

# Obtem e clica no botão Entrar
WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By. XPATH, "//*[@id='btnSubmit']"))).click()

for options in selected_options:
    print(options)
    print(type(options))

    # Obtem e clica no Botão de novo Ticket
    WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class='button-more-icon-container']"))).click()

    # Obtem e clica no Botão de novo Ticket 2
    WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "(//*[@class='dropdown-item-text-movidesk'])[1]"))).click()

    # Obtem e clica no botão solicitante
    WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class='select2-container mv-input-select2-container clients']"))).click()
    time.sleep(5)
    
    # Insere o CDD
    pyautogui.typewrite(options, interval=0.3)
    time.sleep(5)
    navegador.find_element(By.XPATH, "//*[@class='select2-choice']").send_keys(Keys.ENTER)

    # Seleciona o tipo de chamado
    WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class='md-select-treeview-container md-single-select md-small']"))).click()
    time.sleep(5)
    pyautogui.typewrite("Servico de Monitoramento Indisponivel", interval=0.3)
    time.sleep(5)
    WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class='jqx-tree-item-li']"))).click()
    time.sleep(5)

    # Seleciona o tipo de ocorrência
    WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class='select2-container form-control category']"))).click()
    time.sleep(5)
    pyautogui.typewrite("Ocorrencias", interval=0.3)
    time.sleep(5)
    WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class='select2-match']"))).click()
    time.sleep(5)

    # Seleciona o tipo de urgência
    WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class='select2-container form-control urgency']"))).click()
    time.sleep(5)
    pyautogui.typewrite("Urgente", interval=0.3)
    time.sleep(5)
    WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class='select2-results-dept-0 select2-result select2-result-selectable']"))).click()
    time.sleep(5)

    # Seleciona e insere a data na caixa
    WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class='input-mv-new custom-fields hasDatepicker']"))).click()
    time.sleep(5)
    pyautogui.typewrite("02/02/2024", interval=0.3)
    time.sleep(5)
    navegador.find_element(By.XPATH, "//*[@class='input-mv-new custom-fields hasDatepicker']").send_keys(Keys.ENTER)

    # Seleciona e insere a hora 
    navegador.find_element(By.XPATH, "//*[@class='input-mv-new custom-fields hasDatepicker']").send_keys(Keys.TAB)
    time.sleep(5)
    pyautogui.typewrite(start_time, interval=0.3)

    # Seleciona e aplica a macro
    WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class='select2-container macro-button btn-apply-macro macro-select-container']"))).click()
    time.sleep(5)
    pyautogui.typewrite("Ocorrencias", interval=0.3)
    time.sleep(5)
    WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class='select2-results-dept-0 select2-result select2-result-selectable select2-highlighted']"))).click()
    time.sleep(5)

    #Seleciona e insere descrição
    WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@id='tinymce']/table/tbody/tr[5]/td/p[3]/span"))).click()
    time.sleep(5)
    pyautogui.typewrite(additional_info, interval=0.3)
    time.sleep(5)

    # Define a data final do preenchimento
    hora_atual = datetime.now().strftime("%H:%M")
    WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@name='Appointment[0].PeriodEnd']"))).click()
    time.sleep(5)
    pyautogui.typewrite(hora_atual, interval=0.3)
    navegador.find_element(By.XPATH, "//*[@name='Appointment[0].PeriodEnd']").send_keys(Keys.ENTER)

    # Seleciona e define a atividade
    WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@id='select2-chosen-136']"))).click()
    time.sleep(5)
    pyautogui.typewrite("Central Monitoramento", interval=0.3)
    time.sleep(5)
    WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class='select2-result-label']"))).click()
    time.sleep(5)

    # Obter e clicar no botão salvar
    WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[@class='right-group']"))).click()
    time.sleep(5)

    # Obter e clicar no botão continuar editando
    WebDriverWait(navegador, 30).until(expected_conditions.element_to_be_selected((By.XPATH, "//*[text()='Continuar editando']"))).click()
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

time.sleep(30)


