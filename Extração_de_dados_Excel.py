import xlwings as xw
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def ler_dados_da_coluna(file_path, coluna, linha):
    # Abrir o arquivo Excel
    wb = xw.Book(file_path)
    
    # Selecionar a planilha ativa
    planilha = wb.sheets.active
    
    # Ler os dados da coluna 
    dado = planilha.range((linha, coluna)).value
    
    # Retorna os dados
    return dado

def calcular_coluna_mes_atual():
    # Obtém o mês atual
    mes_atual = datetime.now().month
    mes_atual = mes_atual - 1
    
    # Calcula a coluna com base no mês atual
    validacaoColuna = (mes_atual - 1) * 2 + 9
    
    # Se o mês for dezembro, retorna a última coluna possível (23)
    if mes_atual == 12:
        validacaoColuna = 23
    elif mes_atual > 12:
        validacaoColuna = None  # Retorna None para meses inválidos após dezembro
    
    return validacaoColuna

# Obtém a coluna do mês atual
coluna_atual = calcular_coluna_mes_atual()

if coluna_atual is not None:
    print(f"A coluna correspondente ao mês atual é: {coluna_atual}")
else:
    print("Mês inválido após dezembro.")

# Caminho do arquivo
file_path = r'caminho do arquivo .xlsx'

# Variáveis das colunas
colunaEmpresa = 1
colunaUnidade = 2
colunaTipo = 4
colunaAtividade = 6
colunaPrazo = 7
coluna = coluna_atual

# Linha inicial
linha_inicial = 12

data_atual = datetime.now().date()

# Obtém a última linha com dados na coluna O da planilha 'IMPOSTOS-OBRIGAÇÕES'
wb = xw.Book(file_path)
ultima_linha = wb.sheets['IMPOSTOS-OBRIGAÇÕES']['A12'].end('down').row

contador = 0
dadosRecebidos = ""
corpo_email = f"""<html><body><p>Bom dia!</p><br>"""

# Loop sobre as linhas desejadas e imprimir os dados das colunas
for linha in range(linha_inicial, ultima_linha + 1):
    data = ler_dados_da_coluna(file_path, coluna, linha)
    empresa = ler_dados_da_coluna(file_path, colunaEmpresa, linha)
    unidade = ler_dados_da_coluna(file_path, colunaUnidade, linha)
    tipo = ler_dados_da_coluna(file_path, colunaTipo, linha)
    atividade = ler_dados_da_coluna(file_path, colunaAtividade, linha)
    prazo = ler_dados_da_coluna(file_path, colunaPrazo, linha)
    status = ler_dados_da_coluna(file_path, coluna + 1, linha)

    if isinstance(data, datetime) and status == 'Não Iniciado':
        # Converte a data para o formato 'dd/MM/yyyy'
        data_formatada = data.date()
        # Calcula a diferença em dias entre a data atual e a data lida
        diferenca_dias = (data_formatada - data_atual).days
        # Verifica se a data é igual ao dia atual ou se é 1 ou 2 dias antes da data atual
        if diferenca_dias in (0, 1, 2):
            dadosRecebidos = f"Empresa: {empresa} | Unidade: {unidade} | Tipo: {tipo} | Atividade: {atividade} | Prazo: {prazo} | Data: {data_formatada.strftime('%d/%m/%Y')} | Status: {status}"
            print(dadosRecebidos)
            corpo_email += f"""<p>{dadosRecebidos}</p>"""

# Validação se possui dados para enviar no E-mail
if not dadosRecebidos:
    corpo_email += f"""</body>Não há vencimentos para a data de hoje</html>"""
else:
    corpo_email += f"""</body></html>"""

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
mensagem['Subject'] = f"CHECK-LIST FISCAL - OBRIGAÇÕES ACESSÓRIAS E IMPOSTOS {datetime.now().date().strftime("%d/%m/%Y")}" # Assunto

corpo_email += f"""</body></html>"""

mensagem.attach(MIMEText(corpo_email, 'html'))
try:
    servidor_email.sendmail(email, mensagem['To'].split(", "), mensagem.as_string())
except (smtplib.SMTPServerDisconnected, smtplib.SMTPException) as e:
    print(f"Erro ao enviar e-mail: {e}")
    # Reconectar ao servidor SMTP
    servidor_email = conexaoSMTP(email, senha)
    servidor_email.sendmail(email, mensagem['To'].split(", "), mensagem.as_string())
servidor_email.quit()

# Fecha o Excel
wb.close()