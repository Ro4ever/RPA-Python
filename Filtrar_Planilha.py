import pandas as pd
import glob
import os
from datetime import datetime, timedelta

# Abre o arquivo pegando apenas pelo início, ignorando os números após o underline
directory_path = 'Caminho do Arquivo.xlsx'

# Encontra o arquivo filtrando apenas pelo início, por causa dos números aleatórios após o underline
file_pattern = os.path.join(directory_path, 'Ex: roberto_79873218946398.xlsx')
file_list = glob.glob(file_pattern)

# Seleciona o arquivo mais recente dentro da pasta
latest_file = max(file_list, key=os.path.getctime)

df = pd.read_excel(latest_file, header=8)

# Print das colunas
print("Nomes das colunas no DataFrame:")
print(df.columns)

# Validação se está na coluna correta pegando os dados
date_column = 'Coluna que desejar' 

# Verifica e remove os espaços em brancos
df.columns = df.columns.str.strip()

# Conver os dados da coluna e converte para o padrão de data dd/MM/yyyy
df[date_column] = pd.to_datetime(df[date_column], format='%d/%m/%Y')

# Pega a data atual -7 dias
seven_days_ago = datetime.now() - timedelta(days=7)

# Realiza os filtros, mantendo somente os dados de 7 dias atrás a partir da data atual
filtered_df = df[df[date_column] >= seven_days_ago]

# Exibe o DataFrame filtrado
print("DataFrame filtrado para os últimos 7 dias:")
print(filtered_df)

# Formata o arquivo com a data atual
current_date_str = datetime.now().strftime('%d-%m-%Y')
output_file_path = fr'Salva o arquivo com a data atual Ex:C:\user\desktop\{current_date_str}.xlsx'

# Cria o diretório se não existir
os.makedirs(os.path.dirname(output_file_path), exist_ok=True)

# Salva o arquivo
filtered_df.to_excel(output_file_path, index=False)
print(f"DataFrame filtrado salvo em: {output_file_path}")
