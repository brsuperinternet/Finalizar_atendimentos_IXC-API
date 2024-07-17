import requests
import base64
import json  
import openpyxl
from datetime import datetime 
import time

id_a_fechar = []
data_hora_atual = datetime.now().isoformat()
planilha = openpyxl.load_workbook('os.xlsx')   # faz o dowload da planilha Excel com os atendimentos 
planilha_ativa = planilha.active
response_count = 0  

# Define o Host, e a url onde será feito a requisição e o Token de autenticação
host = "provernet.com.br"
url = f"https://{host}/webservice/v1/su_oss_chamado_fechar"
token = "246:31c337437fb3431e4376e69f62c98960fccc9b90509c2da096106bf3e6f5256a".encode('utf-8')

# Ao iniciar o código printa a hora atual para visualização
print(f"Horario da inicialização {data_hora_atual}/n")
print("-"*50)

# Extrai o valor id da planilha e encaminha todos Ids para uma lista
for linha in planilha_ativa.iter_rows(min_row=2,values_only=True):
    if linha[0] is not None:
        id_a_fechar.append(linha[0])


print(id_a_fechar)

# Roda um for construindo realizando uma requisição para cada ID do Array a
for id in id_a_fechar:
    payload = {
        "id_chamado": id,
        "data_inicio": data_hora_atual,
        "data_final": data_hora_atual,
        "mensagem":"O.S Finalizada por API",
        "id_tecnico": "301", # 301 Mateus Felipe Gonçalves
        "gera_comissao":"S",
        "status":"F",
        "data": data_hora_atual,
        "finaliza_processo":"S",
        "id_su_diagnostico":"268",
    }

    # Define o cabeçalho de minha requisição
    headers = {
        'Authorization': 'Basic {}'.format(base64.b64encode(token).decode('utf-8')),
        'Content-Type': 'application/json'
    }
    

    # Realiza uma tratativa de erro para o código não parar quando der algum erro na requisição 
    while True:
        try:
            response = requests.post(url, data=json.dumps(payload), headers=headers, timeout=30)  # Timeout=  aumenta o tempo limite para 30 segundos
            if response.status_code == 200:  
                response_count += 1

                # A cada 1000 requisições da um pause de 5 minutos
                if response_count == 1000:
                    response_count = 0
                    print("1000 O.S finalizados, realizando pausa de 5 minutos")
                    time.sleep(300)
                break
            # Se a resposta não for 200 da um pause de 5 minutos e tenta novamente

            else:
                print('Resposta da API não foi 200, realizando pausa de 5 minutos')
                time.sleep(300)

        # Faz a tratativa do erro que mais ocorria normalmente.
        except requests.exceptions.Timeout:
            print('Tempo limite de conexão excedido, realizando pausa de 5 minutos')
            time.sleep(300)
        except requests.exceptions.RequestException as e:
            print('Erro ao fazer requisição:', e)
            time.sleep(300)

print(response.text)
print("-"*50)
print(f"\nHorario da finalização ", datetime.now().isoformat())
print('\nComando enviado com SUCESSO, FINALIZADO!')