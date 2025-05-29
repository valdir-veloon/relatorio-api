import requests
import pandas as pd
import time
from datetime import datetime
from zoneinfo import ZoneInfo
from dotenv import load_dotenv
from datetime import timedelta
import os

load_dotenv()

api_token = os.getenv("API_TOKEN")

status_mapping = {
    'awaitingPaymentConfirmation': 'Pagamento Confirmado',
    'customerRefused': 'Recusado Cliente',
    'pendingCustomer': 'Pendente Cliente',
    'canceled': 'Cancelada',    
    'institutionRefused': 'Recusado Instituição',
    'pendingInstitution': 'Pendente Instituição',
    'unblockingError': 'Erro de Desbloqueio',
}

PAGE_SIZE = 1500  # Número de itens por página
MAX_ITEMS = 1500 * 5 # Limite de segurança
STARTDATE = "2025-05-01"  # Data de início para a coleta
ENDDATE = "2025-05-31"  # Data de fim para a coleta

def get_api_token():
    """
    Função para obter o token da API.
    """
    try:
        base_url = "https://api.tmjbeneficios.com.br/no-auth/authentication/login"
        headers = {
            "Content-Type": "application/json"
        }
        data = {
            "email": os.getenv("API_EMAIL"),
            "senha": os.getenv("API_PASSWORD")   
        }

        response = requests.post(base_url, headers=headers, json=data)
        
        if not response.ok:
            print('Erro ao chamar a API:', response.text)
            return None
        
        result = response.json()
        return result.get("accessToken")

    except Exception as e:
        print(f"Erro ao obter o token da API: {str(e)}")
        return None

def gerar_planilha():

    # Configuração da API
    base_url = "https://api.tmjbeneficios.com.br/propostas/fgts/listar"
    headers = {
        "Authorization": f"Bearer {get_api_token()}",
        "Content-Type": "application/json"
    }
    
    all_data = []
    processed_ids = set()
    last_key = None

    # Definindo o intervalo de dias
    intervalo_dias = 2
    data_inicio = datetime.strptime(STARTDATE, "%Y-%m-%d")
    data_fim = datetime.strptime(ENDDATE, "%Y-%m-%d")
    
    print("Iniciando coleta de dados da API...")

    # Loop de paginação
    while data_inicio <= data_fim:
        data_intervalo_fim = min(data_inicio + timedelta(days=intervalo_dias - 1), data_fim)
        last_key = None

        print(f"Buscando de {data_inicio.strftime('%Y-%m-%d')} até {data_intervalo_fim.strftime('%Y-%m-%d')}")

        while True:
            params = {
                "limit": PAGE_SIZE,
                "startDate": data_inicio.strftime("%Y-%m-%d"),
                "endDate": data_intervalo_fim.strftime("%Y-%m-%d")
            }
            if last_key:
                params["lastKey"] = last_key

            try:
                response = requests.get(base_url, headers=headers, params=params)
                response.raise_for_status()
                api_data = response.json()

                if not api_data.get("data") or len(api_data["data"]) == 0:
                    break

                for item in api_data["data"]:
                    codigo_operacao = item.get('id', 'N/A')
                    contract_url = item.get('contractURL', 'N/A')
                    phone_number = item.get('customer', {}).get('phoneNumber', 'N/A')
                    number_periods = item.get('reservation', {}).get('numberOfPeriods', 0)

                    if codigo_operacao in processed_ids:
                        continue
                    processed_ids.add(codigo_operacao)
                    # ...restante do processamento...
                    customer_name = item.get('customerName', 'N/A')
                    created_at = item.get('createdAt', 'N/A')
                    try:
                        dt_utc = datetime.fromisoformat(created_at).replace(tzinfo=ZoneInfo("UTC"))
                        dt_sp = dt_utc.astimezone(ZoneInfo("America/Sao_Paulo"))
                        created_at_br = dt_sp.strftime('%d/%m/%Y %H:%M:%S')
                    except Exception:
                        continue
                    original_status = item.get('status', 'N/A')
                    translated_status = status_mapping.get(original_status, 'não encontrado')
                    cpf = item.get('customer', {}).get('cpf', 'N/A')
                    reservation_amount = item.get('reservation', {}).get('reservationAmount', 0)
                    reservation_id = item.get('reservation', {}).get('reservationId', 'N/A')
                    all_data.append({
                        'Código da Operação': codigo_operacao,
                        'CPF': cpf,
                        'Nome': customer_name,
                        'Status': translated_status,
                        'Valor da Reserva': reservation_amount,
                        'Status_Original': original_status,
                        'Reservation ID': reservation_id,
                        'Data': created_at_br,
                        'Phone': phone_number,
                        'contractUrl': contract_url,
                        'originalStatus': original_status,
                        'numberPeriods': number_periods
                    })

                last_key = api_data.get("lastKey")
                if not last_key or last_key == "null":
                    break
                time.sleep(0.8)
            except Exception as e:
                print(f"Erro na requisição à API: {str(e)}")
                break

        data_inicio = data_intervalo_fim + timedelta(days=1)

    print(f"Total de {len(all_data)} itens coletados.")

    if not all_data:
        print("Nenhum dado foi coletado. Verifique a conexão com a API.")
        return

    df = pd.DataFrame(all_data)
    df[['Código da Operação', 'Valor da Reserva']].to_csv('ids_planilha.csv', index=False)
    df = df.drop('Status_Original', axis=1)

    # Criar o arquivo Excel
    try:
        with pd.ExcelWriter('planilha_propostas.xlsx', engine='xlsxwriter') as writer:
            
            # Adicionar os dados principais
            df.to_excel(writer, sheet_name='Propostas', index=False, startrow=0)
            
            # Formatar a planilha
            worksheet = writer.sheets['Propostas']
                
            # Ajustar largura das colunas
            for i, col in enumerate(df.columns):
                column_width = max(df[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(i, i, column_width)
                
        print("Planilha gerada com sucesso!")
        
    except Exception as e:
        print(f"Erro ao gerar a planilha: {str(e)}")

if __name__ == '__main__':
    gerar_planilha()
