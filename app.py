import requests
import pandas as pd
import time
from datetime import datetime
from zoneinfo import ZoneInfo
from dotenv import load_dotenv
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

PAGE_SIZE = 2000  # Número de itens por página
MAX_ITEMS = 2000  # Limite de segurança
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
    last_key = None
    total_items_processed = 0
    
    print("Iniciando coleta de dados da API...")

    # Loop de paginação
    while total_items_processed < MAX_ITEMS:
        # Preparar parâmetros de consulta
        params = {"limit": PAGE_SIZE, "startDate": STARTDATE, "endDate": ENDDATE}
        if last_key:
            params["lastKey"] = last_key
            
        try:
            # Fazer requisição à API
            response = requests.get(base_url, headers=headers, params=params)
            response.raise_for_status()  # Lança exceção para códigos de erro HTTP
            
            # Parsear resposta
            api_data = response.json()
            
            if not api_data.get("data") or len(api_data["data"]) == 0:
                print("Não há mais dados para processar.")
                break
                
            # Processar dados da página atual
            for item in api_data["data"]:
                try:
                    customer_name = item.get('customerName', 'N/A')

                    # Ajustando as data para o formato BR
                    created_at = item.get('createdAt', 'N/A')
                    try:
                        dt_utc = datetime.fromisoformat(created_at).replace(tzinfo=ZoneInfo("UTC"))
                        dt_sp = dt_utc.astimezone(ZoneInfo("America/Sao_Paulo"))
                        created_at_br = dt_sp.strftime('%d/%m/%Y')
                    except Exception:
                        continue
                    
                    original_status = item.get('status', 'N/A')
                    translated_status = status_mapping.get(original_status, 'não encontrado')
                    
                    # Verificar se customer e reservation existem antes de acessar
                    cpf = item.get('customer', {}).get('cpf', 'N/A')
                    reservation_amount = item.get('reservation', {}).get('reservationAmount', 0)
                    reservation_id = item.get('reservation', {}).get('reservationId', 'N/A')
                    
                    all_data.append({
                        'CPF': cpf,
                        'Nome': customer_name,
                        'Status': translated_status,
                        'Valor da Reserva': reservation_amount,
                        'Status_Original': original_status,
                        'Reservation ID': reservation_id,
                        'Data': created_at_br,
                    })
                except Exception as e:
                    print(f"Erro ao processar item: {str(e)}")
                    continue
            
            # Atualizar contador e chave para próxima página
            total_items_processed += len(api_data["data"])
            last_key = api_data.get("lastKey")
            
            print(f"Processados {total_items_processed} itens até o momento.")
            
            # Se não houver próxima página, encerrar o loop
            if last_key == "null" or last_key is None:
                break
                
            # Pequena pausa para não sobrecarregar a API
            time.sleep(0.5)
            
        except Exception as e:
            print(f"Erro na requisição à API: {str(e)}")
            break
    
    print(f"Total de {total_items_processed} itens coletados.")
    
    if not all_data:
        print("Nenhum dado foi coletado. Verifique a conexão com a API.")
        return
        
    # Criar DataFrame com os dados coletados, excluindo a coluna de controle
    df = pd.DataFrame(all_data)
    df = df.drop('Status_Original', axis=1)  # Remove a coluna de status original do DataFrame

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
