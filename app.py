from openpyxl import Workbook
import pandas as pd

status_mapping = {
    'awaitingPaymentConfirmation': 'Pagamento Confirmado',
    'customerRefused': 'Recusado Cliente',
    'pendingCustomer': 'Pendente Cliente',
    'canceled': 'Cancelada',
    'institutionRefused': 'Recusado Instituição',
    'pendingInstitution': 'Pendente Instituição'
}

def gerar_planilha():

    response = {
        "sucesso": True,
        "mensagem": "Propostas listadas com sucesso",
        "totalItems": 1417,
        "data": [
            {
                "customerName": "jose geraldo pereira ribeiro",
                "companyId": "56a241dc-8a7e-4402-83ee-9bb62a32fc46",
                "documentKey": "contracts/unico/contract-78166896672-1745437336437.pdf",
                "status": "pendingInstitution",
                "shouldGenerateCcb": True,
                "createdAt": "2025-04-23T19:31:56.708",
                "nextPayment": "2026-04-01T00:00:00",
                "updatedAt": "2025-04-23T19:42:29.767",
                "documentNumber": "78166896672",
                "contractURL": "https://aces.so/FZhQ9qc",
                "id": "UZ1E9D6708",
                "customer": {
                    "address": "Rua João Batista Figueiredo",
                    "occupation": "Assalariado",
                    "agency": "null",
                    "city": "Aracaju",
                    "accountNumber": "78166896672",
                    "cep": "49001156",
                    "uf": "SE",
                    "bank": "null",
                    "bankIspb": "null",
                    "bankCode": "null",
                    "phoneNumber": "38999667780",
                    "nationality": "Brasileiro",
                    "rg": "65481164",
                    "district": "Aruana",
                    "customerId": "b4c941b4-8bb9-4e9c-a788-c0f34b0d18c3",
                    "cpf": "78166896672",
                    "name": "JOSE GERALDO PEREIRA RIBEIRO",
                    "addressNumber": "2542",
                    "operationType": "0",
                    "email": "5332Munekazu40@hotmail.com",
                    "maritalStatus": "single"
                },
                "reservation": {
                    "interestRate": 0.0165,
                    "totalTransfer": 73.05,
                    "cetMonthly": 2.06,
                    "protocolBlock": "2433554148",
                    "reservationAmount": 160.63,
                    "numberOfPeriods": 10,
                    "subtotalTransfer": 85.89,
                    "totalAmount": 165.6,
                    "cet": 28.3,
                    "reservationId": "a67f0dd4-6d42-485a-91b8-87f4fc051578",
                    "protocolUnblock": "null",
                    "tariff": 10,
                    "hasADisregardedBalance": False,
                    "status": "ready",
                    "periods": [
                        {
                            "iof": 0.13,
                            "amount": 41.36,
                            "total": 29.2,
                            "data": "01/04/2026",
                            "subtotal": 34.3,
                            "fee": 7.06,
                            "reservationAmount": 41.36,
                            "id": "L4X4676708",
                            "iofFlat": 0.96,
                            "status": "pending"
                        },
                    ]
                },
                "insuranceRequired": False
            },
            {
                "customerName": "ana flavia fernandes",
                "companyId": "56a241dc-8a7e-4402-83ee-9bb62a32fc46",
                "documentKey": "contracts/unico/contract-09490855960-1745436395981.pdf",
                "status": "awaitingPaymentConfirmation",
                "shouldGenerateCcb": True,
                "createdAt": "2025-04-23T19:21:40.609",
                "nextPayment": "2025-12-01T00:00:00",
                "updatedAt": "2025-04-23T19:26:49.303",
                "documentNumber": "09490855960",
                "contractURL": "https://aces.so/qFHMlPN",
                "id": "5KAU7I0609",
                "customer": {
                    "address": "Rua João Brícola",
                    "occupation": "Assalariado",
                    "agency": "null",
                    "city": "São Paulo",
                    "accountNumber": "09490855960",
                    "cep": "01014917",
                    "uf": "SP",
                    "bank": "null",
                    "bankIspb": "null",
                    "bankCode": "null",
                    "phoneNumber": "41995658742",
                    "nationality": "Brasileiro",
                    "rg": "311980165",
                    "district": "Centro",
                    "customerId": "85994d4a-8913-4ae7-bd77-00fa4e96f3af",
                    "cpf": "09490855960",
                    "name": "ANA FLAVIA FERNANDES",
                    "addressNumber": "1526",
                    "operationType": "0",
                    "email": "2156Eliderson.Franco@gmail.com",
                    "maritalStatus": "single"
                },
                "reservation": {
                    "interestRate": 0.0165,
                    "totalTransfer": 98.59,
                    "cetMonthly": 2.31,
                    "protocolBlock": "2433514458",
                    "reservationAmount": 159.08,
                    "numberOfPeriods": 10,
                    "subtotalTransfer": 111.84,
                    "totalAmount": 171.05,
                    "cet": 32.22,
                    "reservationId": "8cc0b937-ac76-4692-94af-d9a7bfc8d92c",
                    "protocolUnblock": "null",
                    "tariff": 10,
                    "hasADisregardedBalance": False,
                    "status": "ready",
                    "periods": [
                        {
                            "iof": 0.17,
                            "amount": 50.36,
                            "total": 39.62,
                            "data": "01/12/2025",
                            "subtotal": 44.62,
                            "fee": 5.74,
                            "reservationAmount": 50.36,
                            "id": "6T8OTM0609",
                            "iofFlat": 0.81,
                            "status": "pending"
                        },
                    ]
                },
                "insuranceRequired": False
            }
        ],
        "lastKey": "null",
        "limit": 2000
    }
    # Dicionário para tradução dos status
    status_mapping = {
        'awaitingPaymentConfirmation': 'Pagamento Confirmado',
        'customerRefused': 'Recusado Cliente',
        'pendingCustomer': 'Pendente Cliente',
        'pendingInstitution': 'Pendente Instituição',
        'canceled': 'Cancelado',
        'institutionRefused': 'Recusado Instituição'
    }

    data = []

    for item in response['data']:
        customer_name = item['customerName']
        original_status = item['status']  # Guardamos o status original para uso nas somas
        # Aplica a tradução do status ou mantém o original se não encontrar no dicionário
        translated_status = status_mapping.get(original_status, original_status)
        cpf = item['customer']['cpf']
        reservation_amount = item['reservation']['reservationAmount']

        data.append({
            'CPF': cpf,
            'Nome': customer_name,
            'Status': translated_status,
            'Valor da Reserva': reservation_amount,
            'Status_Original': original_status  # Mantém o status original para uso nas somas
        })

    # Criar DataFrame com os dados coletados, excluindo a coluna de controle
    df = pd.DataFrame(data)
    df = df.drop('Status_Original', axis=1)  # Remove a coluna de status original do DataFrame
    
    # Calcular os totais com base nas regras de status
    total_reservation_amount = sum(
        item['Valor da Reserva'] 
        for item in data 
        if item['Status_Original'] in ['awaitingPaymentConfirmation', 'pendingCustomer']
    )

    total_esteira = sum(
        item['Valor da Reserva'] 
        for item in data 
        if item['Status_Original'] in ['awaitingPaymentConfirmation', 'pendingCustomer', 'pendingInstitution']
    )

    # Criar o arquivo Excel
    try:
        with pd.ExcelWriter('planilha.xlsx', engine='xlsxwriter') as writer:
            # Criar um DataFrame para o resumo
            summary_df = pd.DataFrame([
                ['Total Reservas', f'R$ {total_reservation_amount:.2f}'],
                ['Total Esteira', f'R$ {total_esteira:.2f}']
            ], columns=['Descrição', 'Valor'])
            
            # Adicionar o resumo na planilha (com index=False para remover a coluna de índice)
            summary_df.to_excel(writer, sheet_name='Propostas', index=False, startrow=0)
            
            # Adicionar uma linha em branco
            writer.sheets['Propostas'].write_row(3, 0, [''])
            
            # Adicionar os dados principais (com index=False para remover a coluna de índice)
            df.to_excel(writer, sheet_name='Propostas', index=False, startrow=5)
            
            # Formatar a planilha
            workbook = writer.book
            worksheet = writer.sheets['Propostas']
            
            # Formatar cabeçalhos
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#D9E1F2',
                'border': 1
            })
            
            # Aplicar formato ao cabeçalho
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(5, col_num, value, header_format)
                
            # Ajustar largura das colunas
            for i, col in enumerate(df.columns):
                column_width = max(df[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(i, i, column_width)
                
        print("Planilha gerada com sucesso!")
        
    except Exception as e:
        print(f"Erro ao gerar a planilha: {str(e)}")

if __name__ == '__main__':
    gerar_planilha()
