from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials

# Укажите путь к вашему credentials.json
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Идентификатор таблицы Google Sheets

def create_sheets():
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)
    
    requests = []
    for i in range(22, 31):
        sheet_name = f'День {i}'
        requests.append({
            "addSheet": {
                "properties": {
                    "title": sheet_name
                }
            }
        })
    
    body = {
        'requests': requests
    }
    service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body).execute()
    print("Листы успешно созданы!")

if __name__ == '__main__':
    create_sheets()
