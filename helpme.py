import gspread
from google.oauth2.service_account import Credentials

def main():
    # Carregar credenciais do arquivo JSON
    credentials = Credentials.from_service_account_file('credentials.json', scopes=['https://www.googleapis.com/auth/spreadsheets'])

    # Autorizar com as credenciais
    gc = gspread.authorize(credentials)

    # Abra a planilha pelo seu ID ou URL
    spreadsheet = gc.open_by_key('1S2wPCVcbNIvJzcdNMeOa-9bLjS95q-bw0jGVVaIID6E')
    # OU
    # spreadsheet = gc.open_by_url('your_spreadsheet_url')

    # Selecione a folha pelo nome ou índice
    worksheet = spreadsheet.get_worksheet(0)
    # OU
    # worksheet = spreadsheet.get_worksheet_by_title('Sheet1')

    # Agora você pode manipular a planilha conforme necessário
    cell_value = worksheet.cell(1, 1).value
    print(f'Conteúdo da célula (1, 1): {cell_value}')

if __name__ == '__main__':
    main()
