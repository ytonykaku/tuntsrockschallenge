import os.path
import math
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

def calculate_final_grade(m, naf):
    return math.ceil((m + naf) / 2)

#Função para resgatar a quantidade de aulas, excluindo o trecho digitado em texto antes da quantidade de aulas
def get_total_classes(worksheet):
    total_classes_cell = worksheet[1][0]

    # Extract the number of classes
    total_classes_text = "Total de aulas no semestre: "
    if total_classes_cell.startswith(total_classes_text):
        total_classes_str = total_classes_cell[len(total_classes_text):].strip()
        return int(total_classes_str)
    else:
        raise ValueError("Invalid format for total classes cell")

#Função que atualiza os dados de acordo com o resultado dos alunos
def update_worksheet(worksheet, row, result, result_or_grade, naf=0):
    if result_or_grade == True:
        range_column = 'G'
    else:
        range_column = 'H'
    try:
        student_row = row
        range_name = f'engenharia_de_software!{range_column}{student_row}'
        update_values_request = {
            'values': [[result]],  # Valor a ser atualizado
            'range': range_name,
            'majorDimension': 'ROWS',
        }

        response = worksheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=range_name, body=update_values_request,
                                         valueInputOption='RAW').execute()

        print('Valor atualizado com sucesso!')

    except Exception as e:
        print(f'Error: {e}')

#Função que avalia a qualificação dos alunos
def evaluate_student(row, data, total_classes, worksheet):
    student_id = int(data[row][0])
    p1 = float(data[row][3])  # Coluna D para P1
    p2 = float(data[row][4])  # Coluna E para P2
    p3 = float(data[row][5])  # Coluna F para P3
    attendance = int(data[row][2])  # Coluna C para Faltas

    total_grades = 3
    total_attendance = total_classes * 0.25

    average_grade = (p1 + p2 + p3) / total_grades

    #Avaliação da frequência
    if attendance > total_attendance:
        update_worksheet(worksheet, row + 1, "Reprovado por Falta", True)
        return

    #Avaliação das notas
    if average_grade < 50:
        update_worksheet(worksheet, row + 1, "Reprovado por Nota", True) #Reprovado
    elif 50 <= average_grade < 70:
        # Calculate "Nota para Aprovação Final" for Exame Final
        naf = max(0, 100 - average_grade)
        update_worksheet(worksheet, row + 1, "Exame Final", True)
        update_worksheet(worksheet, row + 1, naf, False)    #Exame final
    else:
        update_worksheet(worksheet, row + 1, "Aprovado", True)  #Aprovado

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

SAMPLE_SPREADSHEET_ID = "1S2wPCVcbNIvJzcdNMeOa-9bLjS95q-bw0jGVVaIID6E"

def main():
  creds = None
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    with open("token.json", "w") as token:
      token.write(creds.to_json())

  try:
    service = build("sheets", "v4", credentials=creds)

    # Chamada da API
    sheet = service.spreadsheets()
    result = (
        sheet.values()
        .get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="engenharia_de_software!A1:H27")
        .execute()
    )

    worksheet = result.get("values", [])

    total_classes = get_total_classes(worksheet)

    #Começando a partir da linha que está o primeiro aluno
    for i in range(3, len(worksheet)):
        evaluate_student(i, worksheet, total_classes, sheet)
  except HttpError as err:
    print(err)

if __name__ == "__main__":
    main()