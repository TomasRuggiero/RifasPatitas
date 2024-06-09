import PyPDF2
import gspread
import pandas as pd
import pytesseract
from PIL import Image
from pydrive.drive import GoogleDrive
from oauth2client.client import GoogleCredentials
from pydrive.auth import GoogleAuth
from gspread_formatting import *

from google.colab import auth
auth.authenticate_user()

from google.auth import default
creds, _ = default()

gc = gspread.authorize(creds)

from google.colab import drive
drive.mount('/content/drive')

gauth = GoogleAuth()
gauth.credentials = GoogleCredentials.get_application_default()
drive = GoogleDrive(gauth)

# Open our new sheet and add some data.
spreadsheet = gc.open("Rifas Patitas").sheet1

folder_id = '1wkq_zmwfWyAupHsccxlm8heFOV7NoQ_r'
spreadsheet_copy = drive.CreateFile({'title': 'Lista de rifas', 'mimeType': 'application/vnd.google-apps.spreadsheet', 'parents': [{'id': folder_id}]})
spreadsheet_copy.Upload()

# Get the new file ID of the copied spreadsheet
spreadsheet_copy_id = spreadsheet_copy['id']
print(spreadsheet_copy_id)

# Open the copied spreadsheet using gspread
spreadsheet_copy = gc.open_by_key(spreadsheet_copy_id)

worksheet = spreadsheet_copy.get_worksheet(0)

# Now you can work with the 'worksheet' object, for example, fetching the data
data = spreadsheet.get_all_values()
print(data)

rowCount = 0

for index, datos in enumerate(data):
  if(index > 1):
    fileId = data[index][1].split("/")[5]
    # Taking the file from google drive

    meta = drive.CreateFile({'id': fileId})
    meta.FetchMetadata(fields='title')

    fileName = meta['title']
    print(fileName)

    archivo = drive.CreateFile({'id': fileId})
    archivo.GetContentFile(fileName)

    # Caso PDF
    if(fileName.lower().endswith('.pdf')):
      reader = PyPDF2.PdfReader(fileName)

      # print the text of the first page
      try:
        arr = reader.pages[0].extract_text().split('\n')
        price = float(arr[5].replace('$','').replace('.',''))
        precioPorVoto = 500

        cantidadDeRifas = price / precioPorVoto

        if(worksheet):
          try:
            while(cantidadDeRifas > 0):
              worksheet.append_row(data[index])
              rowCount += 1
              cantidadDeRifas -= 1
            print("Data added successfully.")
          except Exception as e:
            print(f"An error occurred: {e}")

      except:
        worksheet.append_row(data[index])
        rowCount += 1
        fmt = CellFormat(backgroundColor=Color(1, 0, 0))
        format_cell_range(worksheet, f'A{worksheet.row_count}:B{worksheet.row_count}', fmt)
        print("Hubo un error al procesar la cantidad de rifas de " + data[index][0])

    # Caso Imagen
    elif(fileName.lower().endswith('.jpg') or fileName.lower().endswith('.png')):

      try:
        image = Image.open(fileName)

        pytesseract.pytesseract.tesseract_cmd = '/usr/bin/tesseract'

        text = pytesseract.image_to_string(image)
        pesos = int(text.split("$")[1].strip().split(",")[0].replace(".", "").split()[0])

        valorPorRifa = 500

        cantidadDeRifas = pesos/valorPorRifa

        if(worksheet):
          try:
            while(cantidadDeRifas > 0):
              worksheet.append_row(data[index])
              rowCount += 1
              cantidadDeRifas -= 1
            print("Data added successfully.")
          except Exception as e:
            print(f"An error occurred: {e}")

      except:
        worksheet.append_row(data[index])
        rowCount += 1
        fmt = CellFormat(backgroundColor=Color(1, 0, 0))
        format_cell_range(worksheet, f'A{rowCount}:B{rowCount}', fmt)
        print("Hubo un error al procesar la cantidad de rifas de " + data[index][0])


