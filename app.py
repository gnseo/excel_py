import sys
sys.path.insert(0, './pip')
sys.path.insert(0, './pip_openpyxl')

import xlsxwriter
import json
import base64

from openpyxl import load_workbook

def handler(event, context):

  saveToPath = "/tmp/"

  wb = load_workbook('ON182225.xlsx')
  ws = wb.active
  ws['A6'] = "fjfj"
  wb.save(saveToPath+'test.xlsx')

  excel_file = ""
  with open(saveToPath+'test.xlsx','rb') as f:
    excel_file = f.read()
    excel_file = base64.b64encode(excel_file).decode("utf-8")

  return {
      "statusCode": 200,
      'isBase64Encoded': True,
      'headers': {
          'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          'Access-Control-Allow-Origin': '*',
          'Content-Disposition': 'attachment; filename="test.xlsx"'
      },
      "body": json.dumps(excel_file)#["queryStringParameters"]
      }

def handler_xlsxwriter(event, context):
  print(event)

  saveToPath = "/tmp/"

  workbook = xlsxwriter.Workbook(saveToPath+'hello.xlsx')
  worksheet = workbook.add_worksheet()

  data = [
      ['Apples', 10000, 5000, 8000, 6000],
      ['Pears',   2000, 3000, 4000, 5000],
      ['Bananas', 6000, 6000, 6500, 6000],
      ['Oranges',  500,  300,  200,  700],

  ]

  worksheet.add_table('B3:F7', {'data': data})

  workbook.close()

  excel_file = ""
  with open(saveToPath+'hello.xlsx','rb') as f:
    excel_file = f.read()
    excel_file = base64.b64encode(excel_file).decode("utf-8")

  return {
      "statusCode": 200,
      'isBase64Encoded': True,
      'headers': {
          'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          'Access-Control-Allow-Origin': '*',
          'Content-Disposition': 'attachment; filename="hello.xlsx"'
      },
      "body": json.dumps(excel_file)#["queryStringParameters"]
      }
