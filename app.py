import sys
sys.path.insert(0, './pip')
sys.path.insert(0, './pip_openpyxl')

import xlsxwriter
import json
import base64
import boto3

from openpyxl import load_workbook
s3 = boto3.resource("s3")

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
  q = getQuery(event)

  return_headers = { "Access-Control-Allow-Origin": "*", "Access-Control-Expose-Headers": "" }

  files_suffix = q.get("files_suffix", "")
  file_name = q.get("file_name", None)
  columns = q.get("columns", None)
  arrdata = q.get("arrdata", None)
  if arrdata:
    if file_name is None:
      return {
        'statusCode': 500,
        'headers': return_headers,
        'body': json.dumps({"errorMessage": "file_name is required"}),
      }
    if columns is None:
      return {
        'statusCode': 500,
        'headers': return_headers,
        'body': json.dumps({"errorMessage": "columns is required"}),
      }
  else:
    return {
      'statusCode': 500,
      'headers': return_headers,
      'body': json.dumps({"errorMessage": "arrdata is required"}),
    }

  saveToPath = "/tmp/"
  name_with_suffix = "{0}_{1}".format(file_name,files_suffix)
  workbook = xlsxwriter.Workbook(saveToPath+'{}.xlsx'.format(name_with_suffix))
  worksheet = workbook.add_worksheet()

  last_row = len(arrdata)
  last_col = len(columns) - 1
  worksheet.add_table(0,0,last_row,last_col,{'data': arrdata,'columns':columns})
  #worksheet.add_table(0,0,1,2,{'data': [[1,2]],'columns':[{"header":"Column 1"},{"header":"Column 2"},{"header":"Column 3"}]})

  workbook.close()

  result = {"url": upload_to_s3(name_with_suffix,"jenax/pcr/files/")}

  return {
    'statusCode': 200,
    'headers': return_headers,
    'body': json.dumps(result),
  }

  # excel_file = ""
  # with open(saveToPath+'hello.xlsx','rb') as f:
  #   excel_file = f.read()
  #   excel_file = base64.b64encode(excel_file).decode("utf-8")
  #
  # return {
  #     "statusCode": 200,
  #     'isBase64Encoded': True,
  #     'headers': {
  #         'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  #         'Access-Control-Allow-Origin': '*',
  #         'Content-Disposition': 'attachment; filename="hello.xlsx"'
  #     },
  #     "body": json.dumps(excel_file)#["queryStringParameters"]
  #     }

def upload_to_s3(filename,key_prefix=""):

  s3_bucket_name = "bsg-static-files"
  temp_filename = "{}.xlsx".format(filename)
  key_name = "{}{}".format(key_prefix,temp_filename)

  #with open('/tmp/{}'.format(temp_filename), 'rb') as csvfile:
  #  res = s3.Bucket(s3_bucket_name).put_object(ACL="public-read", Key="{}".format(key_name), Body=csvfile)
  s3.meta.client.upload_file('/tmp/{}'.format(temp_filename), s3_bucket_name, key_name)

  bucket_location = boto3.client('s3').get_bucket_location(Bucket=s3_bucket_name)
  #object_url = "https://s3-{0}.amazonaws.com/{1}/{2}".format(
  #  bucket_location['LocationConstraint'],
  #  s3_bucket_name,
  #  key_name)
  object_url = "https://static.kieat.icu/{0}".format(key_name)

  return object_url

def getQuery(event):
  if event["httpMethod"] == "GET":
    return event["queryStringParameters"]
  else:
    return json.loads(event["body"])
