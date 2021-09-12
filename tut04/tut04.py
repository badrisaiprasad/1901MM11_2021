import os
from os import listdir
import os.path
from os.path import isfile, join
import csv
import openpyxl

def output_by_rollno():
  if(not os.path.exists("output_individual_roll")):
    os.mkdir("output_individual_roll")

  with open('regtable_old.csv', 'r') as f:
    results = []
    for line in f:
        row = line.split(',')

        if(row[0] != "rollno"):
          path = "./output_individual_roll/"+str(row[0])+".csv"
          if(os.path.isfile(path)):
              with open(path, 'a') as pf:
                pf.write(row[0] + ","+row[1]+","+row[3]+","+ row[-1])
                pf.close()
          else:
            with open(path, 'w') as pf:
                pf.write("rollno,register_sem,subno,sub_type \n")
                pf.close()


def output_by_subject():
  if(not os.path.exists("output_by_subject")):
    os.mkdir("output_by_subject")

  with open('regtable_old.csv', 'r') as f:
    results = []
    for line in f:
        row = line.split(',')

        if(row[3] != "subno"):
          path = "./output_by_subject/"+str(row[3])+".csv"
          if(os.path.isfile(path)):
              with open(path, 'a') as pf:
                pf.write(row[0] + ","+row[1]+","+row[3]+","+ row[-1])
                pf.close()
          else:
            with open(path, 'w') as pf:
                pf.write("rollno,register_sem,subno,sub_type \n")
                pf.close()

output_by_rollno()
output_by_subject()



def xlsx_output_by_rollno():
  files = [f for f in listdir('./output_individual_roll/') if isfile(join('./output_individual_roll/', f))]
  for file in files:
    if os.path.splitext(file)[1][1:] == 'csv':
      wb = openpyxl.Workbook()
      ws = wb.active

      with open('./output_individual_roll/' + file) as f:
          reader = csv.reader(f, delimiter=',')
          for row in reader:
              ws.append(row)
      
      file_name = os.path.splitext(file)[0]
      wb.save('./output_individual_roll/' + file_name + '.xlsx')
      os.remove('./output_individual_roll/' + file)


def xlsx_output_by_subject():
  files = [f for f in listdir('./output_by_subject/') if isfile(join('./output_by_subject/', f))]
  for file in files:
    if os.path.splitext(file)[1][1:] == 'csv':
      wb = openpyxl.Workbook()
      ws = wb.active

      with open('./output_by_subject/' + file) as f:
          reader = csv.reader(f, delimiter=',')
          for row in reader:
              ws.append(row)
      
      file_name = os.path.splitext(file)[0]
      wb.save('./output_by_subject/' + file_name + '.xlsx')
      os.remove('./output_by_subject/' + file)

xlsx_output_by_rollno()
xlsx_output_by_subject()


