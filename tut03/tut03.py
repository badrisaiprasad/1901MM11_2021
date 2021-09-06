import os

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