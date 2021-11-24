import csv
from openpyxl import Workbook
def func_fdb_not_submitted():
    req_data = ['rollno','registered_sem','scheduled_sem','subno','Name','email','aemail','contact']
    output = "course_feedback_remaining.xlsx" 
    with open('studentinfo.csv', 'r') as info_f :
          std_info  = list(csv.reader(info_f))
    with open('course_master_dont_open_in_excel.csv', 'r') as master_f :
          c_master  = list(csv.reader(master_f))
    with open('course_registered_by_all_students.csv', 'r') as reg_f:
          c_reg = list(csv.reader(reg_f))
    with open('course_feedback_submitted_by_students.csv', 'r') as fed_f :
          c_fdb  = list(csv.reader(fed_f))

    # creating an excel workbook to fill the required details in it
    workbook = Workbook()
    ans_sheet = workbook.active
    ans_sheet.append(req_data)
    for i in c_reg[1:] :
        ans = []
        roll_num = i[0]
        reg_sem = i[1]
        sch_sem = i[2]
        sub_num = i[3]
        for course in c_master[1:] :
            count = 0
            if sub_num == course[0] :
                nonzero = 0
                ltp = course[2]
                splitted_ltp = ltp.split('-')
                splitted_ltp = [int(i) for i in splitted_ltp]
                # finding nonzeros in ltp
                for i in splitted_ltp :
                    if i != 0 :
                        nonzero = nonzero + 1
                # checking for feedback
                for feedback in c_fdb :
                        if roll_num == feedback[3] :
                             if sub_num == feedback[4] :
                                 count = count + 1
                if count < nonzero :
                    ans.append(roll_num)
                    ans.append(reg_sem)
                    ans.append(sch_sem)
                    ans.append(sub_num)
                    for person in std_info :
                        if roll_num == person[1] :
                          ans.append(person[0])
                          ans.append(person[8])
                          ans.append(person[9])
                          ans.append(person[10])
                    if len(ans) < 8:
                        for i in range(8-len(ans)) :
                            # when student info is not available
                            ans.append('NA_IN_STUDENTINFO')

                if ans != [] :
                    ans_sheet.append(ans)
    # saving the workbook created, in the output file as strict excel sheet
    workbook.save(output)                           
    print(req_data)

func_fdb_not_submitted()