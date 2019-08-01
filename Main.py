import xlrd
import xlwt
import os
import glob
import ctypes

def main(file,count):
       f = xlrd.open_workbook(file)
       sheet = f.sheet_by_index(0)
       rows = sheet.nrows          #50
       cols = sheet.ncols          #16
       questions = []
       ok = exc = exp = 0
       okf = excf = expf = 0
       for i in range(0,cols,2):
              x = sheet.cell(0,i).value
              questions.append(x)

       imp = []
       diff = []
       x =[]
       y = []

       for i in range(0,cols):
              for j in range(1,rows):
                     if i%2 == 0:
                            x.append(sheet.cell(j,i).value)
                     else:
                            y.append(sheet.cell(j,i).value)
              if i%2 == 0:
                     imp.append(x)
              else:
                     diff.append(y)
              x = []
              y = []

       ok_list = []
       ok_avg = []
       exc_list = []
       exc_avg = []
       exp_list = []
       exp_avg = []
       for i in imp:
              for j in range(len(i)):
                     if i[j] == 'OK':
                            ok+=1
                            okf+=int(diff[imp.index(i)][j])
                     elif i[j] == 'Exciting':
                            exc+=1
                            excf+=int(diff[imp.index(i)][j])
                     elif i[j] == 'Expected':
                            exp+=1
                            expf+=int(diff[imp.index(i)][j])

              ok_list.append(ok)
              ok_avg.append(okf/ok)
              exc_list.append(exc)
              exc_avg.append(excf/exc)
              exp_list.append(exp)
              exp_avg.append(expf/exp)

              ok = exc = exp = 0
              okf = excf = expf = 0
       ids = []
       ids.extend([id(questions), id(ok_list), id(ok_avg), id(exc_list),
                   id(exc_avg), id(exp_list), id(exp_avg)])

       save(count, ids, cols//2)

def save(count,ids,n):
       abc = 0
       x = ids[abc]

       f2 = xlwt.Workbook()
       sheet = f2.add_sheet('Sheet 1')
       headings = ["Questions", "Okay", "Average", "Exciting", "Average", "Expected", "Average"]
       for i in range(len(headings)):
              sheet.write(0,i, headings[i])
       for j in range(7):
              x = ids[abc]
              y = ctypes.cast(x, ctypes.py_object).value
              for cn in range(1,n):
                     sheet.write(cn, j, y[cn])

              abc+=1
              if abc > 6:
                     abc = 0
                     
       name = 'data/output/' + 'output' + str(count) + '.xls'
       f2.save(name)
       print(name)


##for filename in os.listdir(os.getcwd()):
##       print(filename)
##
##path = ''
##
##for filename in glob.glob(os.path.join(path, '*.xlsx')):
##       print(filename)

