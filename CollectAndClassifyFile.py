import xlrd
import xlwt
import os 
import shutil

seq = []
workbook = xlrd.open_workbook('C:\\Users\\Autobio-A3517\\Desktop\\sequence.xlsx')
table1 = workbook.sheets()[0]
n_rows = table1.nrows
for i in range(1,n_rows):
    name = table1.cell(i,1).value
    seq.append(name)

def mkdirs(path):
    folder = os.path.exists(path)
    if not folder:
        os.makedirs(path)
    else:
        return None

def CreateHyperlink(path):
    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet('123')
    path = r'C:\Users\Autobio-A3517\Desktop\test1'
    for file in listdir(path):
        path1 = os.path.abspath(os.path.join(path,file))
        link = 'HYPERLINK("%s","%s")'%(str(path),file)
        worksheet.write(1,1,xlwt.Formula(link))
    workbook.save(r'C:\Users\Autobio-A3517\Desktop\test1')

def SearchAndClassify(file_path):
    for file in os.listdir(file_path):
        if os.path.isfile(file_path+'\\'+file):
            for file_name in seq:
                if file_name in file:
                    mkdirs('C:\\Users\\Autobio-A3517\\Desktop\\test1\\'+file_name)
                    shutil.copy(file_path+'\\'+file,'C:\\Users\\Autobio-A3517\\Desktop\\test1\\'+file_name)
                    CreateHyperlink(file_path)
        else:
            SearchAndClassify(file_path+'\\'+file)

''' 将整理的文件路径制作超链接并保存到excel表格中 '''
# workbook = xlwt.Workbook()
# worksheet = workbook.add_sheet('123')
# path = r'C:\\Users\\Autobio-A3517\\Desktop\\456'
# for file in os.listdir(path):
#     file_path = os.path.abspath(os.path.join(path,file))
#     path = os.path.dirname(file_path)
#     link = 'HYPERLINK("%s","%s")'%(str(path),file)
#     worksheet.write(1,1,xlwt.Formula(link))
# workbook.save('C:\\Users\\Autobio-A3517\\Desktop\\test1.xls')

if __name__=='__main__':
    SearchAndClassify(os.path.abspath('C:\\Users\\Autobio-A3517\\Desktop\\test'))
