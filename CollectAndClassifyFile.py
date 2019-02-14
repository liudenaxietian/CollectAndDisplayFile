import xlrd
import xlwt
import os 
import shutil
import time

path_a = r'C:\Users\Autobio-A3517\Desktop\sequence.xlsx'
path_b = r'C:\Users\Autobio-A3517\Desktop\test1.xls'   
path_c = r'C:\Users\Autobio-A3517\Desktop\test'
path_d = r'C:\Users\Autobio-A3517\Desktop\test1'

seq = []
workbook = xlrd.open_workbook(path_a)
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

def CreateHyperlink(file_path):
    def set_style(name,height,bold=False):
        style = xlwt.XFStyle()
        
        font = xlwt.Font()
        font.name = name
        font.bold = bold
        font.colour_index = 4
        font.height = height

        borders= xlwt.Borders()
        borders.left= 6
        borders.right= 6
        borders.top= 6
        borders.bottom= 6

        style.font = font
        style.borders = borders

        return style

    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet('123')
    i= 1
    for file in os.listdir(file_path):
        path = os.path.abspath(os.path.join(file_path,file))
        link='HYPERLINK("%s","%s")'%(str(path),file)
        worksheet.write(i,1,xlwt.Formula(link),set_style('Times New Roman',220,True))
        i += 1
    workbook.save(path_b)

def SearchAndClassify(file_path):
    for file in os.listdir(file_path):
        if os.path.isfile(file_path+'\\'+file):
            for file_name in seq:
                if file_name in file:
                    mkdirs(path_d+'\\'+file_name)
                    shutil.copy(file_path+'\\'+file,path_d+'\\'+file_name)
                    CreateHyperlink(path_d)                   
        else:
            SearchAndClassify(file_path+'\\'+file)

if __name__=='__main__':
    SearchAndClassify(os.path.abspath(path_c))
