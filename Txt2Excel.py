import os
import glob
import sys
import openpyxl

class TxtToExcel:
    def __init__(self, path):
        self._path = path
        self.data = None
        

    def openTxtFile(self):
        try:
            with open(self._path, 'r', encoding='utf-8') as f:
                self.data = f.read()
            return self.data
        except :
            print('error open txt file ...')
            return False
    
    def reshape(self, block):
        self.data = self.data.split(block)
        return self.data

def rwfun(filename):
    filename = filename.replace('\\','/')
    print(filename)
    newfile = path_output + '\\' + os.path.splitext(os.path.basename(filename))[0] + '.xlsx'
    newfile = newfile.replace('\\','/')
    print(newfile)
    te = TxtToExcel(filename)
    
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    
    if not te.openTxtFile(): 
        print('error')
        return
    data = te.reshape('\n')
    for line in data:
        line = line.split(' ')
        sheet.append(line)
    workbook.save(newfile)

def run(text):    
    for filename in Filelist:
        rwfun(filename)

def get_all_files(path):
    Filelist = []
    for home, dirs, files in os.walk(path):
        for filename in files:
            Filelist.append(os.path.join(home,filename)) # 文件名列表，包含完整路径

            # # 文件名列表，只包含文件名
            # Filelist.append(filename)
    return Filelist
        
if __name__ == '__main__':
    path_input = os.path.dirname(sys.argv[0]) + r'\input'
    path_output = os.path.dirname(sys.argv[0]) + r'\output'
    Filelist = get_all_files(path_input)
    run(Filelist)
    #run(path)
    print('over ... ')