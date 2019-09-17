import zipfile
import os,sys,stat
from docx import Document
from pathlib import Path

def unzip(path,zfile):
    file_path=path+os.sep+zfile
    desdir=path+os.sep+zfile[:zfile.index('.zip')]
    srcfile=zipfile.ZipFile(file_path)
    #a = srcfile.encode('cp437').decode('gbk')#先使用cp437编码，然后再使用gbk解码

    for filename in srcfile.namelist():
       # filename=filename.encode('cp437').decode('gbk')
        if filename.endswith('.zip'):
            # if zipfile.is_zipfile(filename):  
            unzip(path,zfile)
            os.remove(os.path.join(path,zfile))
        else:
            srcfile.extract(filename,desdir)
            if filename.endswith('.docx'):
                runDoc(desdir,filename)

def runDoc(path,zfile):
    file_path=path+os.sep+zfile
    # 创建文档对象

    #file_path='/Users/gavin/downloads/test.doc'
    #file_path='/Users/gavin/downloads/aaa.docx'
    #os.chmod(file_path, stat.S_IRWXU)
    document = Document(file_path)
    section = document.sections[0]
    header = section.header
    footer = section.footer
    header.is_linked_to_previous = True
    footer.is_linked_to_previous = True
    document.save(file_path)
   
    
if __name__ == "__main__":
    
    print ("********** 递归解压文件小程序 **********")
    #fileName = input("请输入您要解压的文件")
    fileName ="/Users/gavin/downloads/a.zip"
    path = os.path.dirname(fileName)
    zfile = os.path.basename(fileName)

    unzip(path,zfile)

    print ("*********解压缩完成***************")

