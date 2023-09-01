"""
Purpose:
     a tool create exe for python app

"""
import shutil
import os,sys,time
import PyInstaller.__main__
import KERNEL

PATH_FILE,PY_FILE = os.path.split(os.path.abspath(__file__))
PATH_LIB = os.path.split(PATH_FILE)[0]
#
import importlib.util
import datetime
now = datetime.datetime.now()
#
def deleteFolder(dir_path):
    if not os.path.isdir(dir_path):
        return
    try:
        shutil.rmtree(dir_path)
    except OSError as e:
        print("Error: %s : %s" % (dir_path, e.strerror))
#
def deleteFile(sfile):
    try:
        if os.path.isfile(sfile):
            os.remove(sfile)
    except:
        pass
#
def createVersionFile(pyFile,path):
    """
# UTF-8
#
# For more details about fixed file info 'ffi' see:
# http://msdn.microsoft.com/en-us/library/ms646997.aspx
VSVersionInfo(
  ffi=FixedFileInfo(
    # filevers and prodvers should be always a tuple with four items: (1, 2, 3, 4)
    # Set not needed items to zero 0.
    filevers=(1, 1, 0, 0), #<- Put File version here
    # prodvers=(3, 0, 10, 2),
    # Contains a bitmask that specifies the valid bits 'flags'
    mask=0x3f, #<- Do not change
    # Contains a bitmask that specifies the Boolean attributes of the file.
    flags=0x0, #<- Do not change
    # The operating system for which this file was designed.
    # 0x4 - NT and there is no need to change it.
    OS=0x4, #<- Do not change
    # The general type of file.
    # 0x1 - the file is an application.
    fileType=0x1, #<- Do not change
    # The function of the file.
    # 0x0 - the function is not defined for this fileType
    subtype=0x0, #<- Do not change
    # Creation date and time stamp. Sets it automatically. Do not change.
    date=(0, 0) #<- Do not change
    ),
  kids=[
    StringFileInfo(
      [
      StringTable(
        u'040904b0',
        [StringStruct(u'CompanyName', u'Your name here'),
        StringStruct(u'ProductName', u'name here'),
        StringStruct(u'ProductVersion', u'1.1.0.0'), #<- should be same as filevers
        StringStruct(u'OriginalFilename', u'productname.exe'),
        StringStruct(u'FileDescription', u'Short description goes here'),
        StringStruct(u'LegalCopyright', u'copyright stuff here'),
        StringStruct(u'LegalTrademarks', u'legal stuff here'),])
      ]),
    VarFileInfo([VarStruct(u'Translation', [1033, 1200])])
  ]
)
    """
    sfile = path+'\\version.txt'
    deleteFile(sfile)
    #
    spec = importlib.util.spec_from_file_location('', pyFile)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    #
    versions = KERNEL.__version__.replace('.','')
    CompanyName = 'DHBKHN'
    CopyRight = 'Copyright All rights reserved'
    ProductName = os.path.splitext(os.path.basename(pyFile))[0]+'_' +versions +'.exe'
    LegalTrademarks = 'BKPSA'
    #FileDescription = module.PARSER_INPUTS.usage
    FileDescription = 'Version:'+versions+', Build:' +time.asctime()
    #
    av = KERNEL.__version__.split('.')  #  '1, 1, 0, 5'
    for i in range(len(av),4): # version at least 4 digits
        av.append('0')
    Version = ''
    for v1 in av:
        Version +=','+v1
    Version = Version[1:]
    #
    f = open(sfile, mode='w+')
    #
    f.write("# UTF-8\n")
    f.write("#\n")
    f.write("# For more details about fixed file info 'ffi' see:\n")
    f.write("# http://msdn.microsoft.com/en-us/library/ms646997.aspx\n")
    f.write("VSVersionInfo(\n")
    f.write("  ffi=FixedFileInfo(\n")
    f.write("    # filevers and prodvers should be always a tuple with four items: (1, 2, 3, 4)\n")
    f.write("    # Set not needed items to zero 0.\n")
    f.write("    filevers=(%s), #<- Put File version here\n"%Version)
    f.write("    # prodvers=(3, 0, 10, 2),\n")
    f.write("    # Contains a bitmask that specifies the valid bits 'flags'\n")
    f.write("    mask=0x3f, #<- Do not change\n")
    f.write("    # Contains a bitmask that specifies the Boolean attributes of the file.\n")
    f.write("    flags=0x0, #<- Do not change\n")
    f.write("    # The operating system for which this file was designed.\n")
    f.write("    # 0x4 - NT and there is no need to change it.\n")
    f.write("    OS=0x4, #<- Do not change\n")
    f.write("    # The general type of file.\n")
    f.write("    # 0x1 - the file is an application.\n")
    f.write("    fileType=0x1, #<- Do not change\n")
    f.write("    # The function of the file.\n")
    f.write("    # 0x0 - the function is not defined for this fileType\n")
    f.write("    subtype=0x0, #<- Do not change\n")
    f.write("    # Creation date and time stamp. Sets it automatically. Do not change.\n")
    f.write("    date=(0, 0) #<- Do not change\n")
    f.write("    ),\n")
    f.write("  kids=[\n")
    f.write("    StringFileInfo(\n")
    f.write("      [\n")
    f.write("      StringTable(\n")
    f.write("        u'040904b0',\n")
    f.write("        [StringStruct(u'CompanyName', u'%s'),\n"%CompanyName)
    f.write("        StringStruct(u'ProductName', u'%s'),\n"%ProductName)
    f.write("        StringStruct(u'ProductVersion', u'%s'), #<- should be same as filevers\n"%Version)
    f.write("        StringStruct(u'OriginalFilename', u'%s'),\n"%ProductName)
    f.write("        StringStruct(u'FileDescription', u'%s'),\n"%FileDescription)
    f.write("        StringStruct(u'LegalCopyright', u'%s'),\n"%CopyRight)
    f.write("        StringStruct(u'LegalTrademarks', u'%s'),])\n"%LegalTrademarks)
    f.write("      ]),\n")
    f.write("    VarFileInfo([VarStruct(u'Translation', [1033, 1200])])\n")
    f.write("  ]\n")
    f.write(")\n")
    f.close()
    return sfile
#
def build_APP(pyFile,icon,onefile):
    """
    pyFile: python file path of app
    icon: icon file
    onefile: option create one file .exe or not
    """
    pyFile = os.path.abspath(pyFile)
    baseFile = os.path.splitext(os.path.split(pyFile)[1])[0]
    #
    deleteFolder(PATH_FILE+'\\dist\\'+baseFile)
    deleteFolder(PATH_FILE+'\\build\\'+baseFile)
    deleteFolder(PATH_FILE+'\\build')
    #
    # create folder
    try:
        os.mkdir(PATH_FILE+'\\dist')
    except:
        pass
    #
    versionFile = createVersionFile(pyFile,PATH_FILE+'\\dist')
    #
    if onefile:
        deleteFile('dist\\'+baseFile+'.exe')
    #
    # '--onefile'
    args = [pyFile]
    #
    if onefile:
        args.append('--onefile')
    #
    if os.path.isfile(icon):
        args.extend(['--icon' ,icon])
    #
    args.append('--paths')
    args.append(PATH_LIB)
    args.append('--specpath')
    args.append(PATH_FILE+'\\dist')
    #
    args.append('--distpath')
    args.append(PATH_FILE+'\\dist')
    #
    args.append('--workpath')
    args.append(PATH_FILE+'\\build')
    # --version-file version.txt
    args.append('--version-file')
    args.append(versionFile)

    #
    PyInstaller.__main__.run(args)
    #
    deleteFolder(PATH_FILE+'\\build\\'+baseFile)
    deleteFolder(PATH_FILE+'\\build')
    deleteFile(PATH_FILE+'\\dist\\'+baseFile+'\\'+baseFile+'.exe.manifest')
    deleteFile(PATH_FILE+'\\dist\\'+baseFile+'.spec')
    deleteFile(versionFile)

#
if __name__ == '__main__':
    #
    icon = PATH_FILE+'\\bk.ico'
    onefile = 1  # option create one file .exe or not

    pyFile = PATH_FILE + '\\BKPSA.pyw'
    build_APP(pyFile,icon,onefile)




















