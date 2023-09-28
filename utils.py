__author__    = "Dr. Pham Quang Phuong"
__copyright__ = "Copyright 2022"
__license__   = "All rights reserved"
__email__     = "phuong.phamquang@hust.edu.vn"
__status__    = "Released"
__version__   = "1.0.0.1"
import os
import tempfile,random,string

#
def add2CSV(nameFile,ares,delim):
    """
    append array String to a file CSV
    """
    pathdir = os.path.split(os.path.abspath(nameFile))[0]
    try:
        os.mkdir(pathdir)
    except:
        pass
    #
    if not os.path.isfile(nameFile):
        with open(nameFile, mode='w') as f:
            ew = csv.writer(f, delimiter=delim, quotechar='"',lineterminator="\n")
            for a1 in ares:
                ew.writerow(a1)
            f.close()
    else:
        with open(nameFile, mode='a') as f:
            ew = csv.writer(f, delimiter=delim, quotechar='"',lineterminator="\n", quoting=csv.QUOTE_MINIMAL)
            for a1 in ares:
                ew.writerow(a1)
#
def getIco():
    from win32 import win32gui
    import win32ui
    import win32con
    import win32api
    import os
    opath = get_opath('')
    bmp = opath+'\\bk.bmp'
    ico = opath+'\\bk.ico'
    if os.path.isfile(ico):
        return ico
    ico_x = win32api.GetSystemMetrics(win32con.SM_CXICON)
    ico_y = win32api.GetSystemMetrics(win32con.SM_CYICON)
    large, small = win32gui.ExtractIconEx("BKPSA.exe",0)
    win32gui.DestroyIcon(small[0])
    hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
    hbmp = win32ui.CreateBitmap()
    hbmp.CreateCompatibleBitmap(hdc, ico_x, ico_x)
    hdc = hdc.CreateCompatibleDC()
    hdc.SelectObject(hbmp)
    hdc.DrawIcon((0,0), large[0])
    hbmp.SaveBitmapFile( hdc, bmp)
    ico_x = win32api.GetSystemMetrics(win32con.SM_CXICON)
    ico_y = win32api.GetSystemMetrics(win32con.SM_CYICON)
    large, small = win32gui.ExtractIconEx("BKPSA.exe",0)
    win32gui.DestroyIcon(small[0])
    hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
    hbmp = win32ui.CreateBitmap()
    hbmp.CreateCompatibleBitmap(hdc, ico_x, ico_x)
    hdc = hdc.CreateCompatibleDC()
    hdc.SelectObject(hbmp)
    hdc.DrawIcon((0,0), large[0])
    hbmp.SaveBitmapFile( hdc, bmp)
    from PIL import Image
    img = Image.open(bmp)
    img.save(ico)
    return ico
#
def write2SheetExcel(ws,ra):
    c = 0
    for va in ra:
        c+=1
        for j in range(len(va)):
            ws.cell(c,j+1).value = va[j]
#
def deleteFile(sfile):
    try:
        if os.path.isfile(sfile):
            os.remove(sfile)
    except:
        pass
def get_opath(opath):
    if opath =='':
        opath= os.path.join(tempfile.gettempdir(),'BKPSA')
    try:
        os.mkdir(opath)
    except:
        pass
    return opath
#
def get_file_out(fo,fi,subf,ad,ext):
    """
    get name file output
        fo: name given
        fi: file input (.OLR for example)
        subf: sub folder
        ad: add in the end of file
        ext: extension file output

        check if can write in folder,
        if not=> create in tempo directory
    """
    if fo=='':
        fo1,ext1 = os.path.splitext(fi)
    else:
        fo1,ext1 = os.path.splitext(fo)
        subf = ''
        ad = ''
    #
    if ext=='':
        ext = ext1
    #
    path,sfile = os.path.split(fo1)
    if path=='':
        path = os.path.split(fi)[0]
        if path=='':
            path,sfile = os.path.split(os.path.abspath(fo1))
    # test folder
    if subf!='':
        path = os.path.join(path,subf)
        try:
            os.mkdir(path)
        except:
            pass
    #
    try:
        srandom = ''.join(random.choices(string.ascii_uppercase + string.digits, k=15))
        sf = os.path.join(path,srandom)
        ffile = open(sf, 'w+')
        ffile.close()
        deleteFile(sf)
    except:# create in tempo directory
        path = get_opath('')
        if subf!='':
            path = os.path.join(path,subf)
            try:
                os.mkdir(path)
            except:
                pass
    # test file
    k = -1
    while True:
        k+=1
        if k>0:
            # if k>100:
            ad1 = ad + '_'+str(k)
            # else:
            #     ad1 = ad+'_'+get_String_random(5)
        else:
            ad1 = ad
        #
        fo = os.path.join(path,sfile + ad1+ ext)
        #
        deleteFile(fo)
        if not os.path.isfile(fo):
            return os.path.abspath(fo)


#
def getVal(s0):
    if s0==None or type(s0)==int or type(s0)==float:
        return s0
    sa = s0.split(',')
    res = []
    for s1 in sa:
        try:
            res.append( int(s1) )
        except:
            try:
                res.append( float(s1) )
            except:
                res.append(s1)
    if len(res)==1:
        return res[0]
    return res

#
def toString(v,nRound=5):
    """ convert object/value to String """
    if v is None:
        return 'None'
    t = type(v)
    if t==str:
        if "'" in v:
            return ('"'+v+'"').replace('\n',' ')
        return ("'"+v+"'").replace('\n',' ')
    if t==int:
        return str(v)
    if t==float:
        if v>1.0:
            s1 = str(round(v,nRound))
            return s1[:-2] if s1.endswith('.0') else s1
        elif abs(v)<1e-8:
            return '0'
        s1 ='%.'+str(nRound)+'g'
        return s1 % v
    if t==complex:
        if v.imag>=0:
            return '('+ toString(v.real,nRound)+' +' + toString(v.imag,nRound)+'j)'
        return '('+ toString(v.real,nRound) +' '+ toString(v.imag,nRound)+'j)'
    try:
        return v.toString()
    except:
        pass
    if t in {list,tuple,set}:
        s1=''
        for v1 in v:
            s1+=toString(v1,nRound)+','
        if v:
            s1 = s1[:-1]
        if t==list:
            return '['+s1+']'
        elif t==tuple:
            return '('+s1+')'
        else:
            return '{'+s1+'}'
    if t==dict:
        s1=''
        for k1,v1 in v.items():
            s1+=toString(k1)+':'
            s1+=toString(v1,nRound)+','
        if s1:
            s1 = s1[:-1]
        return '{'+s1+'}'
    return str(v)
