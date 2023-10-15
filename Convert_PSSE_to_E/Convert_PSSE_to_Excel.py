# encoding: utf-8
import openpyxl, os, sys, shutil
import numpy as np
from openpyxl import load_workbook

sys_path_PSSE=r'C:\\Program Files (x86)\\PTI\\PSSE33\\PSSBIN' 
sys.path.append(sys_path_PSSE)
os_path_PSSE=r' C:\\Program Files (x86)\\PTI\\PSSE33\\PSSBIN'  
os.environ['PATH'] += ';' + os_path_PSSE
os.environ['PATH'] += ';' + sys_path_PSSE
import psspy
psspy.psseinit(1000)
from psspy import _i, _f, _s

# class get_data_psse:
# 	def __init__(self):
# 		self.node = {}
# 		self.source = {}
# 		self.line = {}
# 		self.shunt = {}

def array2dict(dict_keys, dict_values):
    tmpdict = {}
    for i in range(len(dict_keys)):
        tmpdict[dict_keys[i]] = dict_values[i]
    return tmpdict

def get_bus_data():

	istrings = ['NUMBER','TYPE', 'AREA', 'ZONE']
	ierr, iarray = psspy.abusint(sid, flag_bus, string = istrings)
	ibuses = array2dict(istrings, iarray)

	rstrings = ['BASE', 'PU', 'KV', 'ANGLED']
	ierr, rarray = psspy.abusreal( string = rstrings)
	rbuses = array2dict(rstrings, rarray)

	cstrings = ['NAME', 'EXNAME']
	ierr, cdata = psspy.abuschar(sid, flag_bus, cstrings)
	cbuses = array2dict(cstrings, cdata)

	res = {}
	for i in range(len(ibuses['NUMBER'])):
		bus = ["NAME", "kV"]
		value = [cbuses['NAME'][i], rbuses['BASE'][i]]
		buses = array2dict(bus, value)
		res[ibuses['NUMBER'][i]] = buses

	return res
	
def get_machine_data():
	istrings = ['NUMBER','STATUS']
	ierr, iarray = psspy.amachint(sid, flag_machine, istrings)
	imachines = array2dict(istrings, iarray)
	print(imachines)

	rstrings = ['PGEN','QGEN','QMAX','QMIN']
	ierr, rarray = psspy.amachreal(sid, flag_machine, rstrings)
	rmachines = array2dict(rstrings, rarray)
	print (rmachines)

	cstrings = ['ID', 'NAME', 'EXNAME']
	ierr, carray = psspy.amachchar(sid, flag_machine, cstrings)
	cmachines = array2dict(cstrings, carray)
	res = {}
	a = 1
	for i in range(len(imachines['NUMBER'])):
		elements = ['BUS_ID', 'NAME', 'FLAG','Pgen', 'Qmax', 'Qmin']
		values = [imachines['NUMBER'][i], cmachines['ID'][i], imachines['STATUS'][i],rmachines['PGEN'][i], rmachines['QMAX'][i], rmachines['QMIN'][i]]
		machines = array2dict(elements, values)
		res[a] = machines
		a += 1 
	return res

def get_plant_data():
	istrings = ['TYPE', 'NUMBER', 'STATUS']
	ierr, iarray = psspy.agenbusint(sid, flag_plant, istrings)
	iplants = array2dict(istrings, iarray)

	rstrings = ['BASE', 'PU', 'ANGLED'] #Kiem tra xem lay input hay lay gia tri da tinh toan
	ierr, rarray = psspy.agenbusreal(sid, flag_plant, rstrings)
	rplants = array2dict(rstrings, rarray)

	res = get_machine_data()
	for keys, value in res.items():
		for i in range(len(iplants['NUMBER'])):
			if res[keys]['BUS_ID'] == iplants['NUMBER'][i]:
				res[keys]['CODE'] = 1 if iplants['TYPE'][i] == 2 else 0
				res[keys]['vGen [pu]'] = rplants['PU'][i]
				res[keys]['aGen [deg]'] = rplants['ANGLED'][i]

	return res

def get_line_data():

	istrings = ['FROMNUMBER', 'TONUMBER', 'STATUS']
	ierr, iarray = psspy.abrnint(sid, _i,  _i, flag_line, _i, istrings)
	ibranches = array2dict(istrings, iarray)
	
	rstrings = ['LENGTH', 'RATEA', 'CHARGING']
	ierr, rarray = psspy.abrnreal(sid, _i, _i, flag_line, _i, rstrings)
	rbranches = array2dict(rstrings, rarray)

	cstrings = ['ID', 'FROMNAME', 'TONAME']
	ierr, carray = psspy.abrnchar(sid, _i, _i, flag_line, _i, cstrings)
	cbranches = array2dict(cstrings, carray)

	xstrings = ['RX']
	ierr, xarray = psspy.abrncplx(sid, _i, _i, flag_line, _i, xstrings)
	xbranches = array2dict(xstrings,xarray)

	res = {}
	a = 1
	for i in range(len(ibranches['FROMNUMBER'])):
		ierr, U_base = psspy.busdat(int(ibranches['FROMNUMBER'][i]), 'BASE')
		Zbase = ((U_base)**2)/S_base
		R_0 = ((xbranches['RX'][i].real)*Zbase)/rbranches['LENGTH'][i]
		X_0 = ((xbranches['RX'][i].imag)*Zbase)/rbranches['LENGTH'][i]
		B_0 = (rbranches['CHARGING'][i])*(10**6)/((Zbase)*(rbranches['LENGTH'][i]))
		keys = ['BUS_ID1', 'BUS_ID2', 'NAME BUS', 'CID', 'FLAG', 'LENGTH [km]', 'RATEA [A]', 'R [Ohm/km]', 'X [Ohm/km]', 'B [microS/km]', 'kV']
		values = [ibranches['FROMNUMBER'][i], ibranches['TONUMBER'][i], cbranches['FROMNAME'][i] + '-' + cbranches['TONAME'][i],
		 cbranches['ID'],ibranches['STATUS'][i], rbranches['LENGTH'][i], rbranches['RATEA'][i], R_0, X_0, B_0, U_base]
		lines = array2dict(keys, values)
		res[a] = lines
		a += 1

	return res

def get_shunt_data(): #chua lay duoc Qshunt

	istrings = ['NUMBER', 'STATUS']
	ierr, iarray = psspy.afxshuntint(sid, flag_shunt, istrings)
	ishunts = array2dict(istrings, iarray)

	xstrings = ['SHUNTAC', 'SHUNTNOM']
	ierr, xarray = psspy.afxshuntcplx(sid, flag_shunt, xstrings)
	xshunts = array2dict(xstrings, xarray)
	
	rstrings = ['SHUNTAC', 'SHUNTNOM']
	ierr, rarray = psspy.afxshuntreal(sid, flag_shunt, rstrings)
	rshunts = array2dict(rstrings, rarray)

	cstrings = ['ID', 'NAME', 'EXNAME']
	ierr, carray = psspy.afxshuntchar(sid, flag_shunt, cstrings)
	cshunts = array2dict(cstrings, carray)

	#Kiểm tra lại chỗ kV xem xét sử dụng ierr, rarray = afxshntbusreal(sid, flag, string)
	res = {}
	a = 1
	for i in range(len(ishunts['NUMBER'])):
		ierr, U_base = psspy.busdat(int(ishunts['NUMBER'][i]), 'BASE')
		ierr, rval = psspy.fxsdt2(ishunts['NUMBER'][i], str(1),'NOM')
		keys = ['BUS_ID', 'NAME', 'kV', 'Qshunt', 'FLAG' ]
		values = [ishunts['NUMBER'][i], cshunts['NAME'][i], U_base, rval, ishunts['STATUS'][i]]
		shunts = array2dict(keys, values)
		res[a] = shunts
		a += 1

	return res

def get_load_data(): #kiem tra lai ID va BUSNUMBER

	istrings = ['NUMBER', 'STATUS']
	ierr, iarray = psspy.alodbusint(sid, flag_load, istrings)
	iloads = array2dict(istrings, iarray)
	print (iloads)

	rstrings = ['BASE', 'MVANOM']
	ierr, rarray = psspy.alodbusreal(sid, flag_load, rstrings)
	rloads = array2dict(rstrings, rarray)

	xstrings = ['MVANOM']
	ierr, xarray = psspy.alodbuscplx(sid, flag_load, xstrings)
	xloads = array2dict(xstrings, xarray)
	print(xloads)

	cstrings = ['NAME', 'EXNAME']
	ierr, carray = psspy.alodbuschar(sid, flag_load, cstrings)
	cloads = array2dict(cstrings, carray)

	res = {}
	for i in range(len(iloads['NUMBER'])):
		keys = ['NAME', 'kV', 'FLAG', 'PLOAD', 'QLOAD', 'MEMO']
		values = [cloads['NAME'][i], rloads['BASE'][i] ,iloads['STATUS'][i], xloads['MVANOM'][i].real,xloads['MVANOM'][i].imag, 'kva,kw']
		loads = array2dict(keys, values)
		res[iloads['NUMBER'][i]]= loads

	return res

def get_X2_data():

	istrings = ['FROMNUMBER', 'TONUMBER', 'STATUS', 'NTPOSN']
	ierr, iarray = psspy.atrnint(sid, _i, 3, flag_x2, _i, istrings)
	ix2 = array2dict(istrings, iarray)

	cstrings = ['ID', 'FROMNAME', 'TONAME', 'XFRNAME']
	ierr, carray = psspy.atrnchar(sid, _i, 3, flag_x2, _i, cstrings)
	cx2 = array2dict(cstrings, carray)

	rstrings = ['SBASE1'] #mba luon mac dinh phia 1 la phia cao ap ???????
	ierr, rarray = psspy.atrnreal(sid, _i, 3, flag_x2, _i, rstrings)
	rx2 = array2dict(rstrings, rarray)

	xstrings = ['RXNOM', 'YMAG']
	ierr, xarray = psspy.atrncplx(sid, _i, 3, flag_x2, _i, xstrings)
	xx2 = array2dict(xstrings, xarray)

	res = {}
	a = 1 
	for i in range(len(ix2['FROMNUMBER'])):
		# lay Uc = Udm
		ierr, U_frombus = psspy.busdat(int(ix2['FROMNUMBER'][i]), 'BASE')
		ierr, U_tobus = psspy.busdat(int(ix2['TONUMBER'][i]), 'BASE')
		kV = U_frombus if U_frombus > U_tobus else U_tobus

		#tinh toan thong so MBA
		Zbase = (kV**2)/S_base
		Sn = rx2['SBASE1'][i]
		R = (xx2['RXNOM'][i].real)*Zbase
		X = (xx2['RXNOM'][i].imag)*Zbase
		G = (xx2['YMAG'][i].real)/Zbase
		B = (xx2['YMAG'][i].imag)/Zbase
		uk = round(((100*X*Sn)/(kV**2)),2)
		pk = round(((R*(Sn**2)*1000)/(kV**2)),1)
		i0 = round(((B*100*(kV**2))/Sn),2)
		p0 = round((G*(kV**2)*100),1)

		keys = ['BUS_ID1', 'BUS_ID2', 'NAME BUS', 'kV', 'CID', 'NAME MBA2', 'FLAG', 'Sn', 'uk [%]', 'pk', 'P0', 'i0 [%]', 'MEMO']
		values = [ix2['FROMNUMBER'][i], ix2['TONUMBER'][i], cx2['FROMNAME'][i] + '-' + cx2['TONAME'][i], kV, cx2['ID'][i], cx2['XFRNAME'][i], ix2['STATUS'][i], Sn, uk, pk, p0, i0, 'kW']
		x2Trans = array2dict(keys, values)
		res[a] = x2Trans
		a += 1 

	return res

def get_X3_data():

	return

def set_data_dict():

	return

def add_data_excel(data_dict, excel_file, sheet_name):
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook[sheet_name]
    row = sheet[2]

    # Duyệt qua từng ô trong hàng thứ 2 để tìm cột
    a = {}
    for cell in row:
        if cell.value is None:
            break
        else:
            value = cell.value
            column = cell.column_letter
            a[value] = column

    # Bắt đầu từ hàng 3 để bỏ qua hàng tiêu đề
    i = 3
    for keys, values in data_dict.items():
        cot = a['ID']
        hang = i
        sheet[cot + str(hang)] = keys
        for key, value in data_dict[keys].items():
            o = a[key] + str(i)
            sheet[o] = value
        i += 1

    workbook.save(excel_file)

if __name__ == '__main__':

	default_file = 'default.xlsx'
	output_file = 'Output.xlsx'
	excel_file = shutil.copy(default_file, output_file)
	# sav_file = 'savnw.sav'
	sav_file = 'file5bus(psse33).sav'

	sid = -1
	flag_bus     = 2    # in-service
	flag_machine = 4    # all machines
	flag_plant   = 2    # in-service
	flag_load    = 4    # for all load buses, including those with only out-ofservice loads
	flag_line    = 2    # chỉ lấy các nhánh không có mba
	flag_shunt   = 4    # for all fixed bus shunts.
	flag_x2      = 2    #for all two-winding transformers.
	flag_swsh    = 1    # in-service
	flag_brflow  = 2    # in-service
	owner_brflow = 1    # bus, ignored if sid is -ve
	ties_brflow  = 5

	ierr = psspy.case(sav_file)
	# psspy.fnsl([0,0,0,1,1,0,99,0])
	S_base = psspy.sysmva()
	
	# print(get_bus_data())
	# print(get_load_data())
	# print(get_line_data())
	# print(get_shunt_data())
	# print(get_load_data())
	# print(get_X2_data())


