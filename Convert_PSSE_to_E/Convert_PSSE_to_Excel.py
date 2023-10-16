# encoding: utf-8
import openpyxl, os, sys, shutil
from openpyxl.styles import Alignment
from openpyxl import load_workbook

sys_path_PSSE=r'C:\\Program Files (x86)\\PTI\\PSSE33\\PSSBIN' 
sys.path.append(sys_path_PSSE)
os_path_PSSE=r' C:\\Program Files (x86)\\PTI\\PSSE33\\PSSBIN'  
os.environ['PATH'] += ';' + os_path_PSSE
os.environ['PATH'] += ';' + sys_path_PSSE
import psspy
psspy.psseinit(1000)
from psspy import _i, _f, _s


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
		value = [cbuses['NAME'][i].strip(), rbuses['BASE'][i]]
		buses = array2dict(bus, value)
		res[ibuses['NUMBER'][i]] = buses

	return res
	
def get_machine_data():

	istrings = ['NUMBER','STATUS']
	ierr, iarray = psspy.amachint(sid, flag_machine, istrings)
	imachines = array2dict(istrings, iarray)

	rstrings = ['PGEN','QGEN','QMAX','QMIN']
	ierr, rarray = psspy.amachreal(sid, flag_machine, rstrings)
	rmachines = array2dict(rstrings, rarray)

	cstrings = ['ID', 'NAME', 'EXNAME']
	ierr, carray = psspy.amachchar(sid, flag_machine, cstrings)
	cmachines = array2dict(cstrings, carray)
	res = {}
	a = 1
	for i in range(len(imachines['NUMBER'])):
		elements = ['BUS_ID', 'NAME', 'FLAG','Pgen', 'Qmax', 'Qmin']
		values = [imachines['NUMBER'][i], cmachines['NAME'][i].strip(), imachines['STATUS'][i],rmachines['PGEN'][i], rmachines['QMAX'][i], rmachines['QMIN'][i]]
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
				res[keys]['kV'] = rplants['BASE'][i]

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

		#tinh toan R0, X0, B0
		ierr, U_base = psspy.busdat(int(ibranches['FROMNUMBER'][i]), 'BASE')
		Zbase = ((U_base)**2)/S_base
		if rbranches['LENGTH'][i] != 0 :
			lenght = rbranches['LENGTH'][i]
			R_0 = ((xbranches['RX'][i].real)*Zbase)/rbranches['LENGTH'][i]
			X_0 = ((xbranches['RX'][i].imag)*Zbase)/rbranches['LENGTH'][i]
			B_0 = (rbranches['CHARGING'][i])*(10**6)/((Zbase)*(rbranches['LENGTH'][i]))

		else:
			lenght = 1
			R_0 = xbranches['RX'][i].real
			X_0 = xbranches['RX'][i].imag
			B_0 = rbranches['CHARGING'][i]

		keys = ['BUS_ID1', 'BUS_ID2', 'NAME BUS', 'CID', 'FLAG', 'LENGTH [km]', 'RATEA [A]', 'R [Ohm/km]', 'X [Ohm/km]', 'B [microS/km]', 'kV']
		values = [ibranches['FROMNUMBER'][i], ibranches['TONUMBER'][i], cbranches['FROMNAME'][i].strip() + '-' + cbranches['TONAME'][i].strip(),
		 int(cbranches['ID'][i]), ibranches['STATUS'][i], lenght, rbranches['RATEA'][i], R_0, X_0, B_0, U_base]
		lines = array2dict(keys, values)
		res[a] = lines
		a += 1

	return res

def get_shunt_data(): 

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
		keys = ['BUS_ID', 'NAME', 'kV', 'Qshunt', 'FLAG', 'MEMO' ]
		values = [ishunts['NUMBER'][i], cshunts['NAME'][i].strip(), U_base, rval.imag, ishunts['STATUS'][i], 'MVar' if a == 1 else '']
		shunts = array2dict(keys, values)
		res[a] = shunts
		a += 1

	return res

def get_load_data(): #kiem tra lai ID va BUSNUMBER

	istrings = ['NUMBER', 'STATUS']
	ierr, iarray = psspy.alodbusint(sid, flag_load, istrings)
	iloads = array2dict(istrings, iarray)

	rstrings = ['BASE', 'MVANOM']
	ierr, rarray = psspy.alodbusreal(sid, flag_load, rstrings)
	rloads = array2dict(rstrings, rarray)

	xstrings = ['MVANOM']
	ierr, xarray = psspy.alodbuscplx(sid, flag_load, xstrings)
	xloads = array2dict(xstrings, xarray)

	cstrings = ['NAME', 'EXNAME']
	ierr, carray = psspy.alodbuschar(sid, flag_load, cstrings)
	cloads = array2dict(cstrings, carray)

	res = {}
	for i in range(len(iloads['NUMBER'])):
		keys = ['NAME', 'kV', 'FLAG', 'PLOAD', 'QLOAD', 'MEMO']
		values = [cloads['NAME'][i].strip(), rloads['BASE'][i] ,iloads['STATUS'][i], xloads['MVANOM'][i].real,xloads['MVANOM'][i].imag, 'kva,kw']
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
		values = [ix2['FROMNUMBER'][i], ix2['TONUMBER'][i], cx2['FROMNAME'][i].strip() + '-' + cx2['TONAME'][i].strip(), kV, int(cx2['ID'][i]), cx2['XFRNAME'][i].strip(), 
		ix2['STATUS'][i], Sn, uk, pk, p0, i0, 'kW(P), MVA(S)']
		x2Trans = array2dict(keys, values)
		res[a] = x2Trans
		a += 1 

	return res

def get_X3_data():

	istrings = ['WIND1NUMBER', 'WIND2NUMBER', 'WIND3NUMBER', 'STATUS']
	ierr, iarray = psspy.atr3int(sid, _i, 3, flag_x3, 2, istrings)
	ix3 = array2dict(istrings, iarray)

	xstrings = ['RX1-2NOM', 'RX2-3NOM', 'RX3-1NOM', 'YMAG']
	ierr, xarray = psspy.atr3cplx(sid, _i, 3, flag_x3, 2, xstrings)
	xx3 = array2dict(xstrings, xarray)

	cstrings = ['ID', 'WIND1NAME', 'WIND2NAME', 'WIND3NAME', 'XFRNAME']
	ierr, carray = psspy.atr3char(sid, _i, 3, flag_x3, 2, cstrings)
	cx3 = array2dict(cstrings, carray)

	rstrings = ['SBASE']
	ierr, rarray = psspy.awndreal(sid, _i, 3, flag_x3, 2, rstrings)
	rx3 = array2dict(rstrings, rarray)
	
	res ={}
	a = 1
	for i in range(len(ix3['WIND3NUMBER'])):

		S_h = max(rx3['SBASE'])

		pk1_2 = xx3['RX1-2NOM'][i].real*1000*(S_h**2)/S_base
		pk2_3 = xx3['RX2-3NOM'][i].real*1000*(S_h**2)/S_base
		pk3_1 = xx3['RX3-1NOM'][i].real*1000*(S_h**2)/S_base

		uk1_2 = xx3['RX1-2NOM'][i].imag*100*S_h/S_base
		uk2_3 = xx3['RX2-3NOM'][i].imag*100*S_h/S_base
		uk3_1 = xx3['RX3-1NOM'][i].imag*100*S_h/S_base

		p_0 = xx3['YMAG'][i].real*1000*S_base
		i_0 = xx3['YMAG'][i].imag*100*S_base/S_h

		keys = ['BUS_ID1', 'BUS_ID2', 'BUS_ID3', 'NAME BUS', 'CID', 'NAME MBA3', 'FLAG', 'Sn1', 'Sn2', 'Sn3', 'uk1-2   [%]', 'uk1-3   [%]', 
		'uk2-3   [%]', 'pk1-2', 'pk1-3', 'pk2-3', 'P0', 'i0     [%]', 'MEMO']

		values = [ix3['WIND1NUMBER'][i], ix3['WIND2NUMBER'][i], ix3['WIND3NUMBER'][i], str(cx3['WIND1NAME'][i].strip()) +'-'+ str(cx3['WIND2NAME'][i].strip()) + '-' + str(cx3['WIND3NAME'][i].strip()), 
		str(cx3['ID'][i]), cx3['XFRNAME'][i].strip(), ix3['STATUS'][i], rx3['SBASE'][0], rx3['SBASE'][1], rx3['SBASE'][2], uk1_2, uk3_1, uk2_3, pk1_2, pk3_1, pk2_3, p_0, i_0, 'MVA, KW' ]
		x3Trans = array2dict(keys, values)

		res[a] = x3Trans
		a += 1

	return res

def set_bus_data_dict(bus_data_dict, load_data_dict):

	for keys, values in bus_data_dict.items():
		for key, value in load_data_dict.items():
			if key == keys:
				bus_data_dict[keys]['QLOAD'] = load_data_dict[key]['QLOAD']
				bus_data_dict[keys]['PLOAD'] = load_data_dict[key]['PLOAD']
				bus_data_dict[keys]['FLAG'] = load_data_dict[key]['FLAG']
				bus_data_dict[keys]['MEMO'] = load_data_dict[key]['MEMO']

	return bus_data_dict

def add_data_excel(data_dict, excel_file, sheet_name):

    workbook = openpyxl.load_workbook(excel_file, read_only=False, data_only=False)
    sheet = workbook[sheet_name]
    row = sheet[2]
    a = {cell.value: cell.column_letter for cell in row if cell.value is not None}
    i = 3
    for keys, values in data_dict.items():
        cot = a.get('ID')
        if cot is None:
            return
        hang = i
        sheet[cot + str(hang)] = keys
        sheet[cot + str(hang)].alignment = Alignment(horizontal='center', vertical='center')
        for key, value in data_dict[keys].items():
            o = a.get(key)
            if o is not None:
                sheet[o + str(i)] = value
                sheet[o + str(i)].alignment = Alignment(horizontal='center', vertical='center')
        i += 1
    workbook.save(excel_file)
    workbook.close()

def Creat_new_excel():
    path = os.getcwd()
    path_default = os.path.join(path, 'default.xlsx')
    path_new = os.path.join(path, 'output.xlsx')

    counter = 1
    while os.path.isfile(path_new):
        base_name, extension = os.path.splitext(path_new)
        path_new = "{0}({1}){2}".format(base_name, counter, extension)
        counter += 1

    shutil.copy(path_default, path_new)

    return path_new

if __name__ == '__main__':

	excel_file = Creat_new_excel()
	sav_file = 'savnw.sav'
	# sav_file = 'file5bus(psse33).sav'

	sid = -1
	flag_bus     = 2    # in-service
	flag_machine = 4    # all machines
	flag_plant   = 2    # in-service
	flag_load    = 4    # for all load buses, including those with only out-ofservice loads
	flag_line    = 2    # chỉ lấy các nhánh không có mba
	flag_shunt   = 4    # for all fixed bus shunts.
	flag_x2      = 2    # for all two-winding transformers.
	flag_x3      = 2    # for all three-winding transformers
	flag_swsh    = 1    # in-service
	flag_brflow  = 2    # in-service
	owner_brflow = 1    # bus, ignored if sid is -ve
	ties_brflow  = 5

	ierr = psspy.case(sav_file)
	# psspy.fnsl([0,0,0,1,1,0,99,0])
	S_base = psspy.sysmva()

	get_machine_data()
	source_data_dict = get_plant_data()
	add_data_excel(source_data_dict, excel_file, 'SOURCE')

	shunts_data_dict = get_shunt_data()
	add_data_excel(shunts_data_dict, excel_file, 'SHUNT')

	lines_data_dict = get_line_data()
	add_data_excel(lines_data_dict, excel_file, 'LINE')

	x2Trans_data_dict = get_X2_data()
	add_data_excel(x2Trans_data_dict, excel_file, 'TRF2')

	bus_data_dict = get_bus_data()
	load_data_dict = get_load_data()
	allbus_data_dict = set_bus_data_dict(bus_data_dict, load_data_dict)
	add_data_excel(allbus_data_dict, excel_file, 'BUS')

	x3Trans_data_dict = get_X3_data()
	add_data_excel(x3Trans_data_dict, excel_file, 'TRF3')


