import os, sys, openpyxl
import shutil 
import pandas as pd
import sqlite3
import math

sys.path.append("D:\\psdistribution")
# KERNEL = "D:\\psdistribution\\KERNEL.py"
from KERNEL import DATAP
sys.stdout.reconfigure(encoding='utf-8')

ELEMENT_ID = 0
TERMINAL_ID = 0

def get_databasefile(excel_file):

    current_dir = os.path.dirname(os.path.abspath(__file__))
    default_folder_path = os.path.join(current_dir, 'DEFAULT FILE')
    base_folder_name = os.path.basename(excel_file).split(".")[0] + 'sin'
    
    index = 1
    while True:
        if index == 1:
            new_folder_path = os.path.join(current_dir, base_folder_name)
        else:
            new_folder_path = os.path.join(current_dir, f'{base_folder_name}({index})')

        if not os.path.exists(new_folder_path):
            break

        index += 1
    shutil.copytree(default_folder_path, new_folder_path)

    for name in os.listdir(new_folder_path):
        item_path = os.path.join(new_folder_path, name)
        if os.path.isfile(item_path):
            file_name, file_extension = os.path.splitext(name)
            # Đổi tên tệp (không đổi định dạng)
            new_name = base_folder_name + file_extension
            os.rename(item_path, os.path.join(new_folder_path, new_name))
        elif os.path.isdir(item_path):
            new_name = base_folder_name + '_files'
            os.rename(item_path, os.path.join(new_folder_path, new_name))

    subfolder_name = base_folder_name + '_files'  
    database_file_name = 'database.db' 
    database_file_path = os.path.join(new_folder_path, subfolder_name, database_file_name)
 
    if os.path.exists(database_file_path):

        return database_file_path

    else:

        return None

def find_and_output_value(database_file, table_name, search_value):
    conn = sqlite3.connect(database_file)
    cursor = conn.cursor()
    cursor.execute(f"SELECT VoltLevel_ID FROM {table_name} WHERE Un = ?", (search_value,))
    
    result = cursor.fetchone()[0]
    return result

def get_input_excel(file_name):

	current_directory = os.path.dirname(os.path.abspath(__file__))
	parent_directory = os.path.dirname(current_directory)
	inputs_folder = parent_directory + '\\inputs'
	file_path = os.path.join(inputs_folder, file_name)
	
	return file_path

def add_data_db(db_file, table_name, data_dict):
    # Kiểm tra xem tệp cơ sở dữ liệu tồn tại
    if not os.path.exists(db_file):
        print(f"Tệp cơ sở dữ liệu '{db_file}' không tồn tại.")
        return

    try:
        conn = sqlite3.connect(db_file)
        cursor = conn.cursor()
        column_names = list(data_dict.keys())
        values = list(data_dict.values())
        query = f"INSERT OR REPLACE INTO {table_name} ({', '.join(column_names)}) VALUES ({', '.join(['?'] * len(column_names))})"
        cursor.execute(query, values)

        conn.commit()

        # print("Dữ liệu đã được chèn hoặc thay thế thành công vào bảng", table_name)

    except sqlite3.Error as e:
        print("Lỗi SQLite:", e)
    finally:
        if conn:
            conn.close()

def get_default_data(templates_db, table_name):

    conn = sqlite3.connect(templates_db)
    cursor = conn.cursor()

    cursor.execute('PRAGMA table_info({})'.format(table_name))
    column_names = [column[1] for column in cursor.fetchall()]
    
    cursor.execute('SELECT * FROM {} LIMIT 1'.format(table_name))
    row = cursor.fetchone()

    if row:
        # Tạo một từ điển cho hàng đầu tiên
        first_row_dict = {}
        for i, value in enumerate(row):
            column_name = column_names[i]
            first_row_dict[column_name] = value
    else:
        first_row_dict = None  

    conn.close() 

    return first_row_dict        

#add voltagelevel và variant vào database
def add_voltagelevel(db_file, kv_values):

	Variant_dict = get_default_data(templates_db, 'Variant')
	add_data_db(db_file, 'Variant', Variant_dict)

	voltagelevel_dict = get_default_data(templates_db, 'VoltageLevel')

	kv_level_dict = {}
	for i in range(len(kv_values)):
		voltagelevel_dict['VoltLevel_ID'] = int(i + 1)
		voltagelevel_dict['Name'] = str(kv_values[i]) + ' kV'
		voltagelevel_dict['ShortName'] = str(kv_values[i]) + ' kV'
		voltagelevel_dict['Un'] = float(kv_values[i])
		voltagelevel_dict['Uop']= float(kv_values[i])
		kv_level_dict[int(i + 1)] = kv_values[i]

		add_data_db(db_file, "VoltageLevel" , voltagelevel_dict)

	return kv_level_dict

#-------------------------------------------------------------------------
#add dữ liệu của bus vào database(gồm Node và GraphicNode, GraphicText):
def add_node_db(node_data_dict, db_file, kv_level_dict ):

	node_dict_db = get_default_data(templates_db, 'Node')

	for keys, values in node_data_dict.items():
		node_dict_db['Node_ID'] = keys
		node_dict_db['Name']= node_data_dict[keys]['NAME']
		node_dict_db['ShortName']= node_data_dict[keys]['NAME']
		node_dict_db['Un']= node_data_dict[keys]['kV']
		for key_kv_value, value_kv_value in kv_level_dict.items():
			if value_kv_value == node_data_dict[keys]['kV']:
				node_dict_db['VoltLevel_ID'] = key_kv_value
		add_data_db(db_file, "Node" , node_dict_db)

	#GraphicNodeeeeeeeeeeeeeeeeeeeeeeeeeeeeee

	GraphicNode_dict_db = get_default_data(templates_db, 'GraphicNode')

	for keys, values in node_data_dict.items():
		GraphicNode_dict_db['GraphicNode_ID'] = keys
		GraphicNode_dict_db['GraphicText_ID1'] = keys
		GraphicNode_dict_db['Node_ID'] = keys
		GraphicNode_dict_db['NodeStartX'] = node_data_dict[keys]['xCoord'] + 100
		GraphicNode_dict_db['NodeStartY'] = node_data_dict[keys]['yCoord'] + 100
		GraphicNode_dict_db['NodeEndX'] = node_data_dict[keys]['xCoord'] + 100
		GraphicNode_dict_db['NodeEndY'] = node_data_dict[keys]['yCoord'] + 100
		add_data_db(db_file, "GraphicNode" , GraphicNode_dict_db)

	#GraphicTextttttttttttttttttttttttttttttttttt

	GraphicText_dict_db = get_default_data(templates_db, 'GraphicText')

	for keys, values in node_data_dict.items():
		GraphicText_dict_db['GraphicText_ID'] = keys
		add_data_db(db_file, "GraphicText" , GraphicText_dict_db)

#-------------------------------------------------------------------------------------
#add dữ liệu cua Infeeder:
def add_source_db(source_data_dict, db_file, kv_level_dict):
	
	global ELEMENT_ID
	global TERMINAL_ID
	Infeeder_dict_db = get_default_data(templates_db, 'Infeeder')
	#Flag_LF(type bus) = 3(swing bus), 7(PV);'u'; 'delta' ; Flag_Limit (Gioi han P,Q,V)

	for key,value in source_data_dict.items():

		a = source_data_dict[key]

		ELEMENT_ID += 1
		TERMINAL_ID += 1
		Element_dict_db['Element_ID'] = ELEMENT_ID
		Element_dict_db['Type'] = 'Infeeder'
		Element_dict_db['Flag_State'] = 1 if a['FLAG'] == 1 else 0
		Element_dict_db['Name'] = 'I' + str(ELEMENT_ID)
		Element_dict_db['ShortName'] = 'I' + str(ELEMENT_ID)
		for key_kv_value, value_kv_value in kv_level_dict.items():
			if value_kv_value == a['kV']:
				Element_dict_db['VoltLevel_ID'] = key_kv_value
		add_data_db(db_file, "Element" , Element_dict_db)

		Infeeder_dict_db['Element_ID'] = ELEMENT_ID
		if a['CODE'] == 0 or a['CODE'] == None:
			Infeeder_dict_db['Flag_Lf'] = 3
			Infeeder_dict_db['u'] = (a['vGen [pu]'])*100 #đơn vị điện áp là %
			Infeeder_dict_db['delta'] = a['aGen [deg]']
			add_data_db(db_file, "Infeeder", Infeeder_dict_db)
		elif a['CODE'] == 1 : #can kiem tra lai
			Infeeder_dict_db['Flag_Lf'] = 11
			Infeeder_dict_db['u'] = a['vGen [pu]']
			Infeeder_dict_db['P'] = a['Pgen']
			Infeeder_dict_db['Flag_LfLimit'] = 0 #Kiem tra lai elif nay 
			add_data_db(db_file, "Infeeder", Infeeder_dict_db)

		Terminal_dict_db['Terminal_ID'] = TERMINAL_ID
		Terminal_dict_db['Element_ID'] = ELEMENT_ID
		Terminal_dict_db['Node_ID'] = a['BUS_ID']
		Terminal_dict_db['TerminalNo'] = 1 
		Terminal_dict_db['Flag_State'] = 1 #switch feeder
		Terminal_dict_db['Flag_Switch'] = 0 # đang mặc định là 0 
		add_data_db(db_file, "Terminal", Terminal_dict_db)

#-------------------------------------------------------------------------------------
#add dữ liệu của Line: 
def add_line_db(line_data_dict, db_file, kv_level_dict):

	global ELEMENT_ID
	global TERMINAL_ID

	Line_dict_db = get_default_data(templates_db, 'Line')
	#Element_ID, l, r, x, c, Un, Ith

	for key in line_data_dict.keys():
		a = line_data_dict[key]
		ELEMENT_ID += 1
		TERMINAL_ID += 1
		Line_dict_db['Element_ID'] = ELEMENT_ID
		Line_dict_db['l'] = a['LENGTH [km]']
		Line_dict_db['r'] = a['R [Ohm/km]']
		Line_dict_db['x'] = a['X [Ohm/km]']
		Line_dict_db['c'] = (a['B [microS/km]'])*(1e-3)/(100*(math.pi))
		Line_dict_db['Un'] = a['kV']
		Line_dict_db['Ith'] = a['RATEA [A]']
		add_data_db(db_file, "Line", Line_dict_db)

		Element_dict_db['Element_ID'] = ELEMENT_ID
		Element_dict_db['Type'] = 'Line'
		Element_dict_db['Flag_State'] = a['FLAG']
		Element_dict_db['Name'] = 'L'+str(ELEMENT_ID)
		Element_dict_db['ShortName'] = 'L'+str(ELEMENT_ID)
		for key_kv_value, value_kv_value in kv_level_dict.items():
			if value_kv_value == a['kV']:
				Element_dict_db['VoltLevel_ID'] = key_kv_value
		add_data_db(db_file, 'Element', Element_dict_db)

		Terminal_dict_db['Terminal_ID'] = TERMINAL_ID
		Terminal_dict_db['Element_ID'] = ELEMENT_ID
		Terminal_dict_db['Node_ID'] = a ['BUS_ID1']
		Terminal_dict_db['Flag_State'] = a['FLAG2']
		Terminal_dict_db['Flag_Switch'] = 0
		Terminal_dict_db['TerminalNo'] = 1
		add_data_db(db_file, "Terminal", Terminal_dict_db)

		TERMINAL_ID += 1
		Terminal_dict_db['Terminal_ID'] = TERMINAL_ID
		Terminal_dict_db['Element_ID'] = ELEMENT_ID
		Terminal_dict_db['Node_ID'] = a ['BUS_ID2']
		Terminal_dict_db['Flag_State'] = a['FLAG3']
		Terminal_dict_db['Flag_Switch'] = 0
		Terminal_dict_db['TerminalNo'] = 2
		add_data_db(db_file, "Terminal", Terminal_dict_db)

#cẦN KIỂM TRA XEM CHỖ CID

#-------------------------------------------------------------------------------------
#add dữ liệu của load 
def add_load_db(bus_data_dict, db_file, kv_level_dict, profile):

	global ELEMENT_ID
	global TERMINAL_ID

	Load_dict_db = get_default_data(templates_db, 'Load')
	#Element_ID, P(MW), Q(MVAr), DayOpSer_ID

	for key in bus_data_dict.keys():

		a = bus_data_dict[key]
		
		if a['PLOAD'] == None and a['QLOAD'] == None :
			continue
		else:
			ELEMENT_ID += 1
			TERMINAL_ID +=1
			Load_dict_db['Element_ID'] =  ELEMENT_ID
			Load_dict_db['P'] = (a['PLOAD'])/1000 
			Load_dict_db['Q'] = (a['QLOAD'])/1000
			Load_dict_db['DayOpSer_ID'] = 0 if a['Load Profile'] == None else profile[a['Load Profile']]
			add_data_db(db_file, "Load", Load_dict_db)

			Element_dict_db['Element_ID'] = ELEMENT_ID
			Element_dict_db['Type'] = 'Load'
			Element_dict_db['Flag_State'] = a['FLAG']
			Element_dict_db['Name'] = 'LO'+str(ELEMENT_ID)
			Element_dict_db['ShortName'] = 'LO'+str(ELEMENT_ID)
			for key_kv_value, value_kv_value in kv_level_dict.items():
				if value_kv_value == a['kV']:
					Element_dict_db['VoltLevel_ID'] = key_kv_value
			add_data_db(db_file, 'Element', Element_dict_db)

			Terminal_dict_db['Terminal_ID'] = TERMINAL_ID
			Terminal_dict_db['Element_ID'] = ELEMENT_ID
			Terminal_dict_db['Node_ID'] = int(key)
			Terminal_dict_db['TerminalNo'] = 1 
			Terminal_dict_db['Flag_State'] = 1 #switch feeder
			Terminal_dict_db['Flag_Switch'] = 0 # đang mặc định là 0 
			add_data_db(db_file, "Terminal", Terminal_dict_db)

#-------------------------------------------------------------------------------------
#add dữ liệu của shunt
def add_shunt_db(shunt_data_dict, db_file, kv_level_dict):

	global ELEMENT_ID
	global TERMINAL_ID

	ShuntReactor_dict_db = get_default_data(templates_db, 'ShuntReactor')
	#ElementID, Sn, Un

	ShuntCondensator_dict_db = get_default_data(templates_db, 'ShuntCondensator')
	#Element_ID, Sn, Un

	for key in shunt_data_dict.keys():
		a =shunt_data_dict[key]
		ELEMENT_ID += 1
		TERMINAL_ID += 1

		ShuntCondensator_dict_db['Element_ID'] = ELEMENT_ID
		ShuntCondensator_dict_db['Sn']  = (a['Qshunt'])/1000 if a['Qshunt'] >= 0 else -(a['Qshunt'])/1000
		ShuntCondensator_dict_db['Un'] = a['kV']
		add_data_db(db_file, "ShuntCondensator", ShuntCondensator_dict_db)
		
		Element_dict_db['Element_ID'] = ELEMENT_ID
		Element_dict_db['Type'] = 'ShuntCondensator'
		Element_dict_db['Flag_State'] = a['FLAG']
		Element_dict_db['Name'] = 'SHC'+str(ELEMENT_ID)
		Element_dict_db['ShortName'] = 'SHC'+str(ELEMENT_ID)
		for key_kv_value, value_kv_value in kv_level_dict.items():
			if value_kv_value == a['kV']:
				Element_dict_db['VoltLevel_ID'] = key_kv_value
		add_data_db(db_file, 'Element', Element_dict_db)

		Terminal_dict_db['Terminal_ID'] = TERMINAL_ID
		Terminal_dict_db['Element_ID'] = ELEMENT_ID
		Terminal_dict_db['Node_ID'] = a['BUS_ID']
		Terminal_dict_db['TerminalNo'] = 1 # 1 terminal
		Terminal_dict_db['Flag_State'] = a['FLAG3'] #switch 
		Terminal_dict_db['Flag_Switch'] = 0 # đang mặc định là 0 
		add_data_db(db_file, "Terminal", Terminal_dict_db)

#-------------------------------------------------------------------------------------
#add dữ liệu của MBA 2 cuộn dây

def add_x2Trans_db(TRF_2_data_dict, db_file, kv_level_dict, node_data_dict):

	global ELEMENT_ID
	global TERMINAL_ID

	TwoWindingTransformer_dict_db =  get_default_data(templates_db, 'TwoWindingTransformer')
	#Element_ID, Un1, Un2, Sn, Smax, uk, ur, Vfe, i0

	for key in TRF_2_data_dict.keys():
		a =TRF_2_data_dict[key]
		ELEMENT_ID += 1
		TERMINAL_ID += 1

		TwoWindingTransformer_dict_db['Element_ID'] = ELEMENT_ID
		TwoWindingTransformer_dict_db['Un1'] = node_data_dict[a['BUS_ID1']]['kV']
		TwoWindingTransformer_dict_db['Un2'] = node_data_dict[a['BUS_ID2']]['kV']
		TwoWindingTransformer_dict_db['Sn'] = a['Sn']
		TwoWindingTransformer_dict_db['Smax'] = a['Sn']	
		TwoWindingTransformer_dict_db['uk'] = 8/100 #std 
		TwoWindingTransformer_dict_db['ur'] = a['uk [%]']
		TwoWindingTransformer_dict_db['Vfe'] = a['P0']
		TwoWindingTransformer_dict_db['i0'] = a['i0 [%]']
		add_data_db(db_file, "TwoWindingTransformer", TwoWindingTransformer_dict_db)

		Element_dict_db['Element_ID'] = ELEMENT_ID
		Element_dict_db['Type'] = 'TwoWindingTransformer'
		Element_dict_db['Flag_State'] = a['FLAG']
		Element_dict_db['Name'] = a['NAME MBA2']
		Element_dict_db['ShortName'] = '2T'+str(ELEMENT_ID)
		for key_kv_value, value_kv_value in kv_level_dict.items():
			if value_kv_value == a['kV']:
				Element_dict_db['VoltLevel_ID'] = key_kv_value
		add_data_db(db_file, 'Element', Element_dict_db)

		Terminal_dict_db['Terminal_ID'] = TERMINAL_ID
		Terminal_dict_db['Element_ID'] = ELEMENT_ID
		Terminal_dict_db['Node_ID'] = a['BUS_ID1']
		Terminal_dict_db['TerminalNo'] = 1 # 1 terminal
		Terminal_dict_db['Flag_State'] = 1 #switch 
		Terminal_dict_db['Flag_Switch'] = 0 # đang mặc định là 0 
		add_data_db(db_file, "Terminal", Terminal_dict_db)
		TERMINAL_ID += 1

		Terminal_dict_db['Terminal_ID'] = TERMINAL_ID
		Terminal_dict_db['Element_ID'] = ELEMENT_ID
		Terminal_dict_db['Node_ID'] = a['BUS_ID2']
		Terminal_dict_db['TerminalNo'] = 2 # 1 terminal
		Terminal_dict_db['Flag_State'] = 1 #switch 
		Terminal_dict_db['Flag_Switch'] = 0 # đang mặc định là 0 
		add_data_db(db_file, "Terminal", Terminal_dict_db)

def add_x3Trans_db(TRF_3_data_dict, db_file):

	return

def add_profile_db(profile_data_dict, templates_db, db_file):

	OpSer_dict_db = get_default_data(templates_db, 'OpSer')
	#OpSer_ID, Name, Shortname, BaseT, Flag_Ser(1: Time series, 4: Operating points), Flag_Typ(1: Factor, 2: Factor P and Q, ...)

	OpSerVal_dict_db = get_default_data(templates_db, 'OpSerVal')
	#OpSerVal_ID, OpSer_ID, OpTime, Factor, Flag_Curve (1:Liên tục, 2: Rời rạc)

	CalcParameter_dict_db = get_default_data(templates_db, 'CalcParameter')
	#LC_StartTime, LC_Duration, LC_TimeStep
 
	OpSer_ID = 1
	OpSerVal_ID = 1

	name_profile = []
	res ={}
	OpTime =[]
	deltaTime = []

	for keys, values in profile_data_dict.items():
		for key in values.keys():
			if key != 'deltaTime' and  key != 'MEMO':
				name_profile.append(key)

	name_profile = list (set(name_profile))			
	# print(name_profile)
	for name in name_profile:
		profile_data = {}
		for keys, values in profile_data_dict.items():
			a = {}
			b = {}
			a ['name'] = name
			a ['OpTime'] = keys  
			a [name] = values[name]
			a ['deltaTime'] = values['deltaTime']
			a ['MEMO'] = values['MEMO']
			a ['OpSer_ID'] = OpSer_ID
			del values[name]
			profile_data[OpSerVal_ID] = a
			OpSerVal_ID += 1

		for key, value in profile_data.items():

			OpTime.append(value['OpTime'])
			deltaTime.append(value['deltaTime'])

			res[value['name']] = value['OpSer_ID']

			OpSerVal_dict_db['OpSerVal_ID'] = key
			OpSerVal_dict_db['OpSer_ID'] = value['OpSer_ID']
			OpSerVal_dict_db['OpTime'] = value ['OpTime']
			OpSerVal_dict_db['Factor'] = value [name]
			add_data_db(db_file, 'OpSerVal', OpSerVal_dict_db)


			OpSer_dict_db['OpSer_ID'] = OpSer_ID
			OpSer_dict_db['Name'] = name
			OpSer_dict_db['Shortname'] = name
			OpSer_dict_db['BaseT'] = max(OpTime)
			OpSer_dict_db['Flag_Ser'] = 1
			OpSer_dict_db['Flag_Typ'] = 1
			add_data_db(db_file, 'OpSer', OpSer_dict_db)

		OpSer_ID += 1

	CalcParameter_dict_db['LC_StartTime'] = min(OpTime)
	CalcParameter_dict_db['LC_Duration'] = max (OpTime) - min(OpTime)
	CalcParameter_dict_db['LC_TimeStep'] = min (deltaTime)
	add_data_db(db_file, 'CalcParameter', CalcParameter_dict_db)

	return res

def main(templates_db, data, db_file):

	node_data_dict = data.abus
	kv_values = []
	for keys, values in node_data_dict.items():
		kv_values. append(node_data_dict[keys]['kV'])
		kv_values = list(set(kv_values))
	kv_level_dict = add_voltagelevel(db_file, kv_values)

	profile_data_dict = data.aprofile
	profile =add_profile_db(profile_data_dict, templates_db, db_file)
	
	add_node_db(node_data_dict, db_file, kv_level_dict)

	source_data_dict = data.asource
	add_source_db(source_data_dict, db_file, kv_level_dict)

	line_data_dict = data.aline
	add_line_db(line_data_dict, db_file, kv_level_dict)

	# print (node_data_dict)
	add_load_db(node_data_dict, db_file, kv_level_dict, profile)

	shunt_data_dict = data.ashunt
	add_shunt_db(shunt_data_dict, db_file, kv_level_dict)

	TRF_2_data_dict = data.atrf2
	add_x2Trans_db(TRF_2_data_dict, db_file, kv_level_dict, node_data_dict)


if __name__ == '__main__':

	# excel_file = 'Inputs12.xlsx'
	excel_file = get_input_excel('Inputs12.xlsx')

	db_file = get_databasefile(excel_file)
	templates_db = 'templates.db'
	data = DATAP(excel_file)
	print(db_file)
	Element_dict_db = get_default_data(templates_db, 'Element')
	#Flag_State : 1(In service)/0(Out service), Name : 'I'+ element_ID, Voltage_level , Flag_Input là gì ????????????????????????????????????

	Terminal_dict_db = get_default_data(templates_db, 'Terminal')
	#Flag_State(Switch line): 0(off), 1(on); Flag_Terminal(Phase): 7(3 phase L123);Flag_Variant: Delete(-1), Operating (1); TerminalNo : Số đầu của element;  Flag_Switch là gì ??????????????????????????

	main(templates_db, data, db_file)
	














	


















