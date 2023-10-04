import os, openpyxl
import shutil 
import pandas as pd
import sqlite3
import math

ELEMENT_ID = 0
TERMINAL_ID = 0

def create_and_rename_folder():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    default_folder_path = os.path.join(current_dir, 'DEFAULT FILE')
    
    new_folder_path = os.path.join(current_dir, 'OUTPUT')
    if os.path.exists(new_folder_path):
        shutil.rmtree(new_folder_path)

    shutil.copytree(default_folder_path, new_folder_path)

    for name in os.listdir(new_folder_path):
        item_path = os.path.join(new_folder_path, name)
        if os.path.isfile(item_path):
            file_name, file_extension = os.path.splitext(name)
            # Đổi tên tệp (điều này chỉ đổi tên, không đổi định dạng)
            new_name = "Output" + file_extension
            os.rename(item_path, os.path.join(new_folder_path, new_name))
        elif os.path.isdir(item_path):
            new_name = "Output"
            os.rename(item_path, os.path.join(new_folder_path, new_name))

    subfolder_name = 'Output'  
    database_file_name = 'database.db' 
    database_file_path = os.path.join(new_folder_path, subfolder_name, database_file_name)
 
    if os.path.exists(database_file_path):
        return database_file_path
    else:
        return None


def creat_dict_from_excel(excel_file, sheet_name):
    try:
        wb = openpyxl.load_workbook(excel_file, data_only=True)

        sheet = wb[sheet_name]

        data_dict = {}
        keys = [cell[0].value for cell in sheet.iter_rows(min_row=3, max_row=sheet.max_row, min_col=1, max_col=1) if cell[0].value is not None]
        data_list = []
        # Lặp qua các hàng từ hàng thứ hai trở đi
        for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1):
            row_values = [cell.value for cell in row[1:]]  # Bỏ đi cột 1
            data_list.append(row_values)
        for sublist in data_list[:]:
            if sublist[0] is None :
                data_list.remove(sublist)
        for key in keys:
            data_dict[key] = dict(zip(data_list[0], data_list[key])) 

        return data_dict
    except Exception as e:
        print("Lỗi:", e)
        return None

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

        print("Dữ liệu đã được chèn hoặc thay thế thành công vào bảng", table_name)

    except sqlite3.Error as e:
        print("Lỗi SQLite:", e)
    finally:
        if conn:
            conn.close()

db_file = create_and_rename_folder()
excel_file = 'Inputs12.xlsx'

bus_data_dict = creat_dict_from_excel(excel_file, 'BUS')
source_data_dict = creat_dict_from_excel(excel_file, 'SOURCE')

# lấy các cấp điện áp từ file excel 
df = pd.read_excel(excel_file, sheet_name='BUS', skiprows=1)
kv_values = list(set(df['kV']))
kv_level_dict = {}
 
#add voltagelevel và variant vào database

Variant_dict = {'Variant_ID': 1, 'VarIndex': 'Base', 'ParentVariant_ID': None, 'Flag_Variant': 1, 'Name': 'Base Variant', 'Revision': None, 'Comment1': None, 'Comment2': None, 'Author': None, 'Created': None, 'ModifiedBy': None, 'Modified': None, 'Pos': 1}
add_data_db(db_file, "Variant", Variant_dict )

voltagelevel_dict = {'VoltLevel_ID': '', 'Variant_ID': 1, 'Flag_Variant': 1, 'Name': '', 'ShortName': '', 'Flag_Volt': 1, 'Un': '', 'Uop': '', 'f': 50.0, 'Temp_Line': 20.0, 'Temp_Cable': 20.0, 'AmbTemp_Line': 20.0, 'AmbTemp_Cable': 20.0, 'TempTrf': 20.0, 'Flag_LimitRes': 0, 'Flag_Sc': 1, 'AddFaultData_ID': 0, 'Flag_CurStp': 1, 'ts': 0.1, 'Flag_Toleranz': 2, 'Flag_Usc': 1, 'c': 1.1, 'Uk': 0.0, 'u_ansi': 1.0, 'Ikmax': 0.0, 'tkr': 1.0, 'Ipmax': 0.0, 'Iamax': 0.0, 'Flag_OptBr': 1, 'Flag_CompPower': 1, 'CosPhi_ind': 0.95, 'CosPhi_kap': -0.95, 'Flag_Balance': 1, 'HarDistLimit_ID': 0, 'hmax_THD': 0, 'fRD': 50.0, 'Flag_Arc_I': 0, 'I_Fkt_R': 1.5, 'I_R': 4.0, 'I_R_X': 0.25, 'Flag_Arc_M': 0, 'M_Fkt_R': 1.5, 'M_R': 4.0, 'M_R_X': 0.25, 'Flag_Arc_K': 0, 'K_Fkt_R': 1.5, 'K_R': 4.0, 'K_R_X': 0.25, 'Flag_Arc_P': 0, 'P_Fkt_R': 1.5, 'P_R': 4.0, 'P_R_X': 0.25, 'Flag_Arc_MHO': 0, 'MHO_Fkt_R': 1.5, 'MHO_R': 4.0, 'MHO_R_X': 0.25, 'Flag_Arc_Comb': 0, 'Comb_Fkt_R': 1.5, 'Comb_R': 4.0, 'Comb_R_X': 0.25, 'LF_Safety_I': 20.0, 'LF_Safety_Phi': 5.0, 'LF_Safety_Z': 20.0, 'Udgr': 0.0, 'Flag_Reliability': 0, 'SwitchBay1_ID': 0, 'SwitchBay2_ID': 0, 'BusbarType_ID': 0, 'LineType_ID': 0, 'CableType_ID': 0, 'TransformerType_ID': 0, 'SupplyType_ID': 0, 'DCInfeederType_ID': 0, 'LoadDurCurve_ID': 0, 'Flag_LP': 3, 'Flag_DCInfeeder': 0, 'cmin': 1.0, 'cmax': 1.1, 'Flag_RelElement': 0, 'SwitchDur_ID': 0, 'dQ_rouch': 10.0, 'dQ_fine': 2.0, 'TempSun_Line': 20.0}
print(kv_values)
for i in range(len(kv_values)):
	voltagelevel_dict['VoltLevel_ID'] = int(i + 1)
	voltagelevel_dict['Name'] = str(kv_values[i]) + ' kV'
	voltagelevel_dict['ShortName'] = str(kv_values[i]) + ' kV'
	voltagelevel_dict['Un'] = float(kv_values[i])
	voltagelevel_dict['Uop']= float(kv_values[i])
	kv_level_dict[int(i + 1)] = kv_values[i]
	print(voltagelevel_dict)
	add_data_db(db_file, "VoltageLevel" , voltagelevel_dict)
print(kv_level_dict)
#-------------------------------------------------------------------------
#add dữ liệu của bus vào database(gồm Node và GraphicNode, GraphicText):

#nodeeeeeeeeeeeeeeeeeeeeeeeeee
def find_and_output_value(database_file, table_name, search_value):
    conn = sqlite3.connect(database_file)
    cursor = conn.cursor()
    cursor.execute(f"SELECT VoltLevel_ID FROM {table_name} WHERE Un = ?", (search_value,))
    
    result = cursor.fetchone()[0]
    return result

node_data = {}
ID_bus = df['ID'].tolist()

for index, row in df.iterrows():

    key = row['ID'] 
    values = row.tolist()
    
    values.pop(0)  
    
    node_data[key] = values
print (ID_bus)
node_dict_db = {'Node_ID': '', 'Variant_ID': 1, 'Flag_Variant': 1, 'Flag_Private': 0, 'Name': '', 'ShortName': '', 'VoltLevel_ID': '', 'Group_ID': 1, 'Zone_ID': 0, 'EcoStation_ID': 0, 'EcoField_ID': 0, 'Flag_Type': 1, 'Busbar_ID': 0, 'Equipment_ID': 0, 'HarPCCdata_ID': 0, 'Stp_ID': 0, 'Ti': None, 'Ts': None, 'Flag_Diagram': 0, 'InclName': '', 'sh': 0.0, 'Flag_Pos': 2, 'hr': 0.0, 'hh': 0.0, 'lat': 0.0, 'lon': 0.0, 'm': 0.0, 'TextVal': None, 'Uref': 0.0, 'Un': '', 'Phi': 0.0, 'Flag_Phase': 7, 'AddFaultData_ID': 0, 'Ik2': 0.0, 'tkr': 1.0, 'Ip': 0.0, 'Iamax': 0.0, 'Uul': 0.0, 'Ull': 0.0, 'Uul1': 0.0, 'Ull1': 0.0, 'Flag_Reliability': 0, 'BusbarType_ID': 0, 'SwitchBay1_ID': 0, 'SwitchBay2_ID': 0, 'Flag_HK': 0, 'Flag_ABW': 0, 'Flag_UM': 0, 'UM_Node_ID': 0, 'T_UM': 0.0, 'Flag_VER': 0, 'VER_Node_ID': 0, 'Flag_VERP': 1, 'T_VER': 0.0, 'p_VER': 0.0, 'ci': 0.0, 'Cs': 0.0, 'cm': 0.0, 'coo': 0.0, 'Tl': 0.0, 'Report_No': 1, 'Flag_Volt': 0, 'RefNode_ID': 0}
print(node_data)
for ID in ID_bus:
	node_dict_db['Node_ID'] = ID
	node_dict_db['Name']= node_data[ID][0]
	node_dict_db['ShortName']= node_data[ID][0]
	node_dict_db['Un']= node_data[ID][1]
	node_dict_db['VoltLevel_ID'] = find_and_output_value(db_file, 'VoltageLevel', node_data[ID][1])
	add_data_db(db_file, "Node" , node_dict_db)

#GraphicNodeeeeeeeeeeeeeeeeeeeeeeeeeeeeee

GraphicNode_dict_db = {'GraphicNode_ID': '', 'Variant_ID': 1, 'Flag_Variant': 1, 'GraphicLayer_ID': 1, 'GraphicType_ID': 1, 'GraphicText_ID1': '', 'GraphicText_ID2': 0, 'Node_ID':'', 'FrgndColor': 0, 'BkgndColor': -1, 'PenStyle': 0, 'PenWidth': 2, 'NodeSize': 4, 'NodeStartX': '', 'NodeStartY': '', 'NodeEndX': '', 'NodeEndY': '', 'SymType': 0, 'Flag': 0, 'GraphicArea_ID': 1}
for ID in ID_bus:
	GraphicNode_dict_db['GraphicNode_ID'] = ID
	GraphicNode_dict_db['GraphicText_ID1'] = ID
	GraphicNode_dict_db['Node_ID'] = ID
	GraphicNode_dict_db['NodeStartX'] = (node_data[ID][8])/100
	GraphicNode_dict_db['NodeStartY'] = (node_data[ID][9])/100
	GraphicNode_dict_db['NodeEndX'] = (node_data[ID][8])/100
	GraphicNode_dict_db['NodeEndY'] = (node_data[ID][9])/100
	add_data_db(db_file, "GraphicNode" , GraphicNode_dict_db)

#GraphicTextttttttttttttttttttttttttttttttttt

GraphicText_dict_db = {'GraphicText_ID': '', 'Variant_ID': 1, 'Flag_Variant': 1, 'GraphicLayer_ID': 1, 'Font': 'Arial', 'FontStyle': 16, 'FontSize': 9, 'TextAlign': 3, 'TextOrient': 0, 'TextColor': 0, 'Visible': 1, 'AdjustAngle': 1, 'Angle': 0.0, 'Pos1': -0.0025, 'Pos2': 0.0, 'Flag': 0, 'RowTextNo': 0, 'AngleTermNo': 0}
for ID in ID_bus:
	GraphicText_dict_db['GraphicText_ID'] = ID
	add_data_db(db_file, "GraphicText" , GraphicText_dict_db)

#-------------------------------------------------------------------------------------

Element_dict_db = {'Element_ID': '', 'Variant_ID': 1, 'Flag_Variant': 1, 'Type': '', 'Flag_Input': 3, 'Flag_Calc': 0, 'Flag_Private': 0, 'Flag_State': '', 'Name': '', 'ShortName': '', 'Description': '', 'VoltLevel_ID': '', 'Group_ID': 1, 'Zone_ID': 0, 'EcoStation_ID': 0, 'EcoField_ID': 0, 'Ti': None, 'Ts': None, 'TextVal': None, 'ci': 0.0, 'Cs': 0.0, 'cm': 0.0, 'coo': 0.0, 'Tl': 0.0, 'Theta_i': 0.0, 'Theta_u': 0.0, 'fCe': 1.0, 'fAkt': 1.0, 'Report_No': 1, 'Metered': 0}
#Flag_State : 1(In service)/0(Out service), Name : 'I'+ element_ID, Voltage_level , Flag_Input là gì ????????????????????????????????????

Terminal_dict_db = {'Terminal_ID': '', 'Variant_ID': 1, 'Flag_Variant': 1, 'Element_ID': '', 'Node_ID': '', 'Flag_State': '', 'TerminalNo': '', 'Flag_Switch': 0, 'Flag_Terminal': 7, 'Flag_Obs': 0, 'Flag_Cur': 0, 'Ir': 0.0, 'Ik2': 0.0, 'Report_No': 1}
#Flag_State(Switch line): 0(off), 1(on); Flag_Terminal(Phase): 7(3 phase L123);Flag_Variant: Delete(-1), Operating (1); TerminalNo : Số đầu của element;  Flag_Switch là gì ??????????????????????????

#-------------------------------------------------------------------------------------
#add dữ liệu cua Infeeder:

Infeeder_dict_db = {'Element_ID': '', 'Variant_ID': 1, 'Flag_Variant': 1, 'Typ_ID': 0, 'Flag_Typ_ID': 0, 'Flag_Typ': 2, 'R': 0.0, 'X': 0.0, 'Rmax': 0.0, 'Xmax': 0.0, 'Rmin': 0.0, 'Xmin': 0.0, 'Sk2': 1000.0, 'R_X': 0.1, 'cact': 1.0, 'Sk2max': 1000.0, 'R_Xmax': 0.1, 'cmax': 1.1, 'Sk2min': 1000.0, 'R_Xmin': 0.1, 'cmin': 1.0, 'xi': 0.0, 'Flag_Lf': '', 'Mpl_ID': 0, 'Start_P': 0.0, 'Start_Q': 0.0, 'Flag_Macro': 0, 'Macro_ID': 0, 'P': 0.0, 'Q': 0.0, 'fP': 1.0, 'fQ': 1.0, 'cosphi': 1.0, 'S': 0.0, 'fS': 1.0, 'I': 0.0, 'fI': 1.0, 'phi': 0.0, 'delta': '', 'u': '', 'delta1': -30.0, 'delta2': -150.0, 'delta3': 90.0, 'u1': 100.0, 'u2': 100.0, 'u3': 100.0, 'Ug': 0.0, 'Rlf': 0.0, 'Xlf': 0.0, 'Flag_Z0': 0, 'Flag_Z0_Input': 1, 'Z0_Z1': 0.0, 'R0_X0': 0.0, 'Z0_Z1max': 0.0, 'R0_X0max': 0.0, 'Z0_Z1min': 0.0, 'R0_X0min': 0.0, 'R0': 0.0, 'X0': 0.0, 'R0max': 0.0, 'X0max': 0.0, 'R0min': 0.0, 'X0min': 0.0, 'Stp_ID': 0, 'DayOpSer_ID': 0, 'YearOpSer_ID': 0, 'WeekOpSer_ID': 0, 'IncrSer_ID': 0, 'Flag_LfLimit': 0, 'ull': 98.0, 'uul': 103.0, 'Pmin': 0.0, 'Pmax': 0.0, 'Qmin': 0.0, 'Qmax': 0.0, 'cosphi_lim': 0.85, 'PowerLimit_ID': 0, 'Flag_ChkType': 1, 'Flag_LimitType': 0, 'Kr': 0.0, 'Flag_Pctrl': 0, 'Qctrl_U_P_ID': 0, 'Flag_Qctrl': 0, 'Qctrl_PF_U_ID': 0, 'Qctrl_PF_P_ID': 0, 'Qctrl_P_Q_ID': 0, 'Qctrl_U_Q_ID': 0, 'Qctrl_Pmax': 1.0, 'Qctrl_Pmin': 0.5, 'Qctrl_PF_Pmin': -0.95, 'Qctrl_PF_Pmax': 0.95, 'Qctrl_Qmin': 0.0, 'Qctrl_Qmax': 0.4843, 'Qctrl_cosphi_c': -0.95, 'Qctrl_cosphi': 0.95, 'Qctrl_U1_c': 97.0, 'Qctrl_U1': 103.0, 'Qctrl_U2_c': 92.0, 'Qctrl_U2': 108.0, 'Flag_CtrlPrior': 0, 'Node_ID': 0, 'u_node': 0.0, 'MasterElm_ID': 0, 'Flag_Har': 1, 'qr': 1.0, 'ql': 0.0, 'HarImp_ID': 0, 'HarVolt_ID': 0, 'HarCur_ID': 0, 'Flag_Reliability': 0, 'SupplyType_ID': 0, 'Flag_ZU': 0, 'Flag_ZUP': 3, 'T_ZU': 0.0, 'Flag_ShdU': 0, 'Flag_ShdP': 0, 'tShd': 0.0, 'Flag_MaxMin': 0, 'Flag_LfCtrl': 0, 'Qctrl_Pstd': 0.0}
#Flag_LF(type bus) = 3(swing bus), 7(PV);'u'; 'delta' ; Flag_Limit (Gioi han P,Q,V)

for key in source_data_dict.keys():
	a = source_data_dict[key]
	ELEMENT_ID += 1
	TERMINAL_ID += 1
	Element_dict_db['Element_ID'] = ELEMENT_ID
	Element_dict_db['Type'] = 'Infeeder'
	Element_dict_db['Flag_State'] = 1 if a['FLAG'] == 1 else 0
	Element_dict_db['Name'] = 'I' + str(ELEMENT_ID)
	Element_dict_db['ShortName'] = 'I' + str(ELEMENT_ID)
	for key, value in kv_level_dict.items():
		if value == a['kV']:
			Element_dict_db['VoltLevel_ID'] = key
	add_data_db(db_file, "Element" , Element_dict_db)

	Infeeder_dict_db['Element_ID'] = ELEMENT_ID
	if a['CODE'] == 0 or None:
		Infeeder_dict_db['Flag_Lf'] = 3
		Infeeder_dict_db['u'] = (a['vGen [pu]'])*100 #đơn vị điện áp là %
		Infeeder_dict_db['delta'] = a['aGen [deg]']
		add_data_db(db_file, "Infeeder", Infeeder_dict_db)
	elif a['CODE'] == 1 : #can kiem tra lai
		Infeeder_dict_db['Flag_Lf'] = 11
		Infeeder_dict_db['u'] = a['vGen [pu]']
		Infeeder_dict_db['P'] = a['Pgen']
		Infeeder_dict_db['Flag_LfLimit'] = 0 #Kiem tra lai elif nay 
		add_data_db(db_file, "Element", Infeeder_dict_db)

	Terminal_dict_db['Terminal_ID'] = TERMINAL_ID
	Terminal_dict_db['Element_ID'] = ELEMENT_ID
	Terminal_dict_db['Node_ID'] = a['BUS_ID']
	Terminal_dict_db['TerminalNo'] = 1 
	Terminal_dict_db['Flag_State'] = 1 #switch feeder
	Terminal_dict_db['Flag_Switch'] = 0 # đang mặc định là 0 
	add_data_db(db_file, "Terminal", Terminal_dict_db)

#-------------------------------------------------------------------------------------
#add dữ liệu của Line:
line_data_dict = creat_dict_from_excel(excel_file, 'LINE')
print (line_data_dict)  

Line_dict_db = {'Element_ID': '', 'Variant_ID': 1, 'Flag_Variant': 1, 'Typ_ID': 0, 'Flag_Typ_ID': 0, 'Flag_LineTyp': 1, 'LineTyp': '', 'Flag_Ll': 0, 'CoupData_ID': 0, 'l': '', 'ParSys': 1.0, 'fr': 1.0, 'r': '', 'x': '', 'c': '', 'va': 0.0, 'fn': 50.0, 'Flag_Cond': 1, 'Un': '', 'Ith': '', 'Ith1': 0.0, 'Ith2': 0.0, 'Ith3': 0.0, 'ElemLoading_ID': 0, 'I1s': 0.0, 'Tend': 0.0, 'q': 0.0, 'LineInfo': '', 'alpha': 0.004, 'Flag_Vart': 1, 'Umax': 0.0, 'd': 50.0, 'da': 50.0, 'Flag_Z0_Input': 1, 'R0_R1': 0.0, 'X0_X1': 0.0, 'r0': 0.0, 'x0': 0.0, 'c0': 0.0, 'rR': 0.0, 'xR': 0.0, 'cR': 0.0, 'q0': 0.0, 'Flag_Ground': 0, 'Flag_Har': 1, 'qr': 1.0, 'ql': 0.0, 'HarImp_ID': 0, 'Flag_ESB': 1, 'Flag_Reliability': 0, 'Flag_SF1': 0, 'Flag_SF2': 0, 'LineType_ID': 0, 'Overload_ID': 0, 'V_S': 0.0, 'Flag_ZU': 0, 'Flag_ZUP': 3, 'T_ZU': 0.0, 'Flag_Tend': 0, 'Flag_Mat': 2, 'fIk': 1.0, 'Flag_Macro': 0, 'Macro_ID': 0, 'LineTemp_ID': 0, 'Flag_Lf': 1}
#Element_ID, l, r, x, c, Un, Ith
for key in line_data_dict.keys():
	a = line_data_dict[key]
	ELEMENT_ID += 1
	TERMINAL_ID += 1
	Line_dict_db['Element_ID'] = ELEMENT_ID
	Line_dict_db['l'] = a['LENGTH [km]']
	Line_dict_db['r'] = a['R [Ohm/km]']
	Line_dict_db['x'] = a['X [Ohm/km]']
	Line_dict_db['c'] = (a['B [microS/km]'])*(1e-3)/(math.pi)
	Line_dict_db['Un'] = a['kV']
	Line_dict_db['Ith'] = a['RATEA [A]']
	add_data_db(db_file, "Line", Line_dict_db)

	Element_dict_db['Element_ID'] = ELEMENT_ID
	Element_dict_db['Type'] = 'Line'
	Element_dict_db['Flag_State'] = a['FLAG']
	Element_dict_db['Name'] = 'L'+str(ELEMENT_ID)
	Element_dict_db['ShortName'] = 'L'+str(ELEMENT_ID)
	for key, value in kv_level_dict.items():
		if value == a['kV']:
			Element_dict_db['VoltLevel_ID'] = key
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
Load_dict_db = {'Element_ID': '', 'Variant_ID': 1, 'Flag_Variant': 1, 'Flag_Load': 1, 'Flag_LoadType': 2, 'Flag_Lf': 1, 'P': '', 'Q': '', 'u': 100.0, 'Ul': 0.0, 'cosphi': 0.0, 'S': 0.0, 'I': 0.0, 'Eapcon': 0.0, 'Erpcon': 0.0, 'P12': 0.0, 'Q12': 0.0, 'P23': 0.0, 'Q23': 0.0, 'P31': 0.0, 'Q31': 0.0, 'P1': 0.0, 'Q1': 0.0, 'P2': 0.0, 'Q2': 0.0, 'P3': 0.0, 'Q3': 0.0, 'E': 0.0, 't': 0.0, 'Eap': 0.0, 'Erp': 0.0, 'fP': 1.0, 'fQ': 1.0, 'fS': 1.0, 'fI': 1.0, 'fEapcon': 1.0, 'fErpcon': 1.0, 'fE': 1.0, 'fEap': 1.0, 'fErp': 1.0, 'Mpl_ID': 0, 'Flag_Macro': 0, 'Macro_ID': 0, 'Load_ID': 0, 'DayOpSer_ID': 2, 'YearOpSer_ID': 0, 'WeekOpSer_ID': 0, 'IncrSer_ID': 0, 'Flag_LA': 0, 'Flag_Z0_Input': 1, 'Z0_Z1': 0.0, 'R0_X0': 0.0, 'R0': 0.0, 'X0': 0.0, 'Stp_ID': 0, 'Pneg': 0.0, 'Qneg': 0.0, 'Flag_Measure': 1, 'P_max': 0.0, 'P_min': 0.0, 'Imax': 0.0, 'I_min': 0.0, 'cos_phi_imax': 0.95, 'cos_phi_imin': 0.75, 'du_min': 0.0, 'du_max': 0.0, 'TransformerTap_ID': 0, 'Flag_Har': 1, 'qr': 1.0, 'ql': 0.0, 'HarImp_ID': 0, 'HarVolt_ID': 0, 'HarCur_ID': 0, 'Flag_I': 1, 'Ireg': 0.0, 'pk': 0.0, 'SatChar_ID': 0, 'ResFlux1': 0.0, 'ResFlux2': 0.0, 'ResFlux3': 0.0, 'Flag_LP': 3, 'CustCnt': 1, 'S_Inst': 0.0, 'S_Peak': 0.0, 'Flag_ShdU': 0, 'Flag_ShdP': 0, 'tShd': 0.0, 'Typ_ID': 0, 'Gang_ID': 0, 'Flag_Typified': 0, 'Flag_Z0': 1}
#Element_ID, P(MW), Q(MVAr)
print (bus_data_dict)
for key in bus_data_dict.keys():
	a = bus_data_dict[key]
	if a['PLOAD'] == None or a['QLOAD'] == None :
		continue
	else:
		ELEMENT_ID += 1
		TERMINAL_ID +=1
		Load_dict_db['Element_ID'] =  ELEMENT_ID
		Load_dict_db['P'] = (a['PLOAD'])/1000 
		Load_dict_db['Q'] = (a['QLOAD'])/1000
		add_data_db(db_file, "Load", Load_dict_db)

		Element_dict_db['Element_ID'] = ELEMENT_ID
		Element_dict_db['Type'] = 'Load'
		Element_dict_db['Flag_State'] = a['FLAG']
		Element_dict_db['Name'] = 'LO'+str(ELEMENT_ID)
		Element_dict_db['ShortName'] = 'LO'+str(ELEMENT_ID)
		for key, value in kv_level_dict.items():
			if value == a['kV']:
				Element_dict_db['VoltLevel_ID'] = key
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
shunt_data_dict = creat_dict_from_excel(excel_file, 'SHUNT')
ShuntReactor_dict_db = {'Element_ID': '', 'Variant_ID': 1, 'Flag_Variant': 1, 'Typ_ID': 0, 'Flag_Typ_ID': 0, 'Flag_Lf': 1, 'Sn': '', 'Vcu': 0.0, 'Vfe': 0.0, 'Un': '', 'Flag_Macro': 0, 'Macro_ID': 0, 'Ci': 0.0, 'Flag_Z0_Input': 1, 'Flag_Z0': 0, 'Z0_Z1': 0.0, 'R0_X0': 0.0, 'R0': 0.0, 'X0': 0.0, 'Stp_ID': 0, 'Flag_roh': 0, 'roh': 0.0, 'rohl': 0.0, 'rohm': 0.0, 'rohu': 0.0, 'deltaS': 0.0, 'Flag_Step': 0, 'Ctrl_OpSer_ID': 0, 'Ctrl_OpPnt_ID': 0, 'Node_ID': 0, 'uul': 103.0, 'ull': 97.0, 'Q_min': 0.0, 'Q_max': 0.0, 'Terminal_ID': 0, 'CosPhi_min': -0.95, 'CosPhi_max': 0.95, 'Flag_Ph': 8, 'Flag_Har': 1, 'qr': 1.0, 'ql': 0.0, 'HarImp_ID': 0, 'SatChar_ID': 0, 'ResFlux1': 0.0, 'ResFlux2': 0.0, 'ResFlux3': 0.0}
#ElementID, Sn, Un
ShuntCondensator_dict_db =  {'Element_ID': '', 'Variant_ID': 1, 'Flag_Variant': 1, 'Typ_ID': 0, 'Flag_Typ_ID': 0, 'Flag_Lf': 1, 'Sn': '', 'Vdi': 0.0, 'Un': '', 'Flag_Macro': 0, 'Macro_ID': 0, 'Ci': 0.0, 'Flag_Z0_Input': 1, 'Flag_Z0': 0, 'Z0_Z1': 0.0, 'R0_X0': 0.0, 'R0': 0.0, 'X0': 0.0, 'Stp_ID': 0, 'Flag_roh': 0, 'roh': 0.0, 'rohl': 0.0, 'rohm': 0.0, 'rohu': 0.0, 'deltaS': 0.0, 'Flag_Step': 0, 'Ctrl_OpSer_ID': 0, 'Ctrl_OpPnt_ID': 0, 'Node_ID': 0, 'uul': 103.0, 'ull': 97.0, 'Q_min': 0.0, 'Q_max': 0.0, 'Terminal_ID': 0, 'CosPhi_min': -0.95, 'CosPhi_max': 0.95, 'Flag_Ph': 8}
#Element_ID, Sn, Un
for key in shunt_data_dict.keys():
	a =shunt_data_dict[key]
	ELEMENT_ID += 1
	TERMINAL_ID += 1
	if a['Qshunt'] >= 0: #(Condensator)
		ShuntCondensator_dict_db['Element_ID'] = ELEMENT_ID
		ShuntCondensator_dict_db['Sn']  = (a['Qshunt'])/1000
		ShuntCondensator_dict_db['Un'] = a['kV']
		add_data_db(db_file, "ShuntCondensator", ShuntCondensator_dict_db)

	else:
		ShuntReactor_dict_db['Element_ID'] = ELEMENT_ID
		ShuntReactor_dict_db['Sn'] = -(a['Qshunt'])/1000
		ShuntReactor_dict_db['Un'] = a['kV']
		add_data_db(db_file, "ShuntReactor", ShuntReactor_dict_db)

	Element_dict_db['Element_ID'] = ELEMENT_ID
	Element_dict_db['Type'] = 'ShuntCondensator'
	Element_dict_db['Flag_State'] = a['FLAG']
	Element_dict_db['Name'] = 'SHC'+str(ELEMENT_ID)
	Element_dict_db['ShortName'] = 'SHC'+str(ELEMENT_ID)
	for key, value in kv_level_dict.items():
		if value == a['kV']:
			Element_dict_db['VoltLevel_ID'] = key
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
TRF_2_data_dict = creat_dict_from_excel(excel_file, "TRF2")
print(TRF_2_data_dict)
TwoWindingTransformer_dict_db = {'Element_ID': '', 'Variant_ID': 1, 'Flag_Variant': 1, 'Typ_ID': 0, 'Flag_Typ_ID': 0, 'Un1': '', 'Un2': '', 'Sn': '', 'Smax': '', 'Smax1': 0.0, 'Smax2': 0.0, 'Smax3': 0.0, 'ElemLoading_ID': 0, 'uk': '', 'ur': '', 'Vfe': '', 'i0': '', 'AddRotate': 0.0, 'VecGrp': 6, 'Flag_Boost': 0, 'Flag_Lf': 1, 'Flag_Macro': 0, 'Macro_ID': 0, 'Flag_Z0_Input': 1, 'C01': 0.0, 'C02': 0.0, 'Z0_Z1': 0.0, 'R0_X0': 0.0, 'R0': 0.0, 'X0': 0.0, 'X0_X1': 0.0, 'R0_R1': 0.0, 'ZABNL': 0.0, 'ZBANL': 0.0, 'ZABSC': 0.0, 'Stp_ID1': 0, 'Stp_ID2': 0, 'Flag_Ct': 0, 'uk_Ct': 16.0, 'ur_Ct': 0.0, 'Flag_ConNode': 0, 'Flag_TapInput': 0, 'TransformerTap_ID': 0, 'rohl': 0.0, 'rohm': 0.0, 'rohu': 0.0, 'alpha': 0.0, 'ukr': 0.0, 'phi': 0.0, 'ukl': 0.0, 'uku': 0.0, 'Flag_Tap': 0, 'roh': 0.0, 'roh1': 0.0, 'roh2': 0.0, 'roh3': 0.0, 'Ctrl_OpSer_ID': 0, 'Ctrl_OpPnt_ID': 0, 'Flag_roh': 0, 'Flag_Ph': 8, 'Flag_Step': 0, 'Node_ID': 0, 'ull': 98.0, 'uul': 103.0, 'Proh': 0.0, 'Proh2': 0.0, 'Qroh': 0.0, 'Qroh2': 0.0, 'TransformerCon_ID': 0, 'CompImp_ID': 0, 'Flag_CompNode': 2, 'CtrlRange_ID': 0, 'MasterElm_ID': 0, 'Flag_Har': 1, 'qr': 1.0, 'ql': 0.0, 'HarImp_ID': 0, 'SatChar_ID': 0, 'ResFlux1': 0.0, 'ResFlux2': 0.0, 'ResFlux3': 0.0, 'tctrl_dyn': 300.0, 'Flag_Reliability': 0, 'Flag_SF1': 0, 'Flag_SF2': 0, 'TransformerType_ID': 0, 'Overload_ID': 0, 'V_S': 0.0, 'Flag_ZU': 0, 'Flag_ZUP': 3, 'T_ZU': 0.0, 'Flag_Inrush': 0, 'U_inrush': 0.0, 'I_inrush': 0.1, 't_inrush': 0.05, 'InCur_ID': 0, 'Flag_Damage': 1, 'c': 1.1, 'UnG': 0.0, 'UnN': 0.0, 'UGmax': 100.0, 'cosphiG': 0.9, 'Flag_Side': 1, 'StpCt_ID': 0, 'RX_ZABNL': 0.0, 'RX_ZBANL': 0.0, 'RX_ZABSC': 0.0}
#Element_ID, Un1, Un2, Sn, Smax, uk, ur, Vfe, i0
for key in TRF_2_data_dict.keys():
	a =TRF_2_data_dict[key]
	ELEMENT_ID += 1
	TERMINAL_ID += 1

	TwoWindingTransformer_dict_db['Element_ID'] = ELEMENT_ID
	TwoWindingTransformer_dict_db['Un1'] = bus_data_dict[a['BUS_ID1']]['kV']
	TwoWindingTransformer_dict_db['Un2'] = bus_data_dict[a['BUS_ID2']]['kV']
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
	for key, value in kv_level_dict.items():
		if value == a['kV']:
			Element_dict_db['VoltLevel_ID'] = key
	add_data_db(db_file, 'Element', Element_dict_db)

	Terminal_dict_db['Terminal_ID'] = TERMINAL_ID
	Terminal_dict_db['Element_ID'] = ELEMENT_ID
	Terminal_dict_db['Node_ID'] = a['BUS_ID1']
	Terminal_dict_db['TerminalNo'] = 1 # 1 terminal
	Terminal_dict_db['Flag_State'] = 1 #switch 
	Terminal_dict_db['Flag_Switch'] = 0 # đang mặc định là 0 
	add_data_db(db_file, "Terminal", Terminal_dict_db)

	Terminal_dict_db['Terminal_ID'] = TERMINAL_ID
	Terminal_dict_db['Element_ID'] = ELEMENT_ID
	Terminal_dict_db['Node_ID'] = a['BUS_ID2']
	Terminal_dict_db['TerminalNo'] = 2 # 1 terminal
	Terminal_dict_db['Flag_State'] = 1 #switch 
	Terminal_dict_db['Flag_Switch'] = 0 # đang mặc định là 0 
	add_data_db(db_file, "Terminal", Terminal_dict_db)
