### ID Node : ID Node
### The remaining IDs follow the Element's ID

# -------------------------------
import sqlite3
import openpyxl
import shutil
from openpyxl.styles import Alignment
import os
import win32com.client as win32
import sys
class Convert:
    def __init__(self,db_file,excel):
        self.wb=openpyxl.load_workbook(excel)

        self.conn,self.cursor=self.connect_db(db_file)
        ##Node ID, Name_node,  Un at Node
        self.Node = self.get_node()
        print('Node\n', self.Node)

        ##Key: Element , Value: Element_ID
        self.Element = self.get_element()
        print('Element\n',self.Element)

        ## Key:Element_ID, Value: Node_ID 
        self.Line =  self.get_line()
        print('Line\n',self.Line)

        ##Key: Node_ID,Value: P,Q
        self.Load = self.get_load()
        print('Load\n',self.Load)

        ## Key Node, Value:Elemnt, Sn
        self.Shunt = self.get_shunt()
        print('Shunt\n',self.Shunt)


        ## Infeeder Element : Node_ID 
        self.Infeeder = self.get_feeder()
        print('Indeeder\n', self.Infeeder)



    def connect_db(self,db_file):
        conn=sqlite3.connect(db_file)
        cursor=conn.cursor()
        return conn,cursor

    def get_element(self):
        ## Get_ element_ID
        res = {}
        data=['Line','Load','ShuntCondensator','Infeeder']
        for item in data:
            sql=f'SELECT Element_ID FROM Element WHERE Type= "{item}"'
            k = self.cursor.execute(sql)
            list1=[]
            for value in k :
                list1.append(value[0])
            res[item]=list1
        return res

    def get_node(self):
        ## 
        data=['Node_ID','Name','Un']
        res ={}
        for item in data:
            sql=f'SELECT {item} FROM Node'
            k = self.cursor.execute(sql)
            list1=[]
            for value in k:
                 list1.append(value[0])
            res[item] = list1

        res1 = {}
        res2 = {}
        for i, item in enumerate(res['Node_ID']):
            res1[item]=res['Name'][i]
            res2[item]=res['Un'][i]
        res3 = {}
        res3['Node_ID'] = res['Node_ID']
        res3['Name'] = res1
        res3['Un'] = res2

        return res3

    def get_line(self):
        ## Get Line
        ## Key: Name Line, Value: Node_ID
        li={}

        length={}
        r={}
        x={}
        res1={}
        for item in self.Element['Line']:
           
            ## Frombus, Tobus 
            sql=f'SELECT Node_ID FROM Terminal WHERE Element_ID= "{item}"'
            k = self.cursor.execute(sql)
            list1=set()
            for value in k:
                list1.add(value[0])
            li[item]=list1

            ##length
            sql = f'SELECT l,r,x FROM Line WHERE Element_ID= "{item}"'
            self.cursor.execute(sql)
            k= self.cursor.fetchall()
            length[item] = k[0][0]
            r[item] = k[0][1]
            x[item] = k[0][2]

        res1['Line'] = li
        res1['Length'] = length
        res1['r'] = r
        res1['x'] = x

        return res1
    def get_load(self):
        res={}
        for item in self.Element['Load']:

            ## Get PQ
            sql=f'SELECT P,Q FROM Load WHERE Element_ID= "{item}"'
            self.cursor.execute(sql)
            PQ= self.cursor.fetchall()
            
            ## Get Node PQ
            sql1=f'SELECT Node_ID FROM Terminal WHERE Element_ID= "{item}"'
            node = self.cursor.execute(sql1)
            node= self.cursor.fetchall()

            # list P Q
            list1=[]
            for value in PQ:
                list1.append(value[0])
                list1.append(value[1])
            res[node[0][0]]=list1
        return res
    def get_feeder(self):
        list1=[]
        #elemnt
        res={}
        #name 
        res1={}
        #delta
        res2={}
        #u
        res3={}
        #final
        res_f={}

        for item in self.Element['Infeeder']:

            ## Get Element : Node ID
            sql=f'SELECT Node_ID FROM Terminal WHERE Element_ID= "{item}"'
            k = self.cursor.execute(sql)   
            for value in k:
                res[item]=value[0]

            ## Get Name{Element : Name }
            sql=f'SELECT Name FROM Element WHERE Element_ID= "{item}"'
            k = self.cursor.execute(sql)
            for value in k:
                res1[item]=value[0]

            ## Get Delta, U Infeeder
            sql=f'SELECT delta,u FROM Infeeder WHERE Element_ID= "{item}"'
            k = self.cursor.execute(sql)
            
            for value in k:
                res2[item]=value[0]
                res3[item]=value[1]
         
                # res1[item]=value[0]

        res_f['Info']=res
        res_f['Name']=res1
        res_f['aGen']=res2
        res_f['vGen']=res3

        return res_f
        
    def get_shunt(self):
        res = {}
        res2 = {}
        res3 = {}
        for item in self.Element['ShuntCondensator']:
            res1 = {}
            ## Get Node ShuntCondensator
            sql=f'SELECT Node_ID FROM Terminal WHERE Element_ID= "{item}"'
            self.cursor.execute(sql)
            node = self.cursor.fetchall()
            list1=[]
            ##Get Sn: Q shunt
            sql=f'SELECT Sn FROM ShuntCondensator WHERE Element_ID= "{item}"'  
            self.cursor.execute(sql)
            Sn = self.cursor.fetchall()

            if node[0][0] in res:
                # Nếu đã tồn tại, thêm giá trị mới vào danh sách tương ứng
                res[node[0][0]][item] = Sn[0][0]
            else:
                # Nếu chưa tồn tại, tạo danh sách mới
                res1[item] = Sn[0][0]
                res[node[0][0]] = res1
            ## Get name 
            sql=f'SELECT Name FROM Element WHERE Element_ID= "{item}"'  
            
            self.cursor.execute(sql)
            name = self.cursor.fetchall()
            res2[item]=name[0][0]
        res3['Info']=res
        res3['Name']=res2
        return res3
    def get_time_series(self):
        return

    def get_name_column(self,sheet):
        number_of_column={}
        column_order = [col[0].column for col in sheet.iter_cols()]
        for row in sheet.iter_rows(min_row=2,max_row=2):
            for col_num, cell in enumerate(row):
                column_name = column_order[col_num]
                cell_value = cell.value
                number_of_column[cell_value]=column_name
        return number_of_column
    def convert_excel_BUS(self,excel):
        ## Sheet BUS
        sheet=self.wb['BUS']
        number_of_column=self.get_name_column(sheet)
        # Tạo kiểu căn giữa
        center_alignment = Alignment(horizontal='center', vertical='center')
        
        ## Coord X,y
        Coord=self.Graphic_node()  
        row=3
        column=1
        for i,node in enumerate(self.Node['Node_ID']):

            ##Node_ID
            self.value_excel(sheet,center_alignment,node,row,number_of_column['ID'])

            ##Bus_Name
            value1 = self.Node['Name'][node]
            self.value_excel(sheet,center_alignment,value1,row,number_of_column['NAME'])

            ##kV
            value1 = self.Node['Un'][node]
            self.value_excel(sheet,center_alignment,value1,row,number_of_column['kV'])
            ##PQ
            # P=0
            # Q=0
            if node in self.Load:
                P = self.Load[node][0]
                Q = self.Load[node][1]
                ##code
                # self.value_excel(sheet,center_alignment,1,row,number_of_column['CODE'])

                self.value_excel(sheet,center_alignment,P,row,number_of_column['PLOAD [kw]'])
                self.value_excel(sheet,center_alignment,Q,row,number_of_column['QLOAD [kvar]'])
            ## X,Y Coord
            self.value_excel(sheet,center_alignment,Coord[node][0],row,number_of_column['xCoord'])
            self.value_excel(sheet,center_alignment,Coord[node][1],row,number_of_column['yCoord'])
            ##Code Infeeder
            # if node in self.Infeeder:
            #     self.value_excel(sheet,center_alignment,3,row,number_of_column['CODE'])
            row+=1
        return
    def convert_excel_LINE(self,excel):
        sheet=self.wb['LINE']

        ## Get_name column
        number_of_column=self.get_name_column(sheet)
        center_alignment = Alignment(horizontal='center', vertical='center')
        row=3
        column=1

        BucklePoint=self.Graphic_Line()
        for key, value in self.Line['Line'].items():

            ## Element line 
            self.value_excel(sheet,center_alignment,key,row,number_of_column['ID'])
            ##frombus tobus
            i=number_of_column['BUS_ID1']
            i1=number_of_column['NAME1']
            for values in value:
                self.value_excel(sheet,center_alignment,values,row,i)
     
                self.value_excel(sheet,center_alignment,self.Node['Name'][values],row,i1)
                i1+=1
                i+=1

            ##Name1

            ##length
            self.value_excel(sheet,center_alignment,self.Line['Length'][key],row,number_of_column['LENGTH'])
            ## R
            self.value_excel(sheet,center_alignment,self.Line['r'][key],row,number_of_column['R [Ohm/km]'])
            ## X
            self.value_excel(sheet,center_alignment,self.Line['x'][key],row,number_of_column['X [Ohm/km]'])
            ## Buckle Point
            if key in BucklePoint:

                self.value_excel(sheet,center_alignment,BucklePoint[key][0],row,number_of_column['xCoord'])
                self.value_excel(sheet,center_alignment,BucklePoint[key][1],row,number_of_column['yCoord'])
    
            ## R
            row+=1
        return

    def convert_excel_SOURCE(self,excel):
        sheet=self.wb['SOURCE']
        number_of_column=self.get_name_column(sheet)
        center_alignment = Alignment(horizontal='center', vertical='center')
        row=3
        column=1
        for key, value in self.Infeeder['Info'].items():

            ## Element line 
            self.value_excel(sheet,center_alignment,key,row,number_of_column['ID'])
            self.value_excel(sheet,center_alignment,value,row,number_of_column['BUS_ID'])
            self.value_excel(sheet,center_alignment,self.Node['Name'][value],row,number_of_column['NAME'])
            self.value_excel(sheet,center_alignment,self.Node['Un'][value],row,number_of_column['kV'])
            self.value_excel(sheet,center_alignment,self.Infeeder['vGen'][key]/100,row,number_of_column['vGen [pu]'])

            self.value_excel(sheet,center_alignment,self.Infeeder['aGen'][key],row,number_of_column['aGen [deg]'])
            row+=1
        return

    def convert_excel_Shunt(self,excel):
        sheet=self.wb['SHUNT']
        number_of_column=self.get_name_column(sheet)
        center_alignment = Alignment(horizontal='center', vertical='center')
        row=3
        column=1
        for key, value in self.Shunt['Info'].items():
            for key1,value1 in self.Shunt['Info'][key].items():
                # Element line 
                self.value_excel(sheet,center_alignment,key1,row,number_of_column['ID'])
                self.value_excel(sheet,center_alignment,key,row,number_of_column['BUS_ID'])
                self.value_excel(sheet,center_alignment,self.Shunt['Name'][key1],row,number_of_column['NAME'])
                self.value_excel(sheet,center_alignment,self.Node['Un'][key],row,number_of_column['kV'])
                self.value_excel(sheet,center_alignment,value1,row,number_of_column['Qshunt [kvar]'])
                row+=1
        return

    def Graphic_node(self):
        res={}

        for node_id in self.Node['Node_ID']:
            sql=f'SELECT NodeStartX,NodeStartY From GraphicNode WHERE Node_ID= "{node_id}"'
            k = self.cursor.execute(sql)
            
            for value in k:
                res[node_id]=value
        return res

    def Graphic_Line(self):
        res={}
        sql=f'SELECT GraphicTerminal_ID,PosX,PosY From GraphicBucklePoint'
        self.cursor.execute(sql)
        k = self.cursor.fetchall()
        for value in k:
        
            sql1=f'SELECT Element_ID From Terminal WHERE Terminal_ID= "{value[0]}"'
    
            k1= self.cursor.execute(sql1)
            k1=self.cursor.fetchall()
            if k1[0][0] in res:

                res[k1[0][0]][0].append(value[1])
                res[k1[0][0]][1].append(value[2])
            else:
                res[k1[0][0]]=[]
                res[k1[0][0]].append([value[1]])
                res[k1[0][0]].append([value[2]])

        for key in res:        
            for i in range(len(res[key])):
                    # Chuyển mỗi danh sách con thành một chuỗi, các phần tử cách nhau bằng dấu cách
                    res[key][i] = ' '.join(map(str, res[key][i]))    
       
        return res
    def value_excel(self,sheet,center_alignment,value,row,column):
        cell = sheet.cell(row, column)
        cell.value = value
        cell.alignment = center_alignment



    def main(self,excel):
        self.convert_excel_BUS(excel)
        self.convert_excel_LINE(excel)
        self.convert_excel_Shunt(excel)
        self.convert_excel_SOURCE(excel)
        self.wb.save(excel)
        self.wb.close()
        return
def Creat_new_excel():
    path = os.getcwd()
    path_default=path+'\\Default.xlsx'
    path_new=path+'\\Result_File.xlsx'

    ## Check File tồn tại hay chưa 
    if os.path.isfile(path_new):
        
        base_name, extension = os.path.splitext(path_new)
        path_new = f"{base_name}_1{extension}"
        counter = 1
        while os.path.isfile(path_new):
            counter += 1
            path_new = f"{base_name}_{counter}{extension}"


    ## Copy File default
    try:
        # Tạo một phiên làm việc với Excel
        excel = win32.gencache.EnsureDispatch('Excel.Application')

        # Mở tệp Excel mẫu
        wb = excel.Workbooks.Open(path_default)

        # Tạo một bản sao của tệp Excel mẫu
        wb.SaveAs(path_new)

        wb.Close()
        excel.Quit()

    except Exception as e:
        print(f"Error: {e}")

    return path_new
def Set_File():

    ## Creat new file 
    path = os.getcwd()
    nguon=path+'\\Default File'
    dich=path
    new_name='File New'
    duong_dan_dich_moi = os.path.join(dich, new_name)
    shutil.copytree(nguon, duong_dan_dich_moi,dirs_exist_ok=True)

    # copy database to newfile
    file=path+'\\database.db'

    duong_dan_db=duong_dan_dich_moi+'\\Default_files'
    shutil.copy(file, duong_dan_db)

    return

if __name__ == '__main__':
    excel=Creat_new_excel()
    db_file='database.db'
    # excel='test.xlsx'
    convert=Convert(db_file,excel)
   
    # convert.Graphic_Line()
    convert.main(excel)
    Set_File()
