import sqlite3
import openpyxl
import sys
import pandas as pd
from openpyxl.styles import Alignment
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

        ## Key Node, Value:Sn
        self.Shunt = self.get_shunt()
        print('Shunt\n',self.Shunt)
        ## Infeeder : Node_ID 
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
        return res

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
            print(k[0][1])
        res1['Line'] = li
        res1['Length'] = length
        res1['r'] = r
        res1['x'] = x
        print(res1)
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
        for item in self.Element['Infeeder']:
            sql=f'SELECT Node_ID FROM Terminal WHERE Element_ID= "{item}"'
            k = self.cursor.execute(sql)   
            for value in k:
                list1.append(value[0])
        return list1
        
    def get_shunt(self):
        res = {}
        
        for item in self.Element['ShuntCondensator']:

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
                res[node[0][0]] += Sn[0][0]
            else:
                res[node[0][0]] = Sn[0][0]          
        return res
    def get_time_series(self):
        return
    def convert_excel_BUS(self,excel):
        ## Sheet BUS
        sheet=self.wb['BUS']
       
        number_of_column={}
        column_order = [col[0].column for col in sheet.iter_cols()]
        for row in sheet.iter_rows(min_row=2,max_row=2):
            for col_num, cell in enumerate(row):
                column_name = column_order[col_num]
                cell_value = cell.value
                number_of_column[cell_value]=column_name
    
        # Tạo kiểu căn giữa
        center_alignment = Alignment(horizontal='center', vertical='center')
        
        row=3
        column=1
        for i,node in enumerate(self.Node['Node_ID']):

            ##Node_ID
            self.value_excel(sheet,center_alignment,node,row,number_of_column['NO'])
            
            ##Bus_Name
            value1 = self.Node['Name'][i]
            self.value_excel(sheet,center_alignment,value1,row,number_of_column['NAME'])

            ##kV
            value1 = self.Node['Un'][i]
            self.value_excel(sheet,center_alignment,value1,row,number_of_column['kV'])

            ##Shunt
            Shunt=0
            if node in self.Shunt:
                Shunt = self.Shunt[node]
            self.value_excel(sheet,center_alignment,Shunt,row,number_of_column['Vscheduled[pu]'])

            ##PQ
            P=0
            Q=0
            if node in self.Load:
                P = self.Load[node][0]
                Q = self.Load[node][1]
                ##code
                self.value_excel(sheet,center_alignment,1,row,number_of_column['CODE'])
            self.value_excel(sheet,center_alignment,P,row,number_of_column['PLOAD[kw]'])
            self.value_excel(sheet,center_alignment,Q,row,number_of_column['QLOAD[kvar]'])
            
            ##Code Infeerder
            if node in self.Infeeder:
                self.value_excel(sheet,center_alignment,3,row,number_of_column['CODE'])
            row+=1

        return




    def convert_excel_LINE(self,excel):
        sheet=self.wb['LINE']

        ## Get_name column
        number_of_column={}
        column_order = [col[0].column for col in sheet.iter_cols()]
        for row in sheet.iter_rows(min_row=2,max_row=2):
            for col_num, cell in enumerate(row):
                column_name = column_order[col_num]
                cell_value = cell.value
                number_of_column[cell_value]=column_name
 
        center_alignment = Alignment(horizontal='center', vertical='center')
        row=3
        column=1
        for key, value in self.Line['Line'].items():

            ## Element line 
            self.value_excel(sheet,center_alignment,key,row,number_of_column['NO'])
            ##frombus tobus
            i=number_of_column['FROMBUS']
            for values in value:
                self.value_excel(sheet,center_alignment,values,row,i)
                i+=1
            ##length
            self.value_excel(sheet,center_alignment,self.Line['Length'][key],row,number_of_column['LENGTH'])
            ## R
            self.value_excel(sheet,center_alignment,self.Line['r'][key],row,number_of_column['R(Ohm)'])
            ## X
            self.value_excel(sheet,center_alignment,self.Line['x'][key],row,number_of_column['X(Ohm)'])
            row+=1
        
        return
    def value_excel(self,sheet,center_alignment,value,row,column):
        cell = sheet.cell(row, column)
        cell.value = value
        cell.alignment = center_alignment

    def main(self,excel):
        self.convert_excel_BUS(excel)
       
        self.convert_excel_LINE(excel)
        self.wb.save(excel)
        self.wb.close()
        return
if __name__ == '__main__':
    db_file='database.db'
    excel='Inputs12bc_1.xlsx'
    convert=Convert(db_file,excel)
    convert.main(excel)