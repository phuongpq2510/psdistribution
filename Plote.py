import sys
import os
import matplotlib.pyplot as plt

from KERNEL import DATAP
sys.stdout.reconfigure(encoding='utf-8')

class Plot():
	def __init__(self,fi):
		self.datap_object = DATAP(fi)
		self.coor_bus = self.get_bus()
		self.line,self.special_line,self.flag = self.get_line()

	def get_bus(self):
		coor_bus= dict()
		abus_data = self.datap_object.abus
		for key in abus_data:
			coor_bus[key] = [abus_data[key]['xCoord'],abus_data[key]['yCoord']]
		return coor_bus
	def get_line(self):
		flag = dict()
		special_line = dict()
		line = dict()
		aline_data = self.datap_object.aline
		for key in aline_data:

			## Line
			line[key] = [aline_data[key]['BUS_ID1'],aline_data[key]['BUS_ID2']]

			#special line
			if aline_data[key]['xCoord'] != None :
				xadd = str(aline_data[key]['xCoord']).split()
				yadd = str(aline_data[key]['yCoord']).split()
				for i in range(len(xadd)):
						special_line.setdefault(key, []).append([float(xadd[i]),float(yadd[i])])

			flag[key] = aline_data[key]['FLAG']
			## flag
			for value in nline_off:
				if value == key :
					flag[key] = 0

		return line,special_line,flag
	def main(self):

		plt.figure(figsize=(10, 6))

		## draw point
		for bus in self.coor_bus.keys():
			x=self.coor_bus[bus][0]
			y=self.coor_bus[bus][1]
			plt.scatter(x, y,color='black')
			plt.annotate(bus,(x+0.1,y+0.02),fontsize=size)




		## draw line
		for i,li in enumerate(self.line.keys()):

			## đường dây
			if li not in self.special_line:
				x=[]
				y=[]
				for bus in self.line[li]:
					x.append(self.coor_bus[bus][0])
					y.append(self.coor_bus[bus][1])

				## Name Line
				self.draw_name_line(x,y,li,i)

				self.plot(x,y,i,li)
				# self.rate_line(bus,li)
			#đường dây gấp khúc
			else:

				for j,spec in enumerate(self.special_line[li]):

					## bus đầu
					if j == 0 :

						bus=self.line[li][0]

						x=[self.coor_bus[bus][0],spec[0]]
						y=[self.coor_bus[bus][1],spec[1]]
						self.plot(x,y,i,li)
						## name_line
						self.draw_name_line(x,y,li,i)

					##bus cuối
					if j==len(self.special_line[li])-1:

						bus=self.line[li][1]
						x=[spec[0],self.coor_bus[bus][0]]
						y=[spec[1],self.coor_bus[bus][1]]
						self.plot(x,y,i,li)
					##bus trung gian
					else:
						x=[self.special_line[li][j][0],self.special_line[li][j+1][0]]
						y=[self.special_line[li][j][1],self.special_line[li][j+1][1]]
						self.plot(x,y,i,li)

		# x=[1,3]
		# y=[0,0]

		# plt.plot(x, y,'k', linestyle='solid',color='r')




	def plot(self,x,y,i,li):
		line_off=()
		color='black'
		if li in line_off:
			plt.plot(x, y, linestyle='--',color=color)
		elif self.flag[li]==1:
			plt.plot(x, y, linestyle='solid',color=color)
		elif self.flag[li]==0:
			plt.plot(x, y, linestyle='--',color=color)
	def draw_name_line(self,x,y,li,i):

		# tọa độ x bằng nhau
		if x[0]==x[1] :
			x1=x[0]+0.25
			y1=(y[0]+y[1])/2
			plt.annotate(li,(x1,y1),fontsize=size,fontstyle='italic')

			plt.plot([x1-0.02,x1+0.3],[y1-0.03,y1-0.03],'k',linestyle='solid')
			print('ok',li)
		## tọa dộ y bằng nhau
		if y[0]==y[1]:
			x1=(x[0]+x[1])/2
			y1=y[1]+0.06
			plt.annotate(li,(x1,y1),fontsize=size,fontstyle='italic')

			plt.plot([x1-0.02,x1+0.3],[y1-0.03,y1-0.03],'k',linestyle='solid')



	def rate_line(self,bus,li):
	    # Tạo các điểm đỉnh của hình chữ nhật
	    x1=0
	    y1=0
	    for bus in self.line[li]:
	    	x1+=self.coor_bus[bus][0]
	    	y1+=self.coor_bus[bus][1]

	    ## tâm của hình vuông
	    x2=x1/2
	    y2=y1/2

	    size1=0.1

	    x3 = [x2-size1, x2+size1, x2+size1, x2-size1,x2-size1 ]
	    y3 = [y2-size1, y2-size1, y2+size1, y2+size1, y2-size1]

	    	# Vẽ đồ thị
	    plt.fill(x3, y3,color='black')
	    plt.axis('equal')
	    plt.grid(True)



def test():
	return

if __name__ == '__main__':
	fi='inputs\\Inputs12.xlsx'
	size=10
	nline_off=[]
	Plot(fi).main()
	plt.show()