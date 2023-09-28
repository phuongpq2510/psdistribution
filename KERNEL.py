__author__    = "Dr. Pham Quang Phuong"
__copyright__ = "Copyright 2023"
__license__   = "All rights reserved"
__email__     = "phuong.phamquang@hust.edu.vn"
__status__    = "Released"
__version__   = "1.1.5"
"""
about: ....
"""
import os,time,math
import openpyxl,csv
import argparse
import utils
PATH_FILE,PY_FILE = os.path.split(os.path.abspath(__file__))
PARSER_INPUTS = argparse.ArgumentParser(epilog= "")
PARSER_INPUTS.usage = 'Distribution network analysis Tools'
PARSER_INPUTS.add_argument('-fi' , help = '*(str) Input file path' , default = '',type=str,metavar='')
PARSER_INPUTS.add_argument('-fo' , help = '*(str) Output file path', default = '',type=str,metavar='')
ARGVS = PARSER_INPUTS.parse_known_args()[0]
#
RATEC = 100/math.sqrt(3)

#
def __checkLoop__(busHnd,bus,br):
    setBusChecked = set()
    setBrChecked = set()
    for o1 in busHnd:
        if o1 not in setBusChecked:
            setBusChecked.add(o1)
            tb1 = {o1}
            #
            for i in range(20000):
                if i==19999:
                    raise Exception('Error in checkLoop()')
                tb2 = set()
                for b1 in tb1:
                    for l1 in bus[b1]:
                        if l1 not in setBrChecked:
                            setBrChecked.add(l1)
                            for bi in br[l1]:
                                if bi!=b1:
                                    if bi in setBusChecked or bi in tb2:
                                        return bi
                                    tb2.add(bi)
                if len(tb2)==0:
                    break #ok finish no loop for this group
                setBusChecked.update(tb2)
                tb1=tb2.copy()
    return None
#
def __findBusConnected__(bi1,bset,lset):
    ## find all bus connected to Bus [b1]
    res = set(bi1)
    ba = set(bi1)
    while True:
        ba2 = set()
        for b1 in ba:
            for li in bset[b1]:
                try:
                    for bi in lset[li]:
                        if bi not in res:
                            ba2.add(bi)
                except:
                    pass
        if ba2:
            res.update(ba2)
            ba=ba2
        else:
            break
    return res
#
def getIsland():

    return
#
def __getLineISL__(busC0,setSlack):
    lineISL = set() # line can not be off => ISLAND
    busC = busC0.copy()
    if len(setSlack)==1:
        while True:
            n1 = len(lineISL)
            for k,v in busC.items():
                if len(v)==1:
                    lineISL.update(v)
            if n1==len(lineISL):
                break
            busc1 = dict()
            for k,v in busC.items():
                if len(v)!=1:
                    busc1[k]=v-lineISL
            busC = busc1.copy()
    else:
        while True:
            n1 = len(lineISL)
            for k,v in busC.items():
                if k not in setSlack:
                    if len(v)==1:
                        lineISL.update(v)
            if n1==len(lineISL):
                break
            busc1 = dict()
            for k,v in busC.items():
                if len(v)!=1:
                    busc1[k]=v-lineISL
            busC = busc1.copy()
    return lineISL,busC
#
def readSetting(wbInput,sh1,nmax=500):
    ws = wbInput[sh1]
    res = {}
    for i in range(1,nmax):# row
        si = ws.cell(i,1).value
        if type(si)==str and not si.startswith('##'):
            for j in range(1,nmax):
                sij = ws.cell(i,j).value
                if type(sij)==str and len(sij)>2 and sij[2]=='_':
                    res[sij] = utils.getVal(ws.cell(i+1,j).value)
    return res

# read 1 sheet excel
def readInput1Sheet(wbInput,sh1,nmax=20000):
    res = {}
    setNo = set()
    try:
        ws = wbInput[sh1]
    except:
        return res
    # dem so dong data
    for i in range(2,nmax):
        vi = ws.cell(i,1).value
        if vi==None:
            k=i
            break
        elif i>2:
            if type(vi)!=int:
                raise Exception ('\nID data must be Integer\n\tsheet: '+sh1+'\n\tline: '+str(i))
            if vi in setNo:
                raise Exception ('\nDuplicate ID data\n\tsheet: '+sh1+'\n\tline: '+str(i))
            else:
                setNo.add(vi)
    #
    for i in range(1,nmax):
        v1 = ws.cell(2,i).value
        if v1==None:
            return res
        va = []
        #
        for i1 in range(3,k):
            va.append( utils.getVal(ws.cell(i1,i).value) )
        res[str(v1)]=va
    return res
#
class DATAP:
    def __init__(self,fi):
        wbInput = openpyxl.load_workbook(os.path.abspath(fi),data_only=True)

        #setting
        self.setting = readSetting(wbInput,'SETTING')
        self.sBase = self.setting['PF_Sbase[kva]']

        # bus
        self.abus = readInput1Sheet(wbInput,'BUS')
        self.asource = readInput1Sheet(wbInput,'SOURCE')
        self.ashunt = readInput1Sheet(wbInput,'SHUNT')
        self.aline = readInput1Sheet(wbInput,'LINE')
        self.atrf2 = readInput1Sheet(wbInput,'TRF2')
        self.atrf3 = readInput1Sheet(wbInput,'TRF3')
        self.aprofile = readInput1Sheet(wbInput,'PROFILE')
        self.ashuntPla = readInput1Sheet(wbInput,'SHUNT_PLACEMENT')
        """
        self.busC1        connect of BUS  {b1:[l1,l2,..]}
        self.busC2                    {b1:{l1,l2,..} set format
        self.braC1        connect of BRANCH (LINE/X2,..)                {l1:[b1,b2] }
        self.busC_loop    bus of loop
        self.brIsland     br Island
        self.brLoop       br Loop
        self.brLine        []
        self.brTrf2        []
        self.brTrf3        []
        self.busSlack      []
        self.busAllLst     []
        self.busAllSet     set()
        self.busLoadSet    set()
        self.brAllSet     set()
        """
        #
        self.__checkData__()

        #-----------------------------------------------------------------------
        self.busAllLst = []     # all bus
        self.busLoadSet = set() # bus with load
        for i in range(len(self.abus['ID'])):
            if self.abus['FLAG'][i]==1:
                self.busAllLst.append(self.abus['ID'][i])
                if (self.abus['PLOAD [kw]'][i]!=None and abs(self.abus['PLOAD [kw]'][i])>0) or (self.abus['QLOAD [kvar]'][i]!=None and abs(self.abus['QLOAD [kvar]'][i])>0):
                    self.busLoadSet.add(self.abus['ID'][i])
        self.busAllSet = set(self.busAllLst)

        #
        self.busSlack = []
        for i in range(len(self.asource['ID'])):
            if self.asource['FLAG'][i]==1:#if gen is active
                if 'PGen [kw]' not in self.asource.keys() or self.asource['PGen [kw]'][i]==None or self.asource['PGen [kw]'][i]==0:
                    self.busSlack.append(self.asource['BUS_ID'][i])
        print('busSlack:',self.busSlack)
        #-----------------------------------------------------------------------
        self.busC1 = {b1:[] for b1 in self.busAllLst} # connect of BUS  {b1:[l1,l2,..]}
        self.busC2 = {b1:set() for b1 in self.busAllLst} #{b1:{l1,l2,..} set format
        self.braC1 = {} # connect of BRANCH (LINE/X2,..)                {l1:[b1,b2] }

        # LINE
        self.brLine = []
        for i in range(len(self.aline['ID'])):
            if self.aline['FLAG'][i]==1:#if line is active
                l1 = self.aline['ID'][i]
                b1 = self.aline['BUS_ID1'][i]
                b2 = self.aline['BUS_ID2'][i]
                self.busC1[b1].append(l1)
                self.busC1[b2].append(l1)
                self.braC1[l1] = [b1,b2]
                self.brLine.append(l1)
        # TRF2
        self.brTrf2 = []
        for i in range(len(self.atrf2['ID'])):
            if self.atrf2['FLAG'][i]==1:#if trf2 is active
                l1 = 100000+self.atrf2['ID'][i]
                b1 = self.atrf2['BUS_ID1'][i]
                b2 = self.atrf2['BUS_ID2'][i]
                self.busC1[b1].append(l1)
                self.busC1[b2].append(l1)
                self.braC1[l1] = [b1,b2]
                self.brTrf2.append(l1)
        # TRF3
        self.brTrf3 = []
        for i in range(len(self.atrf3['ID'])):
            if self.atrf3['FLAG'][i]==1:#if trf3 is active
                l1 = 200000+self.atrf3['ID'][i]
                b1 = self.atrf3['BUS_ID1'][i]
                b2 = self.atrf3['BUS_ID2'][i]
                b3 = self.atrf3['BUS_ID3'][i]
                self.busC1[b1].append(l1)
                self.busC1[b2].append(l1)
                self.busC1[b3].append(l1)
                self.braC1[l1] = [b1,b2,b3]
                self.brTrf3.append(l1)
        #
        for k,v in self.busC1.items():
            self.busC2[k]= set(v)

        #
        self.brIsland,self.busC_loop = __getLineISL__(self.busC2,self.busSlack)
        self.brAllSet = set(self.braC1.keys())
        self.brLoop = self.brAllSet - self.brIsland
        print(self.brIsland)
        #print(self.brLoop)
        #print(self.busC1 )
        r1 = __findBusConnected__(self.busSlack,self.busC1,self.braC1)
        ri = list(self.busAllSet-r1)
        if ri:
            raise Exception('\nCheck Input Data ISLAND found with bus(es): '+str(ri))

    def run1Config_WithObjective(self,lineOff=[],shuntOff=[],varFlag=None,option=None,fo=''):
        if varFlag is not None:
            if len(varFlag)!=self.nVar:
                raise Exception('Error size of varFlag')
            lineOff = self.getLineOff(varFlag[:self.nL])
            shuntOff = self.getShuntOff(varFlag[self.nL:])
        #
        v1 = self.run1Config(set(lineOff),set(shuntOff),fo)
        if v1['FLAG']!='CONVERGENCE':
            obj = math.inf
        else: #RateMax[%]    Umax[pu]    Umin[pu]    Algo_PF    option_PF
            obj = v1['DeltaA']
            # constraint
            obj+=self.setting['RateMax[%]'][1]*max(0, v1['RateMax[%]']-self.setting['RateMax[%]'][0])
            obj+=self.setting['Umax[pu]'][1]*max(0, v1['Umax[pu]']-self.setting['Umax[pu]'][0])
            obj+=self.setting['Umin[pu]'][1]*max(0,-v1['Umin[pu]']+self.setting['Umin[pu]'][0])
            #cosP ycau cosP>0.9
            obj+=self.setting['cosPhiP'][1]*max(0,-v1['cosP']+self.setting['cosPhiP'][0])
            #cosN ycau cosN<-0.95
            obj+=self.setting['cosPhiN'][1]*max(0,v1['cosN']-self.setting['cosPhiN'][0])
        #
        lineOff.sort()
        shuntOff.sort()
        v1['Objective'] = obj
        v1['LineOff'] = lineOff
        v1['ShuntOff'] = shuntOff
        return v1

    #
    def checkLoopIsland(self,brOff):
        # check island/loop multi slack ----------------------------------------
        if brOff.intersection(self.brIsland):
            return 'ISLAND'
        # brOff must be in brLoop
        if len(brOff.intersection(self.brLoop))==0:
            return 'LOOP'

        # check ISLAND
        braC1 = self.braC1.copy()
        for br1 in brOff:
            braC1.pop(br1)

        r1 = __findBusConnected__(self.busSlack,self.busC1,braC1)
        ri = list(self.busAllSet-r1)
        if ri:
            raise Exception('\nCheck Input Data ISLAND found with bus(es): '+str(ri))



##        print(self.busC1)
##        print(braC1)
##        brIsland,busC_loop = __getLineISL__(self.busC2,self.busSlack)
##        self.brAllSet = set(self.braC1.keys())
##        self.brLoop = self.brAllSet - self.brIsland
##        if ri:
##            raise Exception('\nCheck Input Data ISLAND found with bus(es): '+str(ri))
##
##        print(self.busC_loop)
##        #
##        self.setLineHnd = self.setLineHndAll-lineOff
##        #
##        self.lineC = {k:self.LINE[k][:2] for k in self.setLineHnd}
##        self.busC = {b1:set() for b1 in self.setBusHnd}
##        for k,v in self.lineC.items():
##            self.busC[v[0]].add(k)
##            self.busC[v[1]].add(k)
##        #
##        r11 = self.setBusHnd.copy()
##        self.busGroup = []# cac bus tuong ung o cac slack khac nhau
##        for bs1 in self.busSlack:
##            r1 = __findBusConnected__(bs1,self.busC,self.lineC)
##            if len(r1.intersection(self.setSlack))>1:
##                return 'LOOP MULTI SLACK'
##            #
##            self.busGroup.append(r1)
##            r11.difference_update(r1)
##        #
##        if r11:
##            return 'ISLAND'
##        #
##        # LOOP
##        if __checkLoop__(self.bus0ISL,self.busC,self.lineC):
##            return 'LOOP'
        #
        return ''
    #
    def __checkData__(self):
        return

# data for Power Flow
class DATAP_PF(DATAP):
    def __init__(self,fi):
        super().__init__(fi)

    #
    def run1Config(self,brOff=set(),shuntOff=set(),fo=''):
        """ run PF 1 config """
        if type(brOff)!=set:
            brOff = set(brOff)
        if type(shuntOff)!=set:
            shuntOff = set(shuntOff)
        #
        brOff.intersection_update(self.brAllSet)

        #
        if self.setting['PF_Algo']=='PSM':
            return self.__run1ConfigPSM__(brOff,shuntOff,fo)
        return
     #
    def __run1ConfigPSM__(self,brOff,shuntOff,fo=''):
        """
        - result (dict): {'FLAG':,'RateMax%', 'Umax[pu]','Umin[pu]','DeltaA','RateMax%'}
        - FLAG (str): 'CONVERGENCE' or 'DIVERGENCE' or 'LOOP' or 'ISLAND'
        - DeltaA: MWH
        """
        #
        t0 = time.time()
        #
        c1 = self.checkLoopIsland(brOff)
        if c1:
            return {'FLAG':c1}

        if 1:
            return
        # ok run PSM
        self.__lineDirection__()
        #print(self.lineC)
        self.__ordCompute__()
        #print(self.ordc)
        #
        res = {'FLAG':'CONVERGENCE','RateMax[%]':0, 'Umax[pu]':0,'Umin[pu]':100,'DeltaA':0,'cosP':0,'cosN':0}
        # B of Line
        BUSb = {}
        for bri,v in self.LINEb.items():
            if bri not in lineOff:
                bfrom = self.lineC[bri][0]
                bto = self.lineC[bri][1]
                #
                if bfrom in BUSb.keys():
                    BUSb[bfrom]+=v
                else:
                    BUSb[bfrom]=v
                #
                if bto in BUSb.keys():
                    BUSb[bto]+=v
                else:
                    BUSb[bto]=v

        # Shunt
        for k1,v1 in self.BUSbs.items():
            if k1 not in shuntOff:
                if k1 in BUSb.keys():
                    BUSb[k1]+=v1
                else:
                    BUSb[k1]=v1
        #
        if fo:
            add2CSV(fo,[[],[time.ctime()],['PF 1Profile','lineOff',str(list(lineOff)),'shuntOff',str(list(shuntOff))]],',')
            #
            rB = [[],['BUS/Profile']]
            rB[1].extend([bi for bi in self.lstBusHnd])
            #
            rL = [[],['LINE/Profile']]
            rL[1].extend([bi for bi in self.lstLineHnd])
            #
            rG = [[],['GEN/Profile']]
            for bi in self.busSlack:
                 rG[1].append(str(bi)+'_P')
                 rG[1].append(str(bi)+'_Q')
                 rG[1].append(str(bi)+'_cosPhi')
            print('File out saved as:',fo)
        #
        va,ra,cosP,cosN = [],[],[1],[-1]
        for pi in self.profileID:
            res['DeltaA']-=self.loadAll[pi].real
            sa1,va1,dia1 = dict(),dict(),dict()# for 1 profile
            for i1 in range(self.nSlack):# with each slack bus
                bs1 = self.busSlack[i1]
                ordc1 = self.ordc[i1]
                ordv1 = self.ordv[i1]
                setBusHnd1 = self.busGroup[i1]
                vbus = {h1:complex(self.Ubase,0) for h1 in setBusHnd1}
                vbus[bs1] = complex(self.genProfile[pi][bs1],0)
                #
                du,di = dict(),dict()
                s0 = 0
                for ii in range(self.iterMax+1):
                    sbus = {k:v for k,v in self.loadProfile[pi].items() if k in setBusHnd1}
                    # B of Line + Shunt
                    for k1,v1 in BUSb.items():
                        if k1 in setBusHnd1:
                            vv = abs(vbus[k1])
                            sbus[k1] += complex(0, -vv*vv*v1)
                    # cal cong suat nguoc
                    for bri in ordc1:
                        bfrom = self.lineC[bri][0]
                        bto = self.lineC[bri][1]
                        rx = self.LINE[bri][2]
                        #
                        du[bri] = sbus[bto].conjugate()/vbus[bto].conjugate()*rx
                        ib = abs(sbus[bto]/vbus[bto])
                        di[bri] = ib
                        ds1 = ib*ib*rx
                        #
                        if ds1.real>0.2 and ds1.real>sbus[bto].real:# neu ton that lon hon cong suat cua tai
                            return {'FLAG':'DIVERGENCE'}
                        #
                        sbus[bfrom]+=ds1+sbus[bto]
                    # cal dien ap xuoi
                    for bri in ordv1:
                        bfrom = self.lineC[bri][0]
                        bto = self.lineC[bri][1]
                        vbus[bto]=vbus[bfrom]-du[bri]
                    #
                    if abs(s0-sbus[bs1])<self.epsilon:
                        break
                    else:
                        s0 = sbus[bs1]
                    #
                    if ii==self.iterMax:
                        return {'FLAG':'DIVERGENCE'}
                # finish
                # loss P
                res['DeltaA']+=sbus[bs1].real
                # Umax[pu]/Umin[pu]
                va.extend( [abs(v) for v in vbus.values()] )
                #
                try:
                    if sbus[bs1].imag>=0:
                        cosP.append(sbus[bs1].real/abs(sbus[bs1]))
                    else:
                        cosN.append(-sbus[bs1].real/abs(sbus[bs1]))
                except:
                    pass
                # RateMax
                for bri in ordc1:
                    ra.append( di[bri]/self.LINE[bri][3]*RATEC )
                #
                if fo:
                    va1.update(vbus)
                    dia1.update(di)
                    sa1.update(sbus)
            #
            if fo:
                rb1 = [pi]
                rl1 = [pi]
                rg1 = [pi]
                for bi1 in self.lstBusHnd:
                    rb1.append(toString(abs(va1[bi1])/self.Ubase))
                #
                for bri in self.lstLineHnd:
                    try:
                        r1 = dia1[bri]/self.LINE[bri][3]*RATEC
                        rl1.append( toString(r1,2) )
                    except:
                        rl1.append('0')
                #
                for bs1 in self.busSlack:
                    rg1.append(toString(sa1[bs1].real))
                    rg1.append(toString(sa1[bs1].imag))
                    if sa1[bs1].imag>=0:
                        rg1.append(toString(sa1[bs1].real/abs(sa1[bs1]),3))
                    else:
                        rg1.append(toString(-sa1[bs1].real/abs(sa1[bs1]),3))
                #
                rB.append(rb1)
                rL.append(rl1)
                rG.append(rg1)
        #
        va.sort()
        res['Umax[pu]'] = va[-1]/self.Ubase
        res['Umin[pu]'] = va[0]/self.Ubase
        res['RateMax[%]'] = max(ra)
        res['cosP'] = min(cosP)
        res['cosN'] = max(cosN)
        #
        if fo:
            rB.append(['','Umax[pu]',toString(res['Umax[pu]']),'Umin[pu]',toString(res['Umin[pu]']) ])
            add2CSV(fo,rB,',')
            #
            rL.append(['','RateMax[%]',toString(res['RateMax[%]'],2)])
            add2CSV(fo,rL,',')
            #
            rG.append(['','cosPmin',toString(res['cosP'],3),'cosNMax',toString(res['cosN'],3)])
            add2CSV(fo,rG,',')
        #
        return res
    #
    def __ordCompute__(self):
        busC = dict() # connect [LineUp,[LineDown]]
        for h1 in self.setBusHnd:
            busC[h1] = [0,set()]
        #
        for h1,l1 in self.lineC.items():
            busC[l1[1]][0]= h1     # frombus
            busC[l1[0]][1].add(h1) # tobus
        #
        self.ordc,self.ordv = [],[]
        for bs1 in self.busGroup:
            busC1 = {k:v for k,v in busC.items() if k in bs1}
            balr = {h1:True for h1 in bs1}
            sord = set()
            ordc1 = []
            for k,v in busC1.items():
                if len(v[1])==0:
                    if v[0]!=0:
                        ordc1.append(v[0])
                        sord.add(v[0])
                        balr[k]=False
            #
            for ii in range(500):
                for k,v in busC1.items():
                    if balr[k]:
                        if len(v[1]-sord)==0:
                            if k in self.setSlack:
                                break
                            #
                            if v[0]!=0:
                                ordc1.append(v[0])
                            sord.add(v[0])
                            balr[k]=False
            ordv1 = [ordc1[-i-1]  for i in range(len(ordc1))]
            self.ordc.append(ordc1)
            self.ordv.append(ordv1)
    #
    def __lineDirection__(self):
        ba = self.busSlack[:]
        lset = set()
        for ii in range(20000):
            ba2 = []
            for b1 in ba:
                for l1 in self.setLineHnd:
                    if l1 not in lset:
                        if b1==self.lineC[l1][1]:
                            d = self.lineC[l1][0]
                            self.lineC[l1][0] = self.lineC[l1][1]
                            self.lineC[l1][1] = d
                            lset.add(l1)
                            ba2.append(d)
                        elif b1==self.lineC[l1][0]:
                            lset.add(l1)
                            ba2.append(self.lineC[l1][1])
            if len(ba2)==0:
                break
            ba= ba2.copy()
# data for Recloser Optim
class DATAP_REOP(DATAP):
    def __init__(self,fi):
        super().__init__(fi)
    #
    def getData(self):
        # bus
##        self.abus = readInput1Sheet(wbInput,'BUS')
##        self.asource = readInput1Sheet(wbInput,'SOURCE')
##        self.ashunt = readInput1Sheet(wbInput,'SHUNT')
##        self.aline = readInput1Sheet(wbInput,'LINE')
##        self.atrf2 = readInput1Sheet(wbInput,'MBA2')
##        self.atrf3 = readInput1Sheet(wbInput,'MBA3')
##        self.aprofile = readInput1Sheet(wbInput,'PROFILE')
##        self.ashuntPla = readInput1Sheet(wbInput,'SHUNT_PLACEMENT')

        ns = 0
        for i in range(len(self.asource['ID'])):
            b1 = self.asource['BUS_ID'][i]
            if self.asource['FLAG'][i]==1:#if gen is active
                ns+=1
                self.busSlack=b1
        if ns==0:
            raise Exception('\nError: Not found source (Feeder) Bus: check sheet SOURCE')
        if ns>1:
            raise Exception('\nError: Too much source (Feeder) Bus: check sheet SOURCE')
        #
        self.BusC = dict() # connect [LineUp,[LineDown]]
        for b1 in self.busAllLst:
            self.BusC[b1] = [0,[]]
        #
        self.BusD = dict() # Data [pLoad,nLoad,pLoadPU,nLoadPU]
        na,pa = 0,0
        for i in range (len(self.busAllLst)):
            if self.abus['FLAG'][i]==1:
                h1 = self.abus['ID'][i]
                p1 = self.abus['PLOAD [kw]'][i] if self.abus['PLOAD [kw]'][i]!=None else 0
                np1 = self.abus['nLOAD'][i] if self.abus['nLOAD'][i]!=None else 0
                self.BusD[h1] = [p1,np1,0,0]
                pa+=p1
                na+=np1
        #
        for k in self.BusD.keys():
            self.BusD[k][2] = self.BusD[k][0]/pa
            self.BusD[k][3] = self.BusD[k][1]/na
        #print(self.BusD)

        # LINE
        self.LineC = dict() # [fromBus,toBus]
        self.LineD = dict() # [FLAG2,LENGTH,nFault,LengthPU]
        for i in range (len(self.aline['ID'])):
            if self.aline['FLAG'][i]==1:
                h1 = self.aline['ID'][i]
                frombus = self.aline['FROMBUS'][i]
                tobus = self.aline['TOBUS'][i]
                leng = self.aline['LENGTH'][i]
                nFault = self.aline['nFault[per km per year]'][i]
                flag2 = self.aline['FLAG2'][i]
                #
                self.LineC[h1] = [frombus,tobus]
                self.LineD[h1] = [flag2,leng,nFault,0]
##        print(self.LineC)
        return
    #
    def run(self):
        r1 = self.checkIsland()
        if r1:
            raise Exception(r1)

    #
    def checkIsland(self):
        r1 = __findBusConnected__(self.busSlack[0],self.busC1,self.braC1)
        rn = self.busAllSet - r1
        if rn:
            return '\nISLAND FOUND with BUS : '+str(list(rn))
        return ''
#
def test_ReOp():
    ARGVS.fi = 'inputs\\Inputs12.xlsx'
    p1 = DATAP_REOP(ARGVS.fi)
##    p1.getData()
    p1.run()

#
def test_psm():
    # 1 source
    ARGVS.fi = 'inputs\\Inputs12.xlsx'
##    varFlag = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 12, 13, 14, 15, 16,0,1]
    brOff = [100,130,14,15] # {13,14,15}
    shuntOff = []
    #

    p1 = DATAP_PF(ARGVS.fi)
##    v1 = p1.run1Config(brOff,fo=ARGVS.fo)
##    print(v1)
##    v1 = p1.run1Config_WithObjective(lineOff=lineOff,shuntOff=shuntOff,fo=ARGVS.fo)
##    print('time %.5f'%(time.time()-t01))
##    print(v1)
#
if __name__ == '__main__':
    ARGVS.fo = PATH_FILE+'\\res\\res1Config.csv'
    test_psm()
##    test_ReOp()
