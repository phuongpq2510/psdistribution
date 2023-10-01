__author__    = "Dr. Pham Quang Phuong"
__copyright__ = "Copyright 2023"
__license__   = "All rights reserved"
__email__     = "phuong.phamquang@hust.edu.vn"
__status__    = "in Dev"
__version__   = "2.0.3"
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
def checkLoop(busHnd,bus,br):
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
def findBusConnected(bi1,bset,lset):
    ## find all bus connected to Bus [b1]
    if type(bi1)==int:
        res = {bi1}
        ba = {bi1}
    else:
        res = set(bi1)
        ba = set(bi1)
    while True:
        ba2 = set()
        for b1 in ba:
            for li in bset[b1]:
                for bi in lset[li]:
                    if bi not in res:
                        ba2.add(bi)
        if ba2:
            res.update(ba2)
            ba=ba2
        else:
            break
    return res
#
def getIsland(busC0,busSlack,flagSlack=False):
    lineISL = set() # line can not be off => ISLAND
    busC = busC0.copy()
    while True:
        n1 = len(lineISL)
        for k,v in busC.items():
            if flagSlack or k not in busSlack:
                if len(v)==1:
                    lineISL.update(v)
        if n1==len(lineISL):
            break
        busc1 = dict()
        for k,v in busC.items():
            if (not flagSlack and k in busSlack) or len(v)!=1:
                busc1[k]=v-lineISL
        busC = busc1.copy()
    return lineISL,busC

#
class DATAP:
    def __init__(self,fi,Raw=1): # Raw = 0 ignore out of service object
        wbInput = openpyxl.load_workbook(os.path.abspath(fi),data_only=True)

        #setting
        self.setting = utils.readSetting(wbInput,'SETTING')
        # bus
        self.abus = utils.readInput1Sheet(wbInput,'BUS',Raw=Raw)
        self.asource = utils.readInput1Sheet(wbInput,'SOURCE',Raw=Raw)
        self.ashunt = utils.readInput1Sheet(wbInput,'SHUNT',Raw=Raw)
        self.aline = utils.readInput1Sheet(wbInput,'LINE',Raw=Raw)
        self.atrf2 = utils.readInput1Sheet(wbInput,'TRF2',Raw=Raw)
        self.atrf3 = utils.readInput1Sheet(wbInput,'TRF3',Raw=Raw)
        self.aprofile = utils.readInput1Sheet(wbInput,'PROFILE',Raw=Raw)
        self.ashuntPla = utils.readInput1Sheet(wbInput,'SHUNT_PLACEMENT',Raw=Raw)

        """
        self.busC1       connect of BUS                  {b1:{l1,l2,..}
        self.braC1       connect of BRANCH (LINE/X2,..)  {l1:[b1,b2] }
        self.busC0       ignore island bus/branch
        self.braC0       ignore island bus/branch
        self.busC2       ignore brOff
        self.braC2       ignore brOff
        self.braA2       all branch ignore brOff

        self.brIsland     br Island
        self.brLoop       br loop
        self.brLine        []
        self.brTrf2        []
        self.brTrf3        []
        self.busSlack      []
        self.busAllLst     []
        self.busAllSet     set()
        self.busAll0       set()     ignore island bus/branch
        self.brAllSet      set()
        self.shuntAllSet   set()
        """
        if self.setting['GE_PowerUnit'][0] not in {'kw','mw'}:
            raise Exception('Error PowerUnit not in kw,mw')
        self.sbase0 = self.setting['GE_Sbase'][0]
        if self.setting['GE_PowerUnit'][0]=='kw':
            self.sbase = self.sbase0*1e3
        else:
            self.sbase = self.sbase0*1e6
        self.algoPF = self.setting['PF_Algo'][0]
        self.zzero = float(self.setting['GE_ZSwitch'][0])
        #-----------------------------------------------------------------------
        self.busAllLst = list(self.abus.keys())
        self.busAllSet = set(self.busAllLst)

        #
        self.busSlack = []
        for k,v in self.asource.items():
            if v['CODE']==None or v['CODE']==0:
                self.busSlack.append(v['BUS_ID'])
        print('busSlack:',self.busSlack)

        #-----------------------------------------------------------------------
        self.busC1 = {b1:set() for b1 in self.busAllLst}
        self.braC1 = {}

        # LINE
        for k,v in self.aline.items():
            b1,b2 = v['BUS_ID1'],v['BUS_ID2']
            self.busC1[b1].add(k)
            self.busC1[b2].add(k)
            self.braC1[k] = [b1,b2]
        # TRF2
        for k,v in self.atrf2.items():
            b1,b2 = v['BUS_ID1'],v['BUS_ID2']
            self.busC1[b1].add(k)
            self.busC1[b2].add(k)
            self.braC1[k] = [b1,b2]
        # TRF3
        for k,v in self.atrf3.items():
            b1,b2,b3 = v['BUS_ID1'],v['BUS_ID2'],v['BUS_ID3']
            self.busC1[b1].add(k)
            self.busC1[b2].add(k)
            self.busC1[b3].add(k)
            self.braC1[k] = [b1,b2,b3]
        #
        self.brAllSet = set(self.braC1.keys())
        self.shuntAllSet = set(self.ashunt.keys())
        #
        bra,self.busC0 = getIsland(self.busC1,self.busSlack,flagSlack=False)
        self.braC0 = self.braC1.copy()
        for br1 in bra:
            self.braC0.pop(br1)
        #
        self.busAll0 = set(self.busC0.keys())
        #
        self.brIsland,_ = getIsland(self.busC1,self.busSlack,flagSlack=len(self.busSlack)==1)
        self.brLoop = self.brAllSet - self.brIsland
        #
        print('brIsland:',self.brIsland)
        print('brLoop:',self.brLoop)
        #
        r1 = findBusConnected(self.busSlack,self.busC1,self.braC1)
        ri = self.busAllSet-r1
        if ri:
            rp = '\nCheck Input Data ISLAND found with bus(es): '
            for b1 in ri:
                rp+='\n\t'+self.strBus(b1)
            raise Exception(rp)
    #
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
    def checkLoopIsland(self,brOff,verbose=True):
        # check island/loop multi slack ----------------------------------------
        bri = brOff.intersection(self.brIsland)
        if bri:
            if verbose:
                print('\nCheck Input Data ISLAND: ')
                for br1 in bri:
                    print('\t',self.strBranch(br1))
            return 'ISLAND'

        # brOff must be in brLoop
        if len(brOff.intersection(self.brLoop))==0:
            if verbose:
                brt = self.brLoop-brOff
                print('\nCheck Input Data LOOP found with branches: ')
                for br1 in brt:
                    print('\t',self.strBranch(br1))
            return 'LOOP'

        # check ISLAND
        braCx = self.braC0.copy()
        for br1 in brOff:
            braCx.pop(br1)
        #
        busCx = self.busC0.copy()
        for br1 in brOff:
            for b1 in self.braC1[br1]:
                busCx[b1].remove(br1)
        #
        r1i = self.busAll0.copy()
        for bs1 in self.busSlack:
            r1a = findBusConnected(bs1,busCx,braCx)
            if len(r1a.intersection(self.busSlack))>1:
                if verbose:
                    print('\nCheck Input Data LOOP (multi Slack) found with buses: ')
                    for bii in r1a:
                        print('\t',self.strBus(bii))
                return 'LOOP'
            r1i.difference_update(r1a)
        if r1i:
            if verbose:
                print('\nCheck Input Data ISLAND found with bus(es): ')
                for bii in r1i:
                    print('\t',self.strBus(bii))
            return 'ISLAND'

        # CHECK LOOP: method1
        brIsland,_ = getIsland(busCx,self.busSlack,flagSlack=False)
        brt = self.brLoop - brIsland - brOff
        if brt:
            if verbose:
                for b1 in brt:
                    try:
                        r1 = findBusConnected(b1,busCx,braCx)
                        print('\nCheck Input Data LOOP found with buses: ')
                        for bii in r1:
                            print('\t',self.strBus(bii))
                        break
                    except:
                        pass
            return 'LOOP'
        return ''
    #
    def strBranch(self,br1):
        if type(br1)==int:
            try:
                return str(br1)+" LINE: %s"%self.strBus(self.aline[br1]['BUS_ID1'])+' - '+self.strBus(self.aline[br1]['BUS_ID2']) +" '%s'"%str(self.aline[br1]['CID'])
            except:
                raise Exception('Branch not found: '+str(br1))
        return [self.strBranch(bii) for bii in br1]
    #
    def strBus(self,b1):
        if type(b1)==int:
            try:
                return str(b1)+" '"+self.abus[b1]['NAME']+"' "+str(self.abus[b1]['kV'])+' kV'
            except:
                raise Exception('Bus not found: '+str(b1))
        return [self.strBus(bi) for bi in b1]

# data for Power Flow
class DATAP_PF(DATAP):
    def __init__(self,fi):
        super().__init__(fi,0)
        self._getProfile()
        # get RX
        self._getRXB()
    #
    def _getRXB(self):
        self.braRX = dict()
        self.braB = dict() # for Branch
        self.shuntB = dict() # for Shunt
        #
        for k,v in self.aline.items():
            kv = v['kV']
            zbase = kv*kv*10e6/self.sbase
            l1 = v['LENGTH [km]']
            #
            r1 = v['R [Ohm/km]']*l1/zbase
            x1 = v['X [Ohm/km]']*l1/zbase
            if abs(r1)<self.zzero and abs(x1)<self.zzero :
                x1 = self.zzero
                r1 = 0
            self.braRX[k] = complex(r1,x1)
            #
            b1_2 = v['B [microS/km]']*l1/2*1e-6*zbase if v['B [microS/km]']!=None else 0
            if abs(b1_2)>1e-9:
                self.braB[k] = [b1_2,b1_2]
        # for TRF2
        # for TRF3
        #
        for k,v in self.ashunt.items():
            b1 = v['BUS_ID']
            q1 = v['Qshunt']/self.sbase0
            p1 = v['deltaP']/self.sbase0
            self.shuntB[k] = [b1,p1,q1]
        #print('braRX',self.braRX)
        #print('braB',self.braB)
        #print('shuntB',self.shuntB)
    #
    def _getProfile(self):
        # PROFILE
        nameProfile = []
        YesProfile = False
        for k,v in self.aprofile.items():
            for v1 in v.keys():
                if v1 not in {'deltaTime', 'MEMO'}:
                    nameProfile.append(v1)
            break
        #
        if len(nameProfile)>0:
            for k,v in self.abus.items():
                if str(v['Load Profile']) in nameProfile:
                    YesProfile = True
                    break
            #
            for k,v in self.asource.items():
                if str(v['vGen Profile']) in nameProfile or str(v['pGen Profile']) in nameProfile:
                    YesProfile = True
                    break
        if YesProfile:
            self.IDProfile = list(self.aprofile.keys())
            for k01 in [0,'0','00','01','0001','0002','0003']:
                if k01 not in self.aprofile.keys():
                    self.IDProfile.insert(0,k01)
                    break
        else:
            self.IDProfile = [0]
        print('Profile: ',YesProfile,nameProfile,self.IDProfile)
        #
        self.load = {k:dict() for k in self.IDProfile}
        self.vgen = {k:dict() for k in self.IDProfile}
        self.pgen = {k:dict() for k in self.IDProfile}
        #
        for k,v in self.abus.items():
            p1 = self.abus[k]['PLOAD']/self.setting['GE_Sbase'][0] if self.abus[k]['PLOAD']!=None else 0
            q1 = self.abus[k]['QLOAD']/self.setting['GE_Sbase'][0] if self.abus[k]['QLOAD']!=None else 0
            if p1!=0.0 and q1!=0.0:
                pf1 = self.abus[k]['Load Profile']
                if pf1 in {'deltaTime', 'MEMO'}:
                    raise Exception('\nError Load Profile at BUS: '+self.strBus(k)+ '\n\tLoad Profile: '+pf1)
                for pfr1 in self.IDProfile:
                    try:
                        kp = self.aprofile[pfr1][pf1]
                    except:
                        kp =1
                    self.load[pfr1][k] = complex(p1*kp,q1*kp)
        #print(self.load)
        for k,v in self.asource.items():
            k1 = self.asource[k]['BUS_ID']
            v1 = self.asource[k]['vGen [pu]']
            p1 = self.asource[k]['Pgen']/self.sbase0 if self.asource[k]['Pgen']!=None else 0
            pf1 = self.asource[k]['vGen Profile']
            if pf1 in {'deltaTime', 'MEMO'}:
                raise Exception('\nError vGen Profile at BUS: '+self.strBus(k1)+ '\n\tvGen Profile: '+pf1)
            pf2 = self.asource[k]['pGen Profile']
            if pf2 in {'deltaTime', 'MEMO'}:
                raise Exception('\nError pGen Profile at BUS: '+self.strBus(k1)+ '\n\tpGen Profile: '+pf2)
            for pfr1 in self.IDProfile:
                try:
                    kv = self.aprofile[pfr1][pf1]
                except:
                    kv = 1
                #
                try:
                    kp = self.aprofile[pfr1][pf2]
                except:
                    kp = 1
                self.vgen[pfr1][k1] = v1*kv
                self.pgen[pfr1][k1] = p1*kp
        #print(self.vgen)
        #print(self.pgen)
    #
    def run1Config(self,brOff=set(),shuntOff=set(),fo=''):
        """ run PF 1 config """
        if type(brOff)!=set:
            brOff = set(brOff)
        if type(shuntOff)!=set:
            shuntOff = set(shuntOff)
        #
        brOff.intersection_update(self.brAllSet)
        shuntOff.intersection_update(self.shuntAllSet)
        print('run1Config\n\tbrOff:',list(brOff),'\n\tshuntOff:',list(shuntOff))
        #
        if self.algoPF=='PSM':
            return self.__run1ConfigPSM__(brOff,shuntOff,fo)
        #
        if self.algoPF=='GS':
            return self.__run1ConfigGS__(brOff,shuntOff,fo)
        #
        if self.algoPF=='NR':
            return self.__run1ConfigNR__(brOff,shuntOff,fo)
        raise Exception('Error Algo Powerflow PSM,GS,NR')
    #
    def __run1ConfigPSM__(self,brOff,shuntOff,fo=''):
        """
        - result (dict): {'FLAG':,'RateMax%', 'Umax[pu]','Umin[pu]','DeltaA','RateMax%'}
        - FLAG (str): 'CONVERGENCE' or 'DIVERGENCE' or 'LOOP' or 'ISLAND'
        - DeltaA: MWH
        """
        #
        t0 = time.time()
        c1 = self.checkLoopIsland(brOff)
        if c1:
            return {'FLAG':c1}
        #
        iterMax = self.setting['PF_option_PSM'][0]
        epsilon = self.setting['PF_option_PSM'][1]

        #
        self.braC2 = self.braC1.copy()
        for br1 in brOff:
            self.braC2.pop(br1)
        #
        self.busC2 = self.busC1.copy()
        for br1 in brOff:
            for b1 in self.braC1[br1]:
                self.busC2[b1].remove(br1)
        #
        self.brA2 = self.brAllSet-brOff
        self.__lineDirection__()
        ordc,ordv,groupA = self.__ordCompute__()
        #B of Line
        BUSb = {} # for b, p of line,shunt
        for k,v in self.braB.items():
            if k not in brOff:
                for i in range(len(v)):
                    bi = self.braC1[k][i]
                    vi = self.braB[k][i]
                    if vi!=0.0:
                        try:
                            BUSb[bi]-=complex(0,vi)
                        except:
                            BUSb[bi]=complex(0,-vi)
        #
        #print('BUSb',BUSb)
        for k,v in self.shuntB.items():
            if k not in shuntOff:
                bi = v[0]
                if v[1]!=0.0:
                    try:
                        BUSb[bi]+= v[1]
                    except:
                        BUSb[bi]= complex(v[1],0)
                if v[2]!=0.0:
                    try:
                        BUSb[bi]+=complex(0,-v[2])
                    except:
                        BUSb[bi]=complex(0,-v[2])
        print('BUSb',BUSb)
        #
        for ii1 in range(len(self.IDProfile)):
            load1p = self.load[ii1]
            vgen1p = self.vgen[ii1]
            print('vgen1:',vgen1p)
            for ii2 in range(len(self.busSlack)):
                ordc1 = ordc[ii2]
                ordv1 = ordv[ii2]
                busGrp1 = groupA[ii2]
                sbus1 = {k:complex() for k in busGrp1}
                for k,v in load1p.items():
                    if k in busGrp1:
                        sbus1[k] = v
                BUSb1 = {k:v for k,v in BUSb.items() if k in busGrp1}
                #
                print('ordc1',ordc1)
                print('ordv1',ordv1)
                print('busGrp1',busGrp1)
                print('sbus1:',sbus1)
                #
                bs1 = self.busSlack[ii2] # bus Slack
                vg1 = vgen1p[bs1]
                du,di,s0 = dict(),dict(),0
                vbus = {h1:complex(vg1,0) for h1 in busGrp1}
                #
                for ii in range(iterMax+1):
                    sbus = sbus1.copy()
                    for k,v in BUSb1.items():
                        vm = abs(vbus[k])
                        sbus[k] += vm*vm*v
                    # cal cong suat nguoc
                    for bri in ordc1:
                        b1 = self.braC2[bri][0]
                        b2 = self.braC2[bri][1]
                        rx = self.braRX[bri]
                        ib = sbus[b2]/vbus[b2]
                        iba = abs(ib)
                        du[bri] = ib.conjugate()*rx
                        di[bri] = iba
                        ds1 = iba*iba*rx
                        #
                        sbus[b1]+=ds1+sbus[b2]
                    #
                    # cal dien ap xuoi
                    for bri in ordv1:
                        b1 = self.braC2[bri][0]
                        b2 = self.braC2[bri][1]
                        vbus[b2]=vbus[b1]-du[bri]
                    #
                    ep1 = abs(s0-sbus[bs1])
                    print(sbus[bs1],ep1,ii)
                    if ep1<epsilon:
                        break
                    else:
                        s0 = sbus[bs1]
                    #
                    if ii==iterMax:
                        return {'FLAG':'DIVERGENCE'}

        print('run1ConfigPSM: %.6f[s]'%(time.time()-t0))
        return
    #
    def __ordCompute__(self):
        busC = dict() # connect [LineUp,[LineDown]]
        for h1 in self.brAllSet:
            busC[h1] = [0,set()]
        #
        for h1,l1 in self.braC2.items():
            busC[l1[1]][0]= h1     # frombus
            busC[l1[0]][1].add(h1) # tobus
        #
        ordc,ordv,busGroup = [],[],[]
        for b1 in self.busSlack:
            bs1 = findBusConnected(b1,self.busC2,self.braC2)
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
                            if k in self.busSlack:
                                break
                            #
                            if v[0]!=0:
                                ordc1.append(v[0])
                            sord.add(v[0])
                            balr[k]=False
            ordv1 = [ordc1[-i-1]  for i in range(len(ordc1))]
            ordc.append(ordc1)
            ordv.append(ordv1)
            busGroup.append(bs1)
        return ordc,ordv,busGroup
    #
    def __lineDirection__(self):
        ba = list(self.busSlack)
        lset = set()
        for ii in range(20000):
            ba2 = []
            for b1 in ba:
                for l1 in self.brA2:
                    if l1 not in lset:
                        if b1==self.braC2[l1][1]:
                            d = self.braC2[l1][0]
                            self.braC2[l1][0] = self.braC2[l1][1]
                            self.braC2[l1][1] = d
                            lset.add(l1)
                            ba2.append(d)
                        elif b1==self.braC2[l1][0]:
                            lset.add(l1)
                            ba2.append(self.braC2[l1][1])
            if len(ba2)==0:
                break
            ba=ba2.copy()

# data for Recloser Optim
class DATAP_REOP(DATAP):
    def __init__(self,fi):
        super().__init__(fi,0)
    #
    def getData(self):
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
    brOff = [9,12]
    shuntOff = [1]
    #

    p1 = DATAP_PF(ARGVS.fi)
    v1 = p1.run1Config(brOff,shuntOff,fo=ARGVS.fo)
##    print(v1)
##    v1 = p1.run1Config_WithObjective(lineOff=lineOff,shuntOff=shuntOff,fo=ARGVS.fo)
##    print('time %.5f'%(time.time()-t01))
##    print(v1)
#
if __name__ == '__main__':
    ARGVS.fo = PATH_FILE+'\\res\\res1Config.csv'
    test_psm()
##    test_ReOp()
