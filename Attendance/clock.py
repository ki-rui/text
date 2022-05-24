from logg import l
import sys
class clock():
    def __init__(self,str,isbegin=0,isSummer=False):
        self.h=int(str.split(':')[0])
        self.m = int(str.split(':')[1])
        if isbegin==0: return
        self.hs,self.ms=self.norm(isbegin,isSummer)
        if(self.hs==0):
            try:raise Exception("标准化时间为0，clock初始化出现了某些异常")
            except:
                l.error("标准化时间为0，clock初始化出现了某些异常")
                sys.exit(0)
        # if isbegin:

    def norm(self,isbegin,isSummer):
        h0=0
        m0=0
        cha8=self.cha(clock("8:00"))
        cha10 = self.cha(clock("10:00"))
        cha12 = self.cha(clock("12:00"))
        chax1=self.cha(clock("14:30")) if isSummer else self.cha(clock("14:00"))
        chax2=chax1-120
        chax3 = chax2 - 120

        # 12节---------------------------
        if self.h==7 or (cha8>=0 and cha8<=10):
            h0=8
            m0=0
        elif cha8>10 and cha8<=30:
            h0=8
            m0=30 if isbegin==1 else 0
        elif cha8>30 and cha8<=60:
            h0=9 if isbegin==1 else 8
            m0=0
        elif cha8>60 and cha8<=90:
            h0=10 if isbegin==1 else 9
            m0=0
        elif cha10>-30 and cha10<-10:
            h0 = 10 if isbegin==1 else 9
            m0 = 0 if isbegin==1 else 30

        # 34节------------------------
        elif cha10>=-10 and cha10<=10:
            h0 = 10
            m0 = 0
        elif cha10>10 and cha10<=30:
            h0 = 10 if isbegin==1 else 10
            m0 = 30 if isbegin==1 else 0
        elif cha10>30 and cha10<=60:
            h0=11 if isbegin==1 else 10
            m0=0
        elif cha10>60 and cha10<=90:
            h0=12 if isbegin==1 else 11
            m0=0
        elif cha12>-30 and cha12<-10:
            h0 = 12 if isbegin==1 else 11
            m0 = 0 if isbegin==1 else 30

        # 中午------------------------
        elif cha12 >= -10 and cha12 <= 10:
            h0 = 12
            m0 = 0
        elif cha12 > 10 and cha12 <= 20:
            h0 = 12
            m0 = 30 if isbegin == 1 else 0
        elif cha12 > 20 and cha12 <= 35:
            h0 = 12
            m0 = 30 if isbegin == 1 else 0
        elif cha12 > 35 and cha12 <= 60:
            h0 = 13
            m0 = 0
        elif cha12 > 60 and cha12 <= 90:
            h0 = 14 if isbegin == 1 else 13
            m0 = 0
        elif chax1 > -30 and chax1 < -10:
            h0 = 14 if isbegin == 1 else 13
            m0 = 0 if isbegin == 1 else 30

        # 56节
        elif chax1 >= -10 and chax1 <= 10:
            h0 = 14
            m0 = 0
        elif chax1 >10 and chax1 <= 30:
            h0 = 14 if isbegin == 1 else 14
            m0 = 30 if isbegin == 1 else 0
        elif chax1 > 30 and chax1 <= 60:
            h0 = 15 if isbegin == 1 else 14
            m0 = 0
        elif chax1 > 60 and chax1 <= 90:
            h0 = 16 if isbegin == 1 else 15
            m0 = 0
        elif chax2 > -30 and chax2 < -10:
            h0 = 16 if isbegin == 1 else 15
            m0 = 0 if isbegin == 1 else 30

        # 78节
        elif chax2 >= -10 and chax2 <= 10:
            h0 = 16
            m0 = 0
        elif chax2 >10 and chax2 <= 30:
            h0 = 16 if isbegin == 1 else 16
            m0 = 30 if isbegin == 1 else 0
        elif chax2 > 30 and chax2 <= 60:
            h0 = 17 if isbegin == 1 else 16
            m0 = 0
        elif chax2 > 60 and chax2 <= 90:
            h0 = 18 if isbegin == 1 else 17
            m0 = 0
        elif chax3 > -30 and chax3 < -10:
            h0 = 18 if isbegin == 1 else 17
            m0 = 0 if isbegin == 1 else 30

        # 晚上
        elif chax3 >= -10 and chax3<=10:
            h0=18
            m0=0
        elif chax3 >10 and chax3 <= 30:
            h0 = 18 if isbegin == 1 else 18
            m0 = 30 if isbegin == 1 else 0
        elif chax3 > 30 and chax3 <= 60:
            h0 = 19 if isbegin == 1 else 18
            m0 = 0
        elif chax3 > 60 and chax3 <= 90:
            h0 = 20 if isbegin == 1 else 19
            m0 = 0
        elif chax3 > 90 and chax3 < 110:
            h0 = 20 if isbegin == 1 else 19
            m0 = 0 if isbegin == 1 else 30
        elif chax3 >=110:
            h0=20
            m0=0
        return self.add30(h0,m0,isSummer)
    def add30(self,h,m,isSummer):
        if isSummer and h>=14:
            m0=(m+30)%60
            h0=h+(m+30)//60
            return h0,m0
        else: return h,m
    def cha(self,clocks):
        dh=self.h-clocks.h
        dm = self.m - clocks.m
        return dh*60+dm
    def pr(self):
        print(str(self.h)+":"+str(self.m)+",  normtime="+str(self.hs)+":"+str(self.ms))
class timeStage():
    def __init__(self, strs,isSummer=False):
        if strs=='':
            self.duration=0
            return
        spl=strs.split('\n')
        self.whole = False
        if(len(spl)==2): self.whole=True
        else:
            self.duration = 0
            return
        if not (len(spl[0].split(':'))==2 and len(spl[1].split(':'))==2):
            self.duration = 0
            return
        self.begin=clock(spl[0],isbegin=1,isSummer=isSummer)
        if self.whole:
            self.end=clock(spl[1],isbegin=2,isSummer=isSummer)
        else:
            self.end = clock(spl[0], isbegin=2, isSummer=isSummer)
        self.duration=self.chas()
    def chas(self):
        dh=self.end.hs-self.begin.hs
        dm = self.end.ms - self.begin.ms
        return dh + dm/60
    def pr(self):
        print(str(self.begin.h)+":"+str(self.begin.m)+"-",end='')
        if self.whole:
            print(str(self.end.h)+":"+str(self.end.m)+", ",end='')
            print("duration= {:.1f} h".format(self.duration))
if __name__=='__main__':
    dur="12:19\n19:59"
    ts=timeStage(dur)
    ts.pr()