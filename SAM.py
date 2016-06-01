## this script using the following encoding:utf-8
## using: create SAM table according to Io table
## input: the path of IO table(xls), SAM parameters file(txt)
## output: create and write SAM table
## depending on: xlrd,xlwt,os
## defining:
######account: class type to build account
######SAM: class type to create and wtite SAM table

import xlrd
import xlwt
import os


class account:
    "name:text;ac_in:dict;ac_out:dict"
    def __init__(self,name = 'default'):
        self.name = name
        self.ac_in = {}
##        self.ac_in.setdefault(self.name,0)
        self.ac_out = {}
##        self.ac_out.setdefault(self.name,0)
        self.equal = None
        self.balance = 0
        self.test()
    def __linkin(self,account1,numin):
        if account1.name in self.ac_in:
            self.ac_in[account1.name] += numin
        else:
            self.ac_in[account1.name] = numin
        self.test()
    def __linkout(self,account2,numout):
        if account2.name in self.ac_out:
            self.ac_out[account2.name] += numout
        else:
            self.ac_out[account2.name] = numout
        self.test()
    def test(self):
        ac_in,ac_out = 0,0
        for i in self.ac_in:
            ac_in += self.ac_in[i]
        for j in self.ac_out:
            ac_out += self.ac_out[j]
        self.balance = ac_in - ac_out
        self.equal = (ac_in==ac_out)
        if self.equal:
            return self.equal
        else:
##            print "balance of account({0}) is {1}".format(self.name,self.balance)
            return False
        return self.equal,ac_in-ac_out
    def pay(self,account2,num):
        account2.__linkin(self,num)
        self.__linkout(account2,num)
        
    def get(self,account2,num):
        account2.__linkout(self,num)
        self.__linkin(account2,num)

class SAM:
    
    def __init__(self,ac = []):
        self.accounts = ac
        
    def addAccount(self,ac):
        name_list = [x.name for x in self.accounts]
        if ac.name in name_list:
            print "error: name exists!!!"
        else:
            self.accounts.append(ac)
    def delAccount(self,ac):
        name_list = [x.name for x in self.accounts]
        if ac.name not in name_list:
            print "error: no account!!!"
        else:
            self.accounts.remove(ac)
    def isAll(self,error = 0.01):
        sumlist = [x.balance for x in self.accounts]
        return abs(sum(sumlist)) <= error
    def isEqual(self,Ierror=0.1):
        control = True
        for ac in self.accounts:
            if abs(ac.balance) > Ierror:
                control = False
                break
        return control

    def setEqual(self,Ierror=0.1):
### waiting for coding
        equalList = []
        for i in range(len(self.accounts)):
            error = self.accounts[i].balance
            if error > Ierror:
                maxout = 0
                for k in range(i+1,len(self.accounts)):
                    if self.accounts[k].name in self.accounts[i].ac_out:
                        if self.accounts[i].ac_out[self.accounts[k].name] > maxout:
                            maxout = self.accounts[i].ac_out[self.accounts[k].name]
                            ename = self.accounts[k].name
                for j in range(i+1,len(self.accounts)):
                    try:
                        if self.accounts[j].name == ename:
                            self.accounts[i].pay(self.accounts[j],error)
                            del ename
                            break
                    except:
                        print "add pay_link of {0}({1}) to accounts after this list_code".format(self.accounts[i].name,i)
            elif error < -Ierror:
                maxin = 0
                for k in range(i+1,len(self.accounts)):
                    if self.accounts[k].name in self.accounts[i].ac_in:
                        if self.accounts[i].ac_in[self.accounts[k].name] > maxin:
                            maxin = self.accounts[i].ac_in[self.accounts[k].name]
                            ename = self.accounts[k].name
                for j in range(i+1,len(self.accounts)):
                    try:
                        if self.accounts[j].name == ename:
                            self.accounts[i].get(self.accounts[j],-error)
                            del ename
                            break
                    except:
                        print "add get_link of {0}({1}) to accounts after this list_code".format(self.accounts[i].name,i)                
            equalList.append(self.accounts[i].name)
                        
                    
                
            
    def toXls(self,out_path = 'SAM.xls',Ierror = 0.1):
        'writing to excel files, out_path = SAM.xls, Terror = 0.1'
        w = xlwt.Workbook()
        ws = w.add_sheet('SAM')
### writing(waiting for coding)
#### writing title of row and column
        for i in range(len(self.accounts)):
            ws.write(0,i+1,self.accounts[i].name)
            ws.write(i+1,0,self.accounts[i].name)
        ws.write(0,i+2,'total_in')
        ws.write(i+2,0,'total_out')
#### writing text
        for j in range(len(self.accounts)):
            for k in range(len(self.accounts)):
                if self.accounts[k].name in self.accounts[j].ac_in:
                    if self.accounts[j].ac_in[self.accounts[k].name] != 0:
                        ws.write(j+1,k+1,self.accounts[j].ac_in[self.accounts[k].name])
            ws.write(j+1,k+2,sum([self.accounts[j].ac_in[x] for x in self.accounts[j].ac_in])) 
            ws.write(k+2,j+1,sum([self.accounts[j].ac_out[x] for x in self.accounts[j].ac_out]))
        ws.write(j+2,k+2,self.isEqual(Ierror))            
### save
        add_code = 1
        while True:
            if os.path.exists(out_path) == False:
                w.save(out_path)
                break
            elif add_code == 1:
                out_path = out_path.replace('.',str(add_code)+'.')
                add_code += 1
            else:
                out_path = out_path.replace(str(add_code-1),str(add_code))
                add_code += 1                
        



        
if __name__ == '__main__':
### building accounts
    sam = SAM()
    Account = ['producer_goods','producer_activity','factor_labor','factor_others','households','goverment',
               'investment_fixCapital','investment_stocks','ROW_otherProvince','ROW_otherCountry']
    for acc in Account:
        sam.addAccount(account(acc))
    good_activity = 315080871
    sam.accounts[0].get(sam.accounts[1],good_activity)
    good_households = 3471061+43018067
    sam.accounts[0].get(sam.accounts[4],good_households)
    good_gover = 32588415
    sam.accounts[0].get(sam.accounts[5],good_gover)
    good_fixedCapital = 53423981.88
    sam.accounts[0].get(sam.accounts[6],good_fixedCapital)
    good_stocks = 7546872.87
    sam.accounts[0].get(sam.accounts[7],good_stocks)
    good_otherProovince = 151281525.7
    sam.accounts[0].get(sam.accounts[8],good_otherProovince)
    good_otherCountry = 48120929.48
    sam.accounts[0].get(sam.accounts[9],good_otherCountry)
    activity_good = 456216579.1
    sam.accounts[1].get(sam.accounts[0],activity_good)
    labor_activity = 69199664.1
    sam.accounts[2].get(sam.accounts[1],labor_activity)
    othersFactors_activity = 49963635.69
    sam.accounts[3].get(sam.accounts[1],othersFactors_activity)
    household_labor = 69199664.1
    sam.accounts[4].get(sam.accounts[2],household_labor)
    household_otherFactors = 7883763.9
    sam.accounts[4].get(sam.accounts[3],household_otherFactors)
    gover_activity = 21972408.31
    sam.accounts[5].get(sam.accounts[1],gover_activity)
    gover_otherFactor = 42079871.79
    sam.accounts[5].get(sam.accounts[3],gover_otherFactor)
    gover_household = 7284200
    sam.accounts[5].get(sam.accounts[4],gover_household)
    investmentC_household = 23310100
    sam.accounts[6].get(sam.accounts[4],investmentC_household)
    investmentC_goverment = 387480651
    sam.accounts[6].get(sam.accounts[5],investmentC_goverment)
    investmentS_investmentC = 7546872.87
    sam.accounts[7].get(sam.accounts[6],investmentS_investmentC)
    otherP_goods = 124379648.9
    sam.accounts[8].get(sam.accounts[0],otherP_goods)
    otherP_household = 0
    sam.accounts[8].get(sam.accounts[4],otherP_household)
    otherP_investmentC = 108731035
    sam.accounts[8].get(sam.accounts[6],otherP_investmentC)
    otherC_goods = 73935495.96
    sam.accounts[9].get(sam.accounts[0],otherC_goods)

       

### paying and getting relationship build between accounts

##### add to a SAM
##    sam = SAM([a,b,c])
##    sam.accounts[0].pay(sam.accounts[1],25)
##    sam.accounts[1].pay(sam.accounts[2],25)
##    sam.accounts[2].pay(sam.accounts[0],25)
##    sam.accounts[1].pay(sam.accounts[0],30)
##    sam.accounts[2].pay(sam.accounts[1],10)
    print sam.isAll()
    print sam.isEqual()
    sam.setEqual()
    sam.accounts[5].get(sam.accounts[7],1)
    sam.accounts[7].get(sam.accounts[8],1)
    sam.accounts[8].get(sam.accounts[9],1)
    sam.setEqual()
##    sam.toXls('original.xls')
##    os.system('original.xls')
##    sam.setEqual()
##    sam.toXls('equal.xls')
##    os.system('equal.xls')
##    
    
