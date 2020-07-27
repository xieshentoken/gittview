import numpy as np
import pandas as pd

class Gitt():
    def __init__(self, path):
        self.path = path

    def read_data(self):
        data = pd.read_excel(self.path, sheet_name = '记录表')
        t = data['测试时间/Sec']
        data['测试时间/Sec'] = data['测试时间/Sec'] - t.iloc[0]
        self.pristine_data = data
        self.pristine_data.index = np.arange(len(data))

    def cd_divide(self, pristine_data):
        ll = len(pristine_data) - 1
        priQ = pristine_data['比容量/mAh/g']
        priI = pristine_data['电流/mA']
        for x in range(ll):
            if (priQ.iloc[x] > 0)and(priQ.iloc[x+1] == 0)and(priI.iloc[x+1] > 0):
                cd = 0
                dc = x + 1
                #  将充电放电数据分别存在两个DataFrame中，并统一index，重新计算测试时间
                self.discharge_data = pristine_data.iloc[:dc]
                self.discharge_data.index = np.arange(len(self.discharge_data))
                dischargeT = self.discharge_data['测试时间/Sec']
                self.discharge_data['测试时间/Sec'] = self.discharge_data['测试时间/Sec'] - dischargeT.iloc[0]
                self.charge_data = pristine_data.iloc[dc:]
                self.charge_data.index = np.arange(len(self.charge_data))
                chargeT = self.charge_data['测试时间/Sec']
                self.charge_data['测试时间/Sec'] = self.charge_data['测试时间/Sec'] - chargeT.iloc[0]
                break
            elif (priQ.iloc[x] > 0)and(priQ.iloc[x+1] == 0)and(priI.iloc[x+1] < 0):
                cd = x + 1
                dc = 0
                self.discharge_data = pristine_data.iloc[cd:]
                self.discharge_data.index = np.arange(len(self.discharge_data))
                dischargeT = self.discharge_data['测试时间/Sec']
                self.discharge_data['测试时间/Sec'] = self.discharge_data['测试时间/Sec'] - dischargeT.iloc[0]
                self.charge_data = pristine_data.iloc[:cd]
                self.charge_data.index = np.arange(len(self.charge_data))
                chargeT = self.charge_data['测试时间/Sec']
                self.charge_data['测试时间/Sec'] = self.charge_data['测试时间/Sec'] - chargeT.iloc[0]
                break

    def diffus_fit(self, data, tao, DROP, customize_Constant):
        length0 = len(data)-1
        priU = data['电压/V']
        priI = data['电流/mA']
        capacity = data['比容量/mAh/g']
        E_tao = priU.loc[[x for x in range(length0) if (priI.loc[x] != 0)and(priI.loc[x+1] == 0)]]
        E_s = priU.loc[[x for x in range(length0) if (priI.loc[x] == 0)and(priI.loc[x+1] != 0)]]
        E_R = priU.loc[[x+customize_Constant for x in range(length0) if (priI.loc[x] == 0)and(priI.loc[x+1] != 0)]]
        deta_Es = E_s.diff().dropna()   # 前向差分并丢弃第一个空值
        deta_Etao = E_tao.iloc[DROP:DROP-1].values - E_R.iloc[DROP:DROP-1].values
        index = np.arange(len(deta_Es))
        U = E_tao.iloc[DROP:DROP-1]
        U.index = index
        Q = capacity.loc[[x for x in range(length0) if (priI.loc[x] != 0)and(priI.loc[x+1] == 0)]].iloc[DROP:DROP-1]
        Q.index = index
        k = (deta_Es/deta_Etao)**2/tao
        k.index = index
        result = pd.concat([U, k, Q], axis=1)
        result.columns = ['电压/V', 'tEsEt', '比容量/mAh/g']
        return result
