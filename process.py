class Gitt():
    def __init__(self, path):
        self.path = path

    def read_data(self):
        data = pd.read_excel(self.path, sheet_name = '记录表')
        data['测试时间/Sec'] = data['测试时间/Sec'] - data['测试时间/Sec'].iloc[0]
        self.pristine_data = data
        self.pristine_data.index = np.arange(len(data))

    def cd_divide(self, pristine_data):
        ll = len(pristine_data) - 1
        for x in range(ll):
            if (pristine_data['比容量/mAh/g'].iloc[x] > 0)and(pristine_data['比容量/mAh/g'].iloc[x+1] == 0)and(pristine_data['电流/mA'].iloc[x+1] > 0):
                cd = 0
                dc = x
                #  将充电放电数据分别存在两个DataFrame中，并统一index和测试时间
                self.discharge_data = pristine_data.iloc[:dc]
                self.discharge_data.index = np.arange(len(self.discharge_data))
                self.discharge_data['测试时间/Sec'] = self.discharge_data['测试时间/Sec'] - self.discharge_data['测试时间/Sec'].iloc[0]
                self.charge_data = pristine_data.iloc[dc:]
                self.charge_data.index = np.arange(len(self.charge_data))
                self.charge_data['测试时间/Sec'] = self.charge_data['测试时间/Sec'] - self.charge_data['测试时间/Sec'].iloc[0]
                break
            elif (pristine_data['比容量/mAh/g'].iloc[x] > 0)and(pristine_data['比容量/mAh/g'].iloc[x+1] == 0)and(pristine_data['电流/mA'].iloc[x+1] < 0):
                cd = x
                dc = 0
                self.discharge_data = pristine_data.iloc[cd:]
                self.discharge_data.index = np.arange(len(self.discharge_data))
                self.discharge_data['测试时间/Sec'] = self.discharge_data['测试时间/Sec'] - self.discharge_data['测试时间/Sec'].iloc[0]
                self.charge_data = pristine_data.iloc[:cd]
                self.charge_data.index = np.arange(len(self.charge_data))
                self.charge_data['测试时间/Sec'] = self.charge_data['测试时间/Sec'] - self.charge_data['测试时间/Sec'].iloc[0]
                break

    def diffus_fit(self, data, DROP):
        length0 = len(data)-1
        E_tao = data['电压/V'].loc[[x for x in range(length0) if (data['电流/mA'].loc[x] != 0)and(data['电流/mA'].loc[x+1] == 0)]]
        E_s = data['电压/V'].loc[[x for x in range(length0) if (data['电流/mA'].loc[x] == 0)and(data['电流/mA'].loc[x+1] != 0)]]
        E_R = data['电压/V'].loc[[x+customize_Constant for x in range(length0) if (data['电流/mA'].loc[x] == 0)and(data['电流/mA'].loc[x+1] != 0)]]
        deta_Es = E_s.diff().dropna()   # 前向差分并丢弃第一个空值
        deta_Etao = E_tao.iloc[DROP:DROP-1].values - E_R.iloc[DROP:DROP-1].values
        index = np.arange(len(deta_Es))
        U = E_tao.iloc[DROP:DROP-1]
        U.index = index
        Q = data['比容量/mAh/g'].loc[[x for x in range(length0) if (data['电流/mA'].loc[x] != 0)and(data['电流/mA'].loc[x+1] == 0)]]
        Q.index = index
        k = (deta_Es/deta_Etao)**2/tao
        k.index = index
        return pd.concat([U, k, Q], axis=1)
