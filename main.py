from AmmoniaLoadPrediction import AmmoniaLoadPrediction
from year import year
from month import month
from daily import daily
import numpy as np
import pandas as pd

AmmoniaLoadPrediction(1)

# timeseries = np.array([0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334, 365])
timeseries = np.array([0, 31])
#  参数准备
CapaWind = 300  # 风机容量，MW
CapaSolar = 0  # 光伏容量，MW
CapaHE = 145  # 电解槽容量，MW
CapaASR = 12  # 合成氨反应器容量，MW
CapaHB = 2 * 1e5  # 氢缓冲罐容量，Nm3
CapaAB = 1e4  # 氨缓冲罐容量，吨
delta_T = 24  # 合成氨的调节周期，小时
VirtualNH3Demand = pd.read_excel('VirtualNH3Demand.xlsx', index_col=0)

SpotTran_hour = np.zeros([8760, 1])
OnGrid_hour = np.zeros([8760, 1])
HEPowerSchedule_hour = np.zeros([8760, 1])
ASRPowerSchedule_hour = np.zeros([8760, 1])
HSto_hour = np.zeros([8760, 1])
ASto_hour = np.zeros([8760, 1])
ElecTran_hour = np.zeros([8760, 1])

SpotTran_15min = np.zeros([8760*4, 1])
OnGrid_15min = np.zeros([8760*4, 1])
HEPowerSchedule_15min = np.zeros([8760*4, 1])
ASRPowerSchedule_15min = np.zeros([8760*4, 1])
HSto_15min = np.zeros([8760*4, 1])
ASto_15min = np.zeros([8760*4, 1])
ElecTran_15min = np.zeros([8760*4, 1])

for itr in range(len(timeseries)-1):
    DayBegin = timeseries[itr]
    DayEnd = 31
    if DayBegin == 0:
        AStoinitVal = 0  # 氨存量的初始值为0
        HStoinitVal = 0.1 * CapaHB  # 氢储罐初始值
    else:
        AStoinitVal = ASto_15min[DayBegin * 24 * 4 - 1]
        HStoinitVal = HSto_15min[DayBegin * 24 * 4 - 1]
    # 年度+多月滚动优化
    result_year = year(DayBegin, DayEnd, HStoinitVal, AStoinitVal, delta_T, CapaWind, CapaSolar, CapaHE, CapaASR, CapaHB, CapaAB)

    SpotTran_hour[DayBegin*24:DayEnd*24] = result_year[0]
    OnGrid_hour[DayBegin*24:DayEnd*24] = result_year[1]
    HEPowerSchedule_hour[DayBegin*24:DayEnd*24] = result_year[2]
    ASRPowerSchedule_hour[DayBegin*24:DayEnd*24] = result_year[3]
    HSto_hour[DayBegin*24:DayEnd*24] = result_year[4]
    ASto_hour[DayBegin*24:DayEnd*24] = result_year[5]
    ElecTran_hour[DayBegin*24:DayEnd*24] = result_year[6]

    dayseries = np.arange(timeseries[itr], timeseries[itr+1])

    for Day in range(len(dayseries)):
        # 月度+多日滚动优化
        DayBegin = dayseries[Day]
        print(DayBegin)
        DayEnd = timeseries[itr+1]
        if DayBegin == 0:
            HStoinitVal = 0.1*CapaHB
            AStoinitVal = 0
        else:
            HStoinitVal = HSto_15min[DayBegin * 24 * 4 - 1]
            AStoinitVal = ASto_15min[DayBegin * 24 * 4 - 1]
        AStofinalVal = ASto_hour[DayEnd * 24 - 1]

        result_month = month(DayBegin, DayEnd, HStoinitVal, AStoinitVal, AStofinalVal, delta_T,CapaWind, CapaSolar, CapaHE, CapaASR, CapaHB, CapaAB)
        SpotTran_hour[DayBegin * 24:DayEnd * 24] = result_month[0]
        OnGrid_hour[DayBegin * 24:DayEnd * 24] = result_month[1]
        HEPowerSchedule_hour[DayBegin * 24:DayEnd * 24] = result_month[2]
        ASRPowerSchedule_hour[DayBegin * 24:DayEnd * 24] = result_month[3]
        HSto_hour[DayBegin * 24:DayEnd * 24] = result_month[4]
        ASto_hour[DayBegin * 24:DayEnd * 24] = result_month[5]
        ElecTran_hour[DayBegin * 24:DayEnd * 24] = result_month[6]

        HStofinalVal = HSto_hour[(DayBegin+1) * 24 - 1]
        AStofinalVal = ASto_hour[(DayBegin+1) * 24 - 1]
        NH3Disp = ASRPowerSchedule_hour[DayBegin * 24]
        for t in range(96):
            if (DayBegin == 0) & (t == 0):
                AStoinitVal = 0
                HStoinitVal = 0.1*CapaHB
            else:
                HStoinitVal = HSto_15min[DayBegin * 24 * 4 + t - 1]
                AStoinitVal = ASto_15min[DayBegin * 24 * 4 + t - 1]
            result_daily = daily(DayBegin, t, 96, HStoinitVal, AStoinitVal, HStofinalVal, AStofinalVal, NH3Disp, CapaWind, CapaSolar, CapaHE, CapaASR, CapaHB, CapaAB)
            SpotTran_15min[DayBegin * 24 * 4 + t:DayBegin * 24 * 4 + 96] = result_daily[0]
            OnGrid_15min[DayBegin * 24 * 4 + t:DayBegin * 24 * 4 + 96] = result_daily[1]
            HEPowerSchedule_15min[DayBegin * 24 * 4 + t:DayBegin * 24 * 4 + 96] = result_daily[2]
            ASRPowerSchedule_15min[DayBegin * 24 * 4 + t:DayBegin * 24 * 4 + 96] = result_daily[3]
            HSto_15min[DayBegin * 24 * 4 + t:DayBegin * 24 * 4 + 96] = result_daily[4]
            ASto_15min[DayBegin * 24 * 4 + t:DayBegin * 24 * 4 + 96] = result_daily[5]
            ElecTran_15min[DayBegin * 24 * 4 + t:DayBegin * 24 * 4 + 96] = result_daily[6]

pd.DataFrame(SpotTran_hour).to_excel('result/SpotTran_hour.xlsx')
pd.DataFrame(OnGrid_hour).to_excel('result/OnGrid_hour.xlsx')
pd.DataFrame(HEPowerSchedule_hour).to_excel('result/HEPowerSchedule_hour.xlsx')
pd.DataFrame(ASRPowerSchedule_hour).to_excel('result/ASRPowerSchedule_hour.xlsx')
pd.DataFrame(HSto_hour).to_excel('result/HSto_hour.xlsx')
pd.DataFrame(ASto_hour).to_excel('result/ASto_hour.xlsx')
pd.DataFrame(ElecTran_hour).to_excel('result/ElecTran_hour.xlsx')

pd.DataFrame(SpotTran_15min).to_excel('result/SpotTran_15min.xlsx')
pd.DataFrame(OnGrid_15min).to_excel('result/OnGrid_15min.xlsx')
pd.DataFrame(HEPowerSchedule_15min).to_excel('result/HEPowerSchedule_15min.xlsx')
pd.DataFrame(ASRPowerSchedule_15min).to_excel('result/ASRPowerSchedule_15min.xlsx')
pd.DataFrame(HSto_15min).to_excel('result/HSto_15min.xlsx')
pd.DataFrame(ASto_15min).to_excel('result/ASto_15min.xlsx')
pd.DataFrame(ElecTran_15min).to_excel('result/ElecTran_15min.xlsx')