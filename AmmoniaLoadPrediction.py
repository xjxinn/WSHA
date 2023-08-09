def AmmoniaLoadPrediction(FlagNH3Dispatch):
    import numpy as np
    import math
    import scipy.interpolate as si
    import pandas as pd
    # FlagNH3Dispatch：合成氨交割时间尺度指示，1代表按月交割，0代表按周交割

    AmmoniaLoad = 1e5 / 8760 * np.ones([1, 8760])
    AmmoniaLoad_15min = 1e5 / 8760 * np.ones([1, 8760 * 4])
    ActualNH3Demand = np.zeros([1, 8760])
    ActualNH3Demand_15min = np.zeros([1, 8760 * 4])
    VirtualNH3Demand = np.zeros([1, 8760])
    VirtualNH3Demand_15min = np.zeros([1, 8760 * 4])

    if FlagNH3Dispatch == 1:
        timeseries = np.array([0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334, 365]) * 24
        for i in range(len(timeseries) - 1):
            ActualNH3Demand[0, timeseries[i + 1] - 1] = np.sum(AmmoniaLoad[0, timeseries[i]:timeseries[i + 1] + 1])
        timeseries = np.array([0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334, 365]) * 24 * 4
        for i in range(len(timeseries) - 1):
            ActualNH3Demand_15min[0, timeseries[i + 1] - 4:timeseries[i + 1]] = np.sum(
                AmmoniaLoad_15min[0, timeseries[i]:timeseries[i + 1] + 1]) / 4
    else:
        DeliveryPeriod = 7 * 24
        tmp = np.linspace(1, round(365 * 24 / DeliveryPeriod), round(365 * 24 / DeliveryPeriod)) * 7 * 24
        tmp = np.concatenate([[0], tmp, [8760]], axis=0)
        for i in range(len(tmp) - 1):
            ActualNH3Demand[0, int(tmp[i + 1]) - 1] = sum(AmmoniaLoad[0, int(tmp[i]):int(tmp[i + 1])])

        timeseries = np.array([31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334]) * 24
        for i in range(len(timeseries)):
            xq = timeseries[i] - math.floor(timeseries[i] / DeliveryPeriod) * DeliveryPeriod
            x = np.array([0, DeliveryPeriod])
            v = np.array([0, ActualNH3Demand[0, math.ceil(timeseries[i] / DeliveryPeriod) * DeliveryPeriod - 1]])
            f = si.interp1d(x, v, kind='slinear')
            VirtualNH3Demand[0, timeseries[i] - 1] = f(xq)

        for i in range(len(ActualNH3Demand_15min)):
            ActualNH3Demand_15min[i] = ActualNH3Demand[math.floor(i / 4)]
            VirtualNH3Demand_15min[i] = VirtualNH3Demand[math.floor(i / 4)]

    pd.DataFrame(ActualNH3Demand.T).to_excel("ActualNH3Demand.xlsx")
    pd.DataFrame(VirtualNH3Demand.T).to_excel("VirtualNH3Demand.xlsx")
    pd.DataFrame(ActualNH3Demand_15min.T).to_excel("ActualNH3Demand_15min.xlsx")
    pd.DataFrame(VirtualNH3Demand_15min.T).to_excel("VirtualNH3Demand_15min.xlsx")
