def year(DayBegin, DayEnd, HStoinitVal, AStoinitVal, delta_T,CapaWind, CapaSolar, CapaHE, CapaASR, CapaHB, CapaAB):
    import pandas as pd
    import math
    import cvxpy as cp
    import numpy as np

    #  现货交易决策
    pwindToNorm = pd.read_excel('RESData/ResPredpu.xlsx', header=None)
    SpotPrice = pd.read_excel('RESData/ElecPrice.xlsx', header=None)
    psolarToNorm = 0
    VirtualNH3Demand = pd.read_excel('VirtualNH3Demand.xlsx', index_col=0)
    ActualNH3Demand = pd.read_excel('ActualNH3Demand.xlsx', index_col=0)
    ##
    # 模型参数
    RES = np.array(pwindToNorm * CapaWind + psolarToNorm * CapaSolar)
    ConstG2L = 2 / 3 * 1000 / 22.4 * 17 * 1e-6  # 氢气体积转液氨质量的转换系数，吨/标方
    ConstG2LH2 = 1000 / 22.4 * 2 * 1e-6  # 氢气体积转氢气质量的转换系数，吨/标方
    # 规划参数
    TDPrice = 0.0735 * 1e3  # 输配电价
    Supplement = 0.022425 * 1e3  # 政府性基金及附加
    SpotPriceSell = SpotPrice - (TDPrice + Supplement)
    periodNH3Dispatch = delta_T  # 合成氨工况切换周期，典型值1天
    # 运行参数
    # 电解池
    yitaH2 = 5  # 典型电解池效率，典型值5kwh/Nm3
    ratepAEMin = 0.05  # 电解池最低负载率，典型值：单槽 - 20%，集群 - 5%
    RateOverShoot = 0  # 最大超调率，典型值20%
    # 储氢罐
    yitaMin = 0.1
    yitaMax = 1
    # 合成氨
    LoadNH3Min = 0.3  # 最低负载率，典型值30%
    # 定容优化模型
    horizonDispatch = (DayEnd - DayBegin) * 24  # 规划的时间区间
    nDispatch = horizonDispatch  # 规划问题的规模
    nNH3Dispatch = math.floor(horizonDispatch/periodNH3Dispatch)  # 合成氨稳态工况数量
    ## 变量
    varsCtrlDisp = cp.Variable(shape=(nDispatch, 6))  # 上网功率\下网功率\制氢功率\制氨功率\储罐入口流速\储罐出口流速
    varsSpotMarket = cp.Variable(shape=(nDispatch, 1))  # 下网功率来自现货市场，买电为正
    varsNH3Disp = cp.Variable(shape=(nNH3Dispatch, 1))  # 合成氨准稳态工况数量
    varsSto = cp.Variable(shape=(nDispatch + 1, 2))  # 储氢罐状态：氢气含量，Nm3；储氨罐状态：氨含量，吨
    varsAmmoniaMarket = cp.Variable(shape=(nDispatch, 1))  # 氨市场售出速率
    ## 目标函数
    # 年运行收益：上网收益
    fFuncProfit = np.array(SpotPriceSell.iloc[DayBegin*24:DayBegin*24+nDispatch]).T @ varsCtrlDisp[:, 0]
    # 年运行成本：购买网电成本
    fFuncEleCost = np.array(SpotPrice.iloc[DayBegin*24:DayBegin*24+nDispatch]).T @ varsSpotMarket[:, 0]
    fFuncCost = fFuncEleCost

    # 年净运行成本
    fFuncDisp = fFuncCost - fFuncProfit

    # 约束条件
    constraint = []
    # 运行约束
    for itr in range(nDispatch):
    # 功率平衡约束
        constraint += [(RES[DayBegin * 24 + itr] + varsCtrlDisp[itr, 1]) - (varsCtrlDisp[itr, 0] + varsCtrlDisp[itr, 2] + varsCtrlDisp[itr, 3]) >= 0]
    # 储氢罐约束
        constraint += [varsSto[itr + 1, 0] - varsSto[itr, 0] - (varsCtrlDisp[itr, 4] - varsCtrlDisp[itr, 5]) == 0]  # 状态空间方程
        constraint += [yitaMin * CapaHB <= varsSto[itr, 0],
                       varsSto[itr, 0] <= yitaMax * CapaHB]
    # 制氢约束
        constraint += [varsCtrlDisp[itr, 2] / (5e-3) - varsCtrlDisp[itr, 4] == 0]  # 制氢功率 / (5e-3) = 氢气入口流速
        constraint += [ratepAEMin * CapaHE <= varsCtrlDisp[itr, 2],
                       varsCtrlDisp[itr, 2] <= (1 + RateOverShoot) * CapaHE]
    # 储氨罐约束
        constraint += [varsSto[itr + 1, 1] - varsSto[itr, 1] + varsAmmoniaMarket[itr, 0] - varsCtrlDisp[itr, 3] / (3.230e-4) * ConstG2L == 0]  # 状态空间方程
        constraint += [0 <= varsSto[itr, 1],
                       varsSto[itr, 1] <= CapaAB]
    # 制氨约束
        constraint += [varsCtrlDisp[itr, 3] / (3.230e-4) - varsCtrlDisp[itr, 5] == 0]  # 制氨功率 = 氢气出口流速 * (3.230e-4)
        constraint += [LoadNH3Min * CapaASR <= varsCtrlDisp[itr, 3],
                       varsCtrlDisp[itr, 3] <= CapaASR]
    # 上下网电力约束：每一时刻仅允许上网或下网
        M1 = 1000  # M1取值应该大于最大发电功率和最大负荷
        M2 = 1000
        constraint += [0 <= varsCtrlDisp[itr, 0],
                       varsCtrlDisp[itr, 0] <= M1]
        constraint += [0 <= varsCtrlDisp[itr, 1],
                       varsCtrlDisp[itr, 1] <= M2]
    # 下网电力来源约束：下网电力来源于现货市场
        constraint += [varsSpotMarket[:, 0] - varsCtrlDisp[:, 1] == 0]
        constraint += [varsSpotMarket[:, 0] >= 0]
    # 不允许上网
        constraint += [varsCtrlDisp[:, 0] == 0]

    # 合成氨工况切换约束
    ntmp1 = periodNH3Dispatch
    nNH3Dispatch = math.floor(horizonDispatch / periodNH3Dispatch)  # 合成氨稳态工况数量，典型值365
    for itr in range(nNH3Dispatch):
        constraint += [varsCtrlDisp[(itr * ntmp1):(ntmp1 * (itr + 1)), 3] - varsNH3Disp[itr, 0] == 0]

    # 储氢罐初值和终值约束
    constraint += [varsSto[0, 0] - HStoinitVal == 0,
                   varsSto[nDispatch, 0] - yitaMin * CapaHB >= 0]
    # 储氨罐初值和终值约束
    constraint += [varsSto[0, 1] - AStoinitVal == 0,
                   varsSto[nDispatch, 1] - np.array(VirtualNH3Demand)[DayEnd * 24-1, 0] >= 0]
    # 氨市场约束
    for itr in range(nDispatch):
        constraint += [np.array(ActualNH3Demand)[DayBegin*24+itr, 0] - varsAmmoniaMarket[itr, 0] == 0]  # 合成氨需求量 = 合成氨售出量

    # 模型求解
    print('年度优化开始求解')
    prob = cp.Problem(cp.Minimize(fFuncDisp), constraint)
    prob.solve(solver=cp.GUROBI, MIPGap=1e-2, timelimit=1e4)
    if prob.status == 'optimal':
        print('求解结束，获得最优解！')
    elif prob.status == 'infeasible':
        print('优化模型不可行！')
    else:
        print('Timeout, Display the current optimal solution')
    # 后处理
    # 记录原始变量最优解
    SpotTran = varsSpotMarket.value
    OnGrid = varsCtrlDisp[:, 0].value.reshape(-1, 1)
    HEPowerSchedule = varsCtrlDisp[:, 2].value.reshape(-1, 1)
    ASRPowerSchedule = varsCtrlDisp[:, 3].value.reshape(-1, 1)
    HSto = varsSto[1:, 0].value.reshape(-1, 1)
    ASto = varsSto[1:, 1].value.reshape(-1, 1)
    ElecTran = SpotTran - OnGrid

    return SpotTran, OnGrid, HEPowerSchedule, ASRPowerSchedule, HSto, ASto, ElecTran
