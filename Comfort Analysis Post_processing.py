import os
import sys

import matplotlib
import numpy as np
import comtypes.client
import matplotlib.pyplot as plt


# create API helper object
helper = comtypes.client.CreateObject('SAP2000v1.Helper')
helper = helper.QueryInterface(comtypes.gen.SAP2000v1.cHelper)
# Get SapObject
mySapObject = helper.GetObject("CSI.SAP2000.API.SapObject")
SapModel = mySapObject.SapModel
ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()  # 不起作用
##########################
#  Steady-state analysis #
##########################
# ret = SapModel.Results.Setup.SetCaseSelectedForOutput("Steady_Sate")
# ret = SapModel.Results.Setup.SetOptionSteadyState(2, 3)
# [NumberResults, Obj, Elm, LoadCase,
# StepType, StepNum, U1, U2, U3, R1, R2, R3, Remark] = SapModel.Results.JointDispl("PointsGroup", 2)
#
# pos = [i for i in range(0, NumberResults) if StepType[i] == 'Mag at Freq']
# print(pos[0:1])
# plt.plot(StepNum[pos[0]:(pos[-1]+1)], U3[pos[0]:(pos[-1]+1)])
# plt.show()

##########################
#  Walking analysis     #
##########################
ret = SapModel.Results.Setup.SetOptionModalHist(1)
Walking_Frqs = [1.6 + 0.05 * x for x in range(0, 13, 1)]
Walking_Frqs = [round(x, 2) for x in Walking_Frqs]
for frq in Walking_Frqs:
    ret = SapModel.Results.Setup.SetCaseSelectedForOutput(f'Walking{frq}')
nn = len(Walking_Frqs)

[NumberResults, Obj, Elm, LoadCase,
 StepType, StepNum, U1, U2, U3, R1, R2, R3, Remark] = SapModel.Results.JointAccAbs("PointsGroup", 2)

Acc = []
for n in range(0, nn):
    frq = Walking_Frqs[n]
    LoadCase_i = f'Walking{frq}'
    index_i = LoadCase.index(LoadCase_i)
    U3_tem = [abs(j) for j in U3[index_i: index_i+2]]
    Acc.append(max(U3_tem))
plt.plot(Walking_Frqs, Acc)
plt.title('行人荷载扫频计算频率与最大加速度关系图', fontproperties='SimSun', fontsize=15)
plt.xlabel('计算频率(Hz)', fontproperties='SimSun', fontsize=15)
plt.ylabel('各工况最大加速度(m/s$^2$)', fontproperties='SimSun', fontsize=15)
plt.show()

# 找出最大竖向加速度和对应工况
Acc_max = max(Acc)
Frq_max = Walking_Frqs[Acc.index(Acc_max)]
print(Frq_max)
##########################
#  VerCrowd analysis     #
##########################



##########################
#  LatCrowd analysis     #
##########################