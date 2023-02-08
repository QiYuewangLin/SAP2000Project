import os
import sys
import matplotlib
import numpy as np
import comtypes.client
import matplotlib.pyplot as plt
"""
本程序可用于舒适度分析的后处理，包括改模型单位、提取输出不利点稳态分析、行人荷载结果、竖向人群荷载结果和横向人群荷载结果、绘图并保存
本程序目前适用于单点的所有舒适度结果后处理，若有多点需求后续可补充
输入参数包括图片结果的保存路径、横向水平力作用方向。
编写： 林晨豪   修改时间：20230208
"""

# create API helper object
helper = comtypes.client.CreateObject('SAP2000v1.Helper')
helper = helper.QueryInterface(comtypes.gen.SAP2000v1.cHelper)
# Get SapObject
mySapObject = helper.GetObject("CSI.SAP2000.API.SapObject")
SapModel = mySapObject.SapModel
ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
# Change the units to kN m
ret = SapModel.SetPresentUnits(6)

# Save path of output
savepath = 'E:\\工程项目\\绵阳科技城新区科技大会堂\\SAP2000模型\\舒适度分析'
savepath += '\\稳态分析结果'
if not os.path.exists(savepath):
    os.mkdir(savepath)
# LatDirection
LatDirection = 'Y'       # 横向水平力作用方向
[NumberItems, ObjectType, ObjectName, Remark]= SapModel.GroupDef.GetAssignments("PointsGroup")

##########################
#  Steady-state analysis #
##########################
ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
ret = SapModel.Results.Setup.SetCaseSelectedForOutput("Steady_Sate")
ret = SapModel.Results.Setup.SetOptionSteadyState(2, 3)
[NumberResults, Obj, Elm, LoadCase,
StepType, StepNum, U1, U2, U3, R1, R2, R3, Remark] = SapModel.Results.JointDispl(ObjectName[0], 0)  # 仅输出一个点

pos = [i for i in range(0, NumberResults) if StepType[i] == 'Mag at Freq']
plt.figure(1)
plt.plot(StepNum[pos[0]:(pos[-1]+1)], U3[pos[0]:(pos[-1]+1)])
plt.title(f'稳态分析节点{Obj[0]}的频率-竖向位移曲线', fontproperties='SimSun', fontsize=15)
plt.xlabel('频率(Hz)', fontproperties='SimSun', fontsize=15)
plt.ylabel('位移(m)', fontproperties='SimSun', fontsize=15)
plt.tight_layout()
plt.savefig(f'{savepath}\\稳态分析节点{Obj[0]}的位移时程曲线.tiff')

##########################
#  Walking analysis     #
##########################
ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
ret = SapModel.Results.Setup.SetOptionModalHist(1)
Walking_Frqs = [1.6 + 0.05 * x for x in range(0, 13, 1)]
Walking_Frqs = [round(x, 2) for x in Walking_Frqs]
for frq in Walking_Frqs:
    ret = SapModel.Results.Setup.SetCaseSelectedForOutput(f'Walking{frq}')
nn = len(Walking_Frqs)

[NumberResults, Obj, Elm, LoadCase,
 StepType, StepNum, U1, U2, U3, R1, R2, R3, Remark] = SapModel.Results.JointAccAbs(ObjectName[0], 0)

Acc = []
for n in range(0, nn):
    frq = Walking_Frqs[n]
    LoadCase_i = f'Walking{frq}'
    index_i = LoadCase.index(LoadCase_i)
    U3_tem = [abs(j) for j in U3[index_i: index_i+2]]
    Acc.append(max(U3_tem))
plt.figure(2)
plt.plot(Walking_Frqs, Acc)
plt.title('行人荷载扫频计算频率与最大加速度关系图', fontproperties='SimSun', fontsize=15)
plt.xlabel('计算频率(Hz)', fontproperties='SimSun', fontsize=15)
plt.ylabel('各工况最大加速度(m/s$^2$)', fontproperties='SimSun', fontsize=15)
plt.tight_layout()
plt.savefig(f'{savepath}\\行人荷载扫频计算频率与最大加速度关系图.tiff')

f = open(f'{savepath}\\行人荷载扫频计算频率与最大加速度关系.txt', 'w')
for n in range(0, nn):
    f.write(f'{Walking_Frqs[n]} {Acc[n]}\n')
f.close()
# 找出最大竖向加速度和对应工况
Acc_max = max(Acc)
Frq_max = Walking_Frqs[Acc.index(Acc_max)]
ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
ret = SapModel.Results.Setup.SetOptionModalHist(2)
ret = SapModel.Results.Setup.SetCaseSelectedForOutput(f'Walking{Frq_max}')
[NumberResults, Obj, Elm, LoadCase,
 StepType, StepNum, U1, U2, U3, R1, R2, R3, Remark] = SapModel.Results.JointAccAbs(ObjectName[0], 0)
plt.figure(3)
plt.plot(StepNum, U3)
plt.title(f'行人荷载作用下节点{Obj[0]}竖向加速度时程图（{Frq_max}Hz）', fontproperties='SimSun', fontsize=15)
plt.xlabel('时间(s)', fontproperties='SimSun', fontsize=15)
plt.ylabel('最大工况的竖向加速度(m/s$^2$)', fontproperties='SimSun', fontsize=15)
plt.tight_layout()
plt.savefig(f'{savepath}\\行人荷载作用下节点{Obj[0]}竖向加速度时程图（{Frq_max}Hz）.tiff')
print(f'最大行人荷载频率为{round(Frq_max,3)}Hz，最大竖向加速度为{round(Acc_max,3)}m/s2')

f = open(f'{savepath}\\行人荷载作用下节点{Obj[0]}竖向加速度时程图（{Frq_max}Hz）.txt', 'w')
for n in range(0, NumberResults):
    f.write(f'{StepNum[n]} {U3[n]}\n')
f.close()

##########################
#  VerCrowd analysis     #
##########################
ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
ret = SapModel.Results.Setup.SetOptionModalHist(1)
VerCrowd_Frqs_1 = [1.25 + 0.05 * x for x in range(0, 26, 1)]
VerCrowd_Frqs_2 = [2.6 + 0.1 * x for x in range(0, 21, 1)]
VerCrowd_Frqs = VerCrowd_Frqs_1 + VerCrowd_Frqs_2
VerCrowd_Frqs = [round(x, 2) for x in VerCrowd_Frqs]
for frq in VerCrowd_Frqs:
    ret = SapModel.Results.Setup.SetCaseSelectedForOutput(f'VerCrowd{frq}')
nn = len(VerCrowd_Frqs)

[NumberResults, Obj, Elm, LoadCase,
 StepType, StepNum, U1, U2, U3, R1, R2, R3, Remark] = SapModel.Results.JointAccAbs(ObjectName[0], 0)

Acc = []
for n in range(0, nn):
    frq = VerCrowd_Frqs[n]
    LoadCase_i = f'VerCrowd{frq}'
    index_i = LoadCase.index(LoadCase_i)
    U3_tem = [abs(j) for j in U3[index_i: index_i+2]]
    Acc.append(max(U3_tem))
plt.figure(4)
plt.plot(VerCrowd_Frqs, Acc)
plt.title('竖向人群荷载扫频计算频率与最大加速度关系图', fontproperties='SimSun', fontsize=15)
plt.xlabel('计算频率(Hz)', fontproperties='SimSun', fontsize=15)
plt.ylabel('各工况最大加速度(m/s$^2$)', fontproperties='SimSun', fontsize=15)
plt.tight_layout()
plt.savefig(f'{savepath}\\竖向人群荷载扫频计算频率与最大加速度关系图.tiff')

f = open(f'{savepath}\\竖向人群荷载扫频计算频率与最大加速度关系图.txt', 'w')
for n in range(0, nn):
    f.write(f'{VerCrowd_Frqs[n]} {Acc[n]}\n')
f.close()
# 找出最大竖向加速度和对应工况
Acc_max = max(Acc)
Frq_max = VerCrowd_Frqs[Acc.index(Acc_max)]
ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
ret = SapModel.Results.Setup.SetOptionModalHist(2)
ret = SapModel.Results.Setup.SetCaseSelectedForOutput(f'VerCrowd{Frq_max}')
[NumberResults, Obj, Elm, LoadCase,
 StepType, StepNum, U1, U2, U3, R1, R2, R3, Remark] = SapModel.Results.JointAccAbs(ObjectName[0], 0)
plt.figure(5)
plt.plot(StepNum, U3)
plt.title(f'竖向人群荷载作用下节点{Obj[0]}竖向加速度时程图（{Frq_max}Hz）', fontproperties='SimSun', fontsize=15)
plt.xlabel('时间(s)', fontproperties='SimSun', fontsize=15)
plt.ylabel('最大工况的竖向加速度(m/s$^2$)', fontproperties='SimSun', fontsize=15)
plt.tight_layout()
plt.savefig(f'{savepath}\\竖向人群荷载作用下节点{Obj[0]}竖向加速度时程图（{Frq_max}Hz）.tiff')
print(f'最大竖向人群荷载为{round(Frq_max,3)}Hz，最大竖向加速度为{round(Acc_max,3)}m/s2')

f = open(f'{savepath}\\竖向人群荷载作用下节点{Obj[0]}竖向加速度时程图（{Frq_max}Hz）.txt', 'w')
for n in range(0, NumberResults):
    f.write(f'{StepNum[n]} {U3[n]}\n')
f.close()
##########################
#  LatCrowd analysis     #
##########################

ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
ret = SapModel.Results.Setup.SetOptionModalHist(1)
LatCrowd_Frqs = [0.5 + 0.05 * x for x in range(0, 15, 1)]
LatCrowd_Frqs = [round(x, 2) for x in LatCrowd_Frqs]

for frq in LatCrowd_Frqs:
    ret = SapModel.Results.Setup.SetCaseSelectedForOutput(f'LatCrowd{frq}')
nn = len(LatCrowd_Frqs)

[NumberResults, Obj, Elm, LoadCase,
 StepType, StepNum, U1, U2, U3, R1, R2, R3, Remark] = SapModel.Results.JointAccAbs(ObjectName[0], 0)

Acc = []
for n in range(0, nn):
    frq = LatCrowd_Frqs[n]
    LoadCase_i = f'LatCrowd{frq}'
    index_i = LoadCase.index(LoadCase_i)
    if LatDirection == 'X':
        U_tem = [abs(j) for j in U1[index_i: index_i + 2]]
    elif LatDirection == 'Y':
        U_tem = [abs(j) for j in U2[index_i: index_i+2]]
    Acc.append(max(U_tem))
plt.figure(6)
plt.plot(LatCrowd_Frqs, Acc)
plt.title('横向人群荷载扫频计算频率与最大加速度关系图', fontproperties='SimSun', fontsize=15)
plt.xlabel('计算频率(Hz)', fontproperties='SimSun', fontsize=15)
plt.ylabel('各工况最大加速度(m/s$^2$)', fontproperties='SimSun', fontsize=15)
plt.tight_layout()
plt.savefig(f'{savepath}\\横向人群荷载扫频计算频率与最大加速度关系图.tiff')

f = open(f'{savepath}\\横向人群荷载扫频计算频率与最大加速度关系图.txt', 'w')
for n in range(0, nn):
    f.write(f'{LatCrowd_Frqs[n]} {Acc[n]}\n')
f.close()
# 找出最大竖向加速度和对应工况
Acc_max = max(Acc)
Frq_max = LatCrowd_Frqs[Acc.index(Acc_max)]
ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
ret = SapModel.Results.Setup.SetOptionModalHist(2)
ret = SapModel.Results.Setup.SetCaseSelectedForOutput(f'LatCrowd{Frq_max}')
[NumberResults, Obj, Elm, LoadCase,
 StepType, StepNum, U1, U2, U3, R1, R2, R3, Remark] = SapModel.Results.JointAccAbs(ObjectName[0], 0)
plt.figure(7)
if LatDirection == 'X':
    plt.plot(StepNum, U1)
elif LatDirection == 'Y':
    plt.plot(StepNum, U2)
plt.title(f'横向人群荷载作用下节点{Obj[0]}横向加速度时程图（{Frq_max}Hz）', fontproperties='SimSun', fontsize=15)
plt.xlabel('时间(s)', fontproperties='SimSun', fontsize=15)
plt.ylabel('最大工况的横向加速度(m/s$^2$)', fontproperties='SimSun', fontsize=15)
plt.tight_layout()
plt.savefig(f'{savepath}\\横向人群荷载作用下节点{Obj[0]}横向加速度时程图（{Frq_max}Hz）.tiff')
print(f'最大横向人群荷载频率为{round(Frq_max,3)}Hz，最大横向加速度为{round(Acc_max,3)}m/s2')

f = open(f'{savepath}\\横向人群荷载作用下节点{Obj[0]}横向加速度时程图（{Frq_max}Hz）.txt', 'w')
for n in range(0, NumberResults):
    if LatDirection == 'X':
        f.write(f'{StepNum[n]} {U1[n]}\n')
    elif LatDirection == 'Y':
        f.write(f'{StepNum[n]} {U2[n]}\n')
f.close()


plt.show()
