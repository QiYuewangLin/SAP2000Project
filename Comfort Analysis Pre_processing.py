import os
import sys
import numpy as np
"""
本程序可用于舒适度分析的前处理，包括改模型单位、改楼板类型为薄板（非膜）、改楼板弹模、改质量源、改活载数值、改插入点偏移
设置稳态工况、设置单点行人荷载工况、竖向人群荷载工况、设置横向人群荷载工况，但仍需手动操作的是支座约束施加
本程序适用于平面的楼盖、吊桥、连廊（插入点偏移的对象仅选中了某XY平面）
输入参数包括楼板截面名称、板厚、活载数值、弹模放大系数、最不利作用点集合、阻尼比、楼面面积，并需检查楼板支座的设置情况。
编写： 林晨豪 时间：20220910
"""
import comtypes.client
import pandas as pd
# create API helper object
helper = comtypes.client.CreateObject('SAP2000v1.Helper')
helper = helper.QueryInterface(comtypes.gen.SAP2000v1.cHelper)
# Get SapObject
mySapObject = helper.GetObject("CSI.SAP2000.API.SapObject")
SapModel = mySapObject.SapModel

# 输入参数
PlateSectionName = "TM120C30F"  # 楼板截面名称
plate_thickness = 0.12  # m 板厚
LiveLoad = 0.5          # 恒载按照YJK模型数据施加，面上活载室内楼盖0.5kN/m2，连廊和室内天桥0.35kN/m2
E_factor = 1.35         # 楼板弹模放大系数，1.35钢-混凝土组合楼盖，1.2钢筋混凝土楼盖
Disadvantage_PintTag = [2003452]       # 最不利点集合
kesi = 0.02              # 连廊、天桥：钢-混凝土楼盖0.01，钢楼盖0.005，混凝土楼板，0.05；楼盖：混凝土0.05，钢-混凝土0.02~0.05
Area = 9072              # 楼面面积，用于计算人群密度，单位m2

# Calculate Crow_tense
Crow_tense = 10.8*np.sqrt(0.5*kesi/Area)
# Change the units to kN m
ret = SapModel.SetPresentUnits(6)

# Change the E of concrete slab
[ShellType, IncludeDrillingDOF, MatProp, MatAng, Thickness, Bending, Color, Notes, GUID, Remarks]\
    = SapModel.PropArea.GetShell_1(PlateSectionName)
[e, u, a, g, Remarks] = SapModel.PropMaterial.GetMPIsotropic(MatProp)

MyE = e*E_factor          # 1.35钢-混凝土组合楼盖，1.2钢筋混凝土楼盖
MyU = u                   # 混凝土泊松比
MyA = a                   # 混凝土热膨胀系数
MyG = g                   # 混凝土剪切模量
ret = SapModel.PropMaterial.SetMPIsotropic(MatProp, MyE, MyU, MyA, MyG)  # 定义名为C30F的混凝土材料
print("Change the E of concrete slab completed!")

# Change the shell type
ret = SapModel.PropArea.SetShell(PlateSectionName, 1, MatProp, 0, plate_thickness, plate_thickness)
# Change the LoadPat
LoadPat = ["DEAD", "Live"]
SF = [1.0, 1.0]
# Change the mass resource
ret = SapModel.SourceMass.SetMassSource("MSSSRC1", False, False, True, True, 2, LoadPat, SF)
print("Change mass resource completed!")

# Change the live load
Area_Confirm = input("检查楼面是否完整，输入数字1表示完成：")
ret = SapModel.SelectObj.PropertyArea(PlateSectionName)
ret = SapModel.AreaObj.SetLoadUniformToFrame("Ignored", "LIVE", LiveLoad, 10, 2, True, "Global", 2)
ret = SapModel.SelectObj.PropertyArea(PlateSectionName, True)
print("Change live load completed!")

# Change the beam insertionPoint
ret = SapModel.SelectObj.PlaneXY(str(Disadvantage_PintTag[0]))
Offset1 = [0, 0, 0]
Offset2 = [0, 0, 0]
ItemType = 2
ret = SapModel.FrameObj.SetInsertionPoint("Ignore", 8, False, False, Offset1, Offset2, 'Local', ItemType)
ret = SapModel.SelectObj.PlaneXY(str(Disadvantage_PintTag[0]), True)
# Check Support
Boundary_Condition_Confirm = input("检查支座边界是否设置完毕，输入数字1表示完成：")

# Set PiontGroups
ret = SapModel.GroupDef.SetGroup("PointsGroup")
for ponit in Disadvantage_PintTag:
    ret = SapModel.PointObj.SetGroupAssign(str(ponit), "PointsGroup")
###################
# Modal analysis  #
###################

##########################
#  Steady-state analysis #
##########################

NumberItems = 2
Frequency = [0, 10]
Value = [1, 1]
ret = SapModel.Func.FuncSS.SetUser("Steady_Sate", NumberItems, Frequency, Value)

ret = SapModel.LoadPatterns.Add("Steady_Sate", 8)
Value = [0, 0, -1, 0, 0, 0]
ret = SapModel.PointObj.SetLoadForce("PointsGroup", "Steady_Sate", Value, True, "Global", 1)
ret = SapModel.LoadCases.SteadyState.SetCase("Steady_Sate")
ret = SapModel.LoadCases.SteadyState.SetDampConstant("Steady_Sate", 0, 2*kesi)  # 两倍阻尼比
MyFreqModalDev = [0]
MyFreqSpecified = [0]
ret = SapModel.LoadCases.SteadyState.SetFreqData("Steady_Sate", 1.0, 10, 1, True, False, False,
                                                 "MODAL", 2, MyFreqModalDev, 2, MyFreqSpecified)
MyLoadType = ["Load"]
MyLoadName = ["Steady_Sate"]
MyFunc = ["Steady_Sate"]
MySF = [1.0]
MyPhaseAngle = [0]
MyCSys = ["Global"]
MyAng = [0]
ret = SapModel.LoadCases.SteadyState.SetLoads("Steady_Sate", 1, MyLoadType, MyLoadName, MyFunc,
                                              MySF, MyPhaseAngle, MyCSys, MyAng)


##########################
#  Walking analysis     #
##########################
dt = 0.005
Time = 15
Walking_Frqs = [1.6 + 0.05 * x for x in range(0, 13, 1)]
Walking_Frqs = [round(x, 2) for x in Walking_Frqs]
Wakling_dt = [1/72/x for x in Walking_Frqs]
Waling_n = len(Walking_Frqs)
ret = SapModel.LoadPatterns.Add("Walking", 8)
for frq in Walking_Frqs:
    NumberItems = 601
    MyTime = [0.025*x for x in range(0,  NumberItems, 1)]
    Value = [0.35*np.cos(2*np.pi*frq*x)+0.14*np.cos(4*np.pi*frq*x+np.pi/2) + 0.07*np.cos(6*np.pi*frq*x + np.pi/2) for x in MyTime]
    ret = SapModel.Func.FuncTH.SetUser(f'Walking{frq}', NumberItems, MyTime, Value)
    ret = SapModel.LoadCases.ModHistLinear.SetCase(f'Walking{frq}')
    ret = SapModel.LoadCases.ModHistLinear.SetDampConstant(f'Walking{frq}', kesi)
    ret = SapModel.LoadCases.ModHistLinear.SetLoads(f'Walking{frq}', 1, ["Load"], ["Walking"], [f'Walking{frq}'],
                                                    [1], [1], [0], ["Global"], [0])
    ret = SapModel.LoadCases.ModHistLinear.SetTimeStep(f'Walking{frq}', 3000, 0.005)

##########################
#  VerCrowd analysis     #
##########################
dt = 0.005
Time = 15
VerCrowd_Frqs_1 = [1.25 + 0.05 * x for x in range(0, 26, 1)]
Reduce_factor1 = [min(1, (x-1.25)/(1.7-1.25)) for x in VerCrowd_Frqs_1 if x < 2.15]
Reduce_factor1 += [max(0.25, 1-(x-2.1)/(2.3-2.1)) for x in VerCrowd_Frqs_1 if x > 2.1]
VerCrowd_Frqs_2 = [2.6 + 0.1 * x for x in range(0, 21, 1)]
Reduce_factor2 = [min(0.25, 0.25*(1-(x-4.2)/(4.6-4.2))) for x in VerCrowd_Frqs_2]
VerCrowd_Frqs = VerCrowd_Frqs_1 + VerCrowd_Frqs_2
VerCrowd_Frqs = [round(x, 2) for x in VerCrowd_Frqs]
VerCrow_Reduce_factor = Reduce_factor1 + Reduce_factor2
ret = SapModel.LoadPatterns.Add("VerCrowd", 8)
for frq in VerCrowd_Frqs:
    NumberItems = 601
    MyTime = [0.025 * x for x in range(0, NumberItems, 1)]
    Re_fator = VerCrow_Reduce_factor[VerCrowd_Frqs.index(frq)]
    Value = [0.28 * Crow_tense * Re_fator * np.cos(2 * np.pi * frq * x) for x in MyTime]
    ret = SapModel.Func.FuncTH.SetUser(f'VerCrowd{frq}', NumberItems, MyTime, Value)
    ret = SapModel.LoadCases.ModHistLinear.SetCase(f'VerCrowd{frq}')
    ret = SapModel.LoadCases.ModHistLinear.SetDampConstant(f'VerCrowd{frq}', kesi)
    ret = SapModel.LoadCases.ModHistLinear.SetLoads(f'VerCrowd{frq}', 1, ["Load"], ["VerCrowd"], [f'VerCrowd{frq}'],
                                                    [1], [1], [0], ["Global"], [0])
    ret = SapModel.LoadCases.ModHistLinear.SetTimeStep(f'VerCrowd{frq}', 3000, 0.005)
##########################
#  LatCrowd analysis     #
##########################
dt = 0.005
Time = 15
LatCrowd_Frqs = [0.5 + 0.05 * x for x in range(0, 15, 1)]
LatCrowd_Frqs = [round(x, 2) for x in LatCrowd_Frqs]
Reduce_factor = [min(1, (x-0.5)/0.2) for x in LatCrowd_Frqs if x < 1.05]
Reduce_factor += [1-(x-1)/0.2 for x in LatCrowd_Frqs if x > 1]
ret = SapModel.LoadPatterns.Add("LatCrowd", 8)
for frq in LatCrowd_Frqs:
    NumberItems = 601
    MyTime = [0.025 * x for x in range(0, NumberItems, 1)]
    Re_fator = Reduce_factor[LatCrowd_Frqs.index(frq)]
    Value = [0.035 * Crow_tense * Re_fator * np.cos(2 * np.pi * frq * x) for x in MyTime]
    ret = SapModel.Func.FuncTH.SetUser(f'LatCrowd{frq}', NumberItems, MyTime, Value)
    ret = SapModel.LoadCases.ModHistLinear.SetCase(f'LatCrowd{frq}')
    ret = SapModel.LoadCases.ModHistLinear.SetDampConstant(f'LatCrowd{frq}', kesi)
    ret = SapModel.LoadCases.ModHistLinear.SetLoads(f'LatCrowd{frq}', 1, ["Load"], ["LatCrowd"], [f'LatCrowd{frq}'],
                                                    [1], [1], [0], ["Global"], [0])
    ret = SapModel.LoadCases.ModHistLinear.SetTimeStep(f'LatCrowd{frq}', 3000, 0.005)

####################
#  Add Load       #
####################
ret = SapModel.SelectObj.PropertyArea(PlateSectionName)
ret = SapModel.AreaObj.SetLoadUniformToFrame("Ignored", "VerCrowd", -1, 10, 2, True, "Global", 2)
ret = SapModel.AreaObj.SetLoadUniformToFrame("Ignored", "LatCrowd", -1, 5, 2, True, "Global", 2)
ret = SapModel.SelectObj.PropertyArea(PlateSectionName, True)



