import os
import sys
import comtypes.client
import matplotlib.pyplot as plt
import numpy as np
import math
# create API helper object
helper = comtypes.client.CreateObject('SAP2000v1.Helper')
helper = helper.QueryInterface(comtypes.gen.SAP2000v1.cHelper)
# Get SapObject
mySapObject = helper.GetObject("CSI.SAP2000.API.SapObject")
SapModel = mySapObject.SapModel
path = r'E:\\工程项目\\绵阳科技城新区科技大会堂\\SAP2000模型\\时程分析\\时程分析\\'
#######################################################
# 0.Note: follow items should be checked before run !!#
#######################################################
# 1.Waves list and angle list should be writen, parameters myDamping, dS, Vel, T1, T2, Dampd1, Dampd2 for
# Multi-Support Excitation should be determined !!!
# 2.The wave value should be scaled already based on Chinese standard
# as the SF in load case is 1.0, and is useless in Multi-Support Excitation!!!
# 3. The unit should be unified N,mm or kN,m,The Waves, dS, Vel should be set carefully !!!
# 4.The wave file name should be _X.txt _Y.txt _Z.txt for Uniform Excitation
# 5.The wave file name should be _UX.txt _UY.txt _UZ.txt for Multi-Support Excitation

#####################################################################
# 1.Define Model Linear History Time LoadCases (Uniform Excitation) #
#####################################################################
Waves = ['Artwave-1', 'Artwave-2', 'RSN15_KERN_TAF_1', 'RSN316_WESMORL_PTS_2',
         'RSN594_AuSableForks02_04_20_NE_EMMW__2', 'RSN6897_DARFIELD_2', 'RSN3746_CAPEMEND_CBF_1'
         ]  # Wave names
myDamping = 0.04  # Constant Damping , should be checked

# Change the units to N mm
ret = SapModel.SetPresentUnits(9)

for wave in Waves:
    # Define History-Time Function from file (Acceleration)
    ret = SapModel.Func.FuncTH.SetFromFile_1(wave+'_X', path + wave + '_X.txt', 0, 0, 1, 2, True)
    ret = SapModel.Func.FuncTH.SetFromFile_1(wave+'_Y', path + wave + '_Y.txt', 0, 0, 1, 2, True)
    ret = SapModel.Func.FuncTH.SetFromFile_1(wave+'_Z', path + wave + '_Z.txt', 0, 0, 1, 2, True)
    [NumberItems, MyTime, Value, Remark] = SapModel.Func.GetValues(wave+'_X')
    for direcID in ['X', 'Y']:
        # Set CaseName
        # Uniform Excitation
        MyCaseName = wave + '_' + direcID
        ret = SapModel.LoadCases.ModHistLinear.SetCase(MyCaseName)
        # Set DampConstant
        ret = SapModel.LoadCases.ModHistLinear.SetDampConstant(MyCaseName, myDamping)
        # # Set Damp by  period or frequency
        # ret = SapModel.LoadCases.ModHistLinear.SetDampInterpolated("LCASE1", 5, 3, MyTime, MyDamp)
        # # Set Damp Override
        # ret = SapModel.LoadCases.ModHistLinear.SetDampOverrides("LCASE1", 2, MyMode, MyDamp)
        # # Set the proportional modal damping
        # ret = SapModel.LoadCases.ModHistLinear.SetDampProportional("LCASE1", 2, 0, 0, 0.1, 1, 0.05, 0.06)
        # Set Loads
        MyLoadType = ["Accel"]*3
        MyLoadName = ["U1", "U2", "U3"]
        if direcID == 'X':
            MyFunc = [wave + '_X', wave + '_Y', wave + '_Z']
        elif direcID == 'Y':
            MyFunc = [wave + '_Y', wave + '_X', wave + '_Z']
        MySF = [1, 1, 1]
        MyTF = [1]*3
        MyAT = [0]*3
        MyCSys = ["Global"]*3
        MyAng = [0] * 3
        ret = SapModel.LoadCases.ModHistLinear.SetLoads(MyCaseName, len(MyLoadName), MyLoadType, MyLoadName, MyFunc,
                                                        MySF, MyTF, MyAT, MyCSys, MyAng)
        # Set TimeSteps
        ret = SapModel.LoadCases.ModHistLinear.SetTimeStep(MyCaseName, NumberItems, MyTime[1] - MyTime[0])


############################################################################
# 2.Define Model Linear History Time LoadCases  (Multi-Support Excitation) #
############################################################################
# dS = 51000  # m
# Vel = 800  # wave velocity 800m/s
# T1, T2, Dampd1, Dampd2 = 0.8, 0.08, 0.03, 0.03
# # T2要取到后面周期
# for angle in [0]:  # Unit has been changed 60/360*2*pi in cos()
#     for wave in Waves:
#         # Define History-Time Function from file (Displacement)
#         ret = SapModel.Func.FuncTH.SetFromFile_1(wave + '_UX', path + wave + '_UX.txt', 0, 0, 1, 2, True)
#         ret = SapModel.Func.FuncTH.SetFromFile_1(wave + '_UY', path + wave + '_UY.txt', 0, 0, 1, 2, True)
#         ret = SapModel.Func.FuncTH.SetFromFile_1(wave + '_UZ', path + wave + '_UZ.txt', 0, 0, 1, 2, True)
#         [NumberItems, MyTime, Value, Remark] = SapModel.Func.GetValues(wave+'_UX')
#
#         # Define BaseMotionGroup
#         # 1. Get the Coordinates of Supported Points and calculate the relative coordinates Si
#         DOF = ["True"]*6
#         ret = SapModel.SelectObj.SupportedPoints(DOF)
#         [NumberNames, MyName, Remarks] = SapModel.PointObj.GetNameList()
#         TagSelected_Point = []
#         for PointTag in MyName:
#             [sta, Remarks] = SapModel.PointObj.GetSelected(PointTag, "Selected")
#             if sta:
#                 TagSelected_Point.append(PointTag)
#         S_dic = dict()
#         for PointTag in TagSelected_Point:
#             [x, y, z, Remarks] = SapModel.PointObj.GetCoordCartesian(PointTag)
#             S_dic[PointTag] = x*np.cos(angle/360*2*np.pi) + y*np.sin(angle/360*2*np.pi)
#         KeyTag = sorted(S_dic, key=S_dic.__getitem__)
#         S_total = S_dic[KeyTag[-1]]-S_dic[KeyTag[0]]
#         BMGroupNum = math.ceil((S_total+0.1)/dS)
#         # 2.Define Empty BaseMotionGroup and Load Pattern
#         for k in range(1, BMGroupNum + 1):
#             ret = SapModel.LoadPatterns.Add(f"Basmotion_UX_{k}_{angle}", 8, 0, False)  # LTYPE_OTHER = 8
#             ret = SapModel.LoadPatterns.Add(f"Basmotion_UY_{k}_{angle}", 8, 0, False)  # LTYPE_OTHER = 8
#             ret = SapModel.LoadPatterns.Add(f"Basmotion_UZ_{k}_{angle}", 8, 0, False)  # LTYPE_OTHER = 8
#
#         for k in range(1, BMGroupNum + 1):
#             ret = SapModel.GroupDef.SetGroup(f"ExcitaionGroup{k}_{angle}")
#             ret = SapModel.GroupDef.Clear(f"ExcitaionGroup{k}_{angle}")        # removes all assignments from the specified group
#         # 3.Assign joint to BaseMotionGroup and apply Load Pattern
#         for PointTag in TagSelected_Point:
#             n_Group = math.ceil((S_dic[PointTag]+0.1-S_dic[KeyTag[0]])/dS)
#             ret = SapModel.PointObj.SetGroupAssign(PointTag, f"ExcitaionGroup{n_Group}_{angle}")
#             Value = [0]*6
#             Value[0] = 1  # U1 = 1
#             ret = SapModel.PointObj.SetLoadDispl(PointTag, f"Basmotion_UX_{n_Group}_{angle}", Value, True, "Global")
#             Value = [0]*6
#             Value[1] = 1  # U2 = 1
#             ret = SapModel.PointObj.SetLoadDispl(PointTag, f"Basmotion_UY_{n_Group}_{angle}", Value, True, "Global")
#             Value = [0]*6
#             Value[2] = 1  # U3 = 1
#             ret = SapModel.PointObj.SetLoadDispl(PointTag, f"Basmotion_UZ_{n_Group}_{angle}", Value, True, "Global")
#
#         # Set CaseName
#         for direcID in ['X', 'Y']:
#             MyCaseName = f"{wave}_{direcID}_{angle}_Multi_Exc"
#             ret = SapModel.LoadCases.DirHistLinear.SetCase(MyCaseName)
#             # Set the proportional modal damping by period    CaseName, DampType, Da, Db, T1, T2, Dampd1, Dampd2)
#             ret = SapModel.LoadCases.DirHistLinear.SetDampProportional(MyCaseName, 2, 1, 1, T1, T2, Dampd1, Dampd2)
#             # Set Loads
#             MyLoadType = ["Load"]*BMGroupNum*3
#             MyLoadName = [f"Basmotion_{m}_{str(n + 1)}_{angle}" for n in range(0, BMGroupNum) for m in ["UX", "UY", "UZ"]]
#             if direcID == "X":
#                 MyFunc = [wave + '_UX', wave + '_UY', wave + '_UZ']*BMGroupNum
#             elif direcID == "Y":
#                 MyFunc = [wave + '_UY', wave + '_UX', wave + '_UZ'] * BMGroupNum
#             MySF = [1]*BMGroupNum*3  # MySF useless when Load Type item is not Accel !!!
#             MyTF = [1]*BMGroupNum*3
#             # Calculate Arrive Time
#             DT = dS/1000/Vel
#             time = np.linspace(0, (BMGroupNum - 1) * DT, BMGroupNum)
#             MyAT = sorted(list(time)*3)
#             MyCSys = ["Global"]*BMGroupNum*3   # MyCSys useless when Load Type item is not Accel !!!
#             MyAng = [0]*BMGroupNum*3           # Angle useless when Load Type item is not Accel !!!
#
#             ret = SapModel.LoadCases.DirHistLinear.SetLoads(MyCaseName, 3*BMGroupNum, MyLoadType, MyLoadName, MyFunc, MySF, MyTF,
#                                                             MyAT, MyCSys, MyAng)
#             # Set TimeSteps
#             Total_NumberItems = NumberItems + (BMGroupNum - 1) * DT / (MyTime[1] - MyTime[0])
#             ret = SapModel.LoadCases.DirHistLinear.SetTimeStep(MyCaseName, int(Total_NumberItems), MyTime[1] - MyTime[0])

####################
# 3.Define  Output #
####################
# Define PointGroup for Output, 整体指标01，不同分区节点的位移时程曲线
# MyOutputPoints = "OutputPointsGroup"
# ret = SapModel.GroupDef.SetGroup(MyOutputPoints)
# PointTags = [194, 261, 1222, 1235]
# for PointTag in PointTags:
#     SapModel.PointObj.SetGroupAssign(str(PointTag), MyOutputPoints)
#
# # Define FrameGroup for Output, 关键构件指标01，不同关键构件需命名不同的集，可采用软件进行定义，但在后处理时，需相应的修改
# MyOutputFrame = 'OutputFramesGroup01'
# MyOutputFrame = 'OutputFramesGroup02'
# MyOutputFrame = 'OutputFramesGroup03'
# MyOutputFrame = 'OutputFramesGroup04'
# ret = SapModel.GroupDef.SetGroup(MyOutputFrame)
# FrameTags = [44, 46, 48]
# for FrameTag in FrameTags:
#     SapModel.FrameObj.SetGroupAssign(str(FrameTag), MyOutputFrame)



########################
# 4.Analysis LoadCases #
########################
# Change the CaseFlag
# ret = SapModel.Analyze.SetRunCaseFlag("MODAL", True)
# Save model
# ret = SapModel.File.Save(path+"test.sdb")
# Run Analysis
# ret = SapModel.Analyze.RunAnalysis()




