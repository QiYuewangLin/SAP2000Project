import os
import sys
import math
import numpy as np
import comtypes.client
import matplotlib.pyplot as plt
import pandas as pd
# create API helper object
helper = comtypes.client.CreateObject('SAP2000v1.Helper')
helper = helper.QueryInterface(comtypes.gen.SAP2000v1.cHelper)
# Get SapObject
mySapObject = helper.GetObject("CSI.SAP2000.API.SapObject")
SapModel = mySapObject.SapModel
ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
# Change the units to kN m
ret = SapModel.SetPresentUnits(6)

# define group frame
FrameGroupName = ['1层竖向构件', '2层竖向构件', '3层竖向构件', '4层竖向构件', '5层竖向构件', '6层竖向构件']
savePath = 'C:\\Users\\Lenovo\\Desktop\\'
save_dict = {}
for FrameGroup in FrameGroupName:
    save_Displ = []  # 用于储存每层位移
    # get group frame
    [NumberItems, ObjectType, ObjectName, Remark] = SapModel.GroupDef.GetAssignments(FrameGroup)
    GD_name = []  # 广义位移名称

    for j, Frame in enumerate(ObjectName):
        if ObjectType[j] == 2:  # 选择框架构件
            [RDI, RDJ, Remark] = SapModel.FrameObj.GetPoints(Frame)
            # add generalized displacement
            [x1, y1, z1, Remark] = SapModel.PointObj.GetCoordCartesian(RDI)
            [x2, y2, z2, Remark] = SapModel.PointObj.GetCoordCartesian(RDJ)
            if abs(z1-z2) < 1:  # 过滤掉高差小于1m的杆件
                continue
            FrameLength = math.sqrt((x2-x1) ** 2 + (y2-y1) ** 2 + (z2-z1) ** 2)
            SF1 = (-1/FrameLength, 0, 0, 0, 0, 0)
            SF2 = (1 / FrameLength, 0, 0, 0, 0, 0)
            SF3 = (0, -1/FrameLength, 0, 0, 0, 0)
            SF4 = (0, 1 / FrameLength, 0, 0, 0, 0)
            # X方向
            ret = SapModel.GDispl.Delete("GDX" + Frame)   # 删除原有的点数据
            ret = SapModel.GDispl.Add("GDX"+Frame, 1)  # 1表示平动
            GD_name.append("GDX"+Frame)
            ret = SapModel.GDispl.SetPoint("GDX" + Frame, RDJ, SF1)
            ret = SapModel.GDispl.SetPoint("GDX" + Frame, RDI, SF2)
            # Y方向
            ret = SapModel.GDispl.Delete("GDY" + Frame)  # 删除原有的点数据
            GD_name.append("GDY" + Frame)
            ret = SapModel.GDispl.Add("GDY"+Frame, 1)  # 1表示平动
            ret = SapModel.GDispl.SetPoint("GDY" + Frame, RDJ, SF3)
            ret = SapModel.GDispl.SetPoint("GDY" + Frame, RDI, SF4)

    # 提取各工况广义位移角
    [NumberNames, MycaseNames, Remark] = SapModel.LoadCases.GetNameList_1()
    EQ_name = []
    for casename in MycaseNames:
        if casename[-2:] == "_X" or casename[-2:] == "_Y" or casename[-2:] == "EX" or casename[-2:] == "EY":
            EQ_name.append(casename)
            ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
            ret = SapModel.Results.Setup.SetCaseSelectedForOutput(casename)
            ret = SapModel.Results.Setup.SetOptionModalHist(1)  # 1 表示提取极值 2表示逐步提取
            GenDispl = []
            for GDi in GD_name:
                [NumberResults, GD, LoadCase, StepType,
                StepNum, DType, Value, Remark] = SapModel.Results.GeneralizedDispl(GDi)  # 仅输出一个点
                if len(Value) == 2:
                    GenDispl.append(max(abs(Value[0]), abs(Value[1])))   # 每层所有构件的最大层间位移角
                else:
                    GenDispl.append(abs(Value[0]))
            GenDispl_Max = max(GenDispl)                             # 该层的最大层间位移角
            # if casename == 'RSN3746_CAPEMEND_CBF_1_X':
            #     FrameID = GD_name[GenDispl.index(GenDispl_Max)]
            save_Displ.append(GenDispl_Max)

    save_dict[FrameGroup] = save_Displ
table = pd.DataFrame(save_dict)
Mas_disp = table.max(axis=1)
table.insert(0, '工况', EQ_name)
table.to_excel(f"{savePath}SAP2000层间位移角.xlsx", index=False)

# pandas 后处理数据

pd.set_option('display.unicode.east_asian_width', True)
df = pd.read_excel(f"{savePath}SAP2000层间位移角.xlsx")
df_x = df.iloc[range(0, len(df), 2)].T
df_y = df.iloc[range(1, len(df), 2)].T

# X方向层间位移角图
LoadCasename = []
for name in df_x.iloc[0]:
    if len(name) > 2:
        LoadCasename.append(name[:-2])
    else:
        LoadCasename.append("CQC")
F1 = plt.figure(1)
for col in df_x.columns:
    if df_x[col][0] == "EX":
        plt.plot(df_x[col][1:len(df_x.index)], range(1, len(df_x.index), 1), lw=3.5)
    else:
        plt.plot(df_x[col][1:len(df_x.index)], range(1, len(df_x.index), 1))
plt.plot([1/300, 1/300], [1, len(df_x.index)], linestyle='--', color='r', lw=3)
plt.title("X方向层间位移角", fontproperties='SimSun', fontsize=15)
plt.xlabel("层间位移角", fontproperties='SimSun', fontsize=15)
plt.ylabel("楼层", fontproperties='SimSun', fontsize=15)
x = range(1, 10, 1)
x = [lamda*0.00125 for lamda in x]
xlabes = [f'1/{str(int(1/lamda))}' for lamda in x]
plt.xticks(x, xlabes)
plt.legend(LoadCasename,  fontsize=8, loc='upper right')
plt.savefig(f'{savePath}\\X方向层间位移角.tiff')

# Y方向层间位移角图
LoadCasename = []
for name in df_y.iloc[0]:
    if len(name) > 2:
        LoadCasename.append(name[:-2])
    else:
        LoadCasename.append("CQC")
F2 = plt.figure(2)
for col in df_y.columns:
    if df_y[col][0] == "EY":
        plt.plot(df_y[col][1:len(df_y.index)], range(1, len(df_y.index), 1), lw=3.5)
    else:
        plt.plot(df_y[col][1:len(df_y.index)], range(1, len(df_y.index), 1))
plt.plot([1/300, 1/300], [1, len(df_y.index)], linestyle='--', color='r', lw=3)
plt.title("Y方向层间位移角", fontproperties='SimSun', fontsize=15)
plt.xlabel("层间位移角", fontproperties='SimSun', fontsize=15)
plt.ylabel("楼层", fontproperties='SimSun', fontsize=15)
x = range(1, 10, 1)
x = [lamda*0.00125 for lamda in x]
xlabes = [f'1/{str(int(1/lamda))}' for lamda in x]
plt.xticks(x, xlabes)
plt.legend(LoadCasename, fontsize=8, loc='upper right')
plt.savefig(f'{savePath}\\Y方向层间位移角.tiff')
plt.show()

# 楼层最大层间位移角表格
Mas_disp_x = df_x.iloc[1:].max().tolist()
Mas_disp_x.append(sum(Mas_disp_x[1:])/len(Mas_disp_x[1:]))
Mas_disp_x_ratio = [f'{str(int(lamda/Mas_disp_x[0]*100))}%' for lamda in Mas_disp_x]  # 百分比值
Mas_disp_x = [f'1/{str(int(1/lamda))}' for lamda in Mas_disp_x]  # 转化为分数

Mas_disp_y = df_y.iloc[1:].max().tolist()
Mas_disp_y.append(sum(Mas_disp_y[1:])/len(Mas_disp_y[1:]))
Mas_disp_y_ratio = [f'{str(int(lamda/Mas_disp_y[0]*100))}%' for lamda in Mas_disp_y]  # 百分比值
Mas_disp_y = [f'1/{str(int(1/lamda))}' for lamda in Mas_disp_y]  # 转化为分数

LoadCasename[0] = "CQC法"
LoadCasename.append("平均值")
DFdata = pd.DataFrame({"工况": LoadCasename, "X方向": Mas_disp_x, "X方向比值": Mas_disp_x_ratio, "Y方向": Mas_disp_y,
                       "Y方向比值": Mas_disp_y_ratio})
DFdata.to_excel(f"{savePath}SAP2000最大层间位移角对比.xlsx", index=False)