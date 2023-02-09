import math
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

[NumberNames, MycaseNames, Remark] = SapModel.LoadCases.GetNameList_1()
EQ_name = []
BaseForce = {}
BaseForce_x = []
BaseForce_y = []
for casename in MycaseNames:
    if casename[-2:] == "_X" or casename[-2:] == "_Y" or casename[-2:] == "EX" or casename[-2:] == "EY":
        EQ_name.append(casename)
        ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
        ret = SapModel.Results.Setup.SetCaseSelectedForOutput(casename)
        ret = SapModel.Results.Setup.SetOptionModalHist(1)  # 1 表示提取极值 2表示逐步提取
        [NumberResults, LoadCase, StepType, StepNum,
         Fx, Fy, Fz, Mx, My, Mz, gx, gy, gz, Remark] = SapModel.Results.BaseReact()
        if casename[-2:] == "_X" or casename[-2:] == "EX":
            Fx = max([int(abs(x)) for x in list(Fx)])
            BaseForce_x.append(Fx)
        else:
            Fy = max([int(abs(y)) for y in list(Fy)])
            BaseForce_y.append(Fy)

BaseForce_x.append(int(sum(BaseForce_x[1:])/len(BaseForce_x[1:])))
BaseForce_x_ratio = [f'{str(int(lamda/BaseForce_x[0]*100))}%' for lamda in BaseForce_x]  # 百分比值
BaseForce_y.append(int(sum(BaseForce_y[1:])/len(BaseForce_y[1:])))
BaseForce_y_ratio = [f'{str(int(lamda/BaseForce_y[0]*100))}%' for lamda in BaseForce_y]  # 百分比值

# BaseForceData = pd.DataFrame({"工况": LoadCasename, "X方向": BaseForce_x, "X方向比值": BaseForce_x_ratio, "Y方向": BaseForce_y,
#                        "Y方向比值": BaseForce_y_ratio})
BaseForceData = pd.DataFrame({"X方向": BaseForce_x, "X方向比值": BaseForce_x_ratio, "Y方向": BaseForce_y,
                       "Y方向比值": BaseForce_y_ratio})
print(BaseForceData)

