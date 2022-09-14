import os
import sys
import comtypes.client
# e2k文件处理
"""
1.地震波文件、e2k文件、py文件需在同一文件夹
2.选择弹性时程工况或者非线性时程工况
3.目前属于增加新工况，未删去原有e2k文件工况和时程函数
4.反复操作时需要删去原来的，最好备份一个原始转换e2k文件
"""
filepath = r'E:\工程项目\先河沣启时代大厦\先河沣启时代大厦-AB塔\Etabs文件\弹塑性时程分析模型\弹塑性分析模型.e2k'  # 选择e2k文件路径
path = os.getcwd() + '\\'
Waves = ['Artwave-2', 'RSN92_SFERN_WRP', 'RSN1165_KOCAELI_IZT']  # 地震波名称
Damping = 0.05  # 阻尼比
directions = ['_X', '_Y', '_Z']  # 考虑三向地震
AnalysisType = 'Nonlinear'       # 选择弹性时程 'Elastic' or 非线性时程 'Nonlinear'
T = [3.485, 0.5]                 # 非线性时程工况用于定义阻尼比的结构周期'
casedirections = ['_X', '_Y']
FUNCTION_message = str()
LOADCASE_message = str()

# 非线性重力工况，用于非线性时程
LOADCASE_message += f'\n LOADCASE Gravity  TYPE  "Nonlinear Static"  INITCOND  "NONE"  MODALCASE  "Modal"  ' \
                    f'MASSSOURCE  "MsSrc1" \n'
LOADCASE_message += f' LOADCASE Gravity   LOADPAT  "Dead"  SF  1  \n'
LOADCASE_message += f' LOADCASE Gravity   LOADPAT  "Live"  SF  0.5  \n'
LOADCASE_message += f' LOADCASE Gravity   LOADCONTROL  "Full"  DISPLTYPE  "Monitored"  MONITOREDDISPL  ' \
                    f'"Joint"  DISPLMAG  0 DOF  "U1"  JOINT  "3"  "Story1" \n'
LOADCASE_message += f' LOADCASE Gravity  RESULTSSAVED  "Final" \n'
LOADCASE_message += f' LOADCASE Gravity  SOLUTIONSCHEME  "Iterative Events"  ' \
                    f'MAXTOTALSTEPS  200 MAXNULLSTEPS  50 MAXEVENTSPERSTEP  24  \n'
for wave in Waves:
    for direction in directions:
        wavename = wave + direction
        wavepath = path + wave + direction+'.txt'
        fid = open(wavepath, 'r')
        message = fid.readline()

        Loc1 = message.find('DT:')
        npts = message[6:Loc1-1]
        DT = message[Loc1+3:]
        fid.close()
        FUNCTION_message += f'\n FUNCTION {wavename} FUNCTYPE "HISTORY" FILE {wavepath}  DATATYPE "EQUAL"  DT {DT}'
        FUNCTION_message += f' FUNCTION {wavename} HEADERLINES 1  POINTSPERLINE 1  FORMAT "FREE"\n'
    # 线弹性时程分析模态分析
    if AnalysisType == 'Elastic':
        for casedirection in casedirections:
            wavecase = wave + casedirection
            LOADCASE_message += f'\n LOADCASE {wavecase}  TYPE  "Linear Modal History"  MODALCASE  "Modal" \n'
            if casedirection == '_X':
                LOADCASE_message += f' LOADCASE {wavecase}  ACCEL  "U1" FUNC  {wave}_X  SF  1 \n'
                LOADCASE_message += f' LOADCASE {wavecase}  ACCEL  "U2" FUNC  {wave}_Y  SF  1 \n'
            elif casedirection == '_Y':
                LOADCASE_message += f' LOADCASE {wavecase}  ACCEL  "U1" FUNC  {wave}_Y  SF  1 \n'
                LOADCASE_message += f' LOADCASE {wavecase}  ACCEL  "U2" FUNC  {wave}_X  SF  1 \n'
            if len(directions) == 3:
                LOADCASE_message += f' LOADCASE {wavecase}  ACCEL  "U3" FUNC  {wave}_Z  SF  1 \n'
            LOADCASE_message += f' LOADCASE {wavecase}  NUMBEROUTPUTSTEPS  {npts} OUTPUTSTEPSIZE  {DT}'
            LOADCASE_message += f' LOADCASE {wavecase}  MODALDAMPTYPE  "Constant"  CONSTDAMP  {Damping} '
    # 非线性时程分析
    elif AnalysisType == 'Nonlinear':
        for casedirection in casedirections:
            wavecase = wave + casedirection + '_Nonlinear'
            LOADCASE_message += f'\n LOADCASE {wavecase}  TYPE  "Nonlinear Direct Integration History"  ' \
                                f'INITCOND  Gravity  MODALCASE  "Modal"  MASSSOURCE  "MsSrc1" \n'
            if casedirection == '_X':
                LOADCASE_message += f' LOADCASE {wavecase}  ACCEL  "U1" FUNC  {wave}_X  SF  1 \n'
                LOADCASE_message += f' LOADCASE {wavecase}  ACCEL  "U2" FUNC  {wave}_Y  SF  1 \n'
            elif casedirection == '_Y':
                LOADCASE_message += f' LOADCASE {wavecase}  ACCEL  "U1" FUNC  {wave}_Y  SF  1 \n'
                LOADCASE_message += f' LOADCASE {wavecase}  ACCEL  "U2" FUNC  {wave}_X  SF  1 \n'
            if len(directions) == 3:
                LOADCASE_message += f' LOADCASE {wavecase}  ACCEL  "U3" FUNC  {wave}_Z  SF  1 \n'
            LOADCASE_message += f' LOADCASE {wavecase} NLGEOMTYPE  "PDelta" NUMBEROUTPUTSTEPS  {npts} OUTPUTSTEPSIZE  {DT}'
            LOADCASE_message += f' LOADCASE {wavecase} PRODAMPTYPE  "Period"  T1  {T[0]} DAMP1  {Damping} T2  {T[1]} DAMP2 {Damping}' \
                                f' MODALDAMPTYPE  "None" '
            LOADCASE_message += f' LOADCASE {wavecase} SOLUTIONSCHEME  "Iterative Events"  MAXEVENTSPERSTEP  24 ' \
                                f'ITERCONVTOL  0.02 USELINESEARCH  "Yes"'

# 定义塑性铰

# HINGE_message = str()
#
# HINGE_message += f'\n HINGE "Column"  DOF "FiberPM2M3"  FROMSECTION "Yes"  HINGERELLENGTH 0.1\n'
# HINGE_message += f' HINGE "Beam_M"  BEHAVIOR "Deformation Controlled"  DOF "M2"  TYPE "Moment-Rotation"  HYSTYPE "Kinematic" ' \
#                  f' "-E"  -0.025 -0.2 "-D"  -0.015 -0.2 "-C"  -0.015 -1.1 "-B"  0 -1 "B"  0 1 "C"  0.015 1.1  "D"  0.015 0.2 "E"  0.025 0.2\n'
# HINGE_message += f' HINGE "Beam_M"  ROTATIONSFP 1 ROTATIONSFN 1\n'
# HINGE_message += f' HINGE "Beam_M"  IO 0.003 LS 0.012 CP 0.015 -IO -0.003 -LS -0.012 -CP -0.015\n'
# HINGE_message += f' HINGE "Beam_V"  BEHAVIOR "Deformation Controlled"  DOF "V3"  TYPE "Force-Displacement"\n'
# HINGE_message += f' HINGE "Beam_V"  "-E"  -8 -0.2 "-D"  -6 -0.2 "-C"  -6 -1.25 "-B"  0 -1 "B"  0 1 "C"  6 1.25 "D"  6 0.2 "E"  8 0.2\n'
# HINGE_message += f' HINGE "Beam_V"  IO 2 LS 4 CP 6 -IO -2 -LS -4 -CP -6\n'
# HINGE_message += f' HINGE "Brace"  BEHAVIOR "Deformation Controlled"  DOF "P"  TYPE "Force-Displacement"'
# HINGE_message += f' HINGE "Brace"  "-E"  -8 -0.2 "-D"  -6 -0.2 "-C"  -6 -1.25 "-B"  0 -1 "B"  0 1 "C"  6 1.25 "D"  6 0.2 "E"  8 0.2\n'
# HINGE_message += f' HINGE "Brace"  IO 2 LS 4 CP 6 -IO -2 -LS -4 -CP -6\n'

# 写新e2k文件
print((filepath))
fid = open(filepath, 'r')
basedata = fid.read()
# n1 = basedata.find('$ FRAME HINGE PROPERTIES') ## 文件中必须已经有塑性铰列表
n2 = basedata.find('$ FUNCTIONS')
n3 = basedata.find('$ LOAD CASES')
# 将增加内容插入
# basedata_new = basedata[:n1+len('$ FRAME HINGE PROPERTIES')] + HINGE_message + basedata[n1+len('$ FRAME HINGE PROPERTIES')+1:n2+len('$ FUNCTIONS')]\
#                + FUNCTION_message + basedata[n2+len('$ FUNCTIONS')+1:n3+len('$ LOAD CASES')] + \
#                LOADCASE_message + basedata[n3:]
basedata_new = basedata[1:n2+len('$ FUNCTIONS')]\
               + FUNCTION_message + basedata[n2+len('$ FUNCTIONS')+1:n3+len('$ LOAD CASES')] + \
               LOADCASE_message + basedata[n3:]
fid.close()
fid1 = open(filepath, '+w')
fid1.write(basedata_new)
print('e2k文件处理完毕')




