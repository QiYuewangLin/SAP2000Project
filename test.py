import pandas as pd
import matplotlib.pyplot as plt

savePath = 'C:\\Users\\Lenovo\\Desktop\\'
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