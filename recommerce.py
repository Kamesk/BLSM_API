import pandas as pd
import getpass

username = getpass.getuser()

recommerced = pd.read_excel("C:\\Users\\kamesk\\Desktop\\Work_Flow\\Recommerce.xlsx")
Asin_recommerced = recommerced[['ASIN']] 
C = recommerced['units_graded']
D = recommerced['sellable_units']
E = recommerced['unsellable_units']
F = recommerced['dfwd_negative']
G = recommerced['dfwd_positive']
H = recommerced['damage_yes']
I = recommerced['damage_no']
J = recommerced['wi_positive']
K = recommerced['wi_negative']
L = recommerced['ipm_positive']
M = recommerced['ipm_negative']
N = recommerced['cca_positive']
O = recommerced['cca_negative']
P = recommerced['pdc_positive']
Q = recommerced['pdc_negative']

C_sr = pd.Series(C)
D_sr = pd.Series(D)
E_sr = pd.Series(E)
F_sr = pd.Series(F)
G_sr = pd.Series(G)
H_sr = pd.Series(H)
I_sr = pd.Series(I)
J_sr = pd.Series(J)
K_sr = pd.Series(K)
L_sr = pd.Series(L)
M_sr = pd.Series(M)
N_sr = pd.Series(N)
O_sr = pd.Series(O)
P_sr = pd.Series(P)
Q_sr = pd.Series(Q)

units_graded = C_sr
DDEsum = D_sr+E_sr
GFGsum = F_sr+G_sr
HHIsum = H_sr+I_sr
JJKsum = J_sr+K_sr
LLMsum = L_sr+M_sr
NNOsum = N_sr+O_sr
PPQsum = P_sr+Q_sr

R_sr = units_graded
S_sr = D_sr.div(DDEsum)
T_sr = E_sr.div(DDEsum)
U_sr = G_sr.div(GFGsum)
V_sr = H_sr.div(HHIsum)
W_sr = J_sr.div(JJKsum)
X_sr = L_sr.div(LLMsum)
Y_sr = N_sr.div(NNOsum)
Z_sr = P_sr.div(PPQsum)

R_frame = C
s_frame = S_sr.to_frame()
S_frame = s_frame.fillna(0)
S_frame = S_frame[[0]].round(decimals=2)	
t_frame = T_sr.to_frame()
T_frame = t_frame.fillna(0)
T_frame = T_frame[[0]].round(decimals=2)
u_frame = U_sr.to_frame()
U_frame = u_frame.fillna(0)
U_frame = U_frame[[0]].round(decimals=2)
v_frame = V_sr.to_frame()
V_frame = v_frame.fillna(0)
V_frame = V_frame[[0]].round(decimals=2)
w_frame = W_sr.to_frame()
W_frame = w_frame.fillna(0)
W_frame = W_frame[[0]].round(decimals=2)
x_frame = X_sr.to_frame()
X_frame = x_frame.fillna(0)
X_frame = X_frame[[0]].round(decimals=2)
y_frame = Y_sr.to_frame()
Y_frame = y_frame.fillna(0)
Y_frame = Y_frame[[0]].round(decimals=2)
z_frame = Z_sr.to_frame()
Z_frame = z_frame.fillna(0)
Z_frame = Z_frame[[0]].round(decimals=2)

R_result = pd.concat([Asin_recommerced,R_frame], axis=1, join='outer')
R_result['col'] = 'Units Graded - ' + R_result['units_graded'].astype(str) 

S_result = pd.concat([Asin_recommerced,S_frame], axis=1, join='outer')
S_result['col'] = 'Moved to Sellable bin - ' + S_result[[0]].astype(str) + '%'
# print(S_result)
T_result = pd.concat([Asin_recommerced,T_frame], axis=1, join='outer')
T_result['col'] = 'Moved to unsellable bin - ' + T_result[[0]].astype(str) + '%'

U_result = pd.concat([Asin_recommerced,U_frame], axis=1, join='outer')
U_result['col'] = 'Different from Website Description - ' + U_result[[0]].astype(str) + '%'

V_result = pd.concat([Asin_recommerced,V_frame], axis=1, join='outer')
V_result['col'] = 'Damage - ' + V_result[[0]].astype(str) + '%'

W_result = pd.concat([Asin_recommerced,W_frame], axis=1, join='outer')
W_result['col'] = 'Wrong Item - ' + W_result[[0]].astype(str) + '%'

X_result = pd.concat([Asin_recommerced,X_frame], axis=1, join='outer')
X_result['col'] = 'Missing Item - ' + X_result[[0]].astype(str) + '%'

Y_result = pd.concat([Asin_recommerced,Y_frame], axis=1, join='outer')
Y_result['col'] = 'Customer Comments Accuracy% - ' + Y_result[[0]].astype(str) + '%'

Z_result = pd.concat([Asin_recommerced,Z_frame], axis=1, join='outer')
Z_result['col'] = 'Defect Confirmed % - ' + Z_result[[0]].astype(str) + '%'


combined_data = pd.concat([R_result, S_result, T_result, U_result, V_result, X_result, Y_result, Z_result], axis=0)

combined_data.to_csv(r'C:\\Users\\kamesk\\Desktop\\Work_Flow\\combined_recommerce.csv', header=False,index=False)
