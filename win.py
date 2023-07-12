import pandas as pd
import time
import getpass
from csv import reader

username = getpass.getuser()

recommerced = pd.read_excel("C:\\Users\\"+ username +"\\Desktop\\Work_Flow\\Recommerce.xlsx")
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

combined_data.to_csv(r"C:\\Users\\"+ username +"\\Desktop\\Work_Flow\\combined_recommerce.csv", header=False,index=False)



df_serlock = (r"C:\\Users\\"+ username +"\\Desktop\\Work_Flow\\file_serlock.csv")
file_serlock = pd.read_csv(df_serlock, names=['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','Q','R','S','T','U','V','W','X','Y','Z','A1'])
Asin_Sherlock = file_serlock['A']
A1 = file_serlock['C']
A2 = file_serlock['D']
A3 = file_serlock['E']
A4 = file_serlock['F']
A5 = file_serlock['G']
A6 = file_serlock['H']
A7 = file_serlock['I']
A8 = file_serlock['J']
A9 = file_serlock['K']
A10 = file_serlock['L']
A11 = file_serlock['M']
A12 = file_serlock['N']
A13 = file_serlock['O']
A14 = file_serlock['Q']
A15 = file_serlock['S']
A16 = file_serlock['T']
A17 = file_serlock['U']
A18 = file_serlock['V']
A19 = file_serlock['W']
A20 = file_serlock['X']
A21 = file_serlock['Y']
A22 = file_serlock['Z']
A23 = file_serlock['A1']

A1 = A1.fillna(0)
A2 = A2.fillna(0)
A3 = A3.fillna(0)
A4 = A4.fillna(0)
A5 = A5.fillna(0)
A6 = A6.fillna(0)
A7 = A7.fillna(0)
A8 = A8.fillna(0)
A9 = A9.fillna(0)
A10 = A10.fillna(0)
A11 = A11.fillna(0)
A12 = A12.fillna(0)
A13 = A13.fillna(0)
A14 = A14.fillna(0)
A15 = A15.fillna(0)
A16 = A16.fillna(0)
A17 = A17.fillna(0)
A18 = A18.fillna(0)
A19 = A19.fillna(0)
A20 = A20.fillna(0)
A21 = A21.fillna(0)
A22 = A22.fillna(0)
A23 = A23.fillna(0)

R_S = pd.concat([Asin_Sherlock,A1], axis=1, join='inner', ignore_index=True)
S_S = pd.concat([Asin_Sherlock,A2], axis=1, join='inner', ignore_index=True)
T_S = pd.concat([Asin_Sherlock,A3], axis=1, join='inner', ignore_index=True)
U_S = pd.concat([Asin_Sherlock,A4], axis=1, join='inner', ignore_index=True)
V_S = pd.concat([Asin_Sherlock,A5], axis=1, join='inner', ignore_index=True)
W_S = pd.concat([Asin_Sherlock,A6], axis=1, join='inner', ignore_index=True)
X_S = pd.concat([Asin_Sherlock,A7], axis=1, join='inner', ignore_index=True)
Y_S = pd.concat([Asin_Sherlock,A8], axis=1, join='inner', ignore_index=True)
Z_S = pd.concat([Asin_Sherlock,A9], axis=1, join='inner', ignore_index=True)
A_S = pd.concat([Asin_Sherlock,A10], axis=1, join='inner', ignore_index=True)
P_S = pd.concat([Asin_Sherlock,A11], axis=1, join='inner', ignore_index=True)
B_S = pd.concat([Asin_Sherlock,A12], axis=1, join='inner', ignore_index=True)
C_S = pd.concat([Asin_Sherlock,A13], axis=1, join='inner', ignore_index=True)
D_S = pd.concat([Asin_Sherlock,A14], axis=1, join='inner', ignore_index=True)
E_S = pd.concat([Asin_Sherlock,A15], axis=1, join='inner', ignore_index=True)
F_S = pd.concat([Asin_Sherlock,A16], axis=1, join='inner', ignore_index=True)
H_S = pd.concat([Asin_Sherlock,A17], axis=1, join='inner', ignore_index=True)
I_S = pd.concat([Asin_Sherlock,A18], axis=1, join='inner', ignore_index=True)
J_S = pd.concat([Asin_Sherlock,A19], axis=1, join='inner', ignore_index=True)
K_S = pd.concat([Asin_Sherlock,A20], axis=1, join='inner', ignore_index=True)
L_S = pd.concat([Asin_Sherlock,A21], axis=1, join='inner', ignore_index=True)
M_S = pd.concat([Asin_Sherlock,A22], axis=1, join='inner', ignore_index=True)
N_S = pd.concat([Asin_Sherlock,A23], axis=1, join='inner', ignore_index=True)


combined_sherlock = pd.concat([R_S, S_S, T_S, U_S, V_S, X_S, Y_S, Z_S, A_S, B_S, C_S, P_S, D_S, E_S, F_S, H_S, I_S, J_S, K_S, L_S, M_S, N_S], axis=0)
combined_sherlock.to_csv(r"C:\\Users\\"+ username +"\\Desktop\\Work_Flow\\file_Serlock_out.csv", index=False, header=False)
New_sherlock = pd.read_csv(r"C:\\Users\\"+ username +"\\Desktop\\Work_Flow\\file_Serlock_out.csv")
dfx = New_sherlock[pd.to_numeric(New_sherlock['1'], errors='coerce').isna()]
Asin_sherlock = dfx['asin list']
ram = dfx['1']
entire = 'Sherlock- ' + ram.astype(str)
Finalizer = pd.concat([Asin_sherlock, entire], axis=1)
final = Finalizer.rename(columns={'1': 'Checklist'})
pathptd = pd.read_excel(r"C:\\Users\\"+ username +"\\Desktop\\Work_Flow\\CoMP usecases report.xlsx")
path = r"C:\\Users\\"+ username +"\\Desktop\\Work_Flow\\input_WA.xlsx"
indf = pd.read_excel(path)
merger = indf[["asin list", "product_type"]]
check_list = pd.merge(pathptd,merger, on = "product_type", how = 'right')
check_done = check_list.dropna(subset=['Checklist'])
filz = check_done[["asin list", "Checklist"]]
file2 = final.append(filz)
# file2.to_excel("C:\\Users\\"+ username +"\\Desktop\\Work_Flow\\abc.xlsx", index = False)
# drag1 = pd.read_excel("C:\\Users\\"+ username +"\\Desktop\\Work_Flow\\abc.xlsx")
sorted_file = file2.sort_values(by=['asin list'])
sorted_file.to_csv(r"C:\\Users\\"+ username +"\\Desktop\\Work_Flow\\file_Serlockz_out.csv", index=False)
reil = pd.read_csv(r"C:\\Users\\"+ username +"\\Desktop\\Work_Flow\\file_Serlockz_out.csv")
result1 = [reil['asin list'][0]]
result2 = [reil['Checklist'][0]]
for a, b in zip(reil['asin list'][1:], reil['Checklist'][1:]):
    try:
        if a == result1[-1]:
            result2[-1] += "\n" + b
        else:
            result1.append(a)
            result2.append(b)
    except:
        continue
Sherlock_df = pd.DataFrame({'asin list': result1, 'Checklist': result2})
Final_checklist = pd.merge(Sherlock_df,merger, on = "asin list", how = 'right')
Finaliz = Final_checklist.dropna(subset = ['Checklist'])
Finaliz.to_csv(r"C:\\Users\\"+ username +"\\Desktop\\Work_Flow\\Input_try.csv", index =False)
time.sleep(2)
chck = pd.read_csv(r"C:\\Users\\"+ username +"\\Desktop\\Work_Flow\\Input_try.csv")
wa_fil = indf[['FCF estimate','asin list','MP','pg_rollup', 'brand_name','item_name','quartile','gl_product_group_desc', 'replenishment_category_desc', 'gcv', 'ncv', 'conceded_units', 'ncv_exda', 'shipped_units', 'item_type_keyword', 'product_type', 'product_category', 'product_subcategory', 'model_name', 'model_number','OPS','NeCPU', 'ean']].merge(chck[['asin list','Checklist']],on = 'asin list', how='left')
extra_head = pd.read_csv(r"C:\\Users\\"+ username +"\\Desktop\\Work_Flow\\Input_extra.csv")
Final_file = wa_fil[['FCF estimate','asin list','MP','pg_rollup', 'brand_name','item_name', 'quartile','gl_product_group_desc', 'replenishment_category_desc', 'gcv', 'ncv', 'conceded_units', 'ncv_exda', 'shipped_units', 'item_type_keyword', 'product_type', 'product_category', 'product_subcategory', 'model_name', 'model_number','OPS','NeCPU','ean','Checklist']].join(extra_head[["Is_Ncv_Valid","Is_User_valid","Is_Cti_valid","inventory","Date_month", "Date_day","MPA", "Date_year","Concessions_bad_detail_page", "Concessions_Damage","Concessions_Defective",   "Fit/Style", "Ship_Day",    "Return_Day",  "Order_ID", "Concession_reason",  "shipment_ship_day|shipped_units_orig|concession_creation_day|concession_reason|gcv|ncv|ncv_exda|conceded_units|conceded_shipment_item_id|concession_type_code",  "PR Filter",   "vendor_id",   "comments", "URL", "Status",  "ms_link", "ASIN_Picked_at",  "Ticket_Pending_at", "Ticket_resolved_at",  "Closure_code",    "Is ASIN Picked",  "ASIN Picked By",  "Ticket Created",  "Ticket ID",   "Number", "pri_reas_cat",    "Pri_reas_sub_cat"  , "Pri_relev_comments",  "Pri_root_cause",  "Pri_root_cause_owner",    "Pri_action_taken",    "Pri_Results", "Pri_solution",  "Pri_FC",  "Pri_MS_Link", "Pri_ven_code",    "Pri_Incorrect_Attribute", "primary_corrected_Attribute", "Pri_Incorrect_value", "Pri_correct_value",   "Pri_Inc_Attrib_Existed_in",   "Pri_Status",  "Sec_reas_cat",    "Sec_reas_sub_cat",    "Sec_relev_comments",  "Sec_root_cause",  "Sec_root_cause_owner",    "Sec_action_taken",    "Sec_Results", "Sec_solution",    "Sec_FC",  "Sec_MS_Link", "Sec_ven_code",    "Sec_Incorrect_Attribute", "Secondary_corrected_Attribute",   "Sec_Incorrect_value", "Sec_correct_value",   "Sec_Inc_Attrib_Existed_in",   "Sec_Status",  "Ter_reas_cat",    "Ter_reas_sub_cat",    "Ter_relev_comments",  "Ter_root_cause",  "Ter_root_cause_owner",    "Ter_action_taken",    "Ter_Results", "Ter_solution",    "Ter_FC",  "Ter_MS_Link", "Ter_ven_code", "Ter_Incorrect_Attribute", "Ter_corrected_Attribute", "Ter_Incorrect_value", "Ter_correct_value", "Ter_Inc_Attrib_Existed_in", "Ter_Status","Is_Asin_Valid"]], how = 'left')
pathmpa = r"C:\\Users\\"+ username +"\\Desktop\\Work_Flow\\MPA.xlsx"
MPA = pd.read_excel(pathmpa)
new_mpa = MPA.rename(columns = {'asin':'asin list','comment':'MPA'}, inplace=False)
Final_out = pd.read_csv(r"C:\\Users\\"+ username +"\\Desktop\\Work_Flow\\Wk_test.csv")
MPAL = Final_file[['FCF estimate','asin list','MP','pg_rollup','item_name', 'brand_name', 'quartile','gl_product_group_desc', 'replenishment_category_desc', 'gcv', 'ncv', 'conceded_units', 'ncv_exda', 'shipped_units', 'item_type_keyword', 'product_type', 'product_category', 'product_subcategory', 'model_name', 'model_number','OPS','NeCPU','ean','Checklist',"Is_Ncv_Valid","Is_User_valid","Is_Cti_valid","inventory","Date_month", "Date_day","MPA", "Date_year","Concessions_bad_detail_page", "Concessions_Damage","Concessions_Defective",   "Fit/Style", "Ship_Day",    "Return_Day",  "Order_ID", "Concession_reason",  "shipment_ship_day|shipped_units_orig|concession_creation_day|concession_reason|gcv|ncv|ncv_exda|conceded_units|conceded_shipment_item_id|concession_type_code",  "PR Filter",   "vendor_id",   "comments", "URL", "Status",  "ms_link", "ASIN_Picked_at",  "Ticket_Pending_at", "Ticket_resolved_at",  "Closure_code",    "Is ASIN Picked",  "ASIN Picked By",  "Ticket Created",  "Ticket ID",   "Number", "pri_reas_cat",    "Pri_reas_sub_cat"  , "Pri_relev_comments",  "Pri_root_cause",  "Pri_root_cause_owner",    "Pri_action_taken",    "Pri_Results", "Pri_solution",  "Pri_FC",  "Pri_MS_Link", "Pri_ven_code",    "Pri_Incorrect_Attribute", "primary_corrected_Attribute", "Pri_Incorrect_value", "Pri_correct_value",   "Pri_Inc_Attrib_Existed_in",   "Pri_Status",  "Sec_reas_cat",    "Sec_reas_sub_cat",    "Sec_relev_comments",  "Sec_root_cause",  "Sec_root_cause_owner",    "Sec_action_taken",    "Sec_Results", "Sec_solution",    "Sec_FC",  "Sec_MS_Link", "Sec_ven_code",    "Sec_Incorrect_Attribute", "Secondary_corrected_Attribute",   "Sec_Incorrect_value", "Sec_correct_value",   "Sec_Inc_Attrib_Existed_in",   "Sec_Status",  "Ter_reas_cat",    "Ter_reas_sub_cat",    "Ter_relev_comments",  "Ter_root_cause",  "Ter_root_cause_owner",    "Ter_action_taken",    "Ter_Results", "Ter_solution",    "Ter_FC",  "Ter_MS_Link", "Ter_ven_code", "Ter_Incorrect_Attribute", "Ter_corrected_Attribute", "Ter_Incorrect_value", "Ter_correct_value", "Ter_Inc_Attrib_Existed_in", "Ter_Status","Is_Asin_Valid"]].merge(new_mpa[['asin list','MPA']], on = 'asin list', how='left')
MPAL.to_csv(r"C:\\Users\\"+ username +"\\Desktop\\Work_Flow\\Wk_Num_inputsheet_newz.csv", index=False)
user = pd.read_csv(r"C:\\Users\\"+ username +"\\Desktop\\Work_Flow\\Wk_Num_inputsheet_newz.csv", index_col ="brand_name")
try:
    MPA_finalZ = user.drop(["Apple", "Beats", "Ecco", "Lego", "Calvin Klein", "Simple Human", "Pampers", "Levi's", "Amazon Basics", "Havanians"],axis = 0, inplace = False)
    MPA_finalZ =pd.read_csv(r"C:\\Users\\"+ username +"\\Desktop\\Work_Flow\\Wk_Num_inputsheet_newz.csv")
except KeyError:
    MPA_finalZ =pd.read_csv(r"C:\\Users\\"+ username +"\\Desktop\\Work_Flow\\Wk_Num_inputsheet_newz.csv")
MPA_finalZ.loc[MPA_finalZ.MPA_y != None,'PR Filter'] = 'Regular/MPA'
Wk_update = MPA_finalZ[["asin list",  "MP",  "pg_rollup","brand_name",'quartile',"gl_product_group_desc",    "replenishment_category_desc", "Concessions_bad_detail_page", "Concessions_Damage",  "Concessions_Defective",   "Fit/Style",   "gcv", "ncv", "conceded_units",  "ncv_exda", "shipped_units",   "OPS", "Ship_Day",    "Return_Day",  "Order_ID",    "NeCPU",   "Concession_reason",   "shipment_ship_day|shipped_units_orig|concession_creation_day|concession_reason|gcv|ncv|ncv_exda|conceded_units|conceded_shipment_item_id|concession_type_code",   "item_type_keyword","product_type","product_category","product_subcategory", "PR Filter",   "item_name",   "vendor_id",   "model_number",    "ean", "comments",    "Checklist",   "URL", "Status",  "ms_link", "ASIN_Picked_at",  "Ticket_Pending_at",   "Ticket_resolved_at", "Closure_code",    "Is ASIN Picked","ASIN Picked By","Ticket Created","Ticket ID"]]
df_cut = Wk_update['brand_name']
res = df_cut.isin(['Amazon', 'NuPro', 'Blink Home Security', 'eero', 'Ring']).any().any()
if res :
    print("GL elimination combination Found!!!")
    Wk_update.to_csv(r"C:\\Users\\"+ username +"\\Desktop\\Work_Flow\\Wk_Num_inputsheet_new method.csv", index=False)
else:
    print("GL elimination combination not found")
    Wk_update.to_csv(r"C:\\Users\\"+ username +"\\Desktop\\Work_Flow\\Wk_Num_inputsheet_new method.csv", index=False)
print("Process completed")