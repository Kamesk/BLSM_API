import datetime
import time
from selenium import webdriver
import pandas as pd
from common_library import lib_csi_requests
from common_library import lib_selection_central
from constants import lib_constants
from common_library import lib_dr_sku_BLMA
from common_library import lib_remedyAPI
from common_library import lib_fc_research
from common_library import lib_deprecation
import csv
import os
import getpass
from selenium import webdriver
from selenium.webdriver.support.expected_conditions import _find_element
from webdriver_manager.chrome import ChromeDriverManager
from collections import defaultdict

user = os.getlogin()
username = getpass.getuser()
_find_element(By.CLASS_NAME)

find_element

# def write_excel_log(data_dict):
#     d = data_dict
#     list_data = list(d.values())
#     if len(list_data) != 0:
#         f = open("blma_test.csv", 'w', newline='')
#         writer = csv.writer(f)
#         writer.writerow(list_data)

# browser = webdriver.Chrome(executable_path=r"C:\\Users\\"+ username +"\\Desktop\\Paragon\\chromedriver.exe")
browser = webdriver.Chrome()
time.sleep(40)

path = (r"C:\\Users\\"+username+"\\Desktop\\BLMA\\conc2.xlsx")

df = pd.read_excel(path, sheet_name='Sheet1')
print(df)
d = {'UPC':[],'Asins':[],'Fc':[],'Linking_Asins':[],'Status':[],'Resolution':[]}
d = defaultdict(list)
for index, row in df.iterrows():
    print(row)
    # d = {}
    # print(d)
    d['UPC'].append(row['UPC'])
    d['Asins'].append(row['Asins'])
    d['FC'].append(row['FC'])
    # d['upc'] = str(row['UPC'])
    a = row['Asins']
    b = row['FC']
    # d['asins'] = a
    # d['fc'] = str(b)
    print(a)
    print(b)
    upc = str(row['UPC'])
    if len(str(upc)) < 12:
        upc = '{:0>12}'.format(upc)
    print(upc)

    # drskuData1 = lib_dr_sku_BLMA.drSkudataExtraction(browser,str(b),upc)
    # drskuData = drskuData1['linkingAsinList']

    drskuData = lib_dr_sku_BLMA.DRSKU_linkcheck(upc, str(b))
    if drskuData == 'Exception':
        print("exception")
        continue
    # d.append['asins_linking'] = drskuData
    d['Linking_Asins'].append(drskuData)

    drSkuDecision = ''
    if drskuData == 'Exception':
        print("started with  exception")
        continue
    else:
        if len(drskuData) == 0:
            # d['status'] = 'moved to BNL'
            d['status'].append('Moved to BNL')
            d['resolution'].append(lib_constants.corres_1)
            # d['resolution'] = lib_constants.corres_1
            
            print("moved to bnl1")
            

        elif len(drskuData) == 1:
            d['status'].append('Single ASIN')
            d['resolution'].append(lib_constants.corres_2)
            # d['status'] = 'Single ASIN'
            # d['resolution'] = lib_constants.corres_2
            
            # print("moved to bnl2")


        elif len(drskuData) > 2:
            d['status'].append('More than 2 ASINs')
            # d['status'] = 'More than 2 ASINs'
            d['resolution'].append(lib_constants.corres_3)
            # d['resolution'] = lib_constants.corres_3
            
            # print("moved to bnl3")


        elif len(drskuData) == 2:
            print('multiple')
            d['status'].append('2 ASINs')
            
            asinlist = drskuData
            if len(asinlist) == 2:
                asins_gl = []
                for i in asinlist:
                    dict_gl = {}
                    gl_check = lib_csi_requests.gl_check(str(i), '1')
                    if gl_check == 'Exception':
                        continue
                    dict_gl['asin'] = i
                    dict_gl['GL'] = gl_check
                    asins_gl.append(dict_gl)
                asins_gl_df = pd.DataFrame(asins_gl)
                asins_gl_df1 = asins_gl_df[(asins_gl_df['GL'] == 'gl_pantry')]
                asins_gl_df2 = asins_gl_df[(asins_gl_df['GL'] != 'gl_pantry')]
                asinslist = asins_gl_df['asin'].unique().tolist()  # Total ASINs linking to a barcode
                pantry_asin_list = asins_gl_df1['asin'].unique().tolist()  # Number of pantry ASINs
                non_pantry_asin_list = asins_gl_df2['asin'].unique().tolist()  # Number of non pantry ASINs
                print('Pantry:', pantry_asin_list)
                if len(asins_gl_df1['asin']) == 0:
                    print('No pantry ASINs')
                    is_asin_retail = []
                    for i in asinslist:
                        dict_retail = {}
                        retail_check = lib_csi_requests.retail_check(str(i))
                        if len(retail_check) > 0:
                            asin_status = 'retail'
                            # print("retail asin")
                        else:
                            asin_status = 'Non retail'
                            # print("non retail asin")

                        dict_retail['asin'] = i
                        dict_retail['retail_status'] = asin_status
                        is_asin_retail.append(dict_retail)
                    is_asin_retail_df = pd.DataFrame(is_asin_retail)
                    is_asin_retail_df1 = is_asin_retail_df[(is_asin_retail_df['retail_status'] == 'retail')]
                    is_asin_retail_df2 = is_asin_retail_df[
                        (is_asin_retail_df['retail_status'] != 'retail')]
                    asin_retail_list = is_asin_retail_df1['asin'].unique().tolist()
                    asin_3p_list = is_asin_retail_df2['asin'].unique().tolist()
                    if len(asin_3p_list) == 0:
                        # d['status'] = 'No 3P ASIN to merge'
                        d['status'].append('No 3P ASIN to merge')
                        d['resolution'].append('No 3P ASIN to merge')
                        # d['resolution'] = 'No 3P ASIN to Merge'
                        continue
                    print('Retail:', asin_retail_list)
                    print('3P:', asin_3p_list)
                    if len(is_asin_retail_df1) == 0:
                        print('No retail ASINs')
                    elif len(is_asin_retail_df1) == 1:
                        print('Merge')
                        merge_check = lib_deprecation.retail_3p_check(browser, asin_retail_list[0], asin_3p_list[0])
                        if merge_check == 'Exception':
                            continue
                        elif merge_check == 'Merge':
                            if lib_deprecation.deprecation_merge(browser, asin_retail_list[0], asin_3p_list[0]) == 'Exception':
                                continue
                            else:
                                # d['status'] = 'merge_request'
                                d['status'].append('merge_request')
                                # d['resolution'] = lib_constants.corres_11
                                d['resolution'].append(lib_constants.corres_11)
                                


                        else:
                            #Cleave
                            cleave = lib_dr_sku_BLMA.upc_cleave(browser, str(row['UPC'])
                                                          ,asin_3p_list[0],asin_retail_list[0])
                            if cleave == 'Exception':
                                continue
                            elif cleave == 'Success':

                                # d['status'] = 'Cleave'
                                d['status'].append('Cleave')
                                # d['resolution'] = 'Cleave has been done.'
                                d['resolution'].append('Cleave has been done.')
                                

                            elif cleave == 'Partial':
                                # d['status'] = 'Partial_cleave'
                                d['status'].append('Partial_cleave')
                                # d['resolution'] = 'Partial Cleave has been done.'
                                d['resolution'].append('Partial Cleave has been done.')
                                

                            else:
                                # d['status'] = 'Cleave_failure'
                                d['status'].append('Cleave_failure')
                                d['resolution'] = 'Cleave resulted in failure.'
                                d['resolution'].append('Cleave resulted in failure.')
                                



                    elif len(is_asin_retail_df1) > 1:
                        print('Deprecation')
                        duplicate_check = lib_csi_requests.duplicate_asins_check(asin_retail_list[0], asin_retail_list[1])
                        if duplicate_check == 'Exception':
                            continue
                        else:
                            if duplicate_check == 'Duplicate':
                                print('Duplicate retail asins')
                                isbn_list = []
                                for i in asin_retail_list:
                                    dict_isbn = {}
                                    print(type(i))
                                    if type(i) == int:
                                        is_isbn = 'yes'
                                    else:
                                        is_isbn = 'no'

                                    dict_isbn['asin_isbn'] = i
                                    dict_isbn['is_isbn'] = is_isbn
                                    isbn_list.append(dict_isbn)
                                is_isbn_df = pd.DataFrame(isbn_list)
                                is_isbn_df1 = is_isbn_df[(is_isbn_df['is_isbn'] == 'yes')]
                                is_isbn_df2 = is_isbn_df[(is_isbn_df['is_isbn'] == 'no')]
                                yes_isbn_list = is_isbn_df1['asin_isbn'].unique().tolist()
                                no_isbn_list = is_isbn_df2['asin_isbn'].unique().tolist()
                                if len(is_isbn_df1) == 0:
                                    print('No ISBNs')
                                    filter_check = lib_deprecation.filter_check(browser, no_isbn_list[0], no_isbn_list[1])
                                    if filter_check == 'Exception':
                                        continue
                                    else:
                                        if filter_check == 'True':
                                            print('Filter satisfied')
                                            if lib_deprecation.deprecation_merge(browser, no_isbn_list[0], no_isbn_list[1]) == 'Exception':
                                                continue
                                            else:
                                                # d['status'] = 'deprecation_filed'
                                                d['status'].append('deprecation_filed')
                                                # d['resolution'] = lib_constants.corres_18
                                                d['resolution'].append(lib_constants.corres_18)
                                                

                                        else:
                                            # d['status'] = 'Deprecation cannot be filed, filters not satisfied.'
                                            d['status'].append('Deprecation cannot be filed, filters not satisfied.')
                                            # d['resolution'] = 'Deprecation cannot be filed, filters not satisfied.'
                                            d['resolution'].append('Deprecation cannot be filed, filters not satisfied.')
                                            



                                elif len(is_isbn_df1) == 1:
                                    print('One ISBN')
                                    if lib_deprecation.filter_check(browser, no_isbn_list[0], yes_isbn_list[0]) == 'Exception':
                                        continue
                                    else:
                                        if lib_deprecation.filter_check(browser, no_isbn_list[0], yes_isbn_list[0]) == 'True':
                                            if lib_deprecation.deprecation_merge(browser, no_isbn_list[0], yes_isbn_list[0]) == 'Exception':
                                                continue
                                            else:
                                                # d['status'] = 'deprecation_filed'
                                                d['status'].append('deprecation_filed')
                                                # d['resolution'] = lib_constants.corres_18
                                                d['resolution'].append(lib_constants.corres_18)
                                                

                                        else:
                                            # d['status'] = 'Deprecation cannot be filed, filters not satisfied.'
                                            d['status'].append('Deprecation cannot be filed, filters not satisfied.')
                                            d['resolution'].append('Deprecation cannot be filed, filters not satisfied.')
                                            




                                elif len(is_isbn_df1) > 1:
                                    print('Both are ISBNs')
                                    # d['status'] = 'Both are ISBNs'
                                    d['status'].append('Both are ISBNs')
                                    # d['resolution'] = 'Without PO details ISBNs cannot be handled.'
                                    d['resolution'].append('Without PO details ISBNs cannot be handled.')
                                    



                            else:
                                print('Not duplicate asins')
                                selection_asins = lib_selection_central.selection_central(browser, upc, 'US')
                                if len(selection_asins) == 0:
                                    print('No Asins linking in SC')
                                    print('move to manual')
                                elif len(selection_asins) == 1:
                                    print('One Asin linking in SC')
                                    if selection_asins[0] == asin_retail_list[0]:
                                        upc = lib_csi_requests.upc_contribution_fetch(asin_retail_list[1], str(row['UPC']))
                                        a = []
                                        for i in list(upc.keys()):
                                            if 'Amazon.' in i:
                                                a.append('Retail')
                                            elif 'localization' in i:
                                                a.append('Localization')
                                            else:
                                                a.append('3P')
                                        if 'Retail' in a:
                                            print('Contributed by retail, move to manual')
                                            # d['status'] = 'UPC Contributed by retail'
                                            d['status'].append('UPC Contributed by retail')
                                            d['resolution'].append(lib_constants.corres_7)
                                            # d['resolution'] = lib_constants.corres_7
                                            

                                        elif 'Localization' in a:
                                            print('Contributed by localization team')
                                            detail_description = 'Hi,\n\nASIN :' + str(asin_retail_list[1]) + '\n\nUPC:' + str(row['UPC']) + \
                                                                 '\n\nIncorrect contribution is from 3P ASIN Localization. \n\nKindly cleave ASIN ' + str(asin_retail_list[1]) + ' from UPC ' + str(row['UPC']) +\
                                                                 ' and assign it to ASIN ' + str(asin_retail_list[0]) + '.\n\nThanks!'

                                            localization_TT = lib_remedyAPI.create_ticket_api('Catalog',
                                                'ASIN+Localization', 'Localization+Data',
                                                '', '','Barcode+links+incorrect+ASIN',
                                                detail_description,'4', '-',
                                                '-', '3p-asin-localization-team','localization_flip')
                                            if localization_TT == 'Exception':
                                                continue
                                            else:
                                                # d['status'] = 'localization_flip'
                                                d['status'].append('localization_flip')
                                                d['resolution'].append(lib_constants.corres_8 + 'https://tt.amazon.com/'+str(localization_TT))
                                                # d['resolution'] = lib_constants.corres_8 + 'https://tt.amazon.com/'+str(localization_TT)
                                                


                                        else:
                                            #Cleave
                                            cleave = lib_dr_sku_BLMA.upc_cleave(browser, str(row['UPC'])
                                                                                , asin_retail_list[1], asin_retail_list[0])
                                            if cleave == 'Exception':
                                                continue
                                            elif cleave == 'Success':

                                                d['status'].append('Cleave')
                                                d['resolution'].append('Cleave has been done.')
                                                

                                            elif cleave == 'Partial':
                                                d['status'].append('Partial_cleave')
                                                d['resolution'].append('Partial Cleave has been done.')
                                                

                                            else:
                                                d['status'].append('Cleave_failure')
                                                d['resolution'].append('Cleave resulted in failure.')
                                                
                                    else:
                                        upc = lib_csi_requests.upc_contribution_fetch(asin_retail_list[0], str(row['UPC']))
                                        a = []
                                        for i in list(upc.keys()):
                                            if 'Amazon.' in i:
                                                a.append('Retail')
                                            elif 'localization' in i:
                                                a.append('Localization')
                                            else:
                                                a.append('3P')
                                        if 'Retail' in a:
                                            print('Contributed by retail, move to manual')
                                            # d['status'] = 'UPC Contributed by retail'
                                            d['status'].append('UPC Contributed by retail')
                                            d['resolution'].append(lib_constants.corres_7)
                                            # d['resolution'] = lib_constants.corres_7
                                            

                                        elif 'Localization' in a:
                                            print('Contributed by localization team')
                                            detail_description = 'Hi,\n\nASIN :' + str(asin_retail_list[1]) + '\n\nUPC:' + str(row['UPC']) + \
                                                                 '\n\nIncorrect contribution is from 3P ASIN Localization. \n\nKindly cleave ASIN ' + str(asin_retail_list[1]) + ' from UPC ' + str(row['UPC']) +\
                                                                 ' and assign it to ASIN ' + str(asin_retail_list[0]) + '.\n\nThanks!'

                                            localization_TT = lib_remedyAPI.create_ticket_api('Catalog',
                                                'ASIN+Localization', 'Localization+Data',
                                                '', '','Barcode+links+incorrect+ASIN',
                                                detail_description,'4', '-',
                                                '-', '3p-asin-localization-team','localization_flip')
                                            if localization_TT == 'Exception':
                                                continue
                                            else:
                                                d['status'].append('localization_flip')
                                                d['resolution'].append(lib_constants.corres_8 + 'https://tt.amazon.com/'+str(localization_TT))
                                                


                                        else:
                                            #Cleave
                                            cleave = lib_dr_sku_BLMA.upc_cleave(browser, str(row['UPC'])
                                                                                , asin_retail_list[0], asin_retail_list[1])
                                            if cleave == 'Exception':
                                                continue
                                            elif cleave == 'Success':

                                                d['status'].append('Cleave')
                                                d['resolution'].append('Cleave has been done.')
                                                

                                            elif cleave == 'Partial':
                                                d['status'].append('Partial_cleave')
                                                d['resolution'].append('Partial Cleave has been done.')
                                                

                                            else:
                                                d['status'].append('Cleave_failure')
                                                d['resolution'].append('Cleave resulted in failure.')
                                                


                                elif len(selection_asins) > 1:
                                    print('More than 1 Asin linking in SC')
                                    d['status'].append('More than 1 Asin linking in SC')
                                    d['resolution'].append(lib_constants.corres_19)
                                    


                elif len(asins_gl_df1['asin']) == 1:
                    print('One pantry ASIN')

                    time.sleep(5)
                    inventory_core_fc = lib_fc_research.fc_research_check(browser, str(row['FC']), pantry_asin_list[0])
                    print('inventory', inventory_core_fc)
                    if inventory_core_fc == 'Exception':
                        continue
                    elif inventory_core_fc != 'No element found':
                        print('C4')
                        c4 = lib_constants.corres_4a + str(pantry_asin_list[0]) + lib_constants.corres_4a1 + str(non_pantry_asin_list[0]) + '.'
                        d['status'].append('pantry reassign')
                        d['resolution'].append(c4)
                        


                    else:
                        upc_unlink = lib_dr_sku_BLMA.unlinkasintobarcode(browser, str(row['FC']), str(row['UPC']), pantry_asin_list[0])
                        if upc_unlink == 'Exception':
                            continue
                        elif upc_unlink == 'Failure':
                            d['status'].append('Unlink failure')
                            d['resolution'].append(lib_constants.corres_6)
                            
                        elif upc_unlink == 'Success':
                            d['status'].append('Unlink Success')
                            d['resolution'].append(lib_constants.corres_5)
                            

                elif len(asins_gl_df1['asin']) > 1:
                    d['status'].append('Both ASINs are pantry. Cannot process.')
                    d['resolution'].append(lib_constants.corres_20)
                    
print(d)
df1 = pd.DataFrame.from_dict(d, orient='index')
df2 = df1.transpose()
print(df2)
# df1 = pd.DataFrame(d)
df2.to_csv(r"C:\\Users\\"+username+"\\Desktop\\BLMA\\blma_test.csv", index = False)