import cx_Oracle
import xlsxwriter
from datetime import datetime, timedelta
import os
import pandas as pd

# Dependencies: Operator, Config Matrix Values, Liveries, Maintenix Data, Reconfigurations of seats, ovens, galleys, A-CKS, C-CKS intervals,
# Connection to intranet (e.g. via VPN)

# Defining valid aircraft list
aircraft_list = ['HP-1540CMP','HP-1556CMP','HP-1557CMP','HP-1558CMP','HP-1559CMP','HP-1560CMP','HP-1561CMP','HP-1562CMP','HP-1563CMP','HP-1564CMP','HP-1565CMP','HP-1567CMP','HP-1568CMP','HP-1569CMP']

# Requesting and validating correct aircraft input
while True:
    aircraft = input('Type a valid ERJ-190 aircraft registration in format HP-XXXXCMP: ')
    if aircraft_list.count(aircraft) == 1:
        print(f'Generating Aircraft Specification Sheet for {aircraft}...')
        break

# Date of report generation
today = datetime.today().strftime('%d-%b-%y')

# Oracle SQL Queries Definitions
query_ac_id = f'''
SELECT INV_AC_REG.AC_REG_CD, INV_INV.MANUFACT_DT, INV_INV.SERIAL_NO_OEM,
EQP_PART_NO.PART_NO_OEM AS AC_MODEL

FROM INV_AC_REG

INNER JOIN INV_INV ON
INV_AC_REG.INV_NO_ID = INV_INV.INV_NO_ID

INNER JOIN EQP_PART_NO ON
INV_INV.PART_NO_ID = EQP_PART_NO.PART_NO_ID

WHERE AC_REG_CD = '{aircraft}'
'''

query_ac_times = f'''
SELECT 
INV_CURR_USAGE.TSN_QT, INV_CURR_USAGE.DATA_TYPE_ID /*1 = FH, 10 = FC*/

FROM INV_AC_REG 

INNER JOIN INV_CURR_USAGE  ON 
INV_AC_REG.INV_NO_ID = INV_CURR_USAGE.INV_NO_ID 

WHERE AC_REG_CD = '{aircraft}'
'''

query_next_cck = f'''
SELECT EVT_SCHED_DEAD.SCHED_DEAD_DT
FROM TASK_TASK

INNER JOIN SCHED_STASK ON
TASK_TASK.TASK_DB_ID = SCHED_STASK.TASK_DB_ID AND
TASK_TASK.TASK_ID = SCHED_STASK.TASK_ID

INNER JOIN INV_AC_REG ON
SCHED_STASK.MAIN_INV_NO_ID = INV_AC_REG.INV_NO_ID AND
SCHED_STASK.MAIN_INV_NO_DB_ID = INV_AC_REG.INV_NO_DB_ID

INNER JOIN EVT_SCHED_DEAD ON 
sched_stask.sched_db_id = EVT_SCHED_DEAD.event_db_id  AND
sched_stask.sched_id = EVT_SCHED_DEAD.event_id 

INNER JOIN EVT_EVENT ON 
EVT_SCHED_DEAD.EVENT_ID = EVT_EVENT.EVENT_ID AND
EVT_SCHED_DEAD.EVENT_DB_ID = EVT_EVENT.EVENT_DB_ID

WHERE TASK_TASK.TASK_CD = 'C-CK-1 - ERJ190_CM' AND
EVT_EVENT.EVENT_STATUS_CD = 'ACTV' AND
EVT_SCHED_DEAD.USAGE_REM_QT < 1010 AND
INV_AC_REG.AC_REG_CD = '{aircraft}'
'''

query_last_cck = f'''
SELECT 
EVT_SCHED_DEAD.SCHED_DEAD_DT,
EVT_SCHED_DEAD.DATA_TYPE_ID

FROM TASK_TASK

INNER JOIN SCHED_STASK ON
TASK_TASK.TASK_DB_ID = SCHED_STASK.TASK_DB_ID AND
TASK_TASK.TASK_ID = SCHED_STASK.TASK_ID

INNER JOIN INV_AC_REG ON
SCHED_STASK.MAIN_INV_NO_ID = INV_AC_REG.INV_NO_ID AND
SCHED_STASK.MAIN_INV_NO_DB_ID = INV_AC_REG.INV_NO_DB_ID

INNER JOIN EVT_SCHED_DEAD ON 
sched_stask.sched_db_id = EVT_SCHED_DEAD.event_db_id  AND
sched_stask.sched_id = EVT_SCHED_DEAD.event_id 

INNER JOIN EVT_EVENT ON 
EVT_SCHED_DEAD.EVENT_ID = EVT_EVENT.EVENT_ID AND
EVT_SCHED_DEAD.EVENT_DB_ID = EVT_EVENT.EVENT_DB_ID

WHERE TASK_TASK.TASK_CD = 'C-CK-1 - ERJ190_CM' AND
EVT_EVENT.EVENT_STATUS_CD = 'COMPLETE' AND
EVT_SCHED_DEAD.USAGE_REM_QT <= 1095 AND
INV_AC_REG.AC_REG_CD = '{aircraft}'
'''
query_main_assys = f'''
SELECT 
EQP_PART_NO.PART_NO_OEM,
II.SERIAL_NO_OEM,
INV_CURR_USAGE.TSN_QT,
II.CONFIG_POS_SDESC,
INV_CURR_USAGE.DATA_TYPE_ID,
INV_CURR_USAGE.TSO_QT

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
    
INNER JOIN INV_CURR_USAGE ON
    ii.inv_no_db_id = inv_curr_usage.inv_no_db_id AND
    ii.inv_no_id = inv_curr_usage.inv_no_id

WHERE ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND (EQP_BOM_PART.BOM_PART_CD = '71-00-00-00' /*ENGINES*/
OR EQP_BOM_PART.BOM_PART_CD = '49-10-00-00' /*APU*/
OR EQP_BOM_PART.BOM_PART_CD = '32-21-00-02-1' /*NLG*/
OR EQP_BOM_PART.BOM_PART_CD = '32-11-01-01A' /*MLG LH*/
OR EQP_BOM_PART.BOM_PART_CD = '32-11-01-01B' /* MLG RH*/
)
'''

query_avionics = f'''
SELECT DISTINCT
INV_AC_REG.AC_REG_CD,EQP_MANUFACT.MANUFACT_NAME,EQP_PART_NO.PART_NO_OEM,EQP_BOM_PART.BOM_PART_CD, EQP_PART_NO.PART_NO_SDESC

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
WHERE ii.inv_class_cd IN 'TRK' 
AND ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND (EQP_BOM_PART.BOM_PART_CD = '22-11-01-01' /*GUIDANCE PANEL */
OR EQP_BOM_PART.BOM_PART_CD = '23-11-01-01' /*HF RADIO */
OR EQP_BOM_PART.BOM_PART_CD = '23-12-01-03' /*VHF 1-2-3 */
OR EQP_BOM_PART.BOM_PART_CD = '23-12-01-03-3' /*VHF DATA*/
OR EQP_BOM_PART.BOM_PART_CD = '31-31-01-01' /*DVDR*/
OR EQP_BOM_PART.BOM_PART_CD = '31-32-01-01' /*QAR*/
OR EQP_BOM_PART.BOM_PART_CD = '34-51-01-01' /*DME*/
OR EQP_BOM_PART.BOM_PART_CD = '23-24-01-01' /* PRINTER */
OR EQP_BOM_PART.BOM_PART_CD = '34-31-01-01' /* LLRA */
OR EQP_BOM_PART.BOM_PART_CD = '34-42-01-02' /* WXR RADAR */
OR EQP_BOM_PART.BOM_PART_CD = '34-41-01-01' /*GPWS*/
OR EQP_BOM_PART.BOM_PART_CD = '34-43-01-01' /*TCAS*/
OR EQP_BOM_PART.BOM_PART_CD = '34-56-01-01' /*GPS*/
OR EQP_BOM_PART.BOM_PART_CD = '34-52-01-01A' /*XPONDER*/
OR EQP_BOM_PART.BOM_PART_CD = '34-26-01-01' /*IRU*/
OR EQP_BOM_PART.BOM_PART_CD = '34-32-01-01' /*VOR*/
OR EQP_BOM_PART.BOM_PART_CD = '34-32-01-01' /*VHF NAV*/
OR EQP_BOM_PART.BOM_PART_CD = '31-61-01-01' /*DU*/
OR EQP_BOM_PART.BOM_PART_CD = '34-11-01-01' /*ISFD*/

)
'''

#Connecting to Maintenix Oracle Database
dsn_tns = cx_Oracle.makedsn('maintenixdb-test.somoscopa.com', '1521', service_name='COPAT')
conn = cx_Oracle.connect(user='MX_TEST', password='MXT35T2016', dsn=dsn_tns) 

#Executing queries and storing in pandas dataframes

df_ac_id = pd.read_sql(query_ac_id,con = conn)
df_ac_times = pd.read_sql(query_ac_times,con = conn)
df_last_cck = pd.read_sql(query_last_cck,con = conn)
df_next_cck = pd.read_sql(query_next_cck,con = conn)
df_main_assys = pd.read_sql(query_main_assys,con = conn)
df_avionics = pd.read_sql(query_avionics,con = conn)

# Getting aircraft ID information (AIRFRAME SECTION)
ac_model = 'ERJ 190-100 IGW'
ac_rg = aircraft
msn = df_ac_id['SERIAL_NO_OEM'][0]
man_date = (df_ac_id['MANUFACT_DT'][0]).strftime('%d-%b-%y')
filt_fh = df_ac_times['DATA_TYPE_ID'] == 1
filt_fc = df_ac_times['DATA_TYPE_ID'] == 10
ac_tsn_fh = int((df_ac_times[filt_fh])['TSN_QT'])
ac_tsn_fc = int((df_ac_times[filt_fc])['TSN_QT'])
mtw = '111,245'
mtow = '110,231'
mlw = '97,003'
mzfw = '90,169'
mfc = '4,298'
noise_cat = 'STAGE 3'
cat_status = 'CAT II'

# Maintenance Program info
ac_nextcck = df_next_cck['SCHED_DEAD_DT'][0].strftime('%d-%b-%y')
try:
    ac_lastcck = ((df_last_cck['SCHED_DEAD_DT'].nsmallest(1))[1]).strftime('%d-%b-%y')
except:
    ac_lastcck = (datetime.strptime(ac_nextcck,'%d-%b-%y') - timedelta(days=1095)).strftime('%d-%b-%y')

# Main Assys info
eng_model = 'CF34-10E5'
path_ear = r'C:\Users\ggalina\SpecSheetMakerEMB\CF34 & APU Removals.xlsx'
path_cm = r'C:\Users\ggalina\SpecSheetMakerEMB\Config Matrix EMB.xlsx'

# Engine L/H
filt_eng_lh = df_main_assys['CONFIG_POS_SDESC'] == '71-00-00-00 (LH)'
filt_eng_lh_tsn = (df_main_assys['CONFIG_POS_SDESC'] == '71-00-00-00 (LH)') & (df_main_assys['DATA_TYPE_ID'] == 1)
filt_eng_lh_csn = (df_main_assys['CONFIG_POS_SDESC'] == '71-00-00-00 (LH)') & (df_main_assys['DATA_TYPE_ID'] == 10)

eng_lh_sn = int(df_main_assys[filt_eng_lh]['SERIAL_NO_OEM'].values[0])
eng_lh_tsn = int(df_main_assys[filt_eng_lh_tsn]['TSN_QT'].values[0])
eng_lh_csn = int(df_main_assys[filt_eng_lh_csn]['TSN_QT'].values[0])

df_er = pd.read_excel(path_ear,sheet_name = 'CF34 Removals')
try:
    filt_er = (df_er['Shop Visit'] == 'Y') & (df_er['Rem ESN'] == eng_lh_sn)
    df_er1 = df_er.loc[filt_er].nlargest(1,'Hours')
    eng_lh_sv_tsn = int(df_er1.values[0][5])
    eng_lh_sv_csn = int(df_er1.values[0][6])
except:
    eng_lh_sv_tsn = 0
    eng_lh_sv_csn = 0

eng_lh_tslv = eng_lh_tsn - eng_lh_sv_tsn
eng_lh_cslv = eng_lh_csn - eng_lh_sv_csn

# Engine R/H
filt_eng_rh = df_main_assys['CONFIG_POS_SDESC'] == '71-00-00-00 (RH)'
filt_eng_rh_tsn = (df_main_assys['CONFIG_POS_SDESC'] == '71-00-00-00 (RH)') & (df_main_assys['DATA_TYPE_ID'] == 1)
filt_eng_rh_csn = (df_main_assys['CONFIG_POS_SDESC'] == '71-00-00-00 (RH)') & (df_main_assys['DATA_TYPE_ID'] == 10)

eng_rh_sn = int(df_main_assys[filt_eng_rh]['SERIAL_NO_OEM'].values[0])
eng_rh_tsn = int(df_main_assys[filt_eng_rh_tsn]['TSN_QT'].values[0])
eng_rh_csn = int(df_main_assys[filt_eng_rh_csn]['TSN_QT'].values[0])

df_er = pd.read_excel(path_ear,sheet_name = 'CF34 Removals')
try:
    filt_er = (df_er['Shop Visit'] == 'Y') & (df_er['Rem ESN'] == eng_rh_sn)
    df_er1 = df_er.loc[filt_er].nlargest(1,'Hours')
    eng_rh_sv_tsn = int(df_er1.values[0][5])
    eng_rh_sv_csn = int(df_er1.values[0][6])
except:
    eng_rh_sv_tsn = 0
    eng_rh_sv_csn = 0

eng_rh_tslv = eng_rh_tsn - eng_rh_sv_tsn
eng_rh_cslv = eng_rh_csn - eng_rh_sv_csn

# APU
filt_apu_sn = (df_main_assys['CONFIG_POS_SDESC'] == '49-10-00-00')
filt_apu_aot = (df_main_assys['CONFIG_POS_SDESC'] == '49-10-00-00') & (df_main_assys['DATA_TYPE_ID'] == 101017)
filt_apu_acyc = (df_main_assys['CONFIG_POS_SDESC'] == '49-10-00-00') & (df_main_assys['DATA_TYPE_ID'] == 101018)

apu_sn = df_main_assys[filt_apu_sn]['SERIAL_NO_OEM'].values[0]
apu_aot = int(df_main_assys[filt_apu_aot]['TSN_QT'].values[0])
apu_acyc = int(df_main_assys[filt_apu_acyc]['TSN_QT'].values[0])

df_ar = pd.read_excel(path_ear,sheet_name = 'APS 2300 Removals')
try:
    filt_ar = df_ar['Serial No# Off'] == apu_sn
    df_ar1 = df_ar.loc[filt_ar].nlargest(1,'DMM TSN')
    apu_sv_tsn = int(df_ar1.values[0][11])
    apu_sv_csn = int(df_ar1.values[0][12])
except:
    apu_sv_tsn = 0
    apu_sv_csn = 0

apu_tslv = apu_aot - apu_sv_tsn
apu_clsv = apu_acyc - apu_sv_csn

# NLG
filt_nlg_fh = (df_main_assys['CONFIG_POS_SDESC'] == '32-21-00-02-1') & (df_main_assys['DATA_TYPE_ID'] == 1)
filt_nlg_fc = (df_main_assys['CONFIG_POS_SDESC'] == '32-21-00-02-1') & (df_main_assys['DATA_TYPE_ID'] == 10)

nlg_sn = df_main_assys[filt_nlg_fh]['SERIAL_NO_OEM'].values[0]
nlg_tsn = int(df_main_assys[filt_nlg_fh]['TSN_QT'].values[0])
nlg_csn = int(df_main_assys[filt_nlg_fc]['TSN_QT'].values[0])
nlg_tso = int(df_main_assys[filt_nlg_fh]['TSO_QT'].values[0])
nlg_cso = int(df_main_assys[filt_nlg_fc]['TSO_QT'].values[0])

# MLG L/H
filt_mlg_lh_fh = (df_main_assys['CONFIG_POS_SDESC'] == '32-11-01-01A (LH)') & (df_main_assys['DATA_TYPE_ID'] == 1)
filt_mlg_lh_fc = (df_main_assys['CONFIG_POS_SDESC'] == '32-11-01-01B (RH)') & (df_main_assys['DATA_TYPE_ID'] == 10)

mlg_lh_sn = df_main_assys[filt_mlg_lh_fh]['SERIAL_NO_OEM'].values[0]
mlg_lh_tsn = int(df_main_assys[filt_mlg_lh_fh]['TSN_QT'].values[0])
mlg_lh_csn = int(df_main_assys[filt_mlg_lh_fc]['TSN_QT'].values[0])
mlg_lh_tso = int(df_main_assys[filt_mlg_lh_fh]['TSO_QT'].values[0])
mlg_lh_cso = int(df_main_assys[filt_mlg_lh_fc]['TSO_QT'].values[0])

# MLG R/H
filt_mlg_rh_fh = (df_main_assys['CONFIG_POS_SDESC'] == '32-11-01-01B (RH)') & (df_main_assys['DATA_TYPE_ID'] == 1)
filt_mlg_rh_fc = (df_main_assys['CONFIG_POS_SDESC'] == '32-11-01-01B (RH)') & (df_main_assys['DATA_TYPE_ID'] == 10)

mlg_rh_sn = df_main_assys[filt_mlg_rh_fh]['SERIAL_NO_OEM'].values[0]
mlg_rh_tsn = int(df_main_assys[filt_mlg_rh_fh]['TSN_QT'].values[0])
mlg_rh_csn = int(df_main_assys[filt_mlg_rh_fc]['TSN_QT'].values[0])
mlg_rh_tso = int(df_main_assys[filt_mlg_rh_fh]['TSO_QT'].values[0])
mlg_rh_cso = int(df_main_assys[filt_mlg_rh_fc]['TSO_QT'].values[0])

# Getting avionics components information
filt_gp = df_avionics['BOM_PART_CD'] == '22-11-01-01'
filt_hf = df_avionics['BOM_PART_CD'] == '23-11-01-01'
filt_dvdr = df_avionics['BOM_PART_CD'] == '31-31-01-01'
filt_qar = df_avionics['BOM_PART_CD'] == '31-32-01-01'
filt_dme = df_avionics['BOM_PART_CD'] == '34-51-01-01'
filt_printer = df_avionics['BOM_PART_CD'] == '23-24-01-01'
filt_lrra = df_avionics['BOM_PART_CD'] == '34-31-01-01'
filt_wxr = df_avionics['BOM_PART_CD'] == '34-42-01-02'
filt_gpws = df_avionics['BOM_PART_CD'] == '34-41-01-01'
filt_tcas = df_avionics['BOM_PART_CD'] == '34-43-01-01'
filt_gps = df_avionics['BOM_PART_CD'] == '34-56-01-01'
filt_xpnder = df_avionics['BOM_PART_CD'] == '34-52-01-01A'
filt_iru = df_avionics['BOM_PART_CD'] == '34-26-01-01'
filt_vor = df_avionics['BOM_PART_CD'] == '34-32-01-01'
filt_vhfn = df_avionics['BOM_PART_CD'] == '34-32-01-01'
filt_du = df_avionics['BOM_PART_CD'] == '31-61-01-01'
filt_isfd = df_avionics['BOM_PART_CD'] == '34-11-01-01'

gp = df_avionics[filt_gp]['MANUFACT_NAME'].values[0] + ' P/N: ' + df_avionics[filt_gp]['PART_NO_OEM'].values[0]
hf = df_avionics[filt_hf]['MANUFACT_NAME'].values[0] + ' P/N: ' + df_avionics[filt_hf]['PART_NO_OEM'].values[0]
vhf12 = '7026201-801'
vhf3 = '7026201-804'
dvdr = df_avionics[filt_dvdr]['MANUFACT_NAME'].values[0] + ' P/N: ' + df_avionics[filt_dvdr]['PART_NO_OEM'].values[0]
qar = df_avionics[filt_qar]['MANUFACT_NAME'].values[0] + ' P/N: ' + df_avionics[filt_qar]['PART_NO_OEM'].values[0]
dme = df_avionics[filt_dme]['MANUFACT_NAME'].values[0] + ' P/N: ' + df_avionics[filt_dme]['PART_NO_OEM'].values[0]
printer = df_avionics[filt_printer]['MANUFACT_NAME'].values[0] + ' P/N: ' + df_avionics[filt_printer]['PART_NO_OEM'].values[0]
lrra = df_avionics[filt_lrra]['MANUFACT_NAME'].values[0] + ' P/N: ' + df_avionics[filt_lrra]['PART_NO_OEM'].values[0]
wxr = df_avionics[filt_wxr]['MANUFACT_NAME'].values[0] + ' P/N: ' + df_avionics[filt_wxr]['PART_NO_OEM'].values[0]
gpws = df_avionics[filt_gpws]['MANUFACT_NAME'].values[0] + ' P/N: ' + df_avionics[filt_gpws]['PART_NO_OEM'].values[0]
tcas = df_avionics[filt_tcas]['MANUFACT_NAME'].values[0] + ' P/N: ' + df_avionics[filt_tcas]['PART_NO_OEM'].values[0]
gps = df_avionics[filt_gps]['MANUFACT_NAME'].values[0] + ' P/N: ' + df_avionics[filt_gps]['PART_NO_OEM'].values[0]
xpnder = df_avionics[filt_xpnder]['MANUFACT_NAME'].values[0] + ' P/N: ' + df_avionics[filt_xpnder]['PART_NO_OEM'].values[0]
iru = df_avionics[filt_iru]['MANUFACT_NAME'].values[0] + ' P/N: ' + df_avionics[filt_iru]['PART_NO_OEM'].values[0]
vor = df_avionics[filt_vor]['MANUFACT_NAME'].values[0] + ' P/N: ' + df_avionics[filt_vor]['PART_NO_OEM'].values[0]
vhfn = df_avionics[filt_vhfn]['MANUFACT_NAME'].values[0] + ' P/N: ' + df_avionics[filt_vhfn]['PART_NO_OEM'].values[0]
du = df_avionics[filt_du]['MANUFACT_NAME'].values[0] + ' P/N: ' + df_avionics[filt_du]['PART_NO_OEM'].values[0]
isfd = df_avionics[filt_isfd]['MANUFACT_NAME'].values[0] + ' P/N: ' + df_avionics[filt_isfd]['PART_NO_OEM'].values[0]

if aircraft == 'HP-1568CMP' or aircraft == 'HP-1569CMP':
    elt_port = 1
else:
    elt_port = 0

# Setting up output excel file
user_path = os.environ['USERPROFILE']
filename = f'Spec Sheet {aircraft} MSN {msn} ({today}).xlsx'
location = os.path.join(user_path, 'Documents',f'{filename}')

if os.path.isfile(location):# Delete if filename already exists
  os.remove(location)
else:
    pass

workbook = xlsxwriter.Workbook(location)
worksheet = workbook.add_worksheet()
worksheet.set_margins(0.25,0.25,0.4,0.4)
footer = 'Aircraft specifications are intended to be preliminary information only and must be verified prior to sale.'
worksheet.set_footer(footer)

# Setting cells formatting
cell_format1 = workbook.add_format({'bold': True, 'underline':True}) #Title format 
cell_format2 = workbook.add_format()
cell_format2.set_num_format('#,##0') #Thousand separator for numbers
cell_format2.set_align('left') #Left align
cell_format3 = workbook.add_format({'align':'right'})
cell_format4 = workbook.add_format({'bold': True,'align':'right'})
cell_format5 = workbook.add_format({'bold': True,'align':'center'})
cell_format6 = workbook.add_format()
cell_format6.set_bg_color('#D3D3D3') #Section dividers background color
cell_format6.set_border()
cell_format7 = workbook.add_format() #Text wrapping
cell_format7.set_text_wrap()
cell_format8 = workbook.add_format({'align':'left'})

cell_format1.set_left()
cell_formata4 = workbook.add_format()
cell_formata4.set_left()
cell_borders = workbook.add_format()
cell_borders.set_left()

cell_top = workbook.add_format()
cell_top.set_top()

cell_right = workbook.add_format()
cell_right.set_right()

cell_bottom = workbook.add_format()
cell_bottom.set_bottom()

cell_g98 = workbook.add_format()
cell_g98.set_bottom()
cell_g98.set_right()



cell_format6.set_left()
worksheet.write('B4', '', cell_top)
worksheet.write('C3', '', cell_bottom)
worksheet.write('A3', '', cell_bottom)
worksheet.write('D4', '', cell_top)
worksheet.write('E4', '', cell_top)
worksheet.write('F4', '', cell_top)
worksheet.write('G3', '', cell_bottom)
worksheet.write('A53', '', cell_bottom)


worksheet.write('C4', '', cell_borders)
worksheet.write('C5', '', cell_borders)
worksheet.write('C6', '', cell_borders)
worksheet.write('C7', '', cell_borders)
worksheet.write('C8', '', cell_borders)
worksheet.write('C9', '', cell_borders)
worksheet.write('C10', '', cell_borders)
worksheet.write('C11', '', cell_borders)
worksheet.write('C12', '', cell_borders)

n=4
while n <=48:
 worksheet.write(f'G{n}', '', cell_right)
 n = n + 1

for n in ['A','B','C','D','E','F','G']:
    worksheet.write(f'{n}49','',cell_top)

for n in ['B','C','D','E','F','G']:
    worksheet.write(f'{n}54','',cell_top)

worksheet.write('G53','',cell_bottom)

worksheet.write('A99','',cell_top)



n=54
while n <=98:
 worksheet.write(f'G{n}', '', cell_right)
 n = n + 1

for n in ['A','B','C','D','E','F','G']:
    worksheet.write(f'{n}98','',cell_bottom)

n=81
while n <=98:
 worksheet.write(f'A{n}', '', cell_borders)
 n = n + 1

worksheet.write('G98','',cell_g98)

for n in ['A','B','C','D','E','F','G']:
    worksheet.write(f'{n}103','',cell_bottom)
    
n=104
while n <=148:
 worksheet.write(f'G{n}', '', cell_right)
 n = n + 1

for n in ['A','B','C','D','E','F','G']:
    worksheet.write(f'{n}149','',cell_top)

n=133
while n <=148:
 worksheet.write(f'A{n}', '', cell_borders)
 n = n + 1

cell_format_header = workbook.add_format({'bold': True}) #Header format
worksheet.set_column('A:A', 25.5)
worksheet.set_column('B:B', 15)
worksheet.set_column('C:C', 21)
worksheet.set_column('G:G', 6.9)


# Headers Page 1
worksheet.write('A1', 'OPERATOR: COPA AIRLINES',cell_format_header)
worksheet.merge_range('B1:E1', 'AIRCRAFT SPEC SHEET', cell_format5)
worksheet.merge_range('B2:E2', f'MSN: {msn}', cell_format5)
worksheet.merge_range('F1:G1', f'DATE: {today}', cell_format4)
worksheet.write('G2', '1 of 3',cell_format4)

# Aircraft Data
worksheet.write('A4', 'AIRFRAME',cell_format1)
worksheet.write('A5', 'Model......................................................................................',cell_formata4)
worksheet.write('A6','Registry......................................................................................',cell_borders)
worksheet.write('A7','Serial Number......................................................................................',cell_borders)
worksheet.write('A8', 'Manufacturing Date......................................................................................',cell_borders)
worksheet.write('A9', f'Flight Hours ({today})......................................................................................',cell_borders)
worksheet.write('A10', f'Flight Cycles ({today})......................................................................................',cell_borders)
worksheet.write('B5', 'ERJ-190-100 IGW')
worksheet.write('B6', ac_rg)
worksheet.write('B7', msn)
worksheet.write('B8', man_date)
worksheet.write('B9', ac_tsn_fh, cell_format2)
worksheet.write('B10', ac_tsn_fc, cell_format2)
worksheet.write('A11', 'Maximum Taxi Weight......................................................................................',cell_borders)
worksheet.write('A12', 'Maximum Takeoff Weight......................................................................................',cell_borders)
worksheet.write('A13', 'Maximum Landing Weight......................................................................................',cell_borders)
worksheet.write('A14', 'Maximum Zero-Fuel Weight......................................................................................',cell_borders)
worksheet.write('A15', 'Noise Category......................................................................................',cell_borders)
worksheet.write('A16', 'Landing Category Approval......................................................................................',cell_borders)
worksheet.write('A17', '', cell_borders)
worksheet.write('A18', '', cell_borders)
worksheet.write('A19', '', cell_borders)
worksheet.write('B11', mtw)
worksheet.write('B12', mtow)
worksheet.write('B13', mlw)
worksheet.write('B14', mzfw)
worksheet.write('B15', noise_cat)
worksheet.write('B16', cat_status)

# Maintenance Program
worksheet.write('C13', 'MAINTENANCE PROGRAM', cell_format1)
worksheet.write('C14', 'Last C-Check On:', cell_borders)
worksheet.write('C15', 'Next C-Check Due:', cell_borders)
worksheet.write('C16', 'C-Checks every 3 years. First C-CK at 5 years', cell_borders)
worksheet.write('C17', 'A-Checks every 180 days', cell_borders)
worksheet.write('C18', 'Engines under MCPH and Trend Monitoring', cell_borders)
worksheet.write('C19', 'APU under PBH', cell_borders)
try:
    worksheet.write('D14', ac_lastcck)
except NameError:
    worksheet.write('D14', (datetime.strptime(ac_nextcck,'%d-%b-%y') - timedelta(days=1095)).strftime('%d-%b-%y'))

worksheet.write('D15', ac_nextcck)

# Setting aircraft photo
image_path = r'C:/Users/ggalina/SpecSheetMakerEMB/Aircraft Photos/'+f'{aircraft}.jpg'
worksheet.insert_image('C4', image_path,{'object_position': 3, 'y_offset': 1})

worksheet.merge_range('A20:G20', '', cell_format6) #Section divisor

# L/H Engine Data
worksheet.write('A21', 'ENGINE L/H', cell_format1)
worksheet.write('A22', 'Type......................................................................................',cell_borders)
worksheet.write('A23', 'Thrust Rating (lb)......................................................................................',cell_borders)
worksheet.write('A24', 'Serial Number......................................................................................',cell_borders)
worksheet.write('A25', f'TSN ({today})......................................................................................',cell_borders)
worksheet.write('A26', f'CSN ({today})......................................................................................',cell_borders)
worksheet.write('A27', f'TSLSV ({today})......................................................................................',cell_borders)
worksheet.write('A28', f'CSLSV ({today})......................................................................................',cell_borders)
worksheet.write('B22', 'CF34-10E5')
worksheet.write('B23', '17,390')
worksheet.write('B24', eng_lh_sn,cell_format8)
worksheet.write('B25', eng_lh_tsn,cell_format2)
worksheet.write('B26', eng_lh_csn, cell_format2)
worksheet.write('B27', eng_lh_tslv, cell_format2)
worksheet.write('B28', eng_lh_cslv, cell_format2)

#R/H Engine Data
worksheet.write('C21', 'ENGINE R/H', cell_format1)
worksheet.write('C22', 'Type......................................................................................', cell_borders)
worksheet.write('C23', 'Thrust Rating (lb)......................................................................................', cell_borders)
worksheet.write('C24', 'Serial Number......................................................................................', cell_borders)
worksheet.write('C25', f'TSN ({today})......................................................................................', cell_borders)
worksheet.write('C26', f'CSN ({today})......................................................................................', cell_borders)
worksheet.write('C27', f'TSLSV ({today})......................................................................................', cell_borders)
worksheet.write('C28', f'CSLSV ({today})......................................................................................', cell_borders)
worksheet.write('D22', 'CF34-10E5')
worksheet.write('D23', '17,390')
worksheet.write('D24', eng_rh_sn,cell_format8)
worksheet.write('D25', eng_rh_tsn,cell_format2)
worksheet.write('D26', eng_rh_csn, cell_format2)
worksheet.write('D27', eng_rh_tslv, cell_format2)
worksheet.write('D28', eng_rh_cslv, cell_format2)

worksheet.merge_range('A29:G29', '', cell_format6) #Section divisor

#APU Data
worksheet.write('A30', 'APU', cell_format1)
worksheet.write('A31', 'Type......................................................................................',cell_borders)
worksheet.write('A32', 'Serial Number......................................................................................',cell_borders)
worksheet.write('A33', f'TSN ({today})......................................................................................',cell_borders)
worksheet.write('A34', f'CSN ({today})......................................................................................',cell_borders)
worksheet.write('A35', f'TSLSV ({today})......................................................................................',cell_borders)
worksheet.write('A36', '',cell_borders)
worksheet.write('A37', '',cell_borders)
worksheet.write('A38', '',cell_borders)

worksheet.write('B31', 'Collins APS2300')
worksheet.write('B32', apu_sn)
worksheet.write('B33', apu_aot, cell_format2)
worksheet.write('B34', apu_acyc, cell_format2)
worksheet.write('B35', apu_tslv, cell_format2)


#NLG Data

worksheet.write('C30', 'NOSE LANDING GEAR', cell_format1)
worksheet.write('C31', 'Part Number (Liebherr)......................................................................................', cell_borders)
worksheet.write('C32', 'Serial Number......................................................................................', cell_borders)
worksheet.write('C33', f'TSN ({today})......................................................................................', cell_borders)
worksheet.write('C34', f'CSN ({today})......................................................................................', cell_borders)
worksheet.write('C35', f'TSO ({today})......................................................................................', cell_borders)
worksheet.write('C36', f'CSO ({today})......................................................................................', cell_borders)
worksheet.write('C37', 'CBO......................................................................................', cell_borders)
worksheet.write('C38', 'Cycles to Next Overhaul......................................................................................', cell_borders)

worksheet.write('D31', '190-70450-403')
worksheet.write('D32', nlg_sn)
worksheet.write('D33', nlg_tsn, cell_format2)
worksheet.write('D34', nlg_csn, cell_format2)
worksheet.write('D35', nlg_tso, cell_format2)
worksheet.write('D36', nlg_cso, cell_format2)
worksheet.write('D37', '25,000', cell_format2)
worksheet.write('D38', 25000 - nlg_cso, cell_format2)

worksheet.merge_range('A39:G39', '', cell_format6) #Section divisor

# MLG L/H Data
worksheet.write('A40', 'MAIN LANDING GEAR L/H', cell_format1)
worksheet.write('A41', 'Part Number (Goodrich)......................................................................................',cell_borders)
worksheet.write('A42', 'Serial Number......................................................................................',cell_borders)
worksheet.write('A43', f'TSN ({today})......................................................................................',cell_borders)
worksheet.write('A44', f'CSN ({today})......................................................................................',cell_borders)
worksheet.write('A45', f'TSO ({today})......................................................................................',cell_borders)
worksheet.write('A46', f'CSO ({today})......................................................................................',cell_borders)
worksheet.write('A47', 'CBO......................................................................................',cell_borders)
worksheet.write('A48', 'Cycles to Next Overhaul......................................................................................',cell_borders)

worksheet.write('B41', '190-70024-405')
worksheet.write('B42', mlg_lh_sn)
worksheet.write('B43', mlg_lh_tsn, cell_format2)
worksheet.write('B44', mlg_lh_csn, cell_format2)
worksheet.write('B45', mlg_lh_tso, cell_format2)
worksheet.write('B46', mlg_lh_cso, cell_format2)
worksheet.write('B47', '25,000', cell_format2)
worksheet.write('B48', 25000-mlg_lh_cso, cell_format2)

# MLG R/H Data
worksheet.write('C40', 'MAIN LANDING GEAR R/H', cell_format1)
worksheet.write('C41', 'Part Number (Goodrich)......................................................................................', cell_borders)
worksheet.write('C42', 'Serial Number......................................................................................', cell_borders)
worksheet.write('C43', f'TSN ({today})......................................................................................', cell_borders)
worksheet.write('C44', f'CSN ({today})......................................................................................', cell_borders)
worksheet.write('C45', f'TSO ({today})......................................................................................', cell_borders)
worksheet.write('C46', f'CSO ({today})......................................................................................', cell_borders)
worksheet.write('C47', 'CBO......................................................................................', cell_borders)
worksheet.write('C48', 'Cycles to Next Overhaul......................................................................................', cell_borders)

worksheet.write('D41', '190-70024-406')
worksheet.write('D42', mlg_rh_sn)
worksheet.write('D43', mlg_rh_tsn, cell_format2)
worksheet.write('D44', mlg_rh_csn, cell_format2)
worksheet.write('D45', mlg_rh_tso, cell_format2)
worksheet.write('D46', mlg_rh_cso, cell_format2)
worksheet.write('D47', '25,000', cell_format2)
worksheet.write('D48', 25000-mlg_rh_cso, cell_format2)

# Headers Page 2
worksheet.write('A51', f'OPERATOR: COPA AIRLINES',cell_format_header)
worksheet.merge_range('B51:E51', 'AIRCRAFT SPEC SHEET', cell_format5)
worksheet.merge_range('B52:E52', f'MSN: {msn}', cell_format5)
worksheet.merge_range('F51:G51', f'DATE: {today}', cell_format4)
worksheet.write('G52', '2 of 3',cell_format4)

#Avionics Data
worksheet.write('A54', 'AVIONICS',cell_format1)
worksheet.write('A55', 'Guidance Panel GP-750........................................................',cell_borders)
worksheet.write('A56', 'HF Comm Transceiver .....................................................................',cell_borders)
worksheet.write('A57', 'HFDL enabled ..........................................................................',cell_borders)
worksheet.write('A58', 'VHF Digital Radios (Qty = 2) ......................................................................',cell_borders)
worksheet.write('A59', 'VHF Digital Radio Mod A for Data Mode...............................................................',cell_borders)
worksheet.write('A60', 'Battery ..............................................................................................',cell_borders)
worksheet.write('A61', 'Cockpit PC Power System......................................................................',cell_borders)
worksheet.write('A62', 'DVDR (Digital Voice and Data Recorder) ......................................................................',cell_borders)
worksheet.write('A63', 'QAR (Quick Access Recorder) ......................................................................................',cell_borders)
worksheet.write('A64', 'Epic Load Software Version .............................................................................',cell_borders)
worksheet.write('A65', 'APM Options Software ...............................................................................',cell_borders)
worksheet.write('A66', 'DME (Distance Measuring Equipment Module) ......................................................................................',cell_borders)
worksheet.write('A67', 'Printer TP4840 (Control Pedestal) ...............................................',cell_borders)
worksheet.write('A68', 'LRRA (Low Range Radio Altimeter) ...............................................',cell_borders)
worksheet.write('A69', 'Weather Radar Receiver/Transmitter ...............................................',cell_borders)
worksheet.write('A70', 'Enhanced Ground Proximity Warning Module ...............................................',cell_borders)
worksheet.write('A71', 'TCAS Computer................................................................................',cell_borders)
worksheet.write('A72', 'TCAS Software ...............................................................................',cell_borders)
worksheet.write('A73', 'GPS Module ...................................................................................',cell_borders)
worksheet.write('A74', 'ATC Transponder Module ................................................................',cell_borders)
worksheet.write('A75', 'Micro Inertial Reference Unit ...........................................................',cell_borders)
worksheet.write('A76', 'VOR/Marker Beacon Receiver (VIDL Module) ...............................................',cell_borders)
worksheet.write('A77', 'ADF Module ......................................................................................',cell_borders)
worksheet.write('A78', 'Display Unit .....................................................................................',cell_borders)
worksheet.write('A79', 'Integrated Electronic Standby Unit ......................................................................',cell_borders)
worksheet.write('A80', 'In-Flight Entertainment (IFE)......................................................................',cell_borders)

worksheet.write('C55', gp)
worksheet.write('C56', hf)
worksheet.write('C57', 'NO')
worksheet.write('C58', vhf12)
worksheet.write('C59', vhf3)
worksheet.write('C60', '2 SAFT Battery PN 5912855-01')
worksheet.write('C61', 'NO')
worksheet.write('C62', dvdr)
worksheet.write('C63', qar)
worksheet.write('C64', '25.5')
worksheet.write('C65', 'P/N DM60001133-00592')
worksheet.write('C66', dme)
worksheet.write('C67', printer)
worksheet.write('C68', lrra)
worksheet.write('C69', wxr)
worksheet.write('C70', gpws)
worksheet.write('C71', tcas)
worksheet.write('C72', '7.1')
worksheet.write('C73', gps)
worksheet.write('C74', xpnder)
worksheet.write('C75', iru)
worksheet.write('C76', vor)
worksheet.write('C77', 'Not Installed')
worksheet.write('C78', du)
worksheet.write('C79', isfd)
worksheet.write('C80', 'NO')


# Headers Page 3
worksheet.write('A101', f'OPERATOR: COPA AIRLINES',cell_format_header)
worksheet.merge_range('B101:E101', 'AIRCRAFT SPEC SHEET', cell_format5)
worksheet.merge_range('B102:E102', f'MSN: {msn}', cell_format5)
worksheet.merge_range('F101:G101', f'DATE: {today}', cell_format4)
worksheet.write('G102', '3 of 3',cell_format4)

# Interiors Data
worksheet.write('A104', 'INTERIORS',cell_format1)
worksheet.write('A105', 'Passengers BC / TC...................................',cell_borders)
worksheet.write('A106', 'Seats Manuf. BC / TC.................................................',cell_borders)
worksheet.write('A107', 'Seats pitch BC / TC................................................',cell_borders)
worksheet.write('A108', 'Seats recline BC / TC................................................',cell_borders)
worksheet.write('A109', 'Galley G1..................................................',cell_borders)
worksheet.write('A110', 'Galley G2...................................................',cell_borders)
worksheet.write('A111', 'Galley G3...................................................',cell_borders)
worksheet.write('A112', 'Lavatory Configuration.................................................',cell_borders)
worksheet.write('A113', 'Lavatory Manufacturer.................................................',cell_borders)
worksheet.write('A114', 'Ovens G2.........................................................',cell_borders)
worksheet.write('A115', 'Ovens G3...........................................................',cell_borders)
worksheet.write('A116', 'Observer Seats...............................................',cell_borders)
worksheet.write('A117', 'Escape Slides.........................................................',cell_borders)
worksheet.write('A118', 'ELT Fixed / Portable................................................',cell_borders)

worksheet.write('B105', '10/84')
worksheet.write('B106', '4FLIGHT INDUSTRIES/4FLIGHT INDUSTRIES')
worksheet.write('B107', '38/31')
worksheet.write('B108', '6/4')
worksheet.write('B109', 'C&D ZODIAC PN 190-45086-401')
worksheet.write('B110', 'C&D ZODIAC PN 190-42301-401')
worksheet.write('B111', 'C&D ZODIAC PN 190-59701-401')
worksheet.write('B112', '1 FWD & 1 AFT')
worksheet.write('B113', 'C&D Zodiac')
worksheet.write('B114', '1 x Sell PN 8201-11-0000-12')
worksheet.write('B115', '2 x Sell PN 8203-11-0000-12')
worksheet.write('B116', '1')
worksheet.write('B117', 'Goodrich PN 104003-2 & 104005-1')
worksheet.write('B118', f'1/{elt_port}')

worksheet.merge_range('A119:G119', '', cell_format6) #Section divisor


#Systems Data
worksheet.write('A120', 'SYSTEMS',cell_format1)
worksheet.write('A121', 'Auxiliary Fuel Tanks.................................................................',cell_borders)
worksheet.write('A122', 'Main Brakes Type...................................',cell_borders)
worksheet.write('A123', 'Main Brakes Manuf. & P/N.................................................',cell_borders)
worksheet.write('A124', 'Main Wheels Manuf. & PN................................................',cell_borders)
worksheet.write('A125', 'Switch-Dispatch w/ LG Down................................................',cell_borders)
worksheet.write('A126', '22-min Chemical O2 Generators..................................................',cell_borders)
worksheet.write('A127', 'First Obs. Full-Face Mask..................................................',cell_borders)
worksheet.write('A128', 'Potable Water 40-gallon Capacity..................................................',cell_borders)
worksheet.write('A129', 'Nitrogen Generation System..................................................',cell_borders)


worksheet.write('B121', 'NO')
worksheet.write('B122', 'Carbon')
worksheet.write('B123', 'Meggitt P/N 90002340PR/-1PR/-2PR/-4PR/-2')
worksheet.write('B124', 'Meggitt P/N 90002317WT/-1WT/-2WT/WTA/-1WTA-2WTA')
worksheet.write('B125', 'NO')
worksheet.write('B126', 'NO')
worksheet.write('B127', 'YES')
worksheet.write('B128', 'NO')

worksheet.merge_range('A129:G129', '', cell_format6) #Section divisor

#Structures Data
worksheet.write('A130', 'STRUCTURES',cell_format1)
worksheet.write('A131', 'Enhanced Security Cockpit Door.................................................................',cell_borders)
worksheet.write('A132', 'Winglets...................................',cell_borders)

worksheet.write('B131', 'YES')
worksheet.write('B132', 'YES')

# Closing workbook
workbook.close()

# Opening file
os.system(f'"{location}"')