import gspread
import pandas as pd
import numpy as np
from datetime import date

cred_file = 'equipment-dashboard-a2d53f750591.json'
gc = gspread.service_account(cred_file)

database = gc.open('Database')

borneo = database.worksheet('Bulk Borneo')
celebes = database.worksheet('Bulk Celebes')
sumatra = database.worksheet('Bulk Sumatra')
java = database.worksheet('Bulk Java')
dewata = database.worksheet('Bulk Dewata')
karimun = database.worksheet('Bulk Karimun')
# of1 = database.worksheet('Ocean Flow 1')
natuna = database.worksheet('Bulk Natuna')
sumba = database.worksheet('Bulk Sumba')
derawan = database.worksheet('Bulk Derawan')

##-----Function
#-- 1. Pre-Processing
def preprocessing(df):
    
    df = df.replace(r'^\s*$', np.nan, regex=True)
    df[['Last Overhaul GEN01', 'Last Overhaul GEN02', 'Last Overhaul GEN03', 'Last Overhaul GEN04']] = \
    df[['Last Overhaul GEN01', 'Last Overhaul GEN02', 'Last Overhaul GEN03', 'Last Overhaul GEN04']].fillna(date.today())
    df['Last Overhaul GEN01'] = pd.to_datetime(df['Last Overhaul GEN01'], dayfirst = True)
    df['Last Overhaul GEN02'] = pd.to_datetime(df['Last Overhaul GEN02'], dayfirst = True)
    df['Last Overhaul GEN03'] = pd.to_datetime(df['Last Overhaul GEN03'], dayfirst = True)
    df['Last Overhaul GEN04'] = pd.to_datetime(df['Last Overhaul GEN04'], dayfirst = True)


    return(df)

#-- 1. Bulk Borneo
df_borneo = pd.DataFrame(borneo.get_all_records())
df_borneo = df_borneo.replace(r'^\s*$', np.nan, regex=True)
df_borneo = preprocessing(df_borneo)

#-- 2. Bulk Celebes
df_celebes = pd.DataFrame(celebes.get_all_records())
df_celebes = df_celebes.replace(r'^\s*$', np.nan, regex=True)
df_celebes = preprocessing(df_celebes)

#-- 3. Bulk Sumatra
df_sumatra = pd.DataFrame(sumatra.get_all_records())
df_sumatra = df_sumatra.replace(r'^\s*$', np.nan, regex=True)
df_sumatra = preprocessing(df_sumatra)

#-- 4. Bulk Java
df_java = pd.DataFrame(java.get_all_records())
df_java = df_java.replace(r'^\s*$', np.nan, regex=True)
df_java = preprocessing(df_java)

#-- 5. Bulk Dewata
df_dewata = pd.DataFrame(dewata.get_all_records())
df_dewata = df_dewata.replace(r'^\s*$', np.nan, regex=True)
df_dewata = preprocessing(df_dewata)

#-- 6. Bulk Karimun
df_karimun = pd.DataFrame(karimun.get_all_records())
df_karimun = df_karimun.replace(r'^\s*$', np.nan, regex=True)
df_karimun = preprocessing(df_karimun)

#-- 7. Ocean Flow 1
# df_of1 = pd.DataFrame(of1.get_all_records())
# df_of1 = df_of1.replace(r'^\s*$', np.nan, regex=True)
# df_of1 = preprocessing(df_of1)

#-- 8. Bulk Natuna
df_natuna = pd.DataFrame(natuna.get_all_records())
df_natuna = df_natuna.replace(r'^\s*$', np.nan, regex=True)
df_natuna = preprocessing(df_natuna)

#-- 9. Bulk Sumba
df_sumba = pd.DataFrame(sumba.get_all_records())
df_sumba = df_sumba.replace(r'^\s*$', np.nan, regex=True)
df_sumba = preprocessing(df_sumba)

#-- 10. Bulk Derawan
df_derawan = pd.DataFrame(derawan.get_all_records())
df_derawan = df_derawan.replace(r'^\s*$', np.nan, regex=True)
df_derawan = preprocessing(df_derawan)