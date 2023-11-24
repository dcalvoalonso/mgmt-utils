import glob
import numpy as np
import pandas as pd
import datetime

PRAS_FOLDER = '/mnt/c/Users/A620381/OneDrive - Atos/docs/ai_data_robotics/1. HoU/8. PRAs/projects'
PRAS_extension = '.xlsx'
PRA_SHEET_CONTROL = 'Control'
PRA_SHEET_COSTS = 'GANTT-Costs'
PRA_SHEET_COSTS_START_ROW = 3
COLUMN_NAME_WP = 'WP/task name'
COLUMN_NAME_WP_short = 'WP/task'
COLUMN_MANAGEMENT_TASKS = 'ENTER MANAGEMENT / COORDINATION TASKS (DELETE THE ROWS THAT DON\'T PROCEED)'
COLUMN_TECHNICAL_TASKS = 'ENTER TECHNICAL TASKS'
COLUMN_IMPACT_TASKS = 'ENTER IMPACT TASKS'
COLUMN_UNIT = 'Market/Unit'
COLUMN_GCM = 'GCM'
COLUMN_TOTAL = 'TOTAL:'
COLUMN_2021 = 2021
COLUMN_2022 = 2022
COLUMN_2023 = 2023
COLUMN_2024 = 2024
COLUMN_2025 = 2025
COLUMN_2026 = 2026
COLUMN_2027 = 2027
AIDA_UNIT_VALUE = 'AIDAR'

OUTPUT_FILE = 'AIDA_accumulated_efforts.xlsx'
aida_total_efforts = []

def process_pra(pra_file):
    df_pra_control = pd.read_excel(io=pra_file, sheet_name=PRA_SHEET_CONTROL)
    project_name = df_pra_control.iloc[4][4]
    df_pra_costs = pd.read_excel(io=pra_file, sheet_name=PRA_SHEET_COSTS, header=PRA_SHEET_COSTS_START_ROW)
    row_admin_start_index = df_pra_costs.loc[df_pra_costs[COLUMN_NAME_WP] == COLUMN_MANAGEMENT_TASKS].index.astype(int)[0] + 1
    row_admin_end_index = df_pra_costs.loc[df_pra_costs[COLUMN_NAME_WP] == COLUMN_TECHNICAL_TASKS].index.astype(int)[0]
    row_technical_start_index = df_pra_costs.loc[df_pra_costs[COLUMN_NAME_WP] == COLUMN_TECHNICAL_TASKS].index.astype(int)[0] + 1
    row_technical_end_index = df_pra_costs.loc[df_pra_costs[COLUMN_NAME_WP] == COLUMN_IMPACT_TASKS].index.astype(int)[0]
    df_pra_costs_admin = df_pra_costs.iloc[row_admin_start_index : row_admin_end_index]
    df_pra_costs_tech = df_pra_costs.iloc[row_technical_start_index : row_technical_end_index]
    frames = [df_pra_costs_admin, df_pra_costs_tech]
    df_pra_costs = pd.concat(frames)
    df_pra_costs_aidar = df_pra_costs.loc[df_pra_costs[COLUMN_UNIT] == AIDA_UNIT_VALUE]
    
    df_pra_costs_aidar.fillna(0)
    aida_total_efforts_sum = df_pra_costs_aidar.sum(axis=0)
    df_aidar_total_efforts_project = pd.DataFrame({project_name: aida_total_efforts_sum})
    df_aidar_total_efforts_project = df_aidar_total_efforts_project.drop(COLUMN_GCM)

    df_aidar_total_efforts_project = df_aidar_total_efforts_project.T
    return df_aidar_total_efforts_project

for file in glob.glob(PRAS_FOLDER + '/*' + PRAS_extension, recursive=True):
   df_aidar_total_efforts_project = process_pra(file)
   df_aidar_total_efforts_project
   aida_total_efforts.append(df_aidar_total_efforts_project)

# df_aidar_total_efforts_project = process_pra(PRAS_FOLDER + '/PRA_TEMPLATE_20220324_v10_ERATOSTHENES_v04.xlsx')
# aida_total_efforts.append(df_aidar_total_efforts_project)

df_aida_total_efforts = pd.concat(aida_total_efforts)
df_aida_total_efforts = df_aida_total_efforts.drop(COLUMN_TOTAL, axis=1)
df_aida_total_efforts = df_aida_total_efforts.drop(COLUMN_2021, axis=1, errors='ignore')
df_aida_total_efforts = df_aida_total_efforts.drop(COLUMN_2022, axis=1, errors='ignore')
df_aida_total_efforts = df_aida_total_efforts.drop(COLUMN_2023, axis=1, errors='ignore')
df_aida_total_efforts = df_aida_total_efforts.drop(COLUMN_2024, axis=1, errors='ignore')
df_aida_total_efforts = df_aida_total_efforts.drop(COLUMN_2025, axis=1, errors='ignore')
df_aida_total_efforts = df_aida_total_efforts.drop(COLUMN_2026, axis=1, errors='ignore')
df_aida_total_efforts = df_aida_total_efforts.drop(COLUMN_2027, axis=1, errors='ignore')
df_aida_total_efforts = df_aida_total_efforts.drop(COLUMN_UNIT, axis=1, errors='ignore')
df_aida_total_efforts = df_aida_total_efforts.drop(COLUMN_NAME_WP, axis=1, errors='ignore')
df_aida_total_efforts = df_aida_total_efforts.drop(COLUMN_NAME_WP_short, axis=1, errors='ignore')
df_aida_months_cols =  df_aida_total_efforts.columns.tolist()

df_aida_months_cols.sort()
df_aida_total_efforts = df_aida_total_efforts[df_aida_months_cols]
df_aida_total_efforts = df_aida_total_efforts.fillna(0)

df_aida_total_efforts.to_excel(PRAS_FOLDER + '/' + OUTPUT_FILE)  
