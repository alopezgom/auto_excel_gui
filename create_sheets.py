#!/usr/local/bin python3

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import xlwings as xw
from sklearn import linear_model
from sklearn.model_selection import train_test_split

def read_sheet(loc = "/Users/alexlopez/Documents/Documentos-varios/automation_tests/",
               namef = "ejercicio.xlsx", sheet = "Base", nsheet="Datos"):
  df_base = pd.read_excel(loc+namef, sheet_name=sheet, engine='openpyxl')
  ncol = list(df_base.iloc[2])
  df_base.columns = ncol
  df_base.iloc[3:,9] = df_base.iloc[3:,9]*-1
  df_new = df_base.iloc[3:,9:]
  df_new = df_new.filter(["Puntos redimidos", "Partner", "Vertical", "Dia"], axis=1)
  df_new.reset_index(drop=True, inplace=True)
  return df_new, loc+namef, sheet, nsheet

def create_data(df):
  data = pd.pivot_table(df,
                        index = ["Vertical", "Partner"],
                        columns = "Dia",
                        values = "Puntos redimidos",
                        aggfunc=np.sum)
  data.reset_index(inplace=True)
  data.columns.name = None
  data = data.append(data[data.columns[2:].tolist()].sum(),
                     ignore_index=True).fillna(0)
  return data

def base_mil(df):
  ratio1 = []
  ratio2 = []
  for col in df[df.columns[2:].tolist()]:
    ratio1.append((df[col][1]+df[col][3])/(df[col][1]+df[col][3]+df[col][12]))
    ratio2.append(df[col][12]/(df[col][1]+df[col][3]+df[col][12]))
  ratio1 = np.round(ratio1,2)
  ratio2 = np.round(ratio2,2)
  
  for col, r1, r2 in zip(df[df.columns[2:].tolist()], ratio1, ratio2):
    df[col][1] = (df[col][1]+df[col][3])+(df[col][7]*r1)
    df[col][12] = df[col][12]+(df[col][7]*r2)
  df["Partner"][1] = "Aeromexico & OALs"
  
  for col, items in df[df.columns[2:].tolist()].items():
    for i in range(len(items)):
      df[col][i] = np.round(items[i]/1000.0,2)
  
  t1 = np.array(df.iloc[-1, 3:].tolist())
  t2 = np.array(df.iloc[-1, 2:-1].tolist())
  rate_total = np.round(np.subtract(t1,t2)/t2,2)
     
  t1 = np.array(df.iloc[12, 3:].tolist())
  t2 = np.array(df.iloc[12, 2:-1].tolist())
  rate_online = np.round(np.subtract(t1,t2)/t2,2)
  
  t1 = np.array(df.iloc[1, 3:].tolist())
  t2 = np.array(df.iloc[1, 2:-1].tolist())
  rate_aereo = np.round(np.subtract(t1,t2)/t2,2)
  
  np.insert(rate_total, 0, 0)
  np.insert(rate_online, 0, 0)
  np.insert(rate_aereo, 0, 0)
  
  ls_total = ["Rate total"]+list(rate_total)
  ls_online = ["Rate online"]+list(rate_online)
  ls_aereo = ["Rate aereo"]+list(rate_aereo)
  
  rdf = pd.DataFrame([ls_total, ls_online, ls_aereo],
                     columns=[["Rates"]+list(range(2,len(df.columns[2:].tolist())+1))],
                     dtype=float)
  return df, rdf
