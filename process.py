#!/usr/bin/env python2
# -*- coding: utf-8 -*-
"""
Created on Wed Dec 13 23:40:35 2017

@author: alecioc
"""

import pandas as pd

import numpy as np

from sklearn import linear_model

import datetime

import matplotlib
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D
matplotlib.style.use('ggplot')

from conf import plants

for plant in plants:
    
    plant_name = plant["plant_name"]
    
    print plant_name
    
    latitude = plant["latitude"]
    longitude = plant["longitude"]
    surface = plant["surface"]
    power_tot = plant["power_tot"]
    
    df_aa = pd.read_csv("azi_alt_" + plant_name + ".csv")
    df_aa["datetime"] = pd.to_datetime(df_aa.datetime)
    df_aa["azimuth_rad"] = df_aa.azimuth.apply(lambda x: np.deg2rad(x))
    df_aa["altitude_rad"] = df_aa.altitude.apply(lambda x: np.deg2rad(x))
    df_aa["ndy"] = df_aa.datetime.apply(lambda x: x.timetuple().tm_yday) - 1
    df_aa["ndy_"] = np.deg2rad(360.0 / 365.25 * (df_aa["ndy"] - 1))
    df_aa = df_aa.set_index("datetime", drop=False).sort_index()
    
    df_plant = pd.DataFrame()
    
    for year in range(2013, 2017):
    
        print year

        s_enel = pd.read_excel("./Data/Dati_enel.xlsx", sheetname=str(year)).set_index("Impianto").loc[plant_name][:12]
        s_enel.index = range(12)
        
        for month in range(1, 13):
            
#            print month
            
            if month != 12:
                filename = "./Data/" + plant_name + "/" + "asja-" + str(year) + "-" + str(month) + "-01" + "-" + str(year) + "-" + str(month+1) + "-01.csv"
            else:
                filename = "./Data/" + plant_name + "/" + "asja-" + str(year) + "-" + str(month) + "-01" + "-" + str(year+1) + "-1-01.csv"                            

#            print filename
            
            try:
                df_plant = pd.concat([df_plant, pd.read_csv(filename, delimiter=";")])
            except:
                pass
    
        df_plant = df_plant.drop("forecast KWh", axis=1)
        df_plant = df_plant.drop("forecast disponibile KWh", axis=1)
        df_plant["data/ora"] = pd.to_datetime(df_plant["data/ora"])
        df_plant = df_plant.set_index("data/ora")
        df_plant["datetime"] = df_plant.index.values
        if year == 2012:
            df_plant = df_plant.loc["2012-7-17":]
            
            
        df = df_aa.loc[df_aa.index.intersection(df_plant.index)].dropna()
        df_plant = df_plant.loc[df.index.intersection(df_plant.index)].fillna(0.0)
        df = pd.merge(df, df_plant, how='inner', on=['datetime'])
        
        df["irr"] = df["irradianza misurata kW/m2"]
        df["irr_tilt"] = df["irradianza misurata tilt kW/m2"]
          
        def dec (ndy_):
            return 0.006918 - 0.399912 * np.cos(ndy_) + 0.070257 * np.sin(ndy_) - 0.006758 * np.cos(2*ndy_)\
                    + 0.000907 * np.sin(2*ndy_) - 0.002697 * np.cos(3*ndy_) + 0.00148 * np.sin(3*ndy_)
        
        def eot (ndy_):
            return (0.000075 + 0.001868 * np.cos(ndy_) - 0.032077 * np.sin(ndy_) - 0.014615 * np.cos(2*ndy_)\
                    - 0.04089 * np.sin(2*ndy_)) * 229.18 / 60.0
                    
        df["dec"] = df.ndy_.apply(dec)
        df["eot"] = df.ndy_.apply(eot)
        
        df["Ls"] = 15.0
        df["latitude"] = latitude
        df["longitude"] = longitude
        df["latitude_rad"] = np.deg2rad(latitude)
        df["longitude_rad"] = np.deg2rad(longitude)
        df["beta"] = np.deg2rad(30.0)
        df["aw"] = np.deg2rad(0.001)
        isc = 1367
        df["i0"] = isc * (1 + 0.034 * np.cos(2*np.pi*df.ndy/365.25))
        
        df.loc[(df.datetime > datetime.datetime(2012, 3, 25, 2)) & (df.datetime < datetime.datetime(2012, 10, 28, 3)), "ora_legale"] = 1
        df.loc[(df.datetime > datetime.datetime(2013, 3, 31, 2)) & (df.datetime < datetime.datetime(2013, 10, 27, 3)), "ora_legale"] = 1
        df.loc[(df.datetime > datetime.datetime(2014, 3, 30, 2)) & (df.datetime < datetime.datetime(2014, 10, 26, 3)), "ora_legale"] = 1
        df.loc[(df.datetime > datetime.datetime(2015, 3, 29, 2)) & (df.datetime < datetime.datetime(2015, 10, 25, 3)), "ora_legale"] = 1
        df.loc[(df.datetime > datetime.datetime(2016, 3, 27, 2)) & (df.datetime < datetime.datetime(2016, 10, 30, 3)), "ora_legale"] = 1
        df.loc[(df.datetime > datetime.datetime(2017, 3, 26, 2)) & (df.datetime < datetime.datetime(2017, 10, 29, 3)), "ora_legale"] = 1
        
        df["ora_legale"] = df.ora_legale.fillna(0.0)              
        df["solar_time"] = df.datetime.apply(lambda x: x.hour + x.minute/60.0) + (df.longitude-df.Ls)/15.0 + df.eot - df.ora_legale
        df.loc[df.solar_time < 0, "solar_time"] = df.loc[df.solar_time < 0, "solar_time"] + 24
        
        df["omega"] = 15.0 * (12.0 - df.solar_time)
        df["omega_rad"] = df["omega"].apply(lambda x: np.deg2rad(x))
        df["omega_sin"] = df["omega_rad"].apply(lambda x: np.sin(x))
        df["omega_cos"] = df["omega_rad"].apply(lambda x: np.cos(x))
        df["alpha"] = np.arcsin(np.sin(df.dec) * np.sin(np.deg2rad(df.latitude)) + np.cos(df.dec) * df.omega_cos * np.cos(np.deg2rad(df.latitude)) )
        
        df.loc[df.alpha > 0, "thetaz"] = np.arccos(np.sin(df.alpha))
        df.loc[df.alpha < 0, "thetaz"] = np.deg2rad(90.0 - np.rad2deg(df.alpha))
        df["thetaz_cos"] = np.cos(df.thetaz)
        
        def zibeta ():
            return np.arccos((np.sin(df.latitude_rad) * np.cos(df.beta) - np.cos(df.latitude_rad) * np.sin(df.beta) * np.cos(df.aw)) * np.sin(df.dec)\
                             + (np.cos(df.latitude_rad) * np.cos(df.beta) + np.sin(df.latitude_rad) * np.sin(df.beta) * np.cos(df.aw)) * np.cos(df.dec) * np.cos(df.omega_rad)\
                              + np.cos(df.dec) * np.sin(df.beta) * np.sin(df.aw) * np.sin(df.omega_rad))
        
        df["zibeta"] = zibeta()
        df["rdire"] = np.cos(df.zibeta)/np.cos(df.thetaz)
        df["rdiff"] = (1+np.cos(df.beta))/2.0
        df["ralb"] = (1-np.cos(df.beta))/2.0
        df["iextra"] = df.i0 * np.cos(df.zibeta)
        
        df["Mt"] = df.irr/df.iextra
        df.loc[df.Mt < 0.22, "idiff/ior"] = 1.0 - 0.09 * df.Mt
        df.loc[(df.Mt >= 0.22) & (df.Mt < 0.8), "idiff/ior"] = 0.951 - 0.160 * df.Mt + 4.388 * df.Mt**2 - 16.638 * df.Mt**3 + 12.336*df.Mt**4
        df.loc[df.Mt > 0.8, "idiff/ior"] = 0.165
        df["idiff"] = df["idiff/ior"] * df.irr
        df["idire"] = df.irr - df.idiff
        df["hicalc"] = df.idire * df.rdire + df.idiff * df.rdiff + (df.idire + df.idiff) * df.ralb * 0.26
                
        df["irr_mis"] = df.irr / 1000.0
        df["irr_tilt"] = df.irr_tilt / 1000.0
        df["en_prod"] = df["misurato kWh"]
        df["en_prod_m2"] = (df.en_prod / surface)
    
#            df = df.dropna()        
#            df = df.loc[(df.hicalc > 0.0) & (df.hicalc < 3000.0)]
#            df = df.loc[(df.en_prod > 0.0)]
        
        df["t"] = df.datetime
        df["date"] = df.datetime.apply(lambda x: x.date())
    #        df["week"] = df.datetime.apply(lambda x: x.week)
        df["week"] = df.datetime.apply(lambda x: x.year*51+x.week)
        df["week"] -= df.week.min()
    #        df["month"] = df.datetime.apply(lambda x: x.month)
        df["month"] = df.datetime.apply(lambda x: x.year*11+x.month)
        df["month"] -= df.month.min()
        df["year"] = df.datetime.apply(lambda x: x.year)
        df = df.set_index("datetime").dropna()
        df = df[df.year == year]
        
        writer = pd.ExcelWriter("./Processed_Data/" + plant_name + str(year) + ".xlsx", engine='xlsxwriter')
        df.to_excel(writer, sheet_name='all')
    
        def compute_plant_performances (df, group_col):
    
            df_agg = df[[group_col, "irr_mis", "irr_tilt", "hicalc", "en_prod_m2", "en_prod"]].groupby(group_col).sum()
            df_agg["rend"] = df_agg.en_prod_m2 / df_agg.irr_tilt
            df_agg["pr"] = (df_agg.en_prod) / (df_agg.irr_tilt * power_tot)
    
            if group_col == "month":
                df_agg["enprod_enel"] = s_enel
                df_agg["enprod_enel_m2"] = s_enel / surface
                df_agg["rend_enel"] = df_agg.enprod_enel / surface / df_agg.irr_tilt
                df_agg["pr_enel"] = (df_agg.enprod_enel) / (df_agg.irr_tilt * power_tot)
    
            if group_col == "year":
                df_agg["enprod_enel"] = s_enel.sum()
                df_agg["enprod_enel_m2"] = s_enel / surface
                df_agg["rend_enel"] = df_agg.enprod_enel / surface / df_agg.irr_tilt
                df_agg["pr_enel"] = (df_agg.enprod_enel) / (df_agg.irr_tilt * power_tot)
            df_agg[group_col] = df.groupby(group_col).year.unique()
    
            if group_col == "month":
                X = df_agg["irr_tilt"].values.reshape(-1, 1)
                y = df_agg["enprod_enel_m2"].values.reshape(-1, 1)
                model = linear_model.LinearRegression()
                model.fit(X, y)
                df_agg["y_hat_irr_mis"] = model.predict(X)
                df_agg["err_irr_mis"] = (df_agg["en_prod_m2"] - df_agg["y_hat_irr_mis"])
                df_agg["cum_err_irr_mis"] = df_agg["err_irr_mis"].cumsum()
    
                pd.DataFrame([model.coef_[0,0], model.intercept_[0]], index=["coeff", "intercept"])\
                            .to_excel(writer, sheet_name=group_col, 
                                      startcol=20, startrow=0, 
                                      header=None, index=True)
    
            df_agg.to_excel(writer, sheet_name=group_col)
                            
            return df_agg
    
        for group_col in ["date", "week", "month", "year"]:
            
            df_agg = compute_plant_performances(df, group_col)
    
            workbook = writer.book
            worksheet = writer.sheets[group_col]
            
            chart_row = 3
            if group_col == "month":
                max_col = 13
            else:
                max_col = 7
                
            for col in range(1, max_col):
                
                chart = workbook.add_chart({'type': 'line'})            
                max_row = len(df_agg) + 1
                chart.add_series({
                    'name':       [group_col, 0, col],
                    'categories': [group_col, 1, 0, max_row, 0],
                    'values':     [group_col, 1, col, max_row, col],
                })
                if group_col == "date":
                    chart.set_x_axis({'name': 'tempo', 'date_axis': True})
                else:
                    chart.set_x_axis({'name': 'tempo'})
                chart.set_y_axis({'name': df_agg.iloc[:, col-1].name, 'major_gridlines': {'visible': False}})
                worksheet.insert_chart('Q' + str(chart_row), chart)
                chart_row += 16
    
        writer.save()

for plant in plants:
    print plant

    writer = pd.ExcelWriter("./Processed_Data/" + plant["plant_name"] + ".xlsx", engine='xlsxwriter')

    for group_col in ["date", "week", "month", "year"]:
        df_plant_tot = pd.DataFrame()
        print group_col
        for year in range(2013, 2017):
            print year
            df_year = pd.read_excel("./Processed_Data/" + plant["plant_name"] + str(year) + ".xlsx", sheetname=group_col)
            df_year = df_year.iloc[:, :16]
            df_plant_tot = pd.concat([df_plant_tot, df_year], ignore_index=True)

        df_plant_tot.to_excel(writer, sheet_name=group_col)

        workbook = writer.book
        worksheet = writer.sheets[group_col]
        
        chart_row = 3
        if group_col == "month":
            max_col = 13
        else:
            max_col = 7
            
        for col in range(1, max_col):
            
            chart = workbook.add_chart({'type': 'line'})            
            max_row = len(df_plant_tot) + 1
            chart.add_series({
                'name':       [group_col, 0, col],
                'categories': [group_col, 1, 0, max_row, 0],
                'values':     [group_col, 1, col, max_row, col],
            })
            if group_col == "date":
                chart.set_x_axis({'name': 'tempo', 'date_axis': True})
            else:
                chart.set_x_axis({'name': 'tempo'})
            chart.set_y_axis({'name': df_plant_tot.iloc[:, col-1].name, 'major_gridlines': {'visible': False}})
            worksheet.insert_chart('S' + str(chart_row), chart)
            chart_row += 13

    writer.save()
    
#plt.figure(figsize=(15,6))
#df["irr"].plot(alpha=0.5, label="measured")
#df["hicalc"].plot(alpha=0.5, label="projection")
#plt.legend()
#plt.savefig("model.png")
#
#plt.figure(figsize=(15,6))
#df["irr"].iloc[1000:1600].plot(alpha=0.5, label="measured")
#df["hicalc"].iloc[1000:1600].plot(alpha=0.5, label="projection")
#plt.legend()
#plt.savefig("model_zoom_begin.png")
#
#plt.figure(figsize=(15,6))
#df["irr"].iloc[4000:4300].plot(alpha=0.5, label="measured")
#df["hicalc"].iloc[4000:4300].plot(alpha=0.5, label="projection")
#plt.legend()
#plt.savefig("model_zoom_center.png")
#
#plt.figure(figsize=(15,6))
#df["irr"].iloc[7100:7500].plot(alpha=0.5, label="measured")
#df["hicalc"].iloc[7100:7500].plot(alpha=0.5, label="projection")
#plt.legend()
#plt.savefig("model_zoom_end.png")
