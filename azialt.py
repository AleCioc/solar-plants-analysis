#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Dec 13 23:40:35 2017

@author: alecioc
"""

import os
import os.path

import datetime

from pysolar.solar import *

from conf import plants

for plant in plants:

    plant_name = plant["plant_name"]    
    filename = "azi_alt_" + plant_name + ".csv"
    
    if os.path.isfile(filename):
        	os.remove(filename)
    with open(filename,"a+") as data:
        	data.write("datetime,altitude,azimuth\n")
    
    latitude = plant["latitude"]
    longitude = plant["longitude"]

    for year in range(2012, 2018):
        for month in range(1, 13):
            for day in range(1, 32):
                for hour in range(0, 24):
                    for minute in range(0, 60, 60):
                        try:
                            d = datetime.datetime(year, month, day, hour, minute, 0, 0)
                            altitude = get_altitude(latitude, longitude, d)
                            azimuth = get_azimuth(latitude, longitude, d)
                            with open(filename,"a+") as data:
                                data.write(str(d) + "," + str(altitude) + "," + str(azimuth) + "\n")
                        except:
        #                    print (month, day, hour, minute)
                            pass
