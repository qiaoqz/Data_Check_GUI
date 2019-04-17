# -*- coding: utf-8 -*-
"""
Created on Mon Apr  1 15:50:19 2019

@author: Qiao Zhang
"""
import pandas as pd
import numpy as np
import _pickle as pickle

def dump_baseline():
    g = open('client_id_pickle','wb')
    alldata = pd.read_excel("Alldata_3_22_v2.xlsx")
    dumpdata = alldata[["Agency_Baseline","ClientID_Baseline"]]
    agency_baseline = np.array(dumpdata["Agency_Baseline"])
    clientid_baseline = np.array(dumpdata["ClientID_Baseline"])
    pickling = {}
    pickling = {'agency_baseline':agency_baseline, 'clientid_baseline':clientid_baseline}
    pickle.dump(pickling,g)
    g.close()
    
    
    
dump_baseline()