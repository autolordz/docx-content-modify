# -*- coding: utf-8 -*-
"""
Created on Mon Jun 27 10:34:48 2022

@author: Autozhz
"""
import sys,os
import pandas as pd
from CombineDuplicate import CombineDuplicate
from ExportPostal import ExportPostal
from CommomClass import CommomClass
from Global import *
from PrintLog import PrintLog
import traceback

try:
    #%%
    
    # CommomClass1 = CommomClass()
    # CommomClass1.check_df_expand(df_input_name = df_expand_path)
    
    df = pd.DataFrame()
    # CombineDuplicate Optional to Export postal
    CombineDuplicate1 = CombineDuplicate()
    df = CombineDuplicate1.combine_codes_user(
                                    df_input_name = df_expand_path,
                                    df_output_name = df_combine_path,
                                    isSave=1)
    #%%
    # Export portal with df OA
    # data_main_tmp.xlsx
    ExportPostal1 = ExportPostal()
    df = ExportPostal1.generate_postal(
                                df_input = df,
                                df_input_name = df_expand_path,
                                isOpenPath = 1,
                                **kwargs)

except Exception as e:
    PrintLog.log('APP Running Error %s'%traceback.print_exc())

print('End All Processing')

