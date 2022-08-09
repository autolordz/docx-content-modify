# -*- coding: utf-8 -*-
"""
Created on Mon Jul  4 16:42:49 2022

@author: Autozhz
"""

import sys,os
from JdocsMergeOA import JdocsMergeOA
from ProcessOA import ProcessOA
from CommomClass import CommomClass
from Global import *
from PrintLog import PrintLog
import traceback

try:
    #%%
    # Expand OA content and append to file
    ProcessOA1 = ProcessOA()
    df_expand = ProcessOA1.append_record(
                            df1_path = df_oa_path,
                            df2_path = df_expand_path,
                            df_output_name = df_expand_path,
                            is_version2 = is_version2,
                            isSave=1)
    #%%
    # Merge jdocs record to OA 
    JdocsMergeOA1 = JdocsMergeOA()
    df = JdocsMergeOA1.fill_jdoc_oa(
                                    df1_path = jdocs_extract_path,
                                    df2_path = df_expand_path,
                                    df_output_name = df_expand_path,
                                    isSave=1)
    
    CommomClass1 = CommomClass()
    CommomClass1.check_df_expand(df_input_name=df_expand_path)

except Exception as e:
    PrintLog.log('APP Running Error %s'%traceback.print_exc())

print('End All Processing')

