# -*- coding: utf-8 -*-
"""
Created on Mon Jul  4 16:04:41 2022

@author: Autozhz
"""

import sys,os
from JdocsExtract import JdocsExtract
from AppendOA import AppendOA
from Global import *

#%%
# Extract Jdocs content
JdocsExtract1 = JdocsExtract()
df = JdocsExtract1.get_member_info_all(doc_folder,
                                       df_input_name=jdocs_extract_path,
                                       df_output_name=jdocs_extract_path,
                                       isSave=1)
#%%
# Combine OA record and append new OA record
AppendOA1 = AppendOA()
df = AppendOA1.combine_record(oa_all_path,
                              is_version2 = is_version2,
                              df_input_name=df_oa_path,
                              df_output_name=df_oa_path,
                              isSave=1)

AppendOA1.open_path(df_oa_path, isOpenPath = 1, **kwargs)

print('End All Processing')
