# -*- coding: utf-8 -*-
"""
Created on Mon Jul  4 16:10:23 2022

@author: Autozhz
"""

import os

current_path = os.path.dirname(__file__) # os.path.realpath('.') 

is_version2 = 1

user = 'Elsa' 

if is_version2:
    oa_all_path = os.path.join(current_path, 'oa_all\\%s'%user)
else:
    oa_all_path = os.path.join(current_path, 'oa_all')
    
doc_folder = os.path.join(current_path, 'jdoc_all\\%s'%user) 

save_path = os.path.join(current_path, 'tmp\\%s'%user)
os.makedirs(save_path,exist_ok=1)

jdocs_extract_path = os.path.join(save_path, 'df_jdoc_extract.xlsx')
df_oa_path = os.path.join(save_path, 'df_oa_extract.xlsx')
df_expand_path = os.path.join(save_path, 'df_expand.xlsx')
df_combine_path = os.path.join(save_path, 'df_combine.xlsx')

kwargs={}
kwargs['postal_path'] = os.path.join(save_path, 'postal')
os.makedirs(kwargs['postal_path'],exist_ok=1)

kwargs['postal_template'] = os.path.join(current_path, 'template.docx')
