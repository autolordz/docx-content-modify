# -*- coding: utf-8 -*-
"""
Created on Sat Apr 25 18:37:11 2020

@author: autol
"""


#%%
import numpy as np
import pandas as pd

df = pd.DataFrame({'a': np.random.randn(1000),
                        'b': np.random.randn(1000),
                        'N': np.random.randint(100, 1000, (1000)),
                        'x': 'x'})
     

def f(x):
         return x * (x - 1)
   
def integrate_f(a, b, N):
         s = 0
         dx = (b - a) / N
         for i in range(N):
             s += f(a + i * dx)
         return s * dx
     
#%%
%timeit df.apply(lambda x: integrate_f(x['a'], x['b'], x['N']), axis=1)

%prun -l 4 df.apply(lambda x: integrate_f(x['a'], x['b'], x['N']), axis=1)

#%%
%load_ext Cython

#%%
%%cython
def f_plain(x):
    return x * (x - 1)
def integrate_f_plain(a, b, N):
     s = 0
     dx = (b - a) / N
     for i in range(N):
         s += f_plain(a + i * dx)
     return s * dx
     
 #%%
%timeit df.apply(lambda x: integrate_f_plain(x['a'], x['b'], x['N']), axis=1)