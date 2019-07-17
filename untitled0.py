# -*- coding: utf-8 -*-
"""
Created on Mon Apr 15 10:23:45 2019

@author: DanielYuan
"""
from timeit import timeit
from numba import autojit


@autojit
def foo(x,y):    
    s=0
    for i in range(1,1000):
        s+=i
    return s
t = timeit('foo()', 'from __main__ import foo',number=5)
print(t)

