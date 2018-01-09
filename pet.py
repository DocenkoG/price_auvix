# -*- coding: UTF-8 -*-
import os
import os.path
import sys
import configparser

cfg = configparser.ConfigParser(inline_comment_prefixes=('#'))
cfg.read('cfg_auvix.cfg', encoding='utf-8')
cfg.read('confidential.cfg')
for s in cfg.sections():
    print('-----',s)
    for opt in cfg.options(s):
        print(opt,'=', cfg.get(s,opt)) 

print('~~~~~~~~~~~~')
cfg = configparser.ConfigParser()
cfg.read('confidential.cfg')
cfg.read('confidential.cfg')
for s in cfg.sections():
    print('-----',s)
    for opt in cfg.options(s):
        print(opt,'=', cfg.get(s,opt)) 
