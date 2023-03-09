#!/usr/bin/env python3

import os
import sys

    

#WF_BUNDLE = os.getenv('alfred_workflow_bundleid')
WF_BUNDLE = "CAZ"
CACHE_FOLDER = os.path.expanduser('~')+"/Library/Caches/com.runningwithcrayons.Alfred/Workflow Data/"+WF_BUNDLE+"/"
DATA_FOLDER = os.path.expanduser('~')+"/Library/Application Support/Alfred/Workflow Data/"+WF_BUNDLE





if not os.path.exists(DATA_FOLDER):
    os.makedirs(DATA_FOLDER)


            


def log(s, *args):
    if args:
        s = s % args
    print(s, file=sys.stderr)


