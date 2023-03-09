#!/usr/bin/python3 
# giovanni
# Thursday, January 26, 2023
# Overcast ☁️   🌡️+43°F (feels +36°F, 51%) 🌬️→19mph 🌒 Thu Jan 26 15:21:51 2023
# W4Q1 – 26 ➡️ 338 – 260 ❇️ 104

import sys
import json
import os
MYSTRING = os.getenv('myString')


def log(s, *args):
    if args:
        s = s % args
    print(s, file=sys.stderr)




#log (myOut)

def main ():
    result = {"items": []}    
    result["items"].append({
                "title": f"result",
                "subtitle": "subtitle",
                "type": "file",
                
                "valid":'TRUE',
                        
                "arg":MYSTRING})
            
    
    
    #tts = time.time()
    #finalTime = tts - int(startTS)
    #log (f"================================Timestamp end of script (in sec): {finalTime:.2}")
    
    print (json.dumps(result))
    
if __name__ == "__main__":
    log ("brrrr")
    main()
    
