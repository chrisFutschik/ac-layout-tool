def buildTPOC(tpoc, TPOC_D, verifyLayer, sizeCheck, POC_MPE):

# Removes spaces before and after number to limit user input error
    try:
        tpoc[3] = tpoc[3].strip().lower()
    except IndexError:
        pass
    
    if tpoc[3] == '':
        sizeCheck = 1
        if POC_MPE == 'E':
            sizeCheck = 0


    if '"' in tpoc[3]:
        NUM = tpoc[3].replace('"','')        
        if ' ' in NUM:
            a = int(NUM[0])
            try:
                b = int(NUM[1])
                nonum = 0
            except:
                nonum = 1

            if nonum  == 1:
                TPOC_D = a            
            else:  
                TPOC_D = a*10 + b

        if '/' in tpoc[3]:
            NUM = tpoc[3].replace('"','')
            B = int(NUM[-3])
            C = int(NUM[-1])
            TPOC_D+=(B/C)
        elif '.' in tpoc[3]:
            TPOC_D = float(tpoc[3].replace('"',''))
        else:
            TPOC_D = int(tpoc[3].replace('"',''))
	
    if 'mm' in tpoc[3]:
        TPOC_D = int(tpoc[3].replace('mm',''))/25.4

    if 'nw' in tpoc[3]:
        TPOC_D = int(tpoc[3].replace('nw',''))/25.4

    if 'iso' in tpoc[3]:
        TPOC_D = int(tpoc[3].replace('iso',''))/25.4

    if 'kf' in tpoc[3]:
        TPOC_D = int(tpoc[3].replace('kf',''))/25.4    

    #Checking for verify
    try: 
        if 'v' in tpoc[4].lower():
            verifyLayer = 1
        if 'x' in tpoc[4].lower():
            sizeCheck = 2
            TPOC_D = 'x'
    except IndexError: 
        pass
        
    return TPOC_D, verifyLayer, sizeCheck