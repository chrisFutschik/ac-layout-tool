def renWrite(TPOC_RENAME_FILE,POC_MPE,tpoc,verifyLayer,TPOC_D):

    if '"' in tpoc[3]:
        SIZE = tpoc[3].replace('"','in')
        if '/' in tpoc[3]:
            SIZE = SIZE.replace('/','-')

    else:
        SIZE = tpoc[3]

    # Non-verified layer
    if verifyLayer == 1 or TPOC_D == 'x':
        if POC_MPE == 'E':
            TPOC_RENAME_FILE.write('-rename Block "'+tpoc[1]+'v" "'+tpoc[1]+' '+tpoc[2]+'"'"\n\n\n")
        else:		
            TPOC_RENAME_FILE.write('-rename Block "'+tpoc[1]+'v" "'+tpoc[1]+' '+tpoc[2]+' '+SIZE+'"'"\n\n\n")
    # Verified layer
    else:
        if POC_MPE == 'E':
            TPOC_RENAME_FILE.write('-rename Block "'+tpoc[1]+'" "'+tpoc[1]+' '+tpoc[2]+'"'"\n\n\n")
        else:		
            TPOC_RENAME_FILE.write('-rename Block "'+tpoc[1]+'" "'+tpoc[1]+' '+tpoc[2]+' '+SIZE+'"'"\n\n\n")
