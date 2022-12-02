def tpocEquipCheck(blockType,tpoc,n):

    # Assign initial value
    if n == 0:
        blockType = tpoc[0]
    
    # Post inital
    if n!= 0:
        if tpoc[0] == '':
            tpoc[0] = blockType
        elif tpoc[0] != '':
            blockType = tpoc[0]

    return blockType, tpoc[0]