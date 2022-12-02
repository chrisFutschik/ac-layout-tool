from RenCheck import renCheck

def tpocRefCheck(tpoc,refCheck,n,TPOC):
    # Checking for blank TPOC References
    if tpoc[1] == '':
        refCheck = 1
        return tpoc[1], refCheck

    tpoc[1] = renCheck(tpoc[1],1)

    # Adds M to fully numerical Reference number
    if tpoc[1].isnumeric() == 1:
        #print(len)
        if len(tpoc[1]) == 1:
            tpoc[1] = "M0"+tpoc[1] 
        else:
            tpoc[1] = "M"+tpoc[1] 

    # Check for duplicate TPOC References
    x = 0
    for i in TPOC:
        if i[1] == tpoc[1] and n != x and n > x:
            refCheck = 2
        x+=1

    return tpoc[1], refCheck