from RenSize import renSize
import os

def tpocSizeCheck(TypeSIZE,tpoc,POC_MPE,POC_SIZE,POC_SIZE_REN,POC_BLK_NM,sizeCheck,verifyLayer):

# Removes spaces before and after number to limit user input error
    tpoc[3] = tpoc[3].strip().lower()

    for size in TypeSIZE:

        if size[0] == tpoc[3] and size[1] == POC_MPE:
            POC_SIZE = size[0]
            POC_SIZE_REN = renSize(POC_MPE,size[2])
            POC_BLK_NM = size[2]
            sizeCheck = 0
            
        if size[0] == 'ELEC' and POC_MPE == 'E':	
            POC_SIZE = size[0]
            POC_SIZE_REN = renSize(POC_MPE,size[2])
            POC_BLK_NM = size[2]
            sizeCheck = 0

        #Size value left blank error
        if tpoc[3] == '' and POC_MPE != 'E': 
            sizeCheck = 2

     
        
    return POC_SIZE, POC_SIZE_REN, POC_BLK_NM, sizeCheck, verifyLayer