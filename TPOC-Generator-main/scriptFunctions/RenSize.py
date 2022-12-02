#Removes "TPOCX"_ and ".dwg" characters from TPOCs .dwg file names to produce 
# REN ready names

def renSize(types, POC_BLK_NM):
    
    POC_BLK_NM = POC_BLK_NM.replace('.dwg','')

    if types == 'P':
	    POC_BLK_NM = POC_BLK_NM.replace('TPOCD ','')
    if types == 'E':
        POC_BLK_NM = POC_BLK_NM.replace('TPOCE_','')
    if types == 'M':
        POC_BLK_NM = POC_BLK_NM.replace('TPOCM_','')
    
    if types != 'P' and types != 'M' and types != 'E':
        print('ERROR: RenSize Function cannont find type of POC.')

    return POC_BLK_NM
