import os

def tpocGenWrite(TPOC_GEN_FILE, TPOC_D, tpoc, POC_MPE,verifyLayer, n):

    #!!! Space after * ERASE x,x,x" " is very important
    layer = ''

    if POC_MPE == 'P' or POC_MPE == 'M':
        if POC_MPE == 'P':
            layer = 'D'
        elif POC_MPE == 'M':
            layer = 'M'

        if TPOC_D == 'x':
            TPOC_GEN_FILE.write('(command) -INSERT "'+os.getcwd()+'\TPOCs\\'+'TPOC'+layer+'v.dwg" 0,0,0 1   EXPLODE 0,0,0\n')
            TPOC_GEN_FILE.write('(command) -layer set A-EQPM-POC'+layer+'  -WBLOCK "'+os.getcwd()+'\TPOCs\\'+tpoc[1]+'v.dwg" * ERASE 0,0,0 \n')
        else:
            if verifyLayer == 1: 
                TPOC_GEN_FILE.write('(command) -INSERT "'+os.getcwd()+'\TPOCs\\'+'TPOC'+layer+'v.dwg" 0,0,0 1   EXPLODE 0,0,0\n')
                TPOC_GEN_FILE.write('(command) -layer set A-EQPM-POC'+layer+'  CYLINDER 0,0,0 d '+str(TPOC_D)+' a 0,2,0 -WBLOCK "'+os.getcwd()+'\TPOCs\\'+tpoc[1]+'v.dwg" * ERASE '+str(TPOC_D/2)+',0,0 \n')
            else:
                TPOC_GEN_FILE.write('(command) -INSERT "'+os.getcwd()+'\TPOCs\\'+'TPOC'+layer+'.dwg" 0,0,0 1   EXPLODE 0,0,0\n')
                TPOC_GEN_FILE.write('(command) -layer set A-EQPM-POC'+layer+'  CYLINDER 0,0,0 d '+str(TPOC_D)+' a 0,2,0 -WBLOCK "'+os.getcwd()+'\TPOCs\\'+tpoc[1]+'.dwg" * ERASE '+str(TPOC_D/2)+',0,0 \n')
            
    if POC_MPE == 'E':
        if verifyLayer == 1:
            TPOC_GEN_FILE.write('(command) -INSERT "'+os.getcwd()+'\TPOCs\\'+'TPOCEv.dwg" 0,0,0 1   EXPLODE 0,0,0\n')
            TPOC_GEN_FILE.write('(command) -layer set A-EQPM-POCE  -WBLOCK "'+os.getcwd()+'\TPOCs\\'+tpoc[1]+'v.dwg" * ERASE 0,0,0 \n')
        else:
            TPOC_GEN_FILE.write('(command) -INSERT "'+os.getcwd()+'\TPOCs\\'+'TPOCE.dwg" 0,0,0 1   EXPLODE 0,0,0\n')
            TPOC_GEN_FILE.write('(command) -layer set A-EQPM-POCE  -WBLOCK "'+os.getcwd()+'\TPOCs\\'+tpoc[1]+'.dwg" * ERASE 0,0,0 \n')

    TPOC_GEN_FILE.write('(command) ERASE ALL \n')