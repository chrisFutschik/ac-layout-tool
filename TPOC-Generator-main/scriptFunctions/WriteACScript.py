def writeACScript(TPOC_RENAME_FILE,x_cord,tpoc,verifyLayer,TPOC_D):

	if verifyLayer == 1 or TPOC_D == 'x':
		TPOC_RENAME_FILE.write('-insert'+' "'+tpoc[1]+'v.dwg" '+str(x_cord)+',0,0 1 1 0 ')	
	else:
		TPOC_RENAME_FILE.write('-insert'+' "'+tpoc[1]+'.dwg" '+str(x_cord)+',0,0 1 1 0 ')	
	
	if TPOC_D == 'x':
		tpoc[3] = 'NO SIZE'

	TPOC_RENAME_FILE.write(tpoc[1]+' '+tpoc[2]+' '+tpoc[3]+"\n")
	TPOC_RENAME_FILE.write(tpoc[0]+"\n")	
	TPOC_RENAME_FILE.write(tpoc[1]+"\n")
	TPOC_RENAME_FILE.write(tpoc[2]+"\n")
	TPOC_RENAME_FILE.write(tpoc[3]+"\n")
	TPOC_RENAME_FILE.write("\n"+"\n"+"\n"+"\n"+"\n"+"\n"+"\n"+"\n"+"\n"+"\n"+"\n"+"\n"+"\n"+"\n")

