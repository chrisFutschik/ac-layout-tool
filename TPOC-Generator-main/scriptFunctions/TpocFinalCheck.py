def tpocFinalCheck(POC_ERROR_FILE,typeCheck,refCheck,sizeCheck,numIssues,tpoc,n):
# Handling and skipping illegal TPOCs
# error = 1 - naming issue
# error = 2 - size issue

    linenum = n+1

# Not a standard type
    if typeCheck == 1:
        POC_ERROR_FILE.write(str(numIssues)+'. TYPE ERROR (Type 1): ['+tpoc[2]+'] not a standard type -> ['+tpoc[1]+' '+tpoc[2]+' '+tpoc[3]+'] will be omitted [TPOC_INFO.csv - line:'+str(linenum)+'].\n')	
        error = 0
        return 1

# Missing type value
    elif typeCheck == 2: 
        POC_ERROR_FILE.write(str(numIssues)+'. TYPE ERROR (Type 2): No service type value found -> ['+tpoc[1]+' '+'XXX'+' '+tpoc[3]+'] will be omitted [TPOC_INFO.csv - line:'+str(linenum)+'].\n')	
        error = 0
        return 1

# User specified Size and type
    elif typeCheck == 3 and sizeCheck == 2: 
        POC_ERROR_FILE.write(str(numIssues)+'. ERROR OVERRIDE: User specified type ['+tpoc[2]+'] and size [NO SIZE] -> ['+tpoc[1]+' '+tpoc[2]+' NO SIZE] will be used [TPOC_INFO.csv - line:'+str(linenum)+'].\n')	
        error = 0
        return 2

# User custom type
    elif typeCheck == 3: 
        POC_ERROR_FILE.write(str(numIssues)+'. ERROR OVERRIDE: User specified type ['+tpoc[2]+'] -> ['+tpoc[1]+' '+tpoc[2]+' '+tpoc[3]+'] will be used [TPOC_INFO.csv - line:'+str(linenum)+'].\n')	
        error = 0
        return 2

# Missing size value
    elif sizeCheck == 1: 
        POC_ERROR_FILE.write(str(numIssues)+'. SIZE ERROR (Type 1): No size value found -> ['+tpoc[1]+' '+tpoc[2]+' '+'XXX] will be omitted [TPOC_INFO.csv - line:'+str(linenum)+'].\n')	
        error = 0
        return 1

# User custom size
    elif sizeCheck == 2:
        POC_ERROR_FILE.write(str(numIssues)+'. ERROR OVERRIDE: User specified size [NO SIZE]-> ['+tpoc[1]+' '+tpoc[2]+' NO SIZE] will be used [TPOC_INFO.csv - line:'+str(linenum)+'].\n')	
        error = 0
        return 2

# Missing TPOC Ref value
    elif refCheck == 1:
        POC_ERROR_FILE.write(str(numIssues)+'. REFERENCE ERROR (Type 1): No reference number value found -> [XXX '+tpoc[2]+' '+tpoc[3]+'] will be omitted [TPOC_INFO.csv - line:'+str(linenum)+'].\n')	
        error = 0
        return 1

#  TPOC Ref value
    elif refCheck == 2:
        POC_ERROR_FILE.write(str(numIssues)+'. REFERENCE ERROR (Type 2): ['+tpoc[1]+'] has already been specified -> ['+tpoc[1]+' '+tpoc[2]+' '+tpoc[3]+'] will be omitted [TPOC_INFO.csv - line:'+str(linenum)+'].\n')	
        error = 0
        return 1

    else:
        return 0