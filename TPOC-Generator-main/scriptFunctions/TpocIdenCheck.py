def tpocIdenCheck(TypeIDEN,tpoc,POC_MPE,typeCheck):

    tpoc[2] = tpoc[2].strip()

    for row in TypeIDEN:
        if row[1] == tpoc[2]:
            POC_MPE = row[0]
            typeCheck = 0

    if tpoc[2] == '':
        typeCheck = 2

    if typeCheck != 0:

        try: 
            if 'd' in tpoc[4].lower():
                typeCheck = 3
                POC_MPE = 'P'

            if 'm' in tpoc[4].lower():
                typeCheck = 3
                POC_MPE = 'M'
        except IndexError: 
            pass

    return POC_MPE, typeCheck