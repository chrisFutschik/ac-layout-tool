#Code to import csv, organize and output .scr AutoCad Script
#CSV import
#IMPORTANT: Spaces count in the .CSV files
#v2.1

import csv
import os
import shutil
from os import path
from RenCheck import renCheck
from WriteACScript import writeACScript
from RenWrite import renWrite
from TpocIdenCheck import tpocIdenCheck
from TpocEquipCheck import tpocEquipCheck
from TpocRefCheck import tpocRefCheck
from TpocFinalCheck import tpocFinalCheck
from BuildTPOC import buildTPOC
from TpocGenWrite import tpocGenWrite

# Creating counters
n = 0			# Loop counter
error = 0 		# Flag file error
tpocError = 0	# Flag for TPOC errors
numIssues = 1	# Issue counter
numOverride = 0 # Number of overrides counter
numErrors = 0	# Number of errors counter
totalGen = 0	# Total generated counter
blockType = '' 	# empty array for block types

# Creating files 
TPOC_GEN_FILE = open("TPOC_GEN.txt","w") #Creating write file 
TPOC_RENAME_FILE = open("TPOC_RENAME.txt","w") #Creating write file 
POC_FILE = open("TPOC_SCRIPT.txt","w") #Creating write file 
POC_ERROR_FILE = open("ERROR_LOG.txt","w") #Creating error file

# Creating arrays
TPOCS = []
TypeIDEN = []

#Importing data
try:
	with open('TPOC_INFO.csv','r') as file:
		my_reader = csv.reader(file, delimiter=',')
		for row in my_reader:
			TPOCS.append(row)
except FileNotFoundError:
	print('MISSING CSV FILE ERROR: Please create TPOC_INFO.csv file and try again.\n')
	POC_ERROR_FILE.write('No TPOC_INFO.csv file found.')	
	error = 1

#Data check
try:
	for tpoc in TPOCS:
		if tpoc[1] == '':
			a =1 
except IndexError:
	print('CSV DATA ERROR: Please check that there is data in required columns within TPOC_INFO.csv\n')
	POC_ERROR_FILE.write('CSV DATA ERROR: Please check that there is data in required columns within TPOC_INFO.csv')
	POC_ERROR_FILE.close()
	POC_FILE.close()
	TPOC_GEN_FILE.close()
	TPOC_RENAME_FILE.close()
	os.remove('TPOC_RENAME.txt')
	os.remove('TPOC_GEN.txt')
	os.remove('TPOC_SCRIPT.txt')
	exit()


with open('scriptFunctions\TypeIDEN.csv','r') as file:
	my_reader = csv.reader(file, delimiter=',')
	for row in my_reader:
		TypeIDEN.append(row)

# Cleaning Up TPOC folder
dir = os.getcwd()+'\TPOCs'
for f in os.listdir(dir):
	os.remove(os.path.join(dir,f))

shutil.copy(os.getcwd()+'\scriptFunctions\TPOCD.dwg',os.getcwd()+'\TPOCs')
shutil.copy(os.getcwd()+'\scriptFunctions\TPOCDv.dwg',os.getcwd()+'\TPOCs')
shutil.copy(os.getcwd()+'\scriptFunctions\TPOCM.dwg',os.getcwd()+'\TPOCs')
shutil.copy(os.getcwd()+'\scriptFunctions\TPOCMv.dwg',os.getcwd()+'\TPOCs')
shutil.copy(os.getcwd()+'\scriptFunctions\TPOCE.dwg',os.getcwd()+'\TPOCs')
shutil.copy(os.getcwd()+'\scriptFunctions\TPOCEv.dwg',os.getcwd()+'\TPOCs')

# Main program loop
for tpoc in TPOCS: #tpoc here will be the number of rows in the CSV file		
	x_cord = 0
	typeCheck = 1
	sizeCheck = 0
	refCheck = 0
	verifyLayer = 0
	TPOC_D = 0
	TPOC_SIZE = []
	POC_MPE = '' 	  #Type of POC connection M/P/E		

# Checking for valid TPOC REF Numbers
	tpoc[1], refCheck = tpocRefCheck(tpoc,refCheck,n,TPOCS)

# Checking for valid TPOC types
	POC_MPE, typeCheck = tpocIdenCheck(TypeIDEN,tpoc,POC_MPE,typeCheck)

# Checking for valid TPOC sizes

	TPOC_D, verifyLayer, sizeCheck = buildTPOC(tpoc,TPOC_D,verifyLayer,sizeCheck,POC_MPE)

# Creating Headers
	if n == 0:
		TPOC_GEN_FILE.write('(command) ATTDIA 0 ATTREQ 0\n')
		TPOC_GEN_FILE.write('(command) -DWGUNITS 1 2 4 Y Y\n')
		TPOC_GEN_FILE.write('(command) PDMODE 35\n')
		TPOC_GEN_FILE.write('(command) ERASE All \n')
		TPOC_GEN_FILE.write('(command) PURGE A * N\n')
		TPOC_RENAME_FILE.write('(command) ATTDIA 1 ATTREQ 1\n')
			
#Assigns previous equipment name if no equipment type is detected
	blockType, tpoc[0] = tpocEquipCheck(blockType,tpoc,n)

#Incrementing x value by 5 to space created blocks 5in apart
	x_cord = n*5 

# Handling and skipping illegal TPOCs
	tpocError = 0
	tpocError = tpocFinalCheck(POC_ERROR_FILE,typeCheck,refCheck,sizeCheck,numIssues,tpoc,n)

	if tpocError == 1:
		n+=1
		numIssues+=1
		numErrors+=1
		continue 
	elif tpocError == 2:
		numIssues+=1
		numOverride+=1
		totalGen+=1
	else:
		totalGen+=1

#Writing out file
	writeACScript(TPOC_RENAME_FILE,x_cord,tpoc,verifyLayer,TPOC_D)
	tpocGenWrite(TPOC_GEN_FILE, TPOC_D, tpoc, POC_MPE, verifyLayer, n)

# renCheck formats TPOC names to be eligble to AC "REN" command
	tpoc[2] = renCheck(tpoc[2],2)

#Writing to file.
	renWrite(TPOC_RENAME_FILE,POC_MPE, tpoc, verifyLayer,TPOC_D)

#Counters
	n+=1

#Combining TPOCs

POC_FILE.close()
POC_ERROR_FILE.close()
TPOC_GEN_FILE.close()
TPOC_RENAME_FILE.close()

filenames = ['TPOC_GEN.txt','TPOC_RENAME.txt']
with open('TPOC_SCRIPT.txt', 'w') as outfile:
	for fname in filenames:
		with open(fname) as infile:
			outfile.write((infile.read()))

POC_FILE.close()
POC_ERROR_FILE.close()
TPOC_GEN_FILE.close()
TPOC_RENAME_FILE.close()

#try:
os.remove('TPOC_RENAME.txt')
os.remove('TPOC_GEN.txt')
#except:
#	pass

if path.exists('TPOC_SCRIPT.scr'):
	os.remove('TPOC_SCRIPT.scr')
	os.rename(r'TPOC_SCRIPT.txt',r'TPOC_SCRIPT.scr')
else:
	os.rename(r'TPOC_SCRIPT.txt',r'TPOC_SCRIPT.scr')

# Final display message 

if error == 0:
	print('> TOTAL GENERATED: '+str(totalGen)+' out of '+str(n))

	if os.stat("ERROR_LOG.txt").st_size == 0:
		os.remove("ERROR_LOG.txt")
		print("> NO ISSUES FOUND\n")
	else:
		if numErrors != 0:
			print('> ERRORS  FOUND:   '+str(numErrors)+' - (OMMITED FROM FINAL SCRIPT)')
		else:
				print('> ERRORS  FOUND:   '+str(numErrors))
		print('> USER OVERRIDES:  '+str(numOverride))
		print('> ALL ISSUES LOGGED IN "ERROR_LOG.txt"\n')

	print('> *****AUTOCAD SCRIPT SUCCESSFULLY GENERATED*****')
	
	