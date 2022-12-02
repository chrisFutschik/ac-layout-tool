import re

def renCheck(string,types):
# 1 - TPOC REF
# 2 - SERVICE
# 3 - SIZE

# Removes accidental spaces infront or behind string
	if types == 1:
		string = string.strip()


# For AC REN function
	if types == 2:
		string = string.replace(' ','_') 

	if types == 1 or types == 2:
		for i in string:
			if i == '<': 
				string = string.replace('<','')
			if i == '%':		
				string = string.replace('%','')
			if i == '>':
				string = string.replace('>','')
			if i == '=':
				string = string.replace('=','')
			if i == ';':
				string = string.replace(';','')
			if i == ':':
				string = string.replace(':','')
			if i == '"':
				string = string.replace('"','')
			if i == "/":
				string = string.replace('/','')
			if i == "(":
				string = string.replace('(','')
			if i == ")":
				string = string.replace(')','')
			if i == "*":
				string = string.replace('*','')
			if i == "?":
				string = string.replace('?','')
			if i == "\\":
				string = string.replace('\\','')

	return string
