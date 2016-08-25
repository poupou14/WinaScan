#!/usr/bin/python 
import string, sys
from WSParser import WSParser
#from WSWriter import WSWriter

def main():
	jour = 1
	mois = 1
	annee = 2012
	frequenceSvg = 10
	if len(sys.argv) == 2 :
		if sys.argv[1] == "-h" :
			print "user help :"
			print "$ WinaScan.sh <file>"
			exit()
		else :
			wsfile = sys.argv[1]
	elif len(sys.argv) >= 3 :
		print "user help :"
		print "$ WinaScan.sh <file>"
		exit()
	else :
		print "user help :"
		print "$ WinaScan.sh <file>"
		exit()
	myWS = WSParser()
	myWS.readWS(wsfile)
	myWS.writeOuput()	
main()
