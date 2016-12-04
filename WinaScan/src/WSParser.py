#!/usr/bin/python 
from HTMLParser import HTMLParser
import WSDataFormat
import os,string, sys
import urllib
import time
import urllib2
import StringIO
import copy
#import chardet
#### SPECIFIC IMPORT #####
#sys.path.append("../Import/xlrd-0.7.1")
#sys.path.append("../Import/xlwt-0.7.2")
#sys.path.append("../Import/pyexcelerator-0.6.4.1")

from pyExcelerator import *
import xlwt
from xlrd import open_workbook
from xlwt import Workbook,easyxf,Formula,Style
#from lxml import etree
import xlrd

currentGrille = dict({ 'croix_1':[], 'croix_x':[], 'croix_2':[], 'mise':0})
grilleEmpty = True


def onlyascii(char):
    if ord(char) <= 0 or ord(char) > 127: 
	return ''
    else: 
	return char

def isnumber(s):
    try:
        float(s)
        return True
    except ValueError:
        return False

class WSParser:

	def __init__(self): 
		self.__status = False
		self.__nbGames = 0
    		self.__workbook1 = Workbook()
    		self.__grilleSheet = self.__workbook1.add_sheet("Repartition", cell_overwrite_ok=True)
		self.__outPutFileName = "WS.xls"
		self.__fileGrilleCounter = 0

	def readWS(self, file_p):
		# On lit la premiere page puis les suivantes jusqu'a la page vide
		notRead_l = True
		self.__outPutFileName = "WSScan.xls"
		
		myPronoParser = WSGridParser()
		notRead_l = True
		try :
			try :
				#print "Ouverture : %s" % fpUrl_l
				url = open(file_p)
				print "Lecture : %s" % file_p
				myPronoParser.html = url.read()
				notRead_l = False
			#except IOError :
			except IOError:
				notRead_l = True
				print "pb with : %s" % file_p
				print "url read issue"
			url.close()
			myPronoParser.html = filter(onlyascii, myPronoParser.html)
			myPronoParser.feed(myPronoParser.html)	
		except IOError:
			print "problem while reading %s" % file_p
		


	def writeOuput(self):
#		Book to read xls file (output of main_CSVtoXLS)

		index_l = 0
		size_l = len(WSDataFormat.grille['team1'])
		for i in range(0, size_l) :
			p1 = WSDataFormat.grille['croix_1'][i]
			pN = WSDataFormat.grille['croix_x'][i]
			p2 = WSDataFormat.grille['croix_2'][i]
			total = float(p1+pN+p2)
			r1 = p1/total*100
			r2 = p2/total*100
			rN = pN/total*100
			#print "{} vs {} \t{0:.3f}\t{0:.3f}\t{0:.3f}\n".format( WSDataFormat.grille['team1'][i], WSDataFormat.grille['team2'][i], r1, rN, r2)
			print "{} vs {}\t{:10.3f}\t{:10.3f}\t{:10.3f} ".format( WSDataFormat.grille['team1'][i], WSDataFormat.grille['team2'][i], r1,rN,r2)
		#self.__workbook1.save(self.__outPutFileName)

			
class WSGridParser(HTMLParser): 

	def __init__(self): 
		HTMLParser.__init__(self)
		self.__beginOK = False 
		self.__newGridOK = 0 
		self.__gameOK = False 
		self.__nextTeam1 = False
		self.__nextTeam2 = False
		self.__next1 = False
		self.__nextN = False
		self.__next2 = False
		self.__nextMise = False
		self.__nextMontant = False
		self.__game = 0
		self.__grid = []


	def handle_starttag(self, tag, attrs):
		if tag == "table" and len(attrs) == 1 and not self.__beginOK and attrs[0][0] == "class" and attrs[0][1] == "grid-list" :
			self.__beginOK = True
		elif tag == "div" and len(attrs) == 1 and attrs[0][0] == "class" and attrs[0][1] == "small-grid" : # New grid bet
			self.__game = 0
			self.__newGridOK += 1
			#print "grid nO : %s" % self.__newGridOK
			#print WSDataFormat.grille
		elif tag == "div" and len(attrs) == 1 and attrs[0][0] == "class" and attrs[0][1] == "competitor competitor1" and not self.__nextTeam1: # New grid bet
			self.__nextTeam1 = True
			self.__gameOK = True
		elif tag == "div" and len(attrs) == 1 and attrs[0][0] == "class" and attrs[0][1] == "competitor competitor2" and not self.__nextTeam2: # New grid bet
			self.__nextTeam2 = True
		elif tag == "div" and len(attrs) == 1 and attrs[0][0] == "class" and attrs[0][1] == "croix croix_1" :
			self.__next1 = True
		elif tag == "div" and len(attrs) == 1 and attrs[0][0] == "class" and attrs[0][1] == "croix croix_x" :
			self.__nextN = True
		elif tag == "div" and len(attrs) == 1 and attrs[0][0] == "class" and attrs[0][1] == "croix croix_2" :
			self.__next2 = True
		elif tag == "td" and len(attrs) == 0 :
			self.__nextMise = True

		self.__betweenTag = True

	def handle_data(self, data):
		global currentGrille
		if self.__nextTeam1:
			self.__game += 1
			#print "nb game = %d" % self.__game
			#print "data %s" % data
			if self.__newGridOK == 1: # first loop
				#print "1st loop : %s" % data
				WSDataFormat.grille['team1'].append(data)
				WSDataFormat.grille['croix_1'].append(0)
				WSDataFormat.grille['croix_x'].append(0)
				WSDataFormat.grille['croix_2'].append(0)
				currentGrille['croix_1'].append(0)
				currentGrille['croix_x'].append(0)
				currentGrille['croix_2'].append(0)
			else:
				currentGrille['croix_1'][self.__game-1]=0
				currentGrille['croix_x'][self.__game-1]=0
				currentGrille['croix_2'][self.__game-1]=0
			self.__nextTeam1 = False
		elif self.__nextTeam2 and self.__newGridOK == 1 :
			WSDataFormat.grille['team2'].append(data)
			self.__nextTeam2 = False
		elif self.__next1 :
			if data.find("X") >= 0 :
				currentGrille['croix_1'][self.__game-1] = 1
			else:
				currentGrille['croix_1'][self.__game-1] = 0
			self.__next1 = False
		elif self.__nextN :
			if data.find("X") >= 0 :
				currentGrille['croix_x'][self.__game-1] = 1
			else:
				currentGrille['croix_x'][self.__game-1] = 0
			self.__nextN = False
		elif self.__next2 :
			if data.find("X") >= 0 :
				currentGrille['croix_2'][self.__game-1] = 1
			else:
				currentGrille['croix_2'][self.__game-1] = 0
			self.__next2 = False
		elif self.__nextMise and self.__gameOK:
			print "mise =%s" % data
			if data.find("/") >= 0 :
				self.__nextMontant = True
				self.__nextMise = False
		elif self.__nextMontant :
			print "Montant =%s" % data
			# Format the scrapped data
			#print "full dataTmp =-%s-" % dataTmp
			dataTmp = unicode(data, 'utf-8')
			dataTmp.replace(" ","")
			number = True
			endOfStr = False
			sizeStr = len(dataTmp)
			j = 0
			#print "dataTmp =-%s-" % dataTmp
			while not dataTmp[j:j+1].isnumeric() and not endOfStr:
				j+=1
				endOfStr = (j >= sizeStr)
				
			i = j+1
			while not endOfStr and dataTmp[j:i].isnumeric():
				i+=1
				endOfStr = (i >= sizeStr)
			#print "i =%d" % i
			#print "mise =%s" % dataTmp[j:i]
			currentGrille['mise'] = int(dataTmp[j:i])
				# Compute the number of grid
			for i in range(0,self.__game):
				doubleOuTriple = currentGrille['croix_1'][i]+currentGrille['croix_x'][i]+currentGrille['croix_2'][i]
				p1 = currentGrille['mise']*currentGrille['croix_1'][i]/doubleOuTriple
				WSDataFormat.grille['croix_1'][i]+=currentGrille['mise']*currentGrille['croix_1'][i]/doubleOuTriple
				WSDataFormat.grille['croix_x'][i]+=currentGrille['mise']*currentGrille['croix_x'][i]/doubleOuTriple
				WSDataFormat.grille['croix_2'][i]+=currentGrille['mise']*currentGrille['croix_2'][i]/doubleOuTriple
			self.__gameOK = False
			#print "WSDataFormat filled !!!!!!!!!!!"
			#print WSDataFormat.grille
			self.__next2 = False
			self.__nextMontant = False

	#def handle_endtag(self, tag):
		


def open_excel_sheet():
    """ Opens a reference to an Excel WorkBook and Worksheet objects """
    workbook = Workbook()
    worksheet = workbook.add_sheet("Sheet 1")
    return workbook, worksheet

def write_excel_header(worksheet, title_cols):
    """ Write the header line into the worksheet """
    cno = 0
    for title_col in title_cols:
        worksheet.write(0, cno, title_col)
        cno = cno + 1
    return

def write_excel_row(worksheet, rowNumber, columnNumber):
    """ Write a non-header row into the worksheet """
    cno = 0
    for column in columns:
        worksheet.write(lno, cno, column)
        cno = cno + 1
    return

def save_excel_sheet(workbook, output_file_name):
    """ Saves the in-memory WorkBook object into the specified file """
    workbook.save(output_file_name)
    return

