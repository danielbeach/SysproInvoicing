import requests #HTTP module
import os #file sysem module
import xml # xml module
from xml.sax.saxutils import escape #xml escape characters
from lxml import etree #XML parsing
import sys #handle command line arguments (report name)
import datetime
import pyodbc
import smtplib
import csv
from random import randint

jobid = randint(0,1000000)

class GeneralMethods():

	#search between two strings
	def find_between_r( s, first, last ):
		try:
			start = s.rindex( first ) + len( first )
			end = s.rindex( last, start )
			return s[start:end]
		except ValueError:
			return ""
			
	def makefilereplacements():
		replacements = {'&lt;':'<', '&gt;':'>'}#, 'utf-8':'utf-16', 'soap:':''
		lines = []
		with open(rawfile, 'r', encoding='utf-8') as infile:
			for line in infile:
				if line.strip() == '&lt;GlBalanceInventory':
					break
				for line in infile:
					if line.strip() == '</QueryResult></QueryResponse></soap:Body></soap:Envelope>':
						break
					for src, target in replacements.items():
						line = line.replace(src, target)
					lines.append(line)
		with open(newfile, 'w', encoding='utf-8') as outfile:
			for line in lines:
				outfile.write(line)
		infile.close()
		outfile.close()

	#this method will read file into memory
	def readfiletomemory(self):
		with open(self, 'r') as myfile:
			data=myfile.read()
			return data
	
	def sendEmail(self):
		sender = 'noreply@youremail.com'
		receivers = ['helpdesk@youremail.com']

		message = """Subject: SalesOrder Invoicing Error

					Hello, 
					I am the SalesOrder Invoicing program and I have encountered a serious problem.
					Following is the best I know of the error that occured.
					""" + self
		try:
			smtpObj = smtplib.SMTP('insert your smtp server here')
			smtpObj.sendmail(sender, receivers, message)         
		except SMTPException:
			GeneralMethods.sendPythonLogToSQL("Error: unable to send email")
	
	def sendPythonLogToSQL(self):
		try:
			cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER=YOURSERVER;DATABASE=YOURDATABASE;UID=USERID;PWD=PASSWORD') #for test usage
			cursor = cnxn.cursor()
			cursor.execute("INSERT INTO dbo.SalesOrderPythonLog (JobID,Status,LogDateTime) VALUES (?,?,GETDATE())",jobid,self)
			cnxn.commit()
			cnxn.close()
		except Exception as e:
			GeneralMethods.sendEmail("""I am the method sendPythonLogToSQL(), used for inserting appllication actions SQL. I had the following error as far as I can tell....   
										""" + str(e))
			GeneralMethods.sendPythonLogToSQL(str(e))
			quit()
	
	def readinfile(self):
		try:
			if os.path.exists(self) == True: 
				with open(self,'r') as f:
					reader = csv.reader(f)
					listy = list(reader)
			else:
				pass
		except Exception as e:
			GeneralMethods.sendEmail("""I am the method readinfile(), used for reading the csv file from Finance. I had the following error as far as I can tell....   
										""" + str(e))
			GeneralMethods.sendPythonLogToSQL(str(e))
			quit()
		os.remove(self)
		return(listy)
			
class LogInOut(GeneralMethods):

	Operator = 'USERID'
	OperatorPassword = 'PASSWORD'
	
	def __init__(self, Company):
		self.Company = Company
	
	def EnterMordorHere(self):
		company = self
		url="http://yourserver/SYSPROWebServices/utilities.asmx?op=Logon"
		headers = {'content-type': 'text/xml'}
		body = """<?xml version="1.0" encoding="utf-8"?>
			<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
			<soap:Body>
			<Logon xmlns="http://www.syspro.com/ns/utilities/">
			<Operator>""" + self.Operator + """</Operator>
			<OperatorPassword>""" + self.OperatorPassword  + """</OperatorPassword>
			<CompanyId>""" + self.Company +"""</CompanyId>
			<CompanyPassword></CompanyPassword>
			<LanguageCode>ENGLISH_US</LanguageCode>
			<LogLevel>ldNoDebug</LogLevel>
			<EncoreInstance>EncoreInstance_0</EncoreInstance>
			<XmlIn></XmlIn>
			</Logon>
			</soap:Body>
			</soap:Envelope>"""
		try:
			response = requests.post(url,data=body,headers=headers)
			xmltostring = str(response.content)
			GUID = GeneralMethods.find_between_r(xmltostring, "<LogonResult>","</LogonResult>")
			
			if GUID.strip() == "":
				GeneralMethods.sendEmail("""I am the main method that tries to log into Syspro.
					I was unable to log into Syspro Company and get a GUID back. """)
				GeneralMethods.sendPythonLogToSQL("I was unable to log into Syspro Company and get a GUID back")
				quit()
		
		except Exception as e:
			GeneralMethods.sendEmail(str(e))
			GeneralMethods.sendPythonLogToSQL(str(e))
			
		return(GUID.strip())
		
	def ExitMordorHere(self, GUID):
		url="http://yourserver/SYSPROWebServices/utilities.asmx?op=Logoff"
		headers = {'content-type': 'text/xml'}
		body = """<?xml version="1.0" encoding="utf-8"?>
		<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
		<soap:Body>
		<Logoff xmlns="http://www.syspro.com/ns/utilities/">
		<UserId>""" + GUID + """</UserId>
		</Logoff>
		</soap:Body>
		</soap:Envelope>"""
		response = requests.post(url,data=body,headers=headers)
		xmltostring = str(response.content)
		OUT = GeneralMethods.find_between_r(xmltostring, "<LogoffResult>","</LogoffResult>")
		return(OUT.strip())

class Transactions(GeneralMethods):

	def __init__(self, TransactionType):
		self.TransactionType = TransactionType
		
	def reportToBusinessObject(self):
		try:
			if self.TransactionType == "CreateInvoice":
				bizobj = 'SORTIC'
				return bizobj
		except Exception as e:
			GeneralMethods.sendEmail(str(e))
			GeneralMethods.sendPythonLogToSQL(str(e))
	
	@staticmethod
	def sendTransactionRequest(GUID, BO, xmlparam ,xmlin, sorder):
		url="http://yourserver/SYSPROWebServices/transaction.asmx?op=Post"
		headers = {'content-type': 'text/xml','SOAPAction': 'http://www.syspro.com/ns/transaction/Post'}
		body = """<?xml version="1.0" encoding="utf-8"?>
		<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
		<soap:Body>
		<Post xmlns="http://www.syspro.com/ns/transaction/">
		<UserId>""" + GUID + """</UserId>
		<BusinessObject>SORTIC</BusinessObject>
		<XMLParameters>""" + xmlparam + """</XMLParameters>
		<XMLIn>""" + xmlin + """</XMLIn>
		</Post>
		</soap:Body>
		</soap:Envelope>"""
		try:
			response = requests.post(url,data=body,headers=headers)
	
			#look for faults
			if len(GeneralMethods.find_between_r(str(response.content.decode('utf-8')),'<faultstring>','</faultstring>')) > 0:
				GeneralMethods.sendEmail("""I am the method sendTransactionRequest(), i recieved a fault string response from
											the server, meaning my request went through but something unexpected happened that
											made the server unable to read my request and return a fault. I do not consider this
											a fatal error, i will just continue to do this every time I encounter this sales order. 
											I had the following error as far as I can tell....   
										""" + GeneralMethods.find_between_r(response.content.decode('utf-8'),'<faultstring>','</faultstring>'))
				GeneralMethods.sendPythonLogToSQL( str(GeneralMethods.find_between_r(str(response.content.decode('utf-8')),'<faultstring>','</faultstring>')))
				
		except Exception as e:
			GeneralMethods.sendEmail("""I am the method sendTransactionRequest(), used for sending Transaction Requests into
										the Syspro Business Object. I had the following error as far as I can tell....   
										""" + str(e))
			GeneralMethods.sendPythonLogToSQL(str(e))
			quit()
		#return(body)
		return(response.content.decode('utf-8'))
	
	#return a Dictionary of results of SORTIC response parsed from URL encoded XML (which sucks)  FYI. a python Dictionary is a key, value pair list.
	def parseResponseOfSORTIC(self):
		try:
			dict = {} #create empty dictionary 
			ItemsProcessed = GeneralMethods.find_between_r( self, "&lt;ItemsProcessed&gt;", "&lt;/ItemsProcessed&gt;" )
			dict['ItemsProcessed'] = ItemsProcessed
			ItemsInvalid = GeneralMethods.find_between_r( self, "&lt;ItemsInvalid&gt;", "&lt;/ItemsInvalid&gt;" )
			dict['ItemsInvalid'] = ItemsInvalid
			ItemsRejectedWithWarnings = GeneralMethods.find_between_r( self, "&lt;ItemsRejectedWithWarnings&gt;", "&lt;/ItemsRejectedWithWarnings&gt;" )
			dict['ItemsRejectedWithWarnings'] = ItemsRejectedWithWarnings
			ItemsProcessedWithWarnings = GeneralMethods.find_between_r( self, "&lt;ItemsProcessedWithWarnings&gt;", "&lt;/ItemsProcessedWithWarnings&gt;" )
			dict['ItemsProcessedWithWarnings'] = ItemsProcessedWithWarnings
			ErrorNumber = GeneralMethods.find_between_r( self, "&lt;ErrorNumber&gt;", "&lt;/ErrorNumber&gt;" )
			dict['ErrorNumber'] = ErrorNumber
			ErrorDescription = GeneralMethods.find_between_r( self, "&lt;ErrorDescription&gt;", "&lt;/ErrorDescription&gt;" )
			dict['ErrorDescription'] = ErrorDescription
			SentForValidation = "1"
			dict['SentForValidation'] = SentForValidation
			SentForInvoicing = "0"
			dict['SentForInvoicing'] = SentForInvoicing
			InvoiceNumber = GeneralMethods.find_between_r( self, "&lt;InvoiceNumber&gt;", "&lt;/InvoiceNumber&gt;" )
			dict['InvoiceNumber'] = InvoiceNumber
			TrnYear = GeneralMethods.find_between_r( self, "&lt;TrnYear&gt;", "&lt;/TrnYear&gt;" )
			dict['TrnYear'] = TrnYear
			TrnMonth = GeneralMethods.find_between_r( self, "&lt;TrnMonth&gt;", "&lt;/TrnMonth&gt;" )
			dict['TrnMonth'] = TrnMonth
			Register = GeneralMethods.find_between_r( self, "&lt;Register&gt;", "&lt;/Register&gt;" )
			dict['Register'] = Register
			WarningNumber = GeneralMethods.find_between_r( self, "&lt;WarningNumber&gt;", "&lt;/WarningNumber&gt;" )
			dict['WarningNumber'] = WarningNumber
			WarningDescription = GeneralMethods.find_between_r( self, "&lt;WarningDescription&gt;", "&lt;/WarningDescription&gt;" )
			dict['WarningDescription'] = WarningDescription
			Processed = GeneralMethods.find_between_r( self, "&lt;Processed&gt;", "&lt;/Processed&gt;" )
			dict['Processed'] = Processed
			GlYear = GeneralMethods.find_between_r( self, "&lt;GlYear&gt;", "&lt;/GlYear&gt;" )
			dict['GlYear'] = GlYear
			GlPeriod = GeneralMethods.find_between_r( self, "&lt;GlPeriod&gt;", "&lt;/GlPeriod&gt;" )
			dict['GlPeriod'] = GlPeriod
			GlJournal = GeneralMethods.find_between_r( self, "&lt;GlJournal&gt;", "&lt;/GlJournal&gt;" )
			dict['GlJournal'] = GlJournal
		except Exception as e:
			GeneralMethods.sendEmail("""I am the method parseResponseOfSORTIC(), used for parsing the response from Syspro. I had the following error as far as I can tell....   
										""" + str(e))
			GeneralMethods.sendPythonLogToSQL(str(e))
			quit()
		return dict
		
		
	#get document type of sales order from a list of sales orders given by Finance in a csv file.
	def getDocStatus(self):
		try:
			cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER=YOURSERVER;DATABASE=YOURDATABASE;UID=USERID;PWD=PASSWORD') #for production use only
			cursor = cnxn.cursor()
			fullappend = "%" + str(self) + ""
			final = fullappend.replace("['","").replace("']","")
			query = cursor.execute("SELECT SalesOrder,DocumentType FROM dbo.SorMaster WHERE SalesOrder LIKE ?",final)
			SOdict = {}
			for row in cursor.fetchall():
				SOdict.update({row.SalesOrder:row.DocumentType.replace('O','I')})
			cnxn.close()
		except Exception as e:
			GeneralMethods.sendEmail("""I am the method getDocStatus(), used for getting a dictionary of status 8 orders from Syspro along with their document type. I had the following error as far as I can tell....   
										""" + str(e))
			GeneralMethods.sendPythonLogToSQL(str(e))
			quit()
		return SOdict
		
	#this method will create a list of SalesOrders
	@staticmethod
	def getStatus8():
		try:
			cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER=YOURSERVER;DATABASE=YOURDATABASE;UID=USERID;PWD=PASSWORD') #for production use only
			cursor = cnxn.cursor()
			cursor.execute("SELECT SalesOrder,DocumentType FROM dbo.SorMaster WHERE OrderStatus = '8'")
			SOdict = {}
			for row in cursor.fetchall():
				SOdict[row.SalesOrder] = row.DocumentType.replace('O','I')
			cnxn.close()
		except Exception as e:
			GeneralMethods.sendEmail("""I am the method getStatus8(), used for getting a dictionary of status 8 orders from Syspro along with their document type. I had the following error as far as I can tell....   
										""" + str(e))
			GeneralMethods.sendPythonLogToSQL(str(e))
			quit()
		return SOdict
	
	#this method will send a validation result into SQL logging table
	@staticmethod
	def sendValidationToSQL(sorder,DocumentType,ItemsProcessed,ItemsInvalid,ItemsRejectedWithWarnings,ItemsProcessedWithWarnings,ErrorNumber,ErrorDescription,SentForValidation,SentForInvoicing):
		try:
			cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER=YOURSERVER;DATABASE=YOURDATABASE;UID=USERID;PWD=PASSWORD') #for test usage
			cursor = cnxn.cursor()
			cursor.execute("INSERT INTO dbo.SalesOrderInvoiceLogging (SalesOrder,DocumentType,SentForValidationDate,ItemsProcessedForValidation,ItemsInvalidForValidation,ItemsRejectedWithWarningsForValidation,ItemsProcessedWithWarningsForValidation,ErrorNumberForValidation,ErrorDescriptionForValidation,SentForValidation,SentForInvoicing) VALUES (?,?,GETDATE(),?,?,?,?,?,?,?,?)",sorder,DocumentType,ItemsProcessed,ItemsInvalid,ItemsRejectedWithWarnings,ItemsProcessedWithWarnings,ErrorNumber,ErrorDescription,SentForValidation,SentForInvoicing)
			cnxn.commit()
			cnxn.close()
		except Exception as e:
			GeneralMethods.sendEmail("""I am the method sendValidationToSQL(), used for inserting validation results from Syspro into SQL. I had the following error as far as I can tell....   
										""" + str(e))
			GeneralMethods.sendPythonLogToSQL(str(e))
			quit()
		
	#this method will send a invoicing result into SQL logging table
	@staticmethod
	def sendInvoicingToSQL(InvoiceNumber,TrnYear,TrnMonth,Register,WarningNumber,WarningDescription,Processed,GlYear,GlPeriod,GlJournal,ItemsProcessedForInvoicing,ItemsInvalidForInvoicing,ItemsRejectedWithWarningsForInvoicing,ItemsProcessedWithWarningsForInvoicing,ErrorNumberForInvoicing,ErrorDescriptionForInvoicing,sorder):
		try:
			cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER=YOURSERVER;DATABASE=YOURDATABASE;UID=USERID;PWD=PASSWORD') #for test usage
			cursor = cnxn.cursor()
			cursor.execute(""" UPDATE [dbo].[SalesOrderInvoiceLogging]
						SET [SentForInvoicing] = 1,[SentForInvoicingDate] = GETDATE(),[InvoiceNumber] = ?,
						[TrnYear] = ?,[TrnMonth] = ?,[Register] = ?,[WarningNumber] = ?,[WarningDescription] = ?,
						[Processed] = ?,[GlYear] = ?,[GlPeriod] = ?,[GlJournal] = ?,[ItemsProcessedForInvoicing] = ?,
						[ItemsInvalidForInvoicing] = ?,[ItemsRejectedWithWarningsForInvoicing] = ?,
						[ItemsProcessedWithWarningsForInvoicing] = ?,[ErrorNumberForInvoicing] = ?,
						[ErrorDescriptionForInvoicing] = ?
						FROM [dbo].[SalesOrderInvoiceLogging]
						WHERE SalesOrder = ?
							""",InvoiceNumber,TrnYear,TrnMonth,Register,WarningNumber,WarningDescription,Processed,GlYear,GlPeriod,GlJournal,ItemsProcessedForInvoicing,ItemsInvalidForInvoicing,ItemsRejectedWithWarningsForInvoicing,ItemsProcessedWithWarningsForInvoicing,ErrorNumberForInvoicing,ErrorDescriptionForInvoicing,sorder)
			cnxn.commit()
			cnxn.close()
		except Exception as e:
			GeneralMethods.sendEmail("""I am the method sendInvoicingToSQL(), used for inserting invoicing results from Syspro into SQL. I had the following error as far as I can tell....   
										""" + str(e))
			GeneralMethods.sendPythonLogToSQL(str(e))
			quit()
		
		
	#read in xml parameters file to memory and transform for use over HTTP binding protocol. 
	@staticmethod
	def getParams():
		try:
			param_xml = GeneralMethods.readfiletomemory('C:\\YOURPATH\\SORTIC-parameters.xml')
			transformed_xml_parameters = param_xml.replace('<','&lt;').replace('>','&gt;')
			return transformed_xml_parameters
		except Exception as e:
			GeneralMethods.sendEmail("""I am the method getParams(), getting XML parameters for use with Transaction requests. I had the following error as far as I can tell....   
										""" + str(e))
			GeneralMethods.sendPythonLogToSQL(str(e))
			quit()
		
	#method for sending in SalesOrder and getting XMLIN for business ojbect back formatted with salesorder filled in.
	def getXMLin(self):
		try:
			sortic_xml = GeneralMethods.readfiletomemory('C:\\YOURPATH\\SORTIC.xml')
			sortic = sortic_xml.replace('<','&lt;').replace('>','&gt;').replace('xxx',self)
			return sortic
		except Exception as e:
			GeneralMethods.sendEmail("""I am the method getXMLin(), used for getting XML in for use with Transaction requests. I had the following error as far as I can tell....   
										""" + str(e))
			GeneralMethods.sendPythonLogToSQL(str(e))
			quit()
			
	#get invoice number and Terms (looking for CC - Credit Card ) for every sales order. CC Term Invoices have to be sent to Business Object ARSTPY. 
	def getInvoiceWithCCTerms(self):
		try:
			cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER=YOURSERVER;DATABASE=YOURDATABASE;UID=USERID;PWD=PASSWORD') #for test usage
			cursor = cnxn.cursor()
			fullappend = "%" + str(self) + ""
			final = fullappend.replace("['","").replace("']","")
			query = cursor.execute("SELECT t.Invoice,t.Customer FROM dbo.SorMaster sm INNER JOIN dbo.ArTrnSummary t on t.SalesOrder = sm.SalesOrder WHERE sm.OrderStatus = '9' AND t.TermsCode = 'CC' AND sm.SalesOrder LIKE ?",final)
			SOdict = {}
			for row in cursor.fetchall():
				SOdict.update({row.SalesOrder:row.Customer.replace('','')})
			cnxn.close()
		except Exception as e:
			GeneralMethods.sendEmail("""I am the method getInvoiceWithCCTerms(), used for looking up a Invoice with CC Terms from a sales order. I had the following error as far as I can tell....   
										""" + str(e))
			GeneralMethods.sendPythonLogToSQL(str(e))
			quit()
		return SOdict
		
	#method for getting ARSTPY XML in
	def getARSTPYxMLin(GUID,Invoice,Customer):
		try:
			arstpy_xml = GeneralMethods.readfiletomemory('C:\\YOURPATH\\ARSTPY.xml')
			arstpy = arstpy_xml.replace('<','&lt;').replace('>','&gt;').replace('eCig',GUID).replace('PlaceCustomerHere',Customer).replace('PlaceInvoiceHere',Invoice)
			return arstpy
		except Exception as e:
			GeneralMethods.sendEmail("""I am the method getARSTPYxMLin(), used for getting XML in for use with Transaction requests. I had the following error as far as I can tell....   
										""" + str(e))
			GeneralMethods.sendPythonLogToSQL(str(e))
			quit()
	
	@staticmethod
	def sendARSTPYtransactionRequest(GUID, BO ,xmlin):
		url="http://yourserver/SYSPROWebServices/transaction.asmx?op=Post"
		headers = {'content-type': 'text/xml','SOAPAction': 'http://www.syspro.com/ns/transaction/Post'}
		body = """<?xml version="1.0" encoding="utf-8"?>
		<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
		<soap:Body>
		<Post xmlns="http://www.syspro.com/ns/transaction/">
		<UserId>""" + GUID + """</UserId>
		<BusinessObject>ARSTPY</BusinessObject>
		<XMLParameters></XMLParameters>
		<XMLIn>""" + xmlin + """</XMLIn>
		</Post>
		</soap:Body>
		</soap:Envelope>"""
		try:
			response = requests.post(url,data=body,headers=headers)
	
			#look for faults
			if len(GeneralMethods.find_between_r(str(response.content.decode('utf-8')),'<faultstring>','</faultstring>')) > 0:
				GeneralMethods.sendEmail("""I am the method sendARSTPYtransactionRequest(), i recieved a fault string response from
											the server, meaning my request went through but something unexpected happened that
											made the server unable to read my request and return a fault. I do not consider this
											a fatal error, i will just continue to do this every time I encounter this sales order. 
											I had the following error as far as I can tell....   
										""" + GeneralMethods.find_between_r(response.content.decode('utf-8'),'<faultstring>','</faultstring>'))
				GeneralMethods.sendPythonLogToSQL( str(GeneralMethods.find_between_r(str(response.content.decode('utf-8')),'<faultstring>','</faultstring>')))
				
		except Exception as e:
			GeneralMethods.sendEmail("""I am the method sendARSTPYtransactionRequest(), used for sending Transaction Requests into
										the Syspro Business Object. I had the following error as far as I can tell....   
										""" + str(e))
			GeneralMethods.sendPythonLogToSQL(str(e))
			quit()
		#return(body)
		return(response.content.decode('utf-8'))
			