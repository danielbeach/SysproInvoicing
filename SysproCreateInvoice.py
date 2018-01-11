import SysproClass
import datetime
from SysproClass import GeneralMethods,Transactions, LogInOut

#log that the application started to run
try:
	GeneralMethods.sendPythonLogToSQL('I just started running')
except Exception as e:
	GeneralMethods.sendEmail(str(e))
	GeneralMethods.sendPythonLogToSQL("I failed to log that I just started running at... " +str(e))

#create instance of Company i want to log into for use with login and out
try:
	Company = LogInOut('YOURCOMPANY')
except Exception as e:
	GeneralMethods.sendEmail(str(e))
	GeneralMethods.sendPythonLogToSQL("I failed to Log into Syspro with the following error "+str(e))

#log that the application created an instance to log into syspro
try:
	GeneralMethods.sendPythonLogToSQL('I just created an instance of myself')
except Exception as e:
	GeneralMethods.sendEmail(str(e))
	GeneralMethods.sendPythonLogToSQL("I failed to logt that I just created an instance of myself... " +str(e))

# create a login profile instance
GUID = Company.EnterMordorHere()

#log that the application logged into syspro
try:
	GeneralMethods.sendPythonLogToSQL('I just logged into Syspro sucessfully')
except Exception as e:
	GeneralMethods.sendEmail(str(e))
	GeneralMethods.sendPythonLogToSQL("I failed to log that I just logged into Syspro sucessfully... " +str(e))

#create an instance of the transaction type we are doing, in this case Create Invoice
TransactionType = Transactions('CreateInvoice')

#log that the application created instance for Transactions
try:
	GeneralMethods.sendPythonLogToSQL('I just created an instance for Transactions')
except Exception as e:
	GeneralMethods.sendEmail(str(e))
	GeneralMethods.sendPythonLogToSQL("I failed to log that I just created an instance for Transactions... " +str(e))

#return Business Object we will be working with
BO = TransactionType.reportToBusinessObject()

#log that the application returned a Business Object to work with.
try:
	GeneralMethods.sendPythonLogToSQL('I just returned a Business Object i need to work with')
except Exception as e:
	GeneralMethods.sendEmail(str(e))
	GeneralMethods.sendPythonLogToSQL("I failed to log that I just returned a Business Object i need to work with... " +str(e))


#create empty dictonary
SOdict = {}

#return dictonary of SalesOrders in status 8 and document types
SOdict = Transactions.getStatus8()


#log that the application got all status 8 orders and doucment types
try:
	GeneralMethods.sendPythonLogToSQL('I got status 8 orders and doc types')
except Exception as e:
	GeneralMethods.sendEmail(str(e))
	GeneralMethods.sendPythonLogToSQL("I failed to log that I got status 8 orders and doc types... " +str(e))

#create a blank list that will hold invalid SalesOrders that need to be removed from the main
#SOList that gets used to send things through final invoicing. You can't remove the item while
#instead the list because it will screw up python as it already made reference to the list pointers
#when we started the FOR loop
removeSOList = []

#for each SalesOrder in dictonary send through validation methods
try:
	for key,value in SOdict.items():
	
		sorder = key
		DocumentType = value
	
		#method for getting xmlin with salesorder injected into xml
		xmlin = Transactions.getXMLin(sorder)
		
		#log that the application got XMLin ready to send
		try:
			GeneralMethods.sendPythonLogToSQL('I have prepared the XMLin for sales order ' + sorder)
		except Exception as e:
			GeneralMethods.sendEmail(str(e))
			GeneralMethods.sendPythonLogToSQL("I failed to log that I have prepared the XMLin... " +str(e))

		#method for getting params, sets document type and validation only or not
		xmlparam = Transactions.getParams().replace('IsThisValidation?','Y').replace('WhatOrderTypeAmI',DocumentType)
		
		#log that the application got XMLparms ready to send
		try:
			GeneralMethods.sendPythonLogToSQL('I have prepared the XMLparms')
		except Exception as e:
			GeneralMethods.sendEmail(str(e))
			file.write("I failed to log that I have prepared the XMLparms... " +str(e))
		
		#send off sales order for validation only
		response = Transactions.sendTransactionRequest(GUID, BO, xmlparam ,xmlin, sorder)
		
		#log that the application send off the request to the API
		try:
			GeneralMethods.sendPythonLogToSQL('I have sent the API request  ')
		except Exception as e:
			GeneralMethods.sendEmail(str(e))
			GeneralMethods.sendPythonLogToSQL("I failed to log that I have I have send the API request... " +str(e))
			
		#pull out actual response values from XML crud
		responseparsed = Transactions.parseResponseOfSORTIC(response)
		
		#log that the application recieved a response and parased it
		try:
			GeneralMethods.sendPythonLogToSQL('I have recieved a API response and parsed it  ')
		except Exception as e:
			GeneralMethods.sendEmail(str(e))
			GeneralMethods.sendPythonLogToSQL("I failed to log that I have recieved a API response and parsed it... " +str(e))
	
		#write dict values to vars for use in SQL Statement
		ItemsProcessedForValidation = responseparsed['ItemsProcessed']
		ItemsInvalidForValidation = responseparsed['ItemsInvalid']
		ItemsRejectedWithWarningsForValidation = responseparsed['ItemsRejectedWithWarnings']
		ItemsProcessedWithWarningsForValidation = responseparsed['ItemsProcessedWithWarnings']
		ErrorNumberForValidation = responseparsed['ErrorNumber']
		ErrorDescriptionForValidation = responseparsed['ErrorDescription']
		SentForValidation = responseparsed['SentForValidation']
		SentForInvoicing = responseparsed['SentForInvoicing']
	
		#do SQL logging here of validation results
		Transactions.sendValidationToSQL(sorder,DocumentType,ItemsProcessedForValidation,ItemsInvalidForValidation,ItemsRejectedWithWarningsForValidation,ItemsProcessedWithWarningsForValidation,ErrorNumberForValidation,ErrorDescriptionForValidation,SentForValidation,SentForInvoicing)

		
		#log that the application sent validation results to SQL
		try:
			GeneralMethods.sendPythonLogToSQL('I have send the validation results to SQL  ')
		except Exception as e:
			GeneralMethods.sendEmail(str(e))
			GeneralMethods.sendPythonLogToSQL("I failed to log that I have send the validation results to SQL... " +str(e))
	
		#need to remove those items with errors from the SOList list , as in they don't get sent through for invoicing.
		try:
			if ItemsInvalidForValidation != "000000":
				removeSOList.append(sorder)
				print('i removed someone')
		except Exception as e:
			GeneralMethods.sendEmail(str(e))
			GeneralMethods.sendPythonLogToSQL("i am the method to remove sales orders from the list that had errors in validation... " +str(e))
		
		#log that the application removed items if needed from validation list
		try:
			GeneralMethods.sendPythonLogToSQL('I have removed salesorders(s) if nessesary from the list if they had errors  ')
		except Exception as e:
			GeneralMethods.sendEmail(str(e))
			GeneralMethods.sendPythonLogToSQL("I failed to log that I have removed salesorders(s) if nessesary from the list if they had errors... " +str(e))
			
		print(sorder + " was validated")

#this is the exception catch for the over all FOR LOOP
except Exception as e:
	GeneralMethods.sendEmail(str(e))
	GeneralMethods.sendPythonLogToSQL("I failed to enter my Loop to send sales orders through validation... " +str(e))

#log that the application finished sending all orders through validation
try:
	GeneralMethods.sendPythonLogToSQL('I finished sending all orders through validation ')
except Exception as e:
	GeneralMethods.sendEmail(str(e))
	GeneralMethods.sendPythonLogToSQL("I failed to log that I finished sending all orders through validation... " +str(e))


#create new list for invoicing without SalesOrders that had validation errors
SOdictInvoicing = {key:value for key,value in SOdict.items() if key not in removeSOList}

for key,value in SOdictInvoicing.items():
	sorder = key
	DocumentType = value

	#since we have done validation and removed SalesOrders with errors, we now enter the actual invoicing stage
	xmlparam = Transactions.getParams().replace('IsThisValidation?','N').replace('WhatOrderTypeAmI',DocumentType)
		#method for getting xmlin with salesorder injected into xml
	xmlin = Transactions.getXMLin(sorder)
	
	#send off sales order for invoicing with non validation
	response = Transactions.sendTransactionRequest(GUID, BO, xmlparam ,xmlin, sorder)
	#print(response)
	
	#pull out actual response values from XML crud
	responseparsed = Transactions.parseResponseOfSORTIC(response)
	
	#write dict values to vars for use in SQL Statement
	InvoiceNumber = responseparsed['InvoiceNumber']
	TrnYear = responseparsed['TrnYear']
	TrnMonth = responseparsed['TrnMonth']
	Register = responseparsed['Register']
	WarningNumber = responseparsed['WarningNumber']
	WarningDescription = responseparsed['WarningDescription']
	Processed = responseparsed['Processed']
	GlYear = responseparsed['GlYear']
	GlPeriod = responseparsed['GlPeriod']
	GlJournal = responseparsed['GlJournal']
	ItemsProcessedForInvoicing = responseparsed['ItemsProcessed']
	ItemsInvalidForInvoicing = responseparsed['ItemsInvalid']
	ItemsRejectedWithWarningsForInvoicing = responseparsed['ItemsRejectedWithWarnings']
	ItemsProcessedWithWarningsForInvoicing = responseparsed['ItemsProcessedWithWarnings']
	ErrorNumberForInvoicing = responseparsed['ErrorNumber']
	ErrorDescriptionForInvoicing = responseparsed['ErrorDescription']
	
	#do SQL logging here of validation results
	Transactions.sendInvoicingToSQL(InvoiceNumber,TrnYear,TrnMonth,Register,WarningNumber,WarningDescription,Processed,GlYear,GlPeriod,GlJournal,ItemsProcessedForInvoicing,ItemsInvalidForInvoicing,ItemsRejectedWithWarningsForInvoicing,ItemsProcessedWithWarningsForInvoicing,ErrorNumberForInvoicing,ErrorDescriptionForInvoicing,sorder)
	print(sorder + " was invoiced")


#after invoicing is done we need to go back through and find Invoices created for those with CC Terms (Credit Card) so we can send through ARSTPY to capture payment.
for key,value in SOdictInvoicing.items():
	ccInvoiceTerms = {}
	sorder = key
	MeResult = Transactions.getInvoiceWithCCTerms(sorder) #should return invoice
	ccInvoiceTerms.update(MeResult) #stick results in Dict
	for key, value in ccInvoiceTerms.items(): #get results for inserting into ARSTPY
		Invoice = key
		Customer = value
		#return ARSTPY xmlin sending up to Syspro
		ARSTPY = Transactions.getARSTPYxMLin(GUID,Invoice,Customer)
		#send results up for processing credit card transaction.
		response = sendARSTPYtransactionRequest(GUID, BO, ARSTPY)
		#parse response.

try:
	print(Company.ExitMordorHere(GUID))
except Exception as e:
	SysproClass.sendEmail("I am the method ExitMordorHere() that tried to log out of Syspro.. i had the following errors..... " + str(e))
	GeneralMethods.sendPythonLogToSQL("I am the method ExitMordorHere() i failed... " +str(e))

#log that the application finished sending all orders through validation
try:
	GeneralMethods.sendPythonLogToSQL('I logged out of Syspro ')
except Exception as e:
	GeneralMethods.sendEmail(str(e))
	GeneralMethods.sendPythonLogToSQL("I failed to log that I logged out of Syspro... " +str(e))

