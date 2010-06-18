# -*- coding: iso-8859-1 -*-
#
# pyUnoServer - version 2.0
from SimpleXMLRPCServer import SimpleXMLRPCServer, SimpleXMLRPCRequestHandler
import os
import uno
import sys
import time
import logging
import urllib
 
class PyUNOServer(SimpleXMLRPCServer):
 
	def _dispatch(self, method, params):
 
		try:
			func = getattr(self, '' + method)
		except:
			self.logger.error("El Metodo '"+method+"' no esta soportado!")
			raise Exception('method "%s" is not supported' % method)
		else:
			try:
				ret = func(*params)
				self.logger.info("Llamada: %s%s | Retorno: %s" % (method, params, ret))
				return ret
			except:
				self.logger.info("Llamada: %s%s" % (method, params))
				self.logger.error("Error: %s:%s" % sys.exc_info()[:2])
 
	def init(self):
		"""
		This method	starts an Openoffice session and initializes the XMLRPC server
		"""
 
		# Generate some useful constants
		self.EMPTY = uno.Enum("com.sun.star.table.CellContentType", "EMPTY")
		self.TEXT = uno.Enum("com.sun.star.table.CellContentType", "TEXT")
		self.FORMULA = uno.Enum("com.sun.star.table.CellContentType", "FORMULA")
		self.VALUE = uno.Enum("com.sun.star.table.CellContentType", "VALUE")
 
		# Generate a Logger
		self.logger = logging.getLogger('pyUnoServer')
		hdlr = logging.FileHandler('./pyUnoServer.log')
		formatter = logging.Formatter('%(asctime)s %(levelname)s -- %(message)s')
		hdlr.setFormatter(formatter)
		self.logger.addHandler(hdlr) 
		self.logger.setLevel(logging.INFO)
 
		self.sessions = []
		self.sessionSerial = 0
 
		# Start Calc and grep the PID
		self.PID = os.spawnlp(os.P_NOWAIT,"/usr/bin/oocalc","","-accept=socket,host=localhost,port=2002;urp;")
 
		self.logger.info("OOCalc Iniciado con el PID %d" % self.PID)
 
		while 1:
			try: 
				localContext = uno.getComponentContext()
				resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext )
				smgr = resolver.resolve( "uno:socket,host=localhost,port=2002;urp;StarOffice.ServiceManager" )
				remoteContext = smgr.getPropertyValue( "DefaultContext" )
				self.desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",remoteContext)
				self.logger.info("Sistema conectado al socket de UNO")
				break
			except:
				time.sleep(0.5)
 
	def openSession(self, session):
		"""
		This method opens a new session.
		Sample usage from PHP:
		 - $sessionString is your a session ID provided by your program):
 
			$openSession = array( new XML_RPC_VALUE($sessionString, "string"));
			$sessionID = queryXMLRPC($client, "openSession", $openSession);
			return $sessionID;
		"""
 
		for i,x in enumerate(self.sessions):
			if x[0] == session:
				currentSerial = i
				break
		else:
			self.sessions.append([session,[]])
			currentSerial = len(self.sessions)-1
 
		return currentSerial
 
 
	def getSessions(self):
		return self.sessions
 
	def openBook(self, session, bookPath):
		"""
		This method opens a new Openoffice spreadsheet book.
		Sample usage from PHP:
		 - $sessionID is the session's ID returned by openSession()
		 - $path is the path to the book in the server's filesystem
 
		$openBookQuery = array( new XML_RPC_VALUE($sessionID, "int"), new XML_RPC_VALUE($path,"string") );
		$bookID = queryXMLRPC($client, "openBook", $openBookQuery);
		return $bookID;
		"""
 
		currentSerial = session
		bookPath = bookPath.strip()
		print bookPath
 
 
		exists = os.path.exists(bookPath)
		if exists == False:
			raise Exception('El libro %s no existe', bookPath)
		else:
			self.logger.info("se ha encontrado el libro!")
 
		currentBook = -1
 
		for i, book in enumerate(self.sessions[currentSerial][1]):
			try:
 
				bookname = book[0]
				bookhandler = book[1]
				bookmodtime = book[2]
 
				if bookname == bookPath:
					modtime = os.path.getmtime(bookPath)
					if modtime > bookmodtime:
						book[1].dispose()
						self.sessions[currentSerial][1].remove(book)
					else:
						currentBook = i
					break
			except:
				self.logger.error("El libro "+currentBook+" no existe!")
 
 
		if currentBook < 0:
			handler = self.desktop.loadComponentFromURL( "file://" + bookPath ,"_blank",0,())
 
			if handler == None:
				raise Exception('El libro no existe o tiene caracteres extraños...')
 
			modtime = os.path.getmtime(bookPath)
 
			estruct = [bookPath,handler,modtime]
			self.sessions[currentSerial][1].append(estruct)
			currentBook = len(self.sessions[currentSerial][1])-1
 
		return currentBook
 
	def closeBook(self, sessionID, bookID):
 
		books = self.sessions[sessionID][1]
 
		# Close the document. It doesn't look very elegant but thats the way is documented.
		books[bookID][1].dispose()
		self.sessions[sessionID][1].remove(books[bookID])
 
		return 1
 
 
	def getBookSheets(self, sessionID, bookID):
		"""
		This method return an array with the book's worksheets.
		Sample usage from PHP:
		 - $sessionID is the session's ID returned by openSession()
		 - $bookID is the book's ID returned by openBook()
 
		$sheetsQuery = array( new XML_RPC_VALUE($sessionID, "int"), new XML_RPC_VALUE($bookID, "int") );
		$sheets = queryXMLRPC($client, "getBookSheets" , $sheetsQuery);
		return $sheets;
		"""
 
		sheets = self.sessions[sessionID][1][bookID][1].getSheets().createEnumeration()
 
		container = []
		while sheets.hasMoreElements():
			currentSheet = sheets.nextElement()
			container.append(currentSheet.getName().encode("UTF-8"))
 
		return container
 
	def getCellValue (self,sheet,x,y):
 
		cell = sheet.getCellByPosition(x,y)
		cellType  = cell.getType()
 
		if cellType == self.TEXT :
			data = cell.getString().encode("UTF-8")
			if data == None:
				return ""
			return data.encode("UTF-8")
 
		if cellType == self.FORMULA or cellType == self.VALUE :
			data = cell.getValue()
			if data == None:
				return 0
			return cell.getValue()
 
		if cellType == self.EMPTY :
			return ""
 
 
	def getCell(self, session, book, sheet, x, y):
		"""
		This method returns the content of a cell.
		Sample usage from PHP:
		 - $sessionID is the session's ID returned by openSession()
		 - $bookID is the book's ID returned by openBook()
		 - $sheet is the sheet's ID as returned by getBookSheets()
 
		$query = array(	
						new XML_RPC_Value($sessionID, "int"), 
						new XML_RPC_VALUE($bookID, "int"), 
						new XML_RPC_VALUE($sheet, "int"), 
						new XML_RPC_Value($x, "int"), 
						new XML_RPC_Value($y, "string")
					);
 
		$result = queryXMLRPC($client, "getCell", $query);
		return $result;
		"""
 
		sheetref = self.sessions[session][1][book][1].getSheets().getByIndex(sheet)
		valor = self.getCellValue(sheetref,x,y)
		return valor
 
	def setCell(self,session,book,sheet,x,y,value):
		"""
		This method sets the content of a cell.
		Sample usage from PHP:
		 - $sessionID is the session's ID returned by openSession()
		 - $bookID is the book's ID returned by openBook()
		 - $sheet is the sheet's ID as returned by getBookSheets()
 
		$query = array(
						new XML_RPC_Value($sessionID, "int"), 
						new XML_RPC_VALUE($bookID, "int"), 
						new XML_RPC_VALUE($sheet, "int"), 
						new XML_RPC_Value($x, "int"), 
						new XML_RPC_Value($y, "int"),
						new XML_RPC_Value($value, "string")
					);
 
		$result = queryXMLRPC($client, "setCell", $query);
		return $result;
		"""
 
		bookObject = self.sessions[session][1][book][1]
		cell = bookObject.getSheets().getByIndex(sheet).getCellByPosition(x,y)
		cell.setValue(value)
		return 1
 
	#what will change? mmm...
	def massiveSetCell(self, data):
			for valores in data:
				session,sheetid,x,y,valor = valores.split("|")
				code = self.setCell(self.trim(session),self.trim(sheetid),self.trim(x),self.trim(y),valor)
				if code < 0:
					return code
			return 1
 
	def getSheetPreview(self,session,book,sheet):
 
		sheet = self.sessions[session][1][book][1].getSheets().getByIndex(sheet)
		range = sheet.getCellRangeByPosition(0,0,8,23).getDataArray()
		return range
 
 
	def trim(self, cadena):
		return cadena.strip()
 
server = PyUNOServer(("localhost", 8000), allow_none=True)
 
server.init()
server.logger.info("Servicio Inicializado...")
 
try:
	server.serve_forever()
except KeyboardInterrupt:
	server.logger.info("Cerrando el socket...")
	server.server_close()
	os.kill(server.PID,1)
	server.logger.info("Socket cerrado.")
