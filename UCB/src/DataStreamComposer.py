#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from datetime import datetime
# from itertools import zip_longest
from com.sun.star.sheet import CellFlags as cf # 定数
# from com.sun.star.beans import Property  # Struct
# from com.sun.star.ucb import OpenCommandArgument2  # Struct
# from com.sun.star.ucb import OpenMode  # 定数
# from com.sun.star.ucb import Command  # Struct
def macro(documentevent=None):  # 引数は文書のイベント駆動用。 
	doc = XSCRIPTCONTEXT.getDocument() if documentevent is None else documentevent.Source  # ドキュメントのモデルを取得。
	sheet = getNewSheet(doc, "ChidrenRetriever")  # 新規シートの取得。	
	sheet[0, 0].setString("DataStreamComposer - sets the data stream of a document resource.")
	sheet[0, 1].setString("The data stream is obtained from another (the source) document resource before.")
	t = datetime.now().isoformat()
	s = t.split(".")[0]
	print(s.replace("-", "").replace(":", ""))
# 	print(t.isoformat().translate(table))
	
	
	
# 	with open("data.txt", "w", encoding="utf-8") as f:
# 		f.write("sample sample sample sample sample sample sample sample EOF")	
# 	with open("data.txt", "w", encoding="utf-8") as f:
# 		f.write("sample sample sample sample sample sample sample sample EOF")
# 	
# 	
# 	
# 	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
# 	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
# 	workdir = "" 
# 	contenturl
# 	
# 	srcfile = "$(inst)/sdk/examples/DevelopersGuide/UCB/data/data.txt"  # SDKをインストールしていないときはdata.txtへのパスが必要。
# 	pathsubstservice = smgr.createInstanceWithContext("com.sun.star.comp.framework.PathSubstitution", ctx)
# 	srcurl = pathsubstservice.substituteVariables(srcfile, True)  # $(inst)を変換する。fileurlが返ってくる。



# 	contenturl = "file:///"
# 	propnames = ["Title", "IsDocument"]
# 	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
# 	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
# 	ucb =  smgr.createInstanceWithContext("com.sun.star.ucb.UniversalContentBroker", ctx)
# 	content = ucb.queryContent(ucb.createContentIdentifier(contenturl))
# 	if content:
# 		props = []
# 		[props.append(Property(Name=propname, Handle=-1)) for propname in propnames]
# 		arg = OpenCommandArgument2(Mode=OpenMode.ALL, Priority=32768, Properties=props)
# 		command = Command(Name="open", Handle=-1, Argument=arg)
# 		dynamicresultset = content.execute(command, 0, None)
# 		resultset = dynamicresultset.getStaticResultSet()
# 		outputs = []
# 		flg = resultset.first()
# 		while flg:
# 			propsvalues = []
# 			propsvalues.append(resultset.queryContentIdentifierString())
# 			for i in range(1, len(props)+1):
# 				propvalue = resultset.getObject(i, None)
# 				if isinstance(propvalue, bool):  # ブール型の時
# 					propvalue = str(propvalue)  # 文字列に変換する。シートに出力するため。
# 				elif propvalue is None:
# 					propvalue = "[ Property not found ]"
# 				propsvalues.append(propvalue)
# 			outputs.append(propsvalues)
# 			flg = resultset.next()
# 		sheet = getNewSheet(doc, "ChidrenRetriever")  # 新規シートの取得。	
# 		datarows = [("URL:", *["{}:".format(p) for p in propnames])]
# 		datarows.extend(outputs)
# 		rowsToSheet(sheet[2, 0], datarows)  # datarowsをシートに書き出し。代入範囲の列幅も最適化される。
# 		sheet[0, 0].setString("ChildrenRetriever - obtains the children of a folder resource")
# 		sheet[1, 0].setString("Children of resource {}".format(contenturl))		
# 		controller = doc.getCurrentController()  # コントローラの取得。
# 		controller.setActiveSheet(sheet)  # シートをアクティブにする。
# def rowsToSheet(cellrange, datarows):  # 引数のセル範囲を左上端にして一括書き込みして列幅を最適化する。datarowsはタプルのタプル。
# 	datarows = tuple(zip(*zip_longest(*datarows, fillvalue="")))  # 一番長い行の長さに合わせて空文字を代入。
# 	sheet = cellrange.getSpreadsheet()  # セル範囲のあるシートを取得。
# 	cellcursor = sheet.createCursorByRange(cellrange)  # セル範囲のセルカーサーを取得。
# 	cellcursor.collapseToSize(len(datarows[0]), len(datarows))  # (列、行)で指定。セルカーサーの範囲をdatarowsに合せる。
# 	cellcursor.setDataArray(datarows)  # セルカーサーにdatarowsを代入。代入できるのは整数(int、ただしboolを除く)か文字列のみ。
# 	cellcursor.getColumns().setPropertyValue("OptimalWidth", True)  # セルカーサーのセル範囲の列幅を最適化する。行幅は限定サれない。		
def getNewSheet(doc, sheetname):  # docに名前sheetnameのシートを返す。sheetnameがすでにあれば連番名を使う。
	cellflags = cf.VALUE+cf.DATETIME+cf.STRING+cf.ANNOTATION+cf.FORMULA+cf.HARDATTR+cf.STYLES
	sheets = doc.getSheets()  # シートコレクションを取得。
	c = 1  # 連番名の最初の番号。
	newname = sheetname
	while newname in sheets: # 同名のシートがあるとき。sheets[sheetname]ではFalseのときKeyErrorになる。
		if not sheets[newname].queryContentCells(cellflags):  # シートが未使用のとき
			return sheets[sheetname]  # 未使用の同名シートを返す。
		newname = "{}{}".format(sheetname, c)  # 連番名を作成。
		c += 1	
	sheets.insertNewByName(newname, len(sheets))   # 新しいシートを挿入。同名のシートがあるとRuntimeExceptionがでる。
	if "Sheet1" in sheets:  # デフォルトシートがあるとき。
		if not sheets["Sheet1"].queryContentCells(cellflags):  # シートが未使用のとき
			del sheets["Sheet1"]  # シートを削除する。
	return sheets[newname]
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。		
if __name__ == "__main__":  # オートメーションで実行するとき
	def automation():  # オートメーションのためにglobalに出すのはこの関数のみにする。
		import officehelper
		from functools import wraps
		import sys
		from com.sun.star.beans import PropertyValue  # Struct
		from com.sun.star.script.provider import XScriptContext  
		def connectOffice(func):  # funcの前後でOffice接続の処理
			@wraps(func)
			def wrapper():  # LibreOfficeをバックグラウンドで起動してコンポーネントテクストとサービスマネジャーを取得する。
				try:
					ctx = officehelper.bootstrap()  # コンポーネントコンテクストの取得。
				except:
					print("Could not establish a connection with a running office.", file=sys.stderr)
					sys.exit()
				print("Connected to a running office ...")
				smgr = ctx.getServiceManager()  # サービスマネジャーの取得。
				print("Using {} {}".format(*_getLOVersion(ctx, smgr)))  # LibreOfficeのバージョンを出力。
				return func(ctx, smgr)  # 引数の関数の実行。
			def _getLOVersion(ctx, smgr):  # LibreOfficeの名前とバージョンを返す。
				cp = smgr.createInstanceWithContext('com.sun.star.configuration.ConfigurationProvider', ctx)
				node = PropertyValue(Name = 'nodepath', Value = 'org.openoffice.Setup/Product' )  # share/registry/main.xcd内のノードパス。
				ca = cp.createInstanceWithArguments('com.sun.star.configuration.ConfigurationAccess', (node,))
				return ca.getPropertyValues(('ooName', 'ooSetupVersion'))  # LibreOfficeの名前とバージョンをタプルで返す。
			return wrapper
		@connectOffice  # createXSCRIPTCONTEXTの引数にctxとsmgrを渡すデコレータ。
		def createXSCRIPTCONTEXT(ctx, smgr):  # XSCRIPTCONTEXTを生成。
			class ScriptContext(unohelper.Base, XScriptContext):
				def __init__(self, ctx):
					self.ctx = ctx
				def getComponentContext(self):
					return self.ctx
				def getDesktop(self):
					return ctx.getByName('/singletons/com.sun.star.frame.theDesktop')  # com.sun.star.frame.Desktopはdeprecatedになっている。
				def getDocument(self):
					return self.getDesktop().getCurrentComponent()
			return ScriptContext(ctx)  
		XSCRIPTCONTEXT = createXSCRIPTCONTEXT()  # XSCRIPTCONTEXTの取得。
		doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得。
		doctype = "scalc", "com.sun.star.sheet.SpreadsheetDocument"  # Calcドキュメントを開くとき。
# 		doctype = "swriter", "com.sun.star.text.TextDocument"  # Writerドキュメントを開くとき。
		if (doc is None) or (not doc.supportsService(doctype[1])):  # ドキュメントが取得できなかった時またはCalcドキュメントではない時
			XSCRIPTCONTEXT.getDesktop().loadComponentFromURL("private:factory/{}".format(doctype[0]), "_blank", 0, ())  # ドキュメントを開く。ここでdocに代入してもドキュメントが開く前にmacro()が呼ばれてしまう。
		flg = True
		while flg:
			doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得。
			if doc is not None:
				flg = (not doc.supportsService(doctype[1]))  # ドキュメントタイプが確認できたらwhileを抜ける。
		return XSCRIPTCONTEXT
	XSCRIPTCONTEXT = automation()  # XSCRIPTCONTEXTを取得。	
	macro()  # マクロの実行。