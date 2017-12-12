#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import sys
from itertools import zip_longest
from com.sun.star.sheet import CellFlags as cf # 定数
from com.sun.star.beans import Property  # Struct
from com.sun.star.ucb import OpenCommandArgument2  # Struct
from com.sun.star.ucb import OpenMode  # 定数
from com.sun.star.ucb import Command  # Struct
from com.sun.star.ui.dialogs import TemplateDescription  # 定数
from com.sun.star.util import URL  # Struct
from com.sun.star.ui.dialogs import ExecutableDialogResults  # 定数
def macro(documentevent=None):  # 引数は文書のイベント駆動用。 
	doc = XSCRIPTCONTEXT.getDocument() if documentevent is None else documentevent.Source  # ドキュメントのモデルを取得。 
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
	simplefileaccess = smgr.createInstanceWithContext("com.sun.star.ucb.SimpleFileAccess", ctx)  # SimpleFileAccess
	doc_fileurl = doc.getURL()  # データを出力するCalcドキュメントのfileurlを取得。保存していないドキュメントでは空の文字列が返る。
	if doc_fileurl:  # fileurlが取得できた時。
		url = URL(Complete=doc_fileurl)  # URL StructのCompleteアトリビュートにfileurlを代入。
		urltransformer = smgr.createInstanceWithContext("com.sun.star.util.URLTransformer", ctx)  # URLTransformer
		dummy, url = urltransformer.parseStrict(url)  # URLTransformerでURL Structの他のアトリビュートに値を設定。
		templateurl = "".join((url.Protocol, url.Path))  # データを出力するCalcドキュメントのあるフォルダのfileurlを取得。
	else:  # 保存していないCalcドキュメントのとき。
		templateurl = ctx.getByName('/singletons/com.sun.star.util.thePathSettings').getPropertyValue("Work")  # ホームフォルダを取得。
	title = "Select an office document to analyze"  # ファイル選択ダイアログのタイトル。
	templatedescription = TemplateDescription.FILEOPEN_SIMPLE  # ファイル選択ダイアログの種類。
	filters = {'Writer Document': 'odt', 'Calc Document': 'ods'}  # 表示フィルターの辞書。
	filterall = "All Document Files"  # デフォルト表示フィルター名。
	filters[filterall] = ";".join(filters.values())  # まとめたフィルターを辞書に追加。	
	kwargs = {"TemplateDescription": templatedescription, "setTitle": title, "setDisplayDirectory": templateurl, "setCurrentFilter": filterall, "appendFilter": filters}
	filepicker = createFilePicker(ctx, smgr, kwargs)  # ファイル選択ダイアログを取得。		
	if not filepicker.execute()==ExecutableDialogResults.OK:  # ファイル保存ダイアログを表示する。
		return  # ファイルが選択されなかった時は終了。
	fileurl = filepicker.getFiles()[0]  # ダイアログで選択されたファイルのパスを取得。fileurlのタプルで返ってくるので先頭の要素を取得。
	urireferencefactory = smgr.createInstanceWithContext("com.sun.star.uri.UriReferenceFactory", ctx)  # UriReferenceFactory
	urireference = urireferencefactory.parse(fileurl)  # ドキュメントのUriReferenceを取得。
	vndsunstarpkgurlreferencefactory = smgr.createInstanceWithContext("com.sun.star.uri.VndSunStarPkgUrlReferenceFactory", ctx)  # VndSunStarPkgUrlReferenceFactory
	vndsunstarpkgurlreference = vndsunstarpkgurlreferencefactory.createVndSunStarPkgUrlReference(urireference)  # ドキュメントのvnd.sun.star.pkgプロトコールにUriReferenceを変換。
	contenturl = vndsunstarpkgurlreference.getUriReference()  # UriReferenceから文字列のURIを取得。
	propnames = "Title", "IsDocument", "IsFolder", "ContentType"  # 取得するプロパティのタプル。TitleとIsFolderは必須。
	ucb =  smgr.createInstanceWithContext("com.sun.star.ucb.UniversalContentBroker", ctx)  # UniversalContentBroker
	props = [Property(Name=propname, Handle=-1) for propname in propnames]
	arg = OpenCommandArgument2(Mode=OpenMode.ALL, Priority=32768, Properties=props)
	command = Command(Name="open", Handle=-1, Argument=arg)  # NameとArgumentでexecute()の戻り値の型が決まる。
	outputs = []  # 出力行をいれるリスト。
	getProps = createGetProps(simplefileaccess, propnames, ucb, command, outputs)  # 関数getPropsを取得。
	getProps(contenturl)  # outputsに結果の行が入る。
	sheet = getNewSheet(doc, "ChidrenRetriever")  # 新規シートの取得。	
	datarows = [("URL:", *["{}:".format(p) for p in propnames])]
	datarows.extend(outputs)
	rowsToSheet(sheet[2, 0], datarows)  # datarowsをシートに書き出し。代入範囲の列幅も最適化される。
	sheet[0, 0].setString("ChildrenRetriever - obtains the children of a folder resource")
	sheet[1, 0].setString("Children of resource {}".format(contenturl))		
	controller = doc.getCurrentController()  # コントローラの取得。
	controller.setActiveSheet(sheet)  # シートをアクティブにする。
def createGetProps(simplefileaccess, propnames, ucb, command, outputs):
	c = len(propnames) + 1
	isfolderindex = propnames.index("IsFolder") + 1 
	titleindex = propnames.index("Title") + 1
	def getProps(contenturl):
		content = ucb.queryContent(ucb.createContentIdentifier(contenturl))  # fileurlを元にFile Contentsを取得。
		dynamicresultset = content.execute(command, 0, None)  # XDynamicResultSet型が返る。
		resultset = dynamicresultset.getStaticResultSet()
		flg = resultset.first()
		flg2 = False
		while flg:
			url = resultset.queryContentIdentifierString()
			if simplefileaccess.exists(url):
				propsvalues = [url]
				for i in range(1, c):
					propvalue = resultset.getObject(i, None)
					if isinstance(propvalue, bool):  # ブール型の時
						if i==isfolderindex:
							flg2 = propvalue
						propvalue = str(propvalue)  # 文字列に変換する。シートに出力するため。
					elif propvalue is None:
						propvalue = "[ Property not found ]"
					propsvalues.append(propvalue)
				outputs.append(propsvalues)
				if flg2:
					getProps("/".join((contenturl, propsvalues[titleindex])))
			flg = resultset.next()
	return getProps
def rowsToSheet(cellrange, datarows):  # 引数のセル範囲を左上端にして一括書き込みして列幅を最適化する。datarowsはタプルのタプル。
	datarows = tuple(zip(*zip_longest(*datarows, fillvalue="")))  # 一番長い行の長さに合わせて空文字を代入。
	sheet = cellrange.getSpreadsheet()  # セル範囲のあるシートを取得。
	cellcursor = sheet.createCursorByRange(cellrange)  # セル範囲のセルカーサーを取得。
	cellcursor.collapseToSize(len(datarows[0]), len(datarows))  # (列、行)で指定。セルカーサーの範囲をdatarowsに合せる。
	cellcursor.setDataArray(datarows)  # セルカーサーにdatarowsを代入。代入できるのは整数(int、ただしboolを除く)か文字列のみ。
	cellcursor.getColumns().setPropertyValue("OptimalWidth", True)  # セルカーサーのセル範囲の列幅を最適化する。行幅は限定サれない。		
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
def createFilePicker(ctx, smgr, kwargs):  # ファイル選択ダイアログを返す。kwargsはFilePickerのメソッドをキー、引数を値とする辞書。
	key = "TemplateDescription"
	if key in kwargs:  # TemplateDescriptionキーがある時。必須。
		filepicker = smgr.createInstanceWithArgumentsAndContext("com.sun.star.ui.dialogs.FilePicker", (kwargs.pop(key),), ctx)  # キャッシュするとおかしくなる。
		if kwargs:
			key = "appendFilter"
			if key in kwargs:  # appendFilterキーがある時。
				filters = kwargs.pop(key)  # 値を取得。値はフィルター名をキー、拡張子を値とする辞書。
				for uiname in sorted(filters.keys()):
					displayname = uiname if sys.platform.startswith('win') else "{} (*.{})".format(uiname, filters[uiname])  # 表示フィルターの作成。Windowsの場合は拡張子を含めない。
					getattr(filepicker, key)(displayname, "*.{}".format(filters[uiname]))					
			if kwargs:
				[getattr(filepicker, key)(val) for key, val in kwargs.items()]
		return filepicker
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