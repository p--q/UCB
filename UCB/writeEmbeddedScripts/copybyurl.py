#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import glob
import os
def macro():  # マクロでは利用不可。docを取得している行以降のみはマクロで実行可。 
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
	simplefileaccess = smgr.createInstanceWithContext("com.sun.star.ucb.SimpleFileAccess", ctx)  # SimpleFileAccess
	os.chdir("..")  # 一つ上のディレクトリに移動。
	ods = glob.glob("*.ods")[0]  # odsファイルを取得。最初の一つのみ取得。
	systempath = os.path.join(os.getcwd(), ods)  # odsファイルのフルパス。
	doc_fileurl = unohelper.systemPathToFileUrl(systempath)  # fileurlに変換。
	desktop = ctx.getByName('/singletons/com.sun.star.frame.theDesktop')  # デスクトップの取得。
	components = desktop.getComponents()  # ロードしているコンポーネントコレクションを取得。
	for component in components:  # 各コンポーネントについて。
		if hasattr(component, "getURL"):  # スタートモジュールではgetURL()はないため。
			if component.getURL()==doc_fileurl:  # fileurlが一致するとき、ドキュメントが開いているということ。
				doc = XSCRIPTCONTEXT.getDocument()
				transientdocumentsdocumentcontentfactory = smgr.createInstanceWithContext("com.sun.star.frame.TransientDocumentsDocumentContentFactory", ctx)
				transientdocumentsdocumentcontent = transientdocumentsdocumentcontentfactory.createDocumentContent(doc)
				pkgurl = transientdocumentsdocumentcontent.getIdentifier().getContentIdentifier()  # ex. vnd.sun.star.tdoc:/1
				break
	else:  # ドキュメントを開いていない時。__main__.InteractiveAugmentedIOException: Cannot store persistent data!と言われてできない。
		urireferencefactory = smgr.createInstanceWithContext("com.sun.star.uri.UriReferenceFactory", ctx)  # UriReferenceFactory
		urireference = urireferencefactory.parse(doc_fileurl)  # ドキュメントのUriReferenceを取得。
		vndsunstarpkgurlreferencefactory = smgr.createInstanceWithContext("com.sun.star.uri.VndSunStarPkgUrlReferenceFactory", ctx)  # VndSunStarPkgUrlReferenceFactory
		vndsunstarpkgurlreference = vndsunstarpkgurlreferencefactory.createVndSunStarPkgUrlReference(urireference)  # ドキュメントのvnd.sun.star.pkgプロトコールにUriReferenceを変換。
		pkgurl = vndsunstarpkgurlreference.getUriReference()  # UriReferenceから文字列のURIを取得。
	python_fileurl = "/".join((pkgurl, "Scripts/python"))  # ドキュメント内フォルダへのフルパスを取得。		
	if simplefileaccess.exists(python_fileurl):  # 埋め込みマクロフォルダがすでにあるとき
		simplefileaccess.kill(python_fileurl)  # 埋め込みマクロフォルダを削除。
	simplefileaccess.createFolder(python_fileurl)  # 埋め込みマクロフォルダを作成。
	sourcedir = getSource(simplefileaccess)  # コピー元フォルダのfileurlを取得。	
	if sourcedir:  # コピー元フォルダが存在するとき。
		simplefileaccess.copy(sourcedir, python_fileurl)  # 埋め込みマクロフォルダを出力先フォルダにコピーする。
	else:
		print("{} does not exist.".format(sourcedir))	
def getSource(simplefileaccess):  # コピー元フォルダのfileurlを取得する。
	src_path = os.path.join(os.getcwd(), "src")  # srcフォルダのパスを取得。
	src_fileurl = unohelper.systemPathToFileUrl(src_path)  # fileurlに変換。
	sourcedir = "/".join((src_fileurl, "Scripts/python"))
	if simplefileaccess.exists(sourcedir):  # pythonフォルダがすでにあるとき
		return sourcedir
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
		return createXSCRIPTCONTEXT()  # XSCRIPTCONTEXTの取得。
	XSCRIPTCONTEXT = automation()  # XSCRIPTCONTEXTを取得。	
	macro()  # マクロの実行。