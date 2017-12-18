#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import glob
import os
def macro():  
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
	simplefileaccess = smgr.createInstanceWithContext("com.sun.star.ucb.SimpleFileAccess", ctx)  # SimpleFileAccess
	os.chdir("..")  # 一つ上のディレクトリに移動。
	ods = glob.glob("*.ods")[0]  # odsファイルを取得。最初の一つのみ取得。
	systempath = os.path.join(os.getcwd(), ods)  # odsファイルのフルパス。
	doc_fileurl = unohelper.systemPathToFileUrl(systempath)  # fileurlに変換。
	package = smgr. createInstanceWithArgumentsAndContext("com.sun.star.packages.Package", (doc_fileurl,), ctx)  # Package。第2引数はinitialize()メソッドで後でも渡せる。
	docroot = package.getByHierarchicalName("/")  # /Scripts/pythonは不可。
	if not ("Scripts" in docroot and "python" in docroot["Scripts"]):
		print("The embedded macro folder does not exist in {}.".format(ods))
		return	
	python_fileurl = createDest(simplefileaccess)  # 出力先フォルダのfileurlを取得。
	getContents(simplefileaccess, docroot["Scripts"]["python"], python_fileurl) 
def getContents(simplefileaccess, folder, pwd):
	for sub in folder:  # 子要素のオブジェクトについて。
		name = sub.getName()  # 子要素のオブジェクトの名前を取得。
		fileurl = "/".join((pwd, name))  # 出力先のfileurlを取得。
		if sub.supportsService("com.sun.star.packages.PackageFolder"):  # PackageFolderの時はフォルダとして出力。
			if not simplefileaccess.exists(fileurl):
				simplefileaccess.createFolder(fileurl)
			getContents(simplefileaccess, sub, fileurl)  # 子要素のオブジェクトについて同様に出力。
		elif sub.supportsService("com.sun.star.packages.PackageStream"):  # PackageStreamのときはファイルとして出力。
			simplefileaccess.writeFile(fileurl, sub.getInputStream())  # ファイルが存在しなければ新規作成してくれる。			
def createDest(simplefileaccess):
	src_path = os.path.join(os.getcwd(), "src")  # srcフォルダのパスを取得。
	src_fileurl = unohelper.systemPathToFileUrl(src_path)  # fileurlに変換。
	destdir = "/".join((src_fileurl, "Scripts/python"))
	if simplefileaccess.exists(destdir):  # pythonフォルダがすでにあるとき
		simplefileaccess.kill(destdir)  # すでにあるpythonフォルダを削除。	
	simplefileaccess.createFolder(destdir)  # pythonフォルダを作成。
	return destdir
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