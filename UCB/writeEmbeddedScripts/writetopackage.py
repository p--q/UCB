#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import glob
import os
def main():  
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
	simplefileaccess = smgr.createInstanceWithContext("com.sun.star.ucb.SimpleFileAccess", ctx)  # SimpleFileAccess
	os.chdir("..")  # 一つ上のディレクトリに移動。
	source_path = os.path.join(os.getcwd(), "src", "Scripts", "python")  # コピー元フォルダのパスを取得。	
	source_fileurl = unohelper.systemPathToFileUrl(source_path)  # fileurlに変換。	
	if not simplefileaccess.exists(source_fileurl):  # ソースにするフォルダがないときは終了する。
		print("The source macro folder does not exist.")	
		return
	ods = glob.glob("*.ods")[0]  # odsファイルを取得。最初の一つのみ取得。
	systempath = os.path.join(os.getcwd(), ods)  # odsファイルのフルパス。
	doc_fileurl = unohelper.systemPathToFileUrl(systempath)  # fileurlに変換。
	package = smgr.createInstanceWithArgumentsAndContext("com.sun.star.packages.Package", (doc_fileurl,), ctx)  # Package。第2引数はinitialize()メソッドで後でも渡せる。
	docroot = package.getByHierarchicalName("/")  # /Scripts/pythonは不可。
	if "Scripts" in docroot:  # Scriptsフォルダがすでにあるときは削除する。
		del docroot["Scripts"]
	docroot["Scripts"] = package.createInstanceWithArguments((True,))  # ScriptsキーにPackageFolderを挿入。
	docroot["Scripts"]["python"] = package.createInstanceWithArguments((True,))  # pythonキーにPackageFolderを挿入。
	writeScripts(simplefileaccess, package, source_fileurl, docroot["Scripts"]["python"])  # 再帰的にマクロフォルダーにコピーする。
	package.commitChanges()  # ファイルにパッケージの変更を書き込む。manifest.xmlも編集される。フォルダは書き込まれない。
def writeScripts(simplefileaccess, package, source_fileurl, packagefolder):  # コピー元フォルダのパス、出力先パッケージフォルダ。
	for fileurl in simplefileaccess.getFolderContents(source_fileurl, True):  # Trueでフォルダも含む。再帰的ではない。フルパスのfileurlが返る。
		name = fileurl.split("/")[-1]  # 要素名を取得。
		if simplefileaccess.isFolder(fileurl):  # フォルダの時。
			if not name in packagefolder:  # パッケージの同名のPackageFolderがない時。
				packagefolder[name] = package.createInstanceWithArguments((True,))  # キーをnameとするPackageFolderを挿入。
			writeScripts(simplefileaccess, package, fileurl, packagefolder[name])  # 再帰呼び出し。			
		else:
			packagefolder[name] = package.createInstance()  # キーをnameとするPackageStreamを挿入。
# 			packagefolder[name] = package.createInstanceWithArguments((False,))  # キーをnameとするPackageStreamを挿入。
			packagefolder[name].setInputStream(simplefileaccess.openFileRead(fileurl))  # ソースファイルからインプットストリームを取得。
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
	main()  