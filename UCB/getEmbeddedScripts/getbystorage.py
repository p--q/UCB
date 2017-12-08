#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import glob
import os
from com.sun.star.embed import ElementModes  # 定数
def macro():  
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
			if component.getURL()==doc_fileurl:  # fileurlが一致するとき
				documentstorage = component.getDocumentStorage()  # コンポーネントからストレージを取得。
				break
	else:  # ドキュメントが開いていない時。
		storagefactory = smgr.createInstanceWithContext('com.sun.star.embed.StorageFactory', ctx)  # StorageFactory
		documentstorage = storagefactory.createInstanceWithArguments((doc_fileurl, ElementModes.READ))  # odsファイルからストレージを読み取り専用で取得。
	if not "Scripts" in documentstorage:  # ドキュメント内にScriptsフォルダがないときは終了。
		print("The Scripts directory does not exist in {}.".format(ods))
		return
	scriptsstorage = documentstorage.openStorageElement("Scripts", ElementModes.READ)  # ドキュメント内のScriptsストレージを取得。documentstorage["Scripts"]でもよい。ただしストリームは取得できない。
	if not "python" in scriptsstorage:  # pythonフォルダがないときは終了。
		print("The Scripts/python directory does not exist in {}.".format(ods))
		return	
	pythonstorage = scriptsstorage.openStorageElement("python", ElementModes.READ)  # pythonストレージを取得。documentstorage["Scripts"]["python"]はエラーになる。
	dest_dir = createDest(simplefileaccess)  # 出力先フォルダのfileurlを取得。
	getContents(simplefileaccess, pythonstorage, dest_dir)  # 再帰的にストレージの内容を出力先フォルダに展開。
def getContents(simplefileaccess, storage, pwd):  # SimpleFileAccess、ストレージ、出力フォルダのfileurl	
	for name in storage:
		fileurl = "/".join((pwd, name))
		if storage.isStorageElement(name):  # ストレージのときはフォルダとして処理。
			if not simplefileaccess.exists(fileurl):
				simplefileaccess.createFolder(fileurl)
			substrorage = storage.openStorageElement(name, ElementModes.READ)  # サブストレージを取得。
			getContents(simplefileaccess, substrorage, fileurl)
		elif storage.isStreamElement(name):  # ストリームの時はファイルに書き出す。
			stream = storage.cloneStreamElement(name)  # サブストリームを取得。読み取り専用。
			simplefileaccess.writeFile(fileurl, stream.getInputStream())  # ファイルが存在しなければ新規作成してくれる。			
def createDest(simplefileaccess):  # 出力先フォルダのfileurlを取得する。
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