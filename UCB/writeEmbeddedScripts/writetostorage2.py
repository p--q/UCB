#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import glob
import os
from com.sun.star.embed import ElementModes  # 定数
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
	desktop = ctx.getByName('/singletons/com.sun.star.frame.theDesktop')  # デスクトップの取得。
	components = desktop.getComponents()  # ロードしているコンポーネントコレクションを取得。
	for component in components:  # 各コンポーネントについて。
		if hasattr(component, "getURL"):  # スタートモジュールではgetURL()はないため。
			if component.getURL()==doc_fileurl:  # fileurlが一致するとき
				documentstorage = component.getDocumentStorage()  # コンポーネントからストレージを取得。
				break
	else:  # ドキュメントが開いていない時。
		storagefactory = smgr.createInstanceWithContext('com.sun.star.embed.StorageFactory', ctx)  # StorageFactory
		documentstorage = storagefactory.createInstanceWithArguments((doc_fileurl, ElementModes.READWRITE))  # odsファイルからストレージを取得。
		
		
# 	storagefactory = smgr.createInstanceWithContext('com.sun.star.embed.StorageFactory', ctx)
# 	storage = storagefactory.createInstance()
# 	documentstorage = storagefactory.createInstance()

	
	
	if "Scripts" in documentstorage:  # Scriptsフォルダがすでにあるときは削除する。
		documentstorage.removeElement("Scripts")
		documentstorage.commit()
	scriptsstorage = documentstorage.openStorageElement("Scripts", ElementModes.READWRITE)	
	
	
	
	documentstorage.commit()
	pythonstorage = scriptsstorage.openStorageElement("python", ElementModes.READWRITE)	
	scriptsstorage.commit()
	writeScripts(simplefileaccess, source_fileurl, pythonstorage)  # 再帰的にマクロフォルダーにコピーする。
	documentstorage.commit()
def writeScripts(simplefileaccess, source_fileurl, storage):  # コピー元フォルダのパス、出力先ストレージ。
	for fileurl in simplefileaccess.getFolderContents(source_fileurl, True):  # Trueでフォルダも含む。再帰的ではない。フルパスのfileurlが返る。
		name = fileurl.split("/")[-1]  # 要素名を取得。
		if simplefileaccess.isFolder(fileurl):  # フォルダの時。
			if not name in storage:  # ストレージに同名のストレージがない時。
				substorage = storage.openStorageElement(name, ElementModes.READWRITE)	
				storage.commit()
			writeScripts(simplefileaccess, fileurl, substorage)  # 再帰呼び出し。			
		else:
			xstream = storage.openStreamElement(name, ElementModes.WRITE)	
			outputstream = xstream.getOutputStream()
			inputstream = simplefileaccess.openFileRead(fileurl)
			dummy, b = inputstream.readBytes([], simplefileaccess.getSize(fileurl))
			outputstream.writeBytes(b)
			storage.commit()
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