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
		documentstorage = storagefactory.createInstanceWithArguments((doc_fileurl, ElementModes.WRITE))  # odsファイルからストレージを書き込み用で取得。
	if "Scripts" in documentstorage:
		scriptstorage = documentstorage["Scripts"]
	else:
		scriptstorage = documentstorage.openStorageElement("Scripts", ElementModes.WRITE)  # 出力先に同名のストレージの作成。
		scriptstorage.commit()
	if "python" in scriptstorage:
		pythonstrorage = scriptstorage["python"]
	else:
		pythonstrorage = scriptstorage.openStorageElement("python", ElementModes.WRITE) 
		pythonstrorage.commit()
	sourcedir = getSource(simplefileaccess)  # コピー元フォルダのfileurlを取得。	
	if not sourcedir:
		print("{} does not exist.".format(sourcedir))	
		return
	filesystemstoragefactory = smgr.createInstanceWithContext('com.sun.star.embed.FileSystemStorageFactory', ctx)
	filesystemstorage = filesystemstoragefactory.createInstanceWithArguments((sourcedir, ElementModes.READ))  # ファイルシステムストレージを取得。
	toDocumentStorage(filesystemstorage, pythonstrorage)  # 再帰的にストレージの内容を出力先ストレージに展開。
	documentstorage.commit()
def toDocumentStorage(srcstorage, deststorage):  # SimpleFileAccess、ストレージ、出力先ストレージ	
	for name in srcstorage:  # ストレージの各要素名について。
		if srcstorage.isStorageElement(name):  # ストレージの時。
			subdest = deststorage.openStorageElement(name, ElementModes.WRITE)  # 出力先に同名のストレージの作成。
			toDocumentStorage(srcstorage[name], subdest)  # 子要素について同様にする。
			subdest.commit()
		elif srcstorage.isStreamElement(name):  # ストリームの時。
			subdest = deststorage.openStreamElement(name, ElementModes.WRITE)  # 出力先に同名のストリームを作成。
			inputstream = srcstorage[name].getInputStream()  # 読み取るファイルのインプットストリームを取得。
			outputstream = subdest.getOutputStream()  # 書き込むファイルのアウトプットストリームを取得。
			dummy, bytes = inputstream.readBytes([], inputstream.available())  # インプットストリームからデータをすべて読み込む。バイト配列の要素数とバイト配列のタプルが返る。
			outputstream.writeBytes(bytes)  # バイト配列をアウトプットストリームに渡す。	
			
			print(name)
			print(dummy)
				
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