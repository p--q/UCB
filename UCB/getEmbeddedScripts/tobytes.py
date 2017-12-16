#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import glob
import os
import shutil
from com.sun.star.embed import ElementModes  # 定数
def macro():  #   マクロで実行は不可。
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
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
	if not ("Scripts" in documentstorage and "python" in documentstorage["Scripts"]):  # ドキュメント内に埋め込みマクロフォルダがない時は終了する。
		print("The embedded macro folder does not exist in {}.".format(ods))
		return
	dest_path = os.path.join(os.getcwd(), "src", "Scripts", "python")  # 出力先フォルダのパスを取得。
	if os.path.exists(dest_path):  # 出力先フォルダが存在すれば削除する。
		shutil.rmtree(dest_path)
	os.makedirs(dest_path)  # 出力先フォルダを作成。
	scriptsstorage = documentstorage["Scripts"]  # documentstorage["Scripts"]["python"]ではイテレーターになれない。
	getContents(scriptsstorage["python"], dest_path)  # 再帰的にストレージの内容を出力先フォルダに展開。
def getContents(storage, dest):  # ストレージ、出力先フォルダ	
	for name in storage:  # ストレージの各要素名について。
		if storage.isStorageElement(name):  # ストレージの時。
			subdest = os.path.join(dest, name)  # 出力先フォルダのパスを取得。
			if not os.path.exists(subdest):  # 出力先フォルダが存在しなければ作成する。
				os.mkdir(subdest)
			getContents(storage[name], subdest)  # 子要素について同様にする。
		elif storage.isStreamElement(name):  # ストリームの時。
			inputstream = storage[name].getInputStream()  # 読み取るファイルのインプットストリームを取得。
			dummy, b = inputstream.readBytes([], inputstream.available())  # パッシブデータソースからバイトシークエンスを取得。
			inputstream.closeInput()  # 必要?インプットストリームを閉じる。
			filepath = os.path.join(dest, name)  # 書き込み先ファイルのパスの取得。
			with open(filepath, "wb") as f:  # 書き込み先ファイルの作成。
				f.write(bytes(b))  # バイト列をファイルに書き込む。
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