#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import os
from com.sun.star.ucb import OpenCommandArgument2  # Struct
from com.sun.star.ucb import OpenMode  # 定数
from com.sun.star.io import XActiveDataSink
from com.sun.star.ucb import Command  # Struct
def macro():  # オートメーションでのみ実行可。マクロだとカレントディレクトリが使えない。  
	if not os.path.exists(os.path.join("data", "data.txt")):  # インプットストリームを取得するファイルがなければ作成する。
		if not os.path.exists("data"):  # dataフォルダが存在しない時。
			os.mkdir("data")  # dataフォルダを作成。
		os.chdir("data")  # dataフォルダに移動。
		with open("data.txt", 'w', encoding="utf-8") as f:  # インプットストリームを取得するファイルを作成。
			f.write("sample sample sample sample sample sample sample sample EOF")	
		os.chdir("..") # 元のフォルダに戻る。
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	print("""-----------------------------------------------------------------------
DataStreamRetriever - obtains the data stream from a document resource.
-----------------------------------------------------------------------""")
	pwd = unohelper.systemPathToFileUrl(os.getcwd())  # 現在のディレクトリのfileurlを取得。。
	sourceurl =  "/".join((pwd, "data/data.txt"))  # インプットストリームを取得するfileurlを取得。	
	ucb =  smgr.createInstanceWithContext("com.sun.star.ucb.UniversalContentBroker", ctx)  # UniversalContentBroker。
	content = ucb.queryContent(ucb.createContentIdentifier(sourceurl))  # FileContentを取得。
	datasink = MyActiveDataSink()  # XActiveDataSinkを持ったクラスをインスタンス化。
	arg = OpenCommandArgument2(Mode=OpenMode.DOCUMENT, Priority=32768, Sink=datasink)  # SinkにXActiveDataSinkを持ったインスタンスを渡すとアクティブデータシンクとして使える。
	command = Command(Name="open", Handle=-1, Argument=arg)  # コマンド。
	content.execute(command, 0, None)  # コピー元ファイルについてコマンドを実行。setInputStream()でアクティブデータシンクにインプットストリームが渡される。
	inputstream = datasink.getInputStream()  # アクティブデータシンクからインプットストリームを取得。
	if inputstream:  # インプットストリームが取得出来た時。
		txt = "Getting data stream for resource {} succeeded.".format(sourceurl)
		print("""{}
{}""".format(txt, "-"*len(txt)))
		n, buffer = inputstream.readSomeBytes([], 65536)  # 第1引数のリストは戻り値のタプルの第2要素で返ってくる。
		txt = bytes(buffer).decode("utf-8")  # "".join(map(chr, buffer))でもよい。bytearrayをbyteに変換して文字列に変換している。
		print("""Read bytes : {}
Read data (only first 64K displayed):
{}	""".format(n, txt))
class MyActiveDataSink(unohelper.Base, XActiveDataSink):  # アクティブデータシンク。	
	def setInputStream(self, stream):
		self.stream = stream
	def getInputStream(self):
		return self.stream
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