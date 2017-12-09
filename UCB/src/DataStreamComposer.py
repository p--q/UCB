#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import os
from datetime import datetime
from com.sun.star.ucb import OpenCommandArgument2  # Struct
from com.sun.star.ucb import OpenMode  # 定数
from com.sun.star.io import XActiveDataSink
from com.sun.star.ucb import Command  # Struct
from com.sun.star.ucb import InsertCommandArgument  # Struct
def macro():  
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	print("""
-----------------------------------------------------------------
DataStreamComposer - sets the data stream of a document resource.
The data stream is obtained from another (the source) document resource before.
-----------------------------------------------------------------
""")
	t = datetime.now().isoformat().split(".")[0].replace("-", "").replace(":", "")
	targetfile = "".join(("resource-", t))
	with open(targetfile, 'w', encoding="utf-8") as f:
		f.write("This is the content of a sample data file.")
	pwd = unohelper.systemPathToFileUrl(os.getcwd())  # 現在のディレクトリのfileurlを取得。。
	contenturl = "/".join((pwd, targetfile))
	if not os.path.exists(os.path.join("data", "data.txt")):
		if not os.path.exists("data"):
			os.mkdir("data")
		os.chdir("data")
		with open("data.txt", 'w', encoding="utf-8") as f:
			f.write("sample sample sample sample sample sample sample sample EOF")	
		os.chdir("..")
	sourceurl =  "/".join((pwd, "data/data.txt"))	
	ucb =  smgr.createInstanceWithContext("com.sun.star.ucb.UniversalContentBroker", ctx)
	datasink = MyActiveDataSink()
	arg = OpenCommandArgument2(Mode=OpenMode.DOCUMENT, Priority=32768, Sink=datasink)
	command = Command(Name="open", Handle=-1, Argument=arg)
	content = ucb.queryContent(ucb.createContentIdentifier(sourceurl))
	content.execute(command, 0, None)
	data = datasink.getInputStream()
	insertcommandargument = InsertCommandArgument(Data=data, ReplaceExisting=True)
	command = Command(Name="insert", Handle=-1, Argument=insertcommandargument)
	content = ucb.queryContent(ucb.createContentIdentifier(contenturl))
	content.execute(command, 0, None)
	print("""
Setting data stream succeeded.
Source URL: {}
Target URL: {}	
""".format(sourceurl, contenturl))
class MyActiveDataSink(unohelper.Base, XActiveDataSink):	
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