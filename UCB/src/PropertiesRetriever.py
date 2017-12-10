#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import os
from com.sun.star.beans import Property  # Struct
from com.sun.star.ucb import Command  # Struct
def macro():  
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	print("""--------------------------------------------------------------
PropertiesRetriever - obtains property values from a resource.
--------------------------------------------------------------""")
	
	if not os.path.exists(os.path.join("data", "data.txt")):
		if not os.path.exists("data"):
			os.mkdir("data")
		os.chdir("data")
		with open("data.txt", 'w', encoding="utf-8") as f:
			f.write("sample sample sample sample sample sample sample sample EOF")	
		os.chdir("..")
	pwd = unohelper.systemPathToFileUrl(os.getcwd())  # 現在のディレクトリのfileurlを取得。	
	sourceurl =  "/".join((pwd, "data/data.txt"))		
	ucb =  smgr.createInstanceWithContext("com.sun.star.ucb.UniversalContentBroker", ctx)
	content = ucb.queryContent(ucb.createContentIdentifier(sourceurl))
	propnames = "Title", "IsDocument"
	props = [Property(Name=propname, Handle=-1) for propname in propnames]
	command = Command(Name="getPropertyValues", Handle=-1, Argument=props)
	values = content.execute(command, 0, None)
	
	pass
	
	
	
# 	props = ("Title", "".join(("changed-", os.path.basename(contenturl)))),  # あとで使うために順を保持。
# 	propertyvalues = [PropertyValue(Name=key, Handle=-1, Value=val) for key, val in props]
# 	command = Command(Name="setPropertyValues", Handle=-1, Argument=propertyvalues)		
# 	result = content.execute(command, 0, None)  # なぜかIllegalArgumentExceptionがでてる。
# 	txt = "".join("Setting properties of resource ", contenturl)
# 	print("""{}
# {}""".format(txt, "-"*len(txt)))
# 	for i in range(result):
# 		print("Setting property {} succeeded.".format(props[i][0]))
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