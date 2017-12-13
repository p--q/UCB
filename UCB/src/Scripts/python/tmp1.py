def macro():
	doc = XSCRIPTCONTEXT.getDocument()
	controller = doc.getCurrentController()  # コントローラーを取得。
	sheet = controller.getActiveSheet()  # アクティブなシートを取得。
	sheet["A15"].setString("Hello by the embedded script.")
