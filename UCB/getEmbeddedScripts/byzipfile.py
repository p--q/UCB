#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import glob
import os
from zipfile import ZipFile
import shutil
def macro():  
	os.chdir("..")  # 一つ上のディレクトリに移動。
	ods = glob.glob("*.ods")[0]  # odsファイルを取得。最初の一つのみ取得。
	odszip = ZipFile(ods , 'r')
	src_path = os.path.join(os.getcwd(), "src")  # srcフォルダのパスを取得。
	embeddemacro_path = "Scripts/python/"
	output_path = "/".join((src_path, embeddemacro_path))
	if os.path.exists(output_path):
		shutil.rmtree(output_path)
	[odszip.extract(name, path=src_path) for name in odszip.namelist() if name.startswith(embeddemacro_path)]
	print("Extract the embedded macro folder from {}.".format(ods))
if __name__ == "__main__":  # オートメーションで実行するとき	
	macro()
	