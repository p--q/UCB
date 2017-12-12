#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import glob
import os
from zipfile import ZipFile
import shutil
def macro():  # オートメーションでのみ実行可。 
	os.chdir("..")  # 一つ上のディレクトリに移動。
	ods = glob.glob("*.ods")[0]  # odsファイルを取得。最初の一つのみ取得。
	src_path = os.path.join(os.getcwd(), "src")  # srcフォルダのパスを取得。
	embeddemacro_path = "Scripts/python/"  # ドキュメント内のパス。
	output_path = "/".join((src_path, embeddemacro_path))  # 出力先フォルダのパス。
	if os.path.exists(output_path):  # 出力先フォルダが存在する時。
		shutil.rmtree(output_path)  # 出力先フォルダを削除。
	with ZipFile(ods , 'r') as odszip: # odsファイルをZipFileで読み取り専用で開く。
		[odszip.extract(name, path=src_path) for name in odszip.namelist() if name.startswith(embeddemacro_path)]  # Scripts/python/から始まるパスのファイルのみ出力先フォルダに解凍する。
		print("Extract the embedded macro folder from {}.".format(ods))
if __name__ == "__main__":  # オートメーションで実行するとき	
	macro()
	