import os
import sys
import zipfile
from subprocess import PIPE, run
from xml.dom.minidom import parseString

def find_settings_file(docx):
	file_list = ["word_rels/settings.xml.rels", "word/_rels/settings.xml.rels"]
	for i in range(len(file_list)):
		if file_list[i] in docx.namelist():
			return file_list[i]

def find_docx_get_C2(base_dir):
	for i in os.listdir(base_dir):
		path = os.path.join(base_dir, i)
		if ".docx" in path:
			docx = zipfile.ZipFile(path)
			settings_file = find_settings_file(docx)
			xml = parseString(docx.read(settings_file))
			result = run(["md5", path], stdout=PIPE, stderr=PIPE, universal_newlines=True)
			print (result.stdout.replace("\n", ""), "-", xml.getElementsByTagName('Relationship')[0].getAttribute("Target"))
		if os.path.isdir(path):
			find_docx_get_C2(path)

base_dir = sys.argv[1]

if __name__ == "__main__":
	find_docx_get_C2(base_dir)