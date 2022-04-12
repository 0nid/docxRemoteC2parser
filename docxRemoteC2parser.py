import os
import sys
import zipfile
from xml.dom.minidom import parseString

def find_docx_get_C2(base_dir):
	for i in os.listdir(base_dir):
		path = os.path.join(base_dir, i)
		if ".docx" in path:
			docx = zipfile.ZipFile(path)
			if "word/_rels/settings.xml.rels" in docx.namelist() or "word_rels/settings.xml.rels" in docx.namelist():
				xml = parseString(docx.read("word/_rels/settings.xml.rels"))
				print (xml.getElementsByTagName('Relationship')[0].getAttribute("Target"))
			if os.path.isdir(path):
				find_docx_get_C2(path)

base_dir = sys.argv[1]

if __name__ == "__main__":
	find_docx_get_C2(base_dir)