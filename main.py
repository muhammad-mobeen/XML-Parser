import xml.etree.ElementTree as ET
import openpyxl		# pip install openpyxl    (Library for r/w excel files)

class XMLParser:
	def __init__(self, file_path):
		self.file_path = file_path
		self.tree = None
		self.root = None
		self.wb = openpyxl.Workbook()

	def xml_parser(self):
		self.tree = ET.parse(self.file_path)
		self.root = self.tree.getroot()

		for i in self.root.iter():
			print(i.tag)

	def create_xlsx_file(self):
		# sheet = self.wb.worksheets[0]
		# sheet['C3'] = 'Hello World'
		# self.wb.create_sheet('New Sheet')
		self.wb.save('compiler.xlsx')

if __name__ == "__main__":
	file_path = 'compiler.xml'
	parse_agent = XMLParser(file_path)
	parse_agent.xml_parser()