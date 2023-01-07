'''
Author: Muhammad Mobeen
Reg No: 200901097
BS-CS-01 (B)
Compiler Construction Assignment # 3
Submitted to Mam Reeda Saeed

GitHub Repo URL: https://github.com/muhammad-mobeen/XML-Parser
'''

import xml.etree.ElementTree as ET	# XML Parsing Library
import openpyxl						# pip install openpyxl (Library for r/w excel files)
from openpyxl.styles import Font	# Fonts for Excel
import re							# Regex Library

class XMLParser:
	'''
	XMLParser parses XML file and translates fetched data into a nicely formated Excel File\n
	XMLParser accepts file_path as an argument and generates an Excel file after the process finishes.
	'''
	def __init__(self, file_path):
		self.tree = ET.parse(file_path)
		self.root = self.tree.getroot()
		self.wb = openpyxl.Workbook()
		self.sheet = self.wb.worksheets[0]
		self.excel_filename = 'compiler.xlsx'

	def driver(self):
		'''Drives the class operations by calling class functions in order.'''
		self.init_xlsx_file()
		self.xml_parser()
		self.wb.save(self.excel_filename)

	def init_xlsx_file(self):
		'''
		Intailizer for the Excel file.\n
		Parses XML file and fetches the node tags, and with that intiates coloumns in Excel.
		'''
		st_font = Font(bold=True)	# Standard Font for first row
		node_tags = []
		col = 'A'
		for node in self.root.iter():
			if node.tag in node_tags:
				break
			elif node.tag != 'catalog':
				self.sheet[col+'1'].font = st_font
				self.sheet[col+'1'] = node.tag.capitalize()
				node_tags.append(node.tag)
				col = chr(ord(col)+1)

	def xml_parser(self):
		'''Parses the XML file iteratively and uses regex patterns to match and maps the data onto the Excel file concurrently.'''

		# Regex Patterns
		catalog_pattern = '^catalog$'
		book_pattern = '^book$'
		author_pattern = '^author$'
		title_pattern = '^title$'
		genre_pattern = '^genre$'
		price_pattern = '^price$'
		publish_date_pattern = '^publish_date$'
		description_pattern = '^description$'

		current_row = 1		# Keeps track of the pointer, on which row it is

		for node in self.root.iter():	# Loop to parse each node of parsed xml tree
			if re.findall(catalog_pattern, node.tag):
				self.sheet.title = 'Catalog'
			elif re.findall(book_pattern, node.tag):
				current_row += 1
				self.sheet['A'+str(current_row)] = node.get('id')
			elif re.findall(author_pattern, node.tag):
				self.sheet['B'+str(current_row)] = node.text
			elif re.findall(title_pattern, node.tag):
				self.sheet['C'+str(current_row)] = node.text
			elif re.findall(genre_pattern, node.tag):
				self.sheet['D'+str(current_row)] = node.text
			elif re.findall(price_pattern, node.tag):
				self.sheet['E'+str(current_row)] = node.text
			elif re.findall(publish_date_pattern, node.tag):
				self.sheet['F'+str(current_row)] = node.text
			elif re.findall(description_pattern, node.tag):
				self.sheet['G'+str(current_row)] = node.text
			else:
				print("Did not matched to anything!")

if __name__ == "__main__":
	file_path = 'compiler.xml'
	parse_agent = XMLParser(file_path)
	parse_agent.driver()
