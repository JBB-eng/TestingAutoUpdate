#search pubmed

import urllib.request, urllib.parse
import re, cloudscraper, os, openpyxl

from Bio import Entrez
from Bio import Medline
from openpyxl import load_workbook

jias_bool = 1 #set to True if looking for JIAS publications
citation_bool = 1 #set to True to include citation amount for JIAS publications **SIGNIFICANTLY INCREASES PROCESSING TIME***

#SEARCH STRING
search_string = r'("J Int AIDS Soc"[jour]) AND ("2018"[Date - Publication] : "3000"[Date - Publication])) AND ((((Stigma[Title/Abstract]) OR Discrimination[Title/Abstract]) OR Criminalization[Title/Abstract]) OR "Human Rights"[Title/Abstract])'

def GetDownloadPath():
	"""Returns the default downloads path for linux or windows"""
	if os.name == 'nt':
		import winreg
		sub_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
		downloads_guid = '{374DE290-123F-4565-9164-39C4925E467B}'
		with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key) as key:
			location = winreg.QueryValueEx(key, downloads_guid)[0]
		return location
	else:
		return os.path.join(os.path.expanduser('~'), 'downloads')

def GetCitationNumber(url):
	"""Gets number of citations for JIAS publications from Wiley's website"""
	scraper = cloudscraper.create_scraper()
	web_text = scraper.get(url).text
	citation_text = 'Citations: <a href="#citedby-section">'

	for line in web_text.splitlines():
		if citation_text in line:
			index = line.find(citation_text)
			temp = line[index:]
			temp = re.sub(citation_text, '', temp)
			temp = re.sub('</a></span></div>', '', temp)
			citation_number = temp
			break
		else:
			citation_number = 0

	return citation_number

#search Pubmed...
Entrez.email = "hello_world@example.com"  
handle = Entrez.egquery(term=search_string)
record = Entrez.read(handle)
count = 0
for row in record["eGQueryResult"]:
		if row["DbName"]=="pubmed":
			print("Number of articles found with requested search parameters:", row["Count"])
			count = row["Count"]

handle = Entrez.esearch(db="pubmed", term=search_string, retmax=count)
record = Entrez.read(handle)
handle.close()
idlist = record["IdList"]

handle = Entrez.efetch(db="pubmed", id=idlist, rettype="medline", retmode="text")
records = Medline.parse(handle)

records = list(records)

#print results (without citation)...
x=1
for record in records:
	print("(" + str(x) + ")")
	print("Title:", record.get("TI", "?"))
	print("Authors: ", ", ".join(record.get("AU", "?")))
	print("Pub Date: " + record.get("DP", "?")[5:] + " " + record.get("DP", "?")[:-4])
	print("Journal:", record.get("JT", "?"))
	print("DOI:", record.get("LID", "?")[:-6])
	if jias_bool:
		print("Wiley Link: " + "https://onlinelibrary.wiley.com/doi/full/" + record.get("LID", "?")[:-6])
	print("\n")
	x += 1

#add results to excel folder (including citations)...
x=1
filepath = GetDownloadPath() + "\\test_keyword_search.xlsx"
try:
	book = load_workbook(filepath)
except:
	print("creating excel file....")
	wb = openpyxl.Workbook()
	wb.save(filepath)
	book = load_workbook(filepath)
	sheet = book.active
	sheet['A1'] = "No."
	sheet['B1'] = "Title"
	sheet['C1'] = "Authors"
	sheet['D1'] = "Publication Date"
	sheet['E1'] = "DOI"
	sheet['F1'] = "Wiley Link"
	sheet['G1'] = "Citations"

ws = book.worksheets[0]
for record in records:
	for cell in ws["A"]:
		if cell.value is None:
			emptyRow = cell.row
			break
	else:
		emptyRow = cell.row + 1

	title = record.get("TI", "?")
	authors = ", ".join(record.get("AU", "?"))
	pub_date = record.get("DP", "?")[5:] + " " + record.get("DP", "?")[:-4]
	journal_name = record.get("JT", "?")
	doi = record.get("LID", "?")[:-6]
	wiley_link = "http://onlinelibrary.wiley.com/doi/full/" + record.get("LID", "?")[:-6]

	if citation_bool:
		excel_data = [str(x), title, authors, pub_date, doi, wiley_link, GetCitationNumber(wiley_link)]
	else:
		excel_data = [str(x), title, authors, pub_date, doi, wiley_link, "skipped"]
	
	x += 1

	for row in excel_data:
		ws.append(excel_data)
		break

	book.save(filepath)