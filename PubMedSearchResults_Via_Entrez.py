# You should test that your search return results first on the web 
# https://www.ncbi.nlm.nih.gov/dbvar before using them 
# in your python script.  Available dbVar search terms are on the help page 
# (https://www.ncbi.nlm.nih.gov/dbvar/content/help/#entrezsearch).
# For general Entrez help and boolean search see the online book
# (https://www.ncbi.nlm.nih.gov/books/NBK3837/#EntrezHelp.Entrez_Searching_Options)

# This example will make use of these eUtils History Server parameters
# usehistory, WebEnv, and query_key.  It is highly recommended you use them in
# your pipeline and script.

# /usehistory=/
# When usehistory is set to 'y', ESearch will post the UIDs resulting from the
# search operation onto the History server so that they can be used directly in
# a subsequent E-utility call. Also, usehistory must be set to 'y' for ESearch
# to interpret query key values included in term or to accept a WebEnv as input.

# /WebEnv=/
# Web environment string returned from a previous ESearch, EPost or ELink call.
# When provided, ESearch will post the results of the search operation to this
# pre-existing WebEnv, thereby appending the results to the existing
# environment. In addition, providing WebEnv allows query keys to be used in
# term so that previous search sets can be combined or limited. As described
# above, if WebEnv is used, usehistory must be set to 'y' (ie.
# esearch.fcgi?db=dbvar&term=asthma&WebEnv=<webenv string>&usehistory=y)

# /query_key=/
# Integer query key returned by a previous ESearch, EPost or ELink call. When
# provided, ESearch will find the intersection of the set specified by query_key
# and the set retrieved by the query in term (i.e. joins the two with AND). For
# query_key to function, WebEnv must be assigned an existing WebEnv string and
# usehistory must be set to 'y'.

# load python modules
# May require one time install of biopython and xml2dict.


"""
from Bio import Entrez
import xmltodict

# initialize some default parameters
Entrez.email = 'myemail@ncbi.nlm.nih.gov' # provide your email address
db = 'dbvar'                              # set search to dbVar database
paramEutils = { 'usehistory':'Y' }        # Use Entrez search history to cache results

# generate query to Entrez eSearch
eSearch = Entrez.esearch(db=db, term='("blah blah)', **paramEutils)

# get eSearch result as dict object
res = Entrez.read(eSearch)

# take a peek of what's in the result (ie. WebEnv, Count, etc.)
for k in res:
    print (k, "=",  res[k])

paramEutils['WebEnv'] = res['WebEnv']         #add WebEnv and query_key to eUtils parameters to request esummary using  
paramEutils['query_key'] = res['QueryKey']    #search history (cache results) instead of using IdList 
paramEutils['rettype'] = 'xml'                #get report as xml
paramEutils['retstart'] = 0                   #get result starting at 0, top of IdList
paramEutils['retmax'] = 5                     #get next five results

# generate request to Entrez eSummary
result = Entrez.esummary(db=db, **paramEutils)
# get xml result
xml = result.read()
# take a peek at xml
print(xml)

"""
import os
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


#search pubmed

jias = 1
search_string = r'("J Int AIDS Soc"[jour]) AND ("2018"[Date - Publication] : "3000"[Date - Publication])) AND ((((Stigma[Title/Abstract]) OR Discrimination[Title/Abstract]) OR Criminalization[Title/Abstract]) OR "Human Rights"[Title/Abstract])'

from Bio import Entrez
from Bio import Medline
import openpyxl
from openpyxl import load_workbook

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

x=1
for record in records:
	print("(" + str(x) + ")")
	print("Title:", record.get("TI", "?"))
	print("Authors: ", ", ".join(record.get("AU", "?")))
	print("Pub Date: " + record.get("DP", "?")[5:] + " " + record.get("DP", "?")[:-4])
	print("Journal:", record.get("JT", "?"))
	print("DOI:", record.get("LID", "?")[:-6])
	if jias:
		print("Wiley Link: " + "https://onlinelibrary.wiley.com/doi/full/" + record.get("LID", "?")[:-6])
	print("\n")
	x += 1

#testing adding data to excel file
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
	wiley_link = "https://onlinelibrary.wiley.com/doi/full/" + record.get("LID", "?")[:-6]

	excel_data = [str(x), title, authors, pub_date, doi, wiley_link]
	x += 1

	for row in excel_data:
		ws.append(excel_data)
		break

	book.save(filepath)
