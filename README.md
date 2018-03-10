This is a module for extracting specific Aetna insurance plan data from PDF files. See `para01.pdf` - `para09.pdf` for the format of files that this tool is built for.

# Method:
 - Parse each PDF file one at a time. Since they are all the exact same format, the locations of all the data will always be the same. Split the file line by line, and use some predetermined line numbers and indexes for each piece of data to extract everything we need.
 - Convert the given XLSX template to a CSV file.
 - Construct lines to add into the CSV file using the data we extracted, and add those lines in
 - Convert the CSV file we have back to XLSX, overwriting the old XLSX file
 - Remove the temporary CSV file from disk

# Settings
Settings are specified at the top of `extract_info.py`
```python
settings = {
	"xlsx_filename": "BeneFix Small Group Plans.xlsx",
	"xlsx_sheetname": "Blank Upload Template",
	"pdfs": [
		"para01.pdf",
		"para02.pdf",
		"para03.pdf",
		"para05.pdf",
		"para06.pdf",
		"para07.pdf",
		"para08.pdf",
		"para09.pdf"
	],
}
```
 - `xlsx_filename`: Must point to the template file. Usually will have one row containing column headers
 - `xlsx_sheetname`: Must be a sheet in the template file
# Usage
 - Make sure you're using Python 2.7
 - Make a new `virtualenv` if you would like to. 
 - `pip install -r requirements.txt` 
 - Modify `settings` as necessary. This is inside `extract_info.py`
 - `python2.7 extract_info.py`

# Possible Extension
 - Making the tool accept command line arguments, instead of having to modify the source file to change settings. There are two ways this could work:
	 - Accept a list of input files. E.g. `./extract_info.py --files para01.pdf para02.pdf ... para09.pdf`
	 - Accept a single JSON file. E.g. `./extract_info.py --settings settings.json` where `settings.json` contains something like the following:
	 ```json
	{
		settings: {
			"xlsx_filename": "BeneFix Small Group Plans.xlsx",
			"xlsx_sheetname": "Blank Upload Template",
			"pdfs": [
				"para01.pdf",
				"para02.pdf",
				"para03.pdf",
				"para05.pdf",
				"para06.pdf",
				"para07.pdf",
				"para08.pdf",
				"para09.pdf"
			],
		}
	}
	 ```