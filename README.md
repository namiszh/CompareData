## ValidateIdRecords

**What**

This is specialized tool to check whether the records of an excel file are correct or not. The excel file has two columns: Name and ID NO. Also another excel file that contains massive correct data should be provided.

**Why**

My cousin-in-law needs to deal with massive records in her work. Doing that manually is impossible. So I write this tool to help her.

**Dependencies**

* [Python](https://www.python.org/)
* [Click](http://click.pocoo.org/)
* [OpenPyXL](https://pypi.python.org/pypi/openpyxl)

**How**

1. Install dependencies
2. Download script
3. Run the program
    
`python ValidateIdRecords.py`

You will be prompted to enter the correct data file(in excel format) and the file you like to validate(in excel format).
