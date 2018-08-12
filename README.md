# List-Maker
Python program used for creating C&amp;MA Bible Quizzing Lists.

#### About
This program is designed to assist in question writing for C&MA Bible Quizzing
by automatically creating a series of useful lists in a single Excel output file.

The lists it creates are:
* All Verses (Alpha).
* All Verses with no puncuatioon for easy searching (Alpha).
* Concordance (Alpha).
* Unique Words (Alpha).
* Two and Three word Phrases (Alpha).
* List of first five words of all verses (FTVs).
* List of first five words of all valid verse subsections (FTs).
* List of valid Quotations (SITs).

Future updates:
* Valid CVR and CR Phrases

### Getting Started

#### Prerequisites
This project is built using Python 3.6, Xlxswriter, Openpyxl, and Tkinter. Ensure you have them installed and working properly from the links below:

* (https://www.python.org/downloads/release/python-365/)
* (http://xlsxwriter.readthedocs.io/)
* (https://openpyxl.readthedocs.io/en/stable/)
* (https://wiki.python.org/moin/TkInter)

You can also install the packages using [Pip](https://pip.pypa.io/en/latest/quickstart/#quickstart):

* pip install xlsxwriter
* pip install openpyxl
* pip install tkinter

#### Input files
Upon run, the program will ask for an Excel document containing the material. It should have the following columns (with headers):
* Book => The Book name of verse.
* Chapter => The Chapter number of verse.
* Verse => The Verse number of verse.
* Verse Text => The actual Verse Text.

#### Running the program
To run the program, run the Python file. It will then ask for the input file and the output file.
```
ListMaker.py
```
#### Author(s)
* **Chris Lloyd** - *Main Program* - Legoman3267@Gmail.com
* **Andrew Southwick** - *Gui Design*

