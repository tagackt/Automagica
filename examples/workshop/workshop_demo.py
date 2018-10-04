"""
Simple Mouse Example 1
"""

from automagica import *

GetMouseCoordinates()

"""
Simple Mouse Example 2
"""

from automagica import *

x = 100
y = 100
DoubleClickOnPosition(x, y)

"""
Simple Mouse Example - Failsafe
"""

from automagica import *
import random

for i in range(0,10):
    random_X_position = random.randint(300,500)
    random_Y_position = random.randint(300,500)
    DragToPosition(random_X_position, random_Y_position)

"""
Browser Automation - Opening
""""

from automagica import *

browser = ChromeBrowser()

"""
Browser Automation - Closing
""""
from automagica import *

browser = ChromeBrowser()
browser.get('https://bing.com')
title = browser.title
if not "Google" in title:
    browser.close()
    DisplayMessageBox(title, title="Oops!", type="warning")

"""
Browser Automation - Searching on Google
"""
from automagica import *

browser = ChromeBrowser()
browser.get('https://google.com')
# Enter Search Text
browser.find_element_by_xpath('//*[@id="lst-ib"]').send_keys('KPMG')
# Submit
browser.find_element_by_xpath('//*[@id="lst-ib"]').submit()
# Click on the first result
browser.find_elements_by_class_name('r')[0].click()

""" 
Browser Automation - Going to KPMG page
"""
from automagica import *

browser = ChromeBrowser()
browser.get('https://home.kpmg.com')

"""
Finding Google search results with BeautifulSoup
"""
from automagica import *

GetGoogleSearchLinks("KPMG")

"""
Make a list of all KPMG Advisory Employees
"""
from automagica import *

browser = ChromeBrowser()
browser.get('https://home.kpmg.com/be/nl/home/misc/search.html')

# Klik op Mensen in linker menu
browser.find_element_by_xpath('//*[@id="page-content"]/section/div/div/div/section/div[3]/div[1]/div[1]/div/div/ul/li[4]/a/span[2]').click()
Wait(5)

# Vind alle mensen
persons = browser.find_elements_by_class_name('result')

# Initialiseer namenlijst
names = []

# Haal naam uit resultaten
for person in persons:
    name = person.text.splitlines()[0]
    names.append(name)

# Schijf weg naar file
WriteListToFile(names, file="KpmgEmployees.txt")


"""
Reads information from one Excel file and writes it to the other.
"""
from automagica import *

# Read information from example.xlsx in cell A1
workbook = OpenExcelWorkbook('example.xlsx')
worksheet = workbook.active
cell_content = worksheet['A1'].value

# Write information to new file
new_workbook = NewExcelWorkbook()
new_worksheet = new_workbook.active
new_worksheet['A1'] = cell_content

new_workbook.save('result.xlsx')

"""
Copy File, in this case an xlsx
"""
from automagica import *

Copyfile("example.xlsx", "copy.xlsx")    

