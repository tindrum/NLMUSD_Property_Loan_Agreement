#!/usr/bin/env python
# Daniel Henderson

# CHANGE THE LINE ENDING OF THE EXCEL EXPORT!!!
#
# The line endings from Excel are F**ing Windows
# Change Them!!!

from lxml import html
import requests
import re
import codecs
import os
import sys
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch, cm
from reportlab.lib import colors
from reportlab.lib.utils import haveImages, fileName2FSEnc
from reportlab.lib.pagesizes import letter
from mmap import mmap,ACCESS_READ
from xlrd import open_workbook

people = {}
patronRegex = re.compile(r"([^,]+), (.+)")
assetRegex = re.compile(r"(Asset)")

# trying out xlrd
wb = open_workbook('Excel_files/PatronCircReportJob203634.xls')
for s in wb.sheets():
  print 'Sheet:',s.name
  for row in range(s.nrows):
    values = []
    for col in range(s.ncols):
      values.append(s.cell(row,col).value)
    assetFound = assetRegex.search(str(values[6]))
    if assetFound: # may need to test this too: values[0] == "" and 
      patronName = patronRegex.match(values[1])
      if patronName:
        patronString = patronName.group(2) + " " + patronName.group(1)
      else:
        patronString = values[1]
      if patronString in people:
        patron = people[patronString]
        patron.append([values[11], values[9]])
      else:
        people[patronString]=[[values[11], values[9]]]
      # print(patronString + ":   " + values[9] + " :: " + values[11])


p = people.items()
for person in p:
  print(person[0])
  for i in person[1]:
    print(i[0] + ": " + i[1])


ipad_slot = 4
comp_slot = 0
proj_slot = 2
elmo_slot = 1
printer_slot = 5


p = people.items()
for person in p:
  # print(person[0])
  while len(person[1]) < 14:
    person[1].insert(0, ["", ""])
  index = 0
  for i in person[1]:
    ipad = re.search(r"iPad|ipad|IPAD|Ipad", i[0])
    comp = re.search(r"iMac|Mac|Laptop|MacBook|Desktop", i[0])
    elmo = re.search(r"Elmo|Camera|camera", i[0])
    proj = re.search(r"LCD|DLP|PowerLite", i[0])
    printer = re.search(r"printer|Printer", i[0])
    
    if ipad:
      matchItem, other = person[1][index], person[1][ipad_slot]
      person[1][ipad_slot], person[1][index] = matchItem, other
    elif comp:
      matchItem, other = person[1][index], person[1][comp_slot]
      person[1][comp_slot], person[1][index] = matchItem, other
    elif elmo:
      matchItem, other = person[1][index], person[1][elmo_slot]
      person[1][elmo_slot], person[1][index] = matchItem, other
    elif proj:
      matchItem, other = person[1][index], person[1][proj_slot]
      person[1][proj_slot], person[1][index] = matchItem, other
    elif printer:
      matchItem, other = person[1][index], person[1][printer_slot]
      person[1][printer_slot], person[1][index] = matchItem, other
    index += 1
      
assetMatch = re.compile(r"(\d{1,2}/\d{1,2}/\d{4})[^\d\w]+(.*)[^\d\w]*((12|11|10|0?[1-9])/\d{1,2}/\d{4})")
nameMatch = re.compile(r"^([A-Za-z].+) \[.*\]$")


p = people.items()
for person in p:
  print(person[0])
  for i in person[1]:
    print(i[1] + ": " + i[0])

p = people.items()

# superimpose on PDF
# outputPDFname = 'PLA_form_' + borrowerName + "_r04.pdf"
outputPDFname = "Property_Loan_Agreement_01.pdf"
c = canvas.Canvas(outputPDFname , pagesize=letter)
# c = canvas.Canvas("PLA_anderson4.pdf", pagesize="letter")
# from reportlab.lib.units import inch
# c.translate(inch, inch)
for person in p:
  
  c.setFont("Helvetica", 14)

  # background x, y position (from center?)
  x_pos, y_pos = -0, -425
  bg_width = 610
  b_x, b_y = 170, 629
  sch_x, sch_y = 170, 650

  item_x, item_startY = 262, 591
  item_yoffset = 19

  uitem_x, uitem_startY = 162, 486
  uitem_yoffset = 14

  # report_string_values = [[borrowerName, b_x, b_y], ["test spot", i1_x, i1_y], ["test spot", i2_x, i2_y]]

  c.drawImage("PLA_background.tiff",x_pos, y_pos, width=bg_width, preserveAspectRatio=True) # x_pos and w_pos are # pixels from bl origin

  c.drawString(b_x, b_y, person[0])
  c.drawString(sch_x, sch_y, "Waite Middle School")

  c.setFont("Helvetica", 8)

  printCount = 0 # counter for loop for item printing
  maxOnFirstPage = 5 # max items to print on front of sheet
  backPage = False
  
  for i in person[1]:
    if not backPage:
      c.drawString(item_x, item_startY, (i[0] + "    " + i[1]))
      item_startY -= item_yoffset
    else:
      if i[0] != "" and i[1] != "":
        if printCount == 40:
          c.setFont("Helvetica", 8)
          item_x, item_startY = 150, 692
        if printCount == 76:
          c.setFont("Helvetica", 8)
          item_x, item_startY = 300, 741
        if printCount == 114:
          c.setFont("Helvetica", 7)
          item_x, item_startY = 450, 692
        if printCount == 150:
          c.setFont("Helvetica", 7)
          item_x, item_startY = 600, 741
        c.drawString(item_x, item_startY, (i[0] + "     Ser# " + i[1]))
        item_startY -= item_yoffset
      else:
        printCount += 1
        continue
    if printCount == maxOnFirstPage:
      c.drawString(item_x, item_startY, "(continued on back of page)")
      c.showPage()
      item_x, item_startY = 50, 700
      c.drawString(item_x, item_startY + 20, "(continued from front of page)")
      backPage = True
    printCount += 1
  c.setFont("Helvetica", 12)
  c.drawString(item_x, item_startY, "** (End of list) **")
  


  # draw school name
  c.showPage()
c.save()
