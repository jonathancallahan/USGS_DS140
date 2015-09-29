#!/usr/bin/python

"""
Mazama_print_statistics.py

A script to convert the Excel Spreadsheets that make up the
USGS DataSeries 140 dataset and convert the contents of the 
worksheets into ASCII CSV files for ingest by other software:

The xlrd module (0.6.0a1 or higher) for reading Excel files  must
be installed before this script can be run.  It is available at:

  http://www.lexicon.net/sjmachin/xlrd.htm
"""

import xlrd
import sys, re

########################################
# get_row_data
# 
# Reads part of a row of numeric data from a worksheet, converting
# string values to float where appropriate.
#
# From help(xlrd):
#   XL_CELL_EMPTY   = 0
#   XL_CELL_TEXT    = 1
#   XL_CELL_NUMBER  = 2
#   XL_CELL_DATE    = 3
#   XL_CELL_BOOLEAN = 4
#   XL_CELL_ERROR   = 5
#   XL_CELL_BLANK   = 6

def get_row_data(workbook, sheet, row, colhi):
  result = []
  row_types = sheet.row_types(row)
  row_values = sheet.row_values(row)

  # Go through every cell in this row and check it's type.
  for col in range(0,colhi):
    cell_type = row_types[col]
    cell_value = row_values[col]

    # Empty cells are converted to 'na'.
    if cell_type == xlrd.XL_CELL_EMPTY:
      value = "na"
      cell_type = xlrd.XL_CELL_TEXT

    # Cells with text are converted directly to float.
    elif cell_type == xlrd.XL_CELL_TEXT:
      value = cell_value
      value = float(cell_value)
      cell_type = xlrd.XL_CELL_NUMBER

    # Cells with numbers do not need conversion.
    elif cell_type == xlrd.XL_CELL_NUMBER:
      value = cell_value

    # Cells of any other type are considered errors.
    else:
      print("UNKNOWN data type in row %d, col %d" % (row,col))
      print("    cell type = " + str(cell_type))
      sys.exit(1)

    result.append((cell_type, value))
  return result


########################################
# convert_file
# 
# Reads in an Excel file for a particular mineral and
# converts the contents to a CSV file

def convert_file(mineral,logfile):

  mineral_xls = mineral + '.xls'

  try:
    workbook = xlrd.open_workbook(mineral_xls, logfile=logfile)
  except xlrd.XLRDError:
    print >> logfile, "*** Open failed: %s: %s" % sys.exc_info()[:2]
  except:
    print >> logfile, "*** Open failed: %s: %s" % sys.exc_info()[:2]

  mineral_csv = "USGS_2011_" + mineral + ".csv"
  ###csv = open(mineral_csv,'w')
  sheet = workbook.sheet_by_index(0) # python index 0 = worksheet 1

  # The header row is typically row 5 but occasionally another row (eg. ironsteel.xls)
  # Search for the header row by looking for 'Year' in the first column.
  for i in range(0,200):
    if sheet.row_values(i)[0] == 'Year':
      header_row = i
      break;

  # Get all the titles and create associated names.
  # Harmonize non-standard names where appropriate.
  titles = sheet.row_values(header_row)
  colhi = len(titles)
  names = []
  for i in range(0,colhi):
    title = titles[i]
    title = title.strip()                          # remove leading/following whitespace
    title = re.sub("\s" , " ", title)              # replace any whitespacae with a single space
    title = re.sub("\s+" , " ", title)             # replace multilpe spaces with a single space
    titles[i] = title
    if title == 'Unit value ($/t)':
      names.append('unit_value')
    elif title == 'Unit value (98$/t)':
      names.append('unit_value_1998')
    elif title == 'Net import reliance (%)':       # found in aluminum.xls
      names.append('net_import_reliance')
    elif title == 'Unit value $/t':                # found in asbestos.xls
      names.append('unit_value')
    elif title == 'Unit value 98$/t':              # found in asbestos.xls
      names.append('unit_value_1998')
    else:
      names.append(title.lower().replace(' ','_'))

  titles_string = ','.join(titles)
  names_string = ','.join(names) 

  ###print("Working on " + mineral_csv)
  
  # debugging lines
  ###print(titles_string)
  ###print(names_string)
  if ('production' in names) and ('imports' in names) and ('exports' in names) and ('apparent_consumption' in names):
    print(mineral)
  else:
    print(mineral + " is missing one or more variables")
  """
  if 'production' not in names:
    print("\t missing production")
  if 'imports' not in names:
    print("\t missing imports")
  if 'exports' not in names:
    print("\t missing exports")
  if 'apparent_consumption' not in names:
    print("\t missing apparent_consumption")
  if 'world_production' not in names:
    print("\t missing world_production")
  """
  """
  csv.write("DC.title      = ASCII CSV version of ...\n")
  csv.write("file URL      = http://mazamascience.com/Minerals/USGS_2011_" + mineral + ".csv\n")
  csv.write("original data = http://minerals.usgs.gov/ds/2005/140/" + mineral + ".xls\n")
  csv.write("units         = metric tons\n")
  csv.write("\n")
  csv.write(titles_string + "\n")
  csv.write(names_string + "\n")

  # Data begin after the header_row and continue for up to current_year-1900 years
  # We will check the type of the first column to determine when to stop ingesting data
  for row in range(header_row+1,200):

    # Stop ingesting data if you run out of rows
    try:
      types = sheet.row_types(row)
      values = sheet.row_values(row)
    except IndexError:
      break;

    # Stop ingesting data when the Year column no longer contains numbers.
    if types[0] != xlrd.XL_CELL_NUMBER:
      break;

    # From help(xlrd):
    #   XL_CELL_EMPTY   = 0
    #   XL_CELL_TEXT    = 1
    #   XL_CELL_NUMBER  = 2
    #   XL_CELL_DATE    = 3
    #   XL_CELL_BOOLEAN = 4
    #   XL_CELL_ERROR   = 5
    #   XL_CELL_BLANK   = 6

    # First, validate and fix the data in the cells of this row
    for col in range(0,colhi):

      # Empty cells are converted to 'na'.
      if types[col] == xlrd.XL_CELL_EMPTY:
        values[col] = 'na'
        types[col] = xlrd.XL_CELL_TEXT
  
      # Cells with text are converted directly to float.
      elif types[col] == xlrd.XL_CELL_TEXT:

        value = values[col].strip()

        # Catch non-empty cells with a blank space in them
        if len(value) == 0:
          values[col] = 'na'
          types[col] = xlrd.XL_CELL_TEXT

        # Special case for aluminum, net_import_reliance
        elif mineral == 'aluminum' and names[col] == 'net_import_reliance' and value == 'E':
          values[col] = 'na'
          types[col] = xlrd.XL_CELL_TEXT

        # Special case for tantalum which has '1470*' in [2005,'World production']
        elif mineral == 'tantalum' and value == '1470*':
          values[col] = 1470.0
          types[col] = xlrd.XL_CELL_NUMBER

        # Everything else we convert to float
        else: 
          try:
            value[col] = float(value)
          except:
            print("Cannot convert value '%s' to float in row %d, col %d" % (value,row+1,col+1))
          types[col] = xlrd.XL_CELL_NUMBER
  
      # Cells with numbers do not need conversion.
      elif types[col] == xlrd.XL_CELL_NUMBER:
        pass
  
      # Cells of any other type are considered errors.
      else:
        print("UNKNOWN data type in row %d, col %d" % (row,col))
        print("    cell type = " + str(types[col]))
        sys.exit(1)

    # Second, print out the values with appropriate formatting
    for col in range(0,colhi):

      if types[col] == xlrd.XL_CELL_NUMBER:
        if col == 0: # Year
          csv.write("%d" % int(values[col]))
        else:
          csv.write(",%.1f" % values[col])

      elif types[col] == xlrd.XL_CELL_TEXT:
        csv.write(",\"%s\"" % values[col])

      else:
        print("UNKNOWN cell_type %d in column %d" % (types[col],col))
        sys.exit(1)

    csv.write("\n")

  print("Finished with " + mineral + " workbook.")
  """

################################################################################

def main():

  logfile = open('Mazama_2009.log', 'w')

#  convert_file('antimony',logfile)

  convert_file('abrasivesmanufactured',logfile)
  convert_file('abrasivesnatural',logfile)
  convert_file('agriculture',logfile)
  convert_file('aluminum',logfile)
  convert_file('antimony',logfile)
  convert_file('arsenic',logfile)
  convert_file('asbestos',logfile)
  convert_file('barite',logfile)
  convert_file('bauxitealumina',logfile)
  convert_file('beryllium',logfile)
  convert_file('bismuth',logfile)
  convert_file('boron',logfile)
  convert_file('bromine',logfile)
  convert_file('cadmium',logfile)
  convert_file('cement',logfile)
  convert_file('cesium',logfile)
  convert_file('chromium',logfile)
  convert_file('clay',logfile)
  convert_file('coalcombustionproducts',logfile)
  convert_file('cobalt',logfile)
  convert_file('columbium',logfile)
  convert_file('copper',logfile)
  convert_file('diamondindustrial',logfile)
  convert_file('diatomite',logfile)
  convert_file('feldspar',logfile)
  convert_file('fluorspar',logfile)
  convert_file('gallium',logfile)
  convert_file('garnet',logfile)
  convert_file('gemstones',logfile)
  convert_file('germanium',logfile)
  convert_file('gold',logfile)
  convert_file('graphite',logfile)
  convert_file('gypsum',logfile)
  convert_file('hafnium',logfile)
  convert_file('helium',logfile)
  convert_file('indium',logfile)
  convert_file('iodine',logfile)
  convert_file('ironore',logfile)
  convert_file('ironoxide',logfile)
  convert_file('ironsteelscrap',logfile)
  convert_file('ironsteelslag',logfile)
  convert_file('ironsteel',logfile)
  convert_file('kyanite',logfile)
  convert_file('lead',logfile)
  convert_file('lime',logfile)
  convert_file('lithium',logfile)
  convert_file('magnesiumcompounds',logfile)
  convert_file('magnesium',logfile)
  convert_file('manganese',logfile)
  convert_file('mercury',logfile)
  convert_file('micascrap',logfile)
  convert_file('micasheet',logfile)
  convert_file('molybdenum',logfile)
  convert_file('nickel',logfile)
  convert_file('nitrogen',logfile)
  convert_file('organics',logfile)
  convert_file('peat',logfile)
  convert_file('perlite',logfile)
  convert_file('phosphate',logfile)
  convert_file('platinum',logfile)
  convert_file('potash',logfile)
  convert_file('pumice',logfile)
  convert_file('quartzcrystal',logfile)
  convert_file('rareearths',logfile)
  convert_file('rhenium',logfile)
  convert_file('salt',logfile)
  convert_file('sandgravelconstruction',logfile)
  convert_file('sandgravelindustrial',logfile)
  convert_file('selenium',logfile)
  convert_file('silicon',logfile)
  convert_file('silver',logfile)
  convert_file('sodaash',logfile)
  convert_file('sodiumsulfate',logfile)
  convert_file('stonecrushed',logfile)
  convert_file('stonedimension',logfile)
  convert_file('strontium',logfile)
  convert_file('sulfur',logfile)
  convert_file('talc',logfile)
  convert_file('tantalum',logfile)
  convert_file('tellurium',logfile)
  convert_file('thallium',logfile)
  convert_file('thorium',logfile)
  convert_file('tin',logfile)
  convert_file('titaniumdioxide',logfile)
  convert_file('titaniummineral',logfile)
  convert_file('titanium',logfile)
  convert_file('tungsten',logfile)
  convert_file('vanadium',logfile)
  convert_file('vermiculite',logfile)
  convert_file('wollastonite',logfile)
  convert_file('wood',logfile)
  convert_file('zinc',logfile)
  convert_file('zirconium',logfile)

################################################################################

if __name__ == "__main__":
  main()

