#!/usr/bin/python

"""
Mazama_USGS_DS140_2011.py

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

    mineral_csv = mineral + ".csv"
    csv = open(mineral_csv,'w')
    sheet = workbook.sheet_by_index(0) # python index 0 = worksheet 1

    # In 2011, the nickel.xls file doesn't include 'Year' in the header line so the test below won't work
    if mineral == 'nickel':
        header_row = 4
    else:
        # The header row is typically row 5 but occasionally another row (eg. ironsteel.xls)
        # Search for the header row by looking for 'Year' in the first column.
        for row in range(0,200):
            if sheet.row_values(row)[0] == 'Year':
                header_row = row
                break;

    # Get all the titles and create associated names.
    # Harmonize non-standard names where appropriate.
    titles = sheet.row_values(header_row)
    colhi = len(titles)
    names = []
    for col in range(0,colhi):
        title = titles[col]
        title = title.strip()                          # remove leading/following whitespace
        title = re.sub("\s" , " ", title)              # replace any whitespacae with a single space
        title = re.sub("\s+" , " ", title)             # replace multilpe spaces with a single space
        titles[col] = '"' + title + '"' 
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

    print("Working on " + mineral_csv)
    """
  # debugging lines
  print(titles_string)
  print(names_string)
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

    csv.write("DC.title      = ASCII CSV version of " + mineral + ".xls file\n")
    csv.write("file URL      = http://mazamascience.com/Minerals/USGS/DS140/2011/" + mineral + ".csv\n")
    csv.write("original data = http://minerals.usgs.gov/ds/2005/140/" + mineral + ".xls\n")
    csv.write("units         = metric tons\n")
    csv.write("\n")
    csv.write(titles_string + "\n")
    csv.write(names_string + "\n")

    # Check first year and fill in missing values if first year > 1900
    year = 1900
    first_year = sheet.row_values(header_row+1)[0]
    while (year < first_year):
        for col in range(0,colhi):
            if col == 0:
                csv.write('%d' % year)
            else:
                csv.write(',"na"')
        csv.write('\n')
        year += 1


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
        if types[0] == xlrd.XL_CELL_NUMBER:
            year = values[0]
        else:
            break;

        # From help(xlrd):
        #   XL_CELL_EMPTY   = 0
        #   XL_CELL_TEXT    = 1
        #   XL_CELL_NUMBER  = 2
        #   XL_CELL_DATE    = 3
        #   XL_CELL_BOOLEAN = 4
        #   XL_CELL_ERROR   = 5
        #   XL_CELL_BLANK   = 6

        # Validate and fix the data in the cells of this row.
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
                    if value == 'NA':
                        values[col] = 'na'
                        types[col] = xlrd.XL_CELL_NUMBER
                    elif value == 'W':
                        values[col] = 'na'
                        types[col] = xlrd.XL_CELL_TEXT
                        print("Converting row %d, column %d to 'na'" % (row,col))
                    else:
                        try:
                            values[col] = float(value)
                            types[col] = xlrd.XL_CELL_NUMBER
                        except:
                            print("Cannot convert value '%s' to float in row %d, col %d" % (value,row+1,col+1))

            # Cells with numbers do not need conversion.
            elif types[col] == xlrd.XL_CELL_NUMBER:
                value = values[col]
                if value == 'W':
                    values[col] = 'na'
                    types[col] = xlrd.XL_CELL_TEXT
                pass

            # Cells of any other type are considered errors.
            else:
                print("UNKNOWN data type in row %d, col %d" % (row,col))
                print("    cell type = " + str(types[col]))
                sys.exit(1)

        #### Special case for beryllium which is missing the row for the year 2000.
        #### Before we write out the results for 2001, insert the 2000 results -- all missing values
        ###if (mineral == 'beryllium') and (year == 2001):
            ###csv.write("2000")
            ###for col in range(1,colhi):
                ###csv.write(",\"na\"")
            ###csv.write("\n")

        # Print out the values with appropriate formatting.
        for col in range(0,colhi):

            if types[col] == xlrd.XL_CELL_NUMBER:
                if col == 0: # Year
                    csv.write("%d" % int(values[col]))
                else:
                    if values[col] == 'na':
                        csv.write(",\"na\"")
                    else:
                        csv.write(",%.1f" % values[col])

            elif types[col] == xlrd.XL_CELL_TEXT:
                csv.write(",\"%s\"" % values[col])

            else:
                print("UNKNOWN cell_type %d in column %d" % (types[col],col))
                sys.exit(1)


        csv.write("\n")

    # Check first year and fill in missing values if first year > 1900
    last_year = 2020
    year += 1
    while (year <= last_year):
        for col in range(0,colhi):
            if col == 0:
                csv.write('%d' % year)
            else:
                csv.write(',"na"')
        csv.write('\n')
        year += 1

    print("Finished with " + mineral + " workbook.")


########################################
# convert_use_file
# 
# Reads in an Excel file for a particular mineral and
# converts the contents to a CSV file

def convert_use_file(mineral,logfile):

    mineral_xls = mineral + '.xls'

    try:
        workbook = xlrd.open_workbook(mineral_xls, logfile=logfile)
    except xlrd.XLRDError:
        print >> logfile, "*** Open failed: %s: %s" % sys.exc_info()[:2]
    except:
        print >> logfile, "*** Open failed: %s: %s" % sys.exc_info()[:2]

    mineral_csv = mineral + ".csv"
    csv = open(mineral_csv,'w')
    sheet = workbook.sheet_by_index(0) # python index 0 = worksheet 1

    # The header row is typically row 5 but occasionally another row (eg. ironsteel.xls)
    # Search for the header row by looking for 'Year' in the first column.
    for row in range(0,200):
        if sheet.row_values(row)[0] == 'Year':
            header_row = row
            break;

    # Get all the titles and create associated names.
    # Harmonize non-standard names where appropriate.
    titles = sheet.row_values(header_row)
    colhi = len(titles)
    names = []
    for col in range(0,colhi):
        title = titles[col]
        title = title.strip()                          # remove leading/following whitespace
        title = re.sub("\s" , " ", title)              # replace any whitespacae with a single space
        title = re.sub("\s+" , " ", title)             # replace multilpe spaces with a single space
        titles[col] = '"' + title + '"' 
        if mineral=='stonecrushed-use' and col == 1:
            titles[col] = '"Coarse aggregate"' 
            names.append('coarse_aggregate')
        elif mineral=='stonecrushed-use' and col == 3:
            titles[col] = '"Fine aggregate"' 
            names.append('fine_aggregate')
        else:
            names.append(title.lower().replace(' ','_').replace(',','_').replace('(','_').replace(')','_'))

    titles_string = ','.join(titles)
    names_string = ','.join(names) 

    print("Working on " + mineral_csv)
    """
  # debugging lines
  print(titles_string)
  print(names_string)
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

    # Fix up issues with problem files

    # TODO:  boron-use.xls is a mess because 'Year' and all the titles appear in row 5 EXCEPT for 'Fire retardants'
    # TODO:  that have subtitles in row 6.  The data start in row 7.  Don't include it for now.
    # TODO:  cement-use.xls has the same issues
    # TODO:  clayskaolin-use.xls has the same issues
    # TODO:  claysmisc-use.xls has the same issues
    # TODO:  gypsum-use.xls has the same issues
    # TODO:  ironsteelslag-use.xls has the same issues
    # TODO:  mercury-use.xls has the same issues
    # TODO:  sandgravelindustrial-use.xls has the same issues
    # TODO:  sodaash-use.xls has the same issues
    # TODO:  stonedimension-use.xls has the same issues

    if mineral == 'claysbentonite-use':
        header_row += 1

    csv.write("DC.title      = ASCII CSV version of " + mineral + ".xls file\n")
    csv.write("file URL      = http://mazamascience.com/Minerals/USGS/DS140/2011/" + mineral + ".csv\n")
    csv.write("original data = http://minerals.usgs.gov/ds/2005/140/" + mineral + ".xls\n")
    csv.write("units         = metric tons\n")
    csv.write("\n")
    csv.write(titles_string + "\n")
    csv.write(names_string + "\n")

    # Check first year and fill in missing values if first year > 1975
    year = 1975
    first_year = sheet.row_values(header_row+1)[0]
    while (year < first_year):
        for col in range(0,colhi):
            if col == 0:
                csv.write('%d' % year)
            else:
                csv.write(',"na"')
        csv.write('\n')
        year += 1


    # Data begin after the header_row and continue for up to current_year-1975 years
    # We will check the type of the first column to determine when to stop ingesting data
    for row in range(header_row+1,200):

        # Stop ingesting data if you run out of rows
        try:
            types = sheet.row_types(row)
            values = sheet.row_values(row)
        except IndexError:
            break;

        # Stop ingesting data when the Year column no longer contains numbers.
        if types[0] == xlrd.XL_CELL_NUMBER:
            year = values[0]
        else:
            break;

        # From help(xlrd):
        #   XL_CELL_EMPTY   = 0
        #   XL_CELL_TEXT    = 1
        #   XL_CELL_NUMBER  = 2
        #   XL_CELL_DATE    = 3
        #   XL_CELL_BOOLEAN = 4
        #   XL_CELL_ERROR   = 5
        #   XL_CELL_BLANK   = 6

        # Validate and fix the data in the cells of this row.
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

                # Special case for various worksheets which use 'W' for 'Withheld' values
                elif value == 'W':
                    values[col] = 'na'
                    types[col] = xlrd.XL_CELL_TEXT
                    print("Converting row %d, column %d to 'na'" % (row,col))

                # Special case for strontium.xls (and others) which uses 'NA' instead of XL_CELL_EMPTY
                elif value == 'NA':
                    values[col] = 'na'
                    types[col] = xlrd.XL_CELL_TEXT
                    print("Converting row %d, column %d to 'na'" % (row,col))

                # Everything else we convert to float
                else: 
                    if value == 'NA':
                        values[col] = 'na'
                        types[col] = xlrd.XL_CELL_TEXT
                        print("Converting row %d, column %d to 'na'" % (row,col))
                    elif value == 'W':
                        values[col] = 'na'
                        types[col] = xlrd.XL_CELL_TEXT
                        print("Converting row %d, column %d to 'na'" % (row,col))
                    else:
                        try:
                            values[col] = float(value)
                            types[col] = xlrd.XL_CELL_NUMBER
                        except:
                            print("Cannot convert value '%s' to float in row %d, col %d" % (value,row+1,col+1))

            # Cells with numbers do not need conversion.
            elif types[col] == xlrd.XL_CELL_NUMBER:
                pass

            # Cells of any other type are considered errors.
            else:
                print("UNKNOWN data type in row %d, col %d" % (row,col))
                print("    cell type = " + str(types[col]))
                sys.exit(1)
        """
    # Special case for beryllium which is missing the row for the year 2000.
    # Before we write out the results for 2001, insert the 2000 results -- all missing values
    if (mineral == 'beryllium') and (year == 2001):
      csv.write("2000")
      for col in range(1,colhi):
        csv.write(",\"na\"")
      csv.write("\n")
    """
        # Print out the values with appropriate formatting.
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

    # Check first year and fill in missing values if first year > 1975
    last_year = 2020
    year += 1
    while (year <= last_year):
        for col in range(0,colhi):
            if col == 0:
                csv.write('%d' % year)
            else:
                csv.write(',"na"')
        csv.write('\n')
        year += 1

    print("Finished with " + mineral + " workbook.")


################################################################################

def main():

    logfile = open('Mazama_2011.log', 'w')

    # Convert the 'mineral.xls' files
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
    ### convert_file('columbium',logfile) ### Switched to 'niobium'
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
    convert_file('niobium',logfile)
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

    # Convert the 'mineral-use.xls' files
    convert_use_file('aluminum-use',logfile)
    convert_use_file('antimony-use',logfile)
    convert_use_file('arsenic-use',logfile)
    convert_use_file('asbestos-use',logfile)
    convert_use_file('bauxite-use',logfile)
    convert_use_file('beryllium-use',logfile)
    convert_use_file('bismuth-use',logfile)
    ###convert_use_file('boron-use',logfile)
    convert_use_file('cadmium-use',logfile)
    ###convert_use_file('cement-use',logfile)
    convert_use_file('chromium-use',logfile)
    convert_use_file('claysball-use',logfile)
    convert_use_file('claysbentonite-use',logfile)
    convert_use_file('claysfire-use',logfile)
    convert_use_file('claysfullers-use',logfile)
    ###convert_use_file('clayskaolin-use',logfile)
    ###convert_use_file('claysmisc-use',logfile)
    convert_use_file('cobalt-use',logfile)
    convert_use_file('columbium-use',logfile)
    convert_use_file('copper-use',logfile)
    convert_use_file('diamondindustrial-use',logfile)
    convert_use_file('diatomite-use',logfile)
    convert_use_file('feldspar-use',logfile)
    convert_use_file('fluorspar-use',logfile)
    convert_use_file('gallium-use',logfile)
    convert_use_file('garnet-use',logfile)
    convert_use_file('germanium-use',logfile)
    convert_use_file('gold-use',logfile)
    convert_use_file('graphite-use',logfile)
    ###convert_use_file('gypsum-use',logfile)
    convert_use_file('helium-use',logfile)
    convert_use_file('indium-use',logfile)
    convert_use_file('ironore-use',logfile)
    convert_use_file('ironoxide-use',logfile)
    ###convert_use_file('ironsteelslag-use',logfile)
    convert_use_file('ironsteel-use',logfile)
    convert_use_file('lead-use',logfile)
    convert_use_file('lime-use',logfile)
    convert_use_file('magnesiumcompounds-use',logfile)
    convert_use_file('magnesium-use',logfile)
    convert_use_file('manganese-use',logfile)
    ###convert_use_file('mercury-use',logfile)
    convert_use_file('mica-use',logfile)
    convert_use_file('molybdenum-use',logfile)
    convert_use_file('nickel-use',logfile)
    convert_use_file('nitrogen-use',logfile)
    convert_use_file('peat-use',logfile)
    convert_use_file('perlite-use',logfile)
    convert_use_file('phosphate-use',logfile)
    convert_use_file('pumice-use',logfile)
    convert_use_file('salt-use',logfile)
    convert_use_file('sandgravelconstruction-use',logfile)
    ###convert_use_file('sandgravelindustrial-use',logfile)
    convert_use_file('selenium-use',logfile)
    convert_use_file('silicon-use',logfile)
    convert_use_file('silver-use',logfile)
    ###convert_use_file('sodaash-use',logfile)
    convert_use_file('stonecrushed-use',logfile)
    ###convert_use_file('stonedimension-use',logfile)
    convert_use_file('strontium-use',logfile)
    convert_use_file('sulfur-use',logfile)
    convert_use_file('talc-use',logfile)
    convert_use_file('tantalum-use',logfile)
    convert_use_file('tellurium-use',logfile)
    convert_use_file('tin-use',logfile)
    convert_use_file('titaniumdioxide-use',logfile)
    convert_use_file('titanium-use',logfile)
    convert_use_file('tungsten-use',logfile)
    convert_use_file('vanadium-use',logfile)
    convert_use_file('zinc-use',logfile)

################################################################################

if __name__ == "__main__":
    main()

