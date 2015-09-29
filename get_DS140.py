#!/usr/bin/env python
#
# Python script to download all Excel spreadsheets that make up the USGS dataset:
#   "Historical Statistics for Mineral Commodoties in the United States, Data Series 2005-140"

import urllib
from BeautifulSoup import BeautifulSoup

location = "http://minerals.usgs.gov/ds/2005/140/"
page = urllib.urlopen(location)
soup = BeautifulSoup(page)
    
# Find every occurrence of <a href="...">XLS</a> and download the file pointed to by href="...".
for link in soup.findAll('a'):
    if link.string == 'XLS':
        filename = link.get('href')
        print("Retrieving " + filename)
        url = location + filename
        urllib.urlretrieve(url,filename)
