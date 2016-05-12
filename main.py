#!/usr/bin/env python

__author__ = "Martin Rydén"
__copyright__ = "Copyright 2016, Martin Rydén"
__license__ = "GNU"
__version__ = "1.0.1"
__email__ = "pemryd@gmail.com"

import re
import requests
from bs4 import BeautifulSoup
from GoogleScraper import scrape_with_config, GoogleSearchError
import collections
import time

tries = 0
# Input Excel formula
while True:
    formula = input('\nFunction: ')
    if(formula[0] != '=') or (len(formula) > 150):
        if(tries == 0):
            print('Wrong input format. Try again.')
            tries += 1
        elif(tries == 1):
            print('The formula must begin with "=" and cannot exceed 150 characters.')
            tries += 1
        else:
            break
    else:
        break
#formula = '=INDEX(A1:Z10,MATCH("product",A1:A10,0),MATCH("month",A1:Z1,0))'

startTime = time.time()

# Split by and keep any non-alphanumeric delimiter, filter blanks
def split_formula(f):
    split = list(filter(None, re.split('([^\\w.":!$<>%&/\s])', f)))
    return split
    
dl_formula = split_formula(formula)


# List of functions
functions = []

def find_functions(fl, flist):
    with open('excel_functions.txt') as f:
        lines = f.read().splitlines()
        for fc in fl:
            for f in lines:
                if(fc == f):
                    flist.append(fc)

find_functions(dl_formula, functions)

# Set regex pattern:
# variables: any non-alpha-numeric char, with exceptions
# separators: any non-alpha-numeric char
# tbd: take another look at these and figure out wtf is actually going on here
variables = re.compile('([\w\.":!$<>%&/\s]+)',re.I)
separators = re.compile(r'^\W+',re.I)

#tbd: add operators:
# = (equal sign), >= (greater than or equal to sign), <= (less than or equal to sign),
# <> (not equal to sign),  (space)

# Dictionary to keep track of which formula element is const, var, or sep
dl_type = collections.defaultdict(list)

wc = [] # Same as dl_formula but vars substituted for wilcards
found_functions = [] # Store elements of formula which are known functions
sep = [] # List of separators in formula

# Appends formula elements according to above description
for f in dl_formula:
    if f in functions:
        wc.append(f)
        found_functions.append(f)
    else:
        a = variables.sub("*", f)
        wc.append(a)
        if(not re.search(variables, f)):
            sep.append(f)

# Appends element type to dict with appropriate element type key
for f in dl_formula:
    if(f in found_functions):
        dl_type['const'].append(f)
    elif(f in sep):
        dl_type['sep'].append(f)
    else:
        dl_type['var'].append(f)

# Join the wilcard formula, add an extra wildcard for good luck
wcf = ''.join(wc)+"*"

print('\nSearching for substituted formula:\n%s' % wcf)

#### Scrape google for top hits ####

config = {
    'use_own_ip': 'True',
    'keyword': ('excel formula -youtube -stackoverflow -"stack exchange" -site:knowexcel.com %s') % wcf,
    'search_engines': ['duckduckgo'],
    'num_pages_for_keyword': 1,
    'scrape_method': 'http-async',
    'do_caching': 'True',
    'print_results': 'summarize'
}

urls = [] # List of result URLs

# Begin scraping
try:
    search = scrape_with_config(config)

except GoogleSearchError as e:
    pass

# Manually set max urls generated, since the built-in function is a bit wonky
max_results = 25
r = 0
# Results - append URL for each hit to urls list
for serp in search.serps:
    for link in serp.links:
        if(r < max_results):
            urls.append(link.link)
            r += 1
        
#### Parse into BS4 ####

# This dict will be used to store each hit with an id as key
# and chosen web element, its URL, and a total score (sum of found elements)
ranking = collections.defaultdict(dict)

# Web elements to look for
elements = ['pre', 'p', 'ul', 'td', 'h1', 'h2', 'h3', 'h4']


# Searches a web page for an element, stores matches in dict
def find_elements(element):
        for p in (soup.find_all(element)):
            if(all(x in p.getText() for x in found_functions)):
                matches[element] += 1

web_id = 0

for url in urls:
    web_id += 1 # Gives an id to each web hit
    matches = collections.defaultdict(int) # New dict for each url
    try:
        r = requests.get(url)
        soup = BeautifulSoup(r.content, "html.parser") # Parses the page
    
        # Iterate through each chosen element, which are counted
        # using the find_elements function
        for e in elements:
            find_elements(e)
       
        try:
            stitle = soup.title.string
        except:
            stitle = "No title"
     
        print('\nFound matches in "%s".\n \
        URL: %s' % (stitle,url))
    
        ##
        
        # Adds each element an its number of matches to ranking dict
        # Also adds the URL for reference
        ranking[web_id] = (matches)
        ranking[web_id]['url'] = (url)
    except:
        continue


# Sums total count of elements per web hit into a score
# This score will used in decided which page is more likely
# to contain useful data
for k,v in ranking.items():
    score = 0
    for e in elements:
        score += v[e]
        ranking[k]['xscore'] = (score)

# Now, the ranking dict should look something like this:

# defaultdict(dict,
#             {1: defaultdict(int,
#                          {'p': 0,
#                           'pre': 0,
#                           'ul': 0,
#                           'url': 'http://www.example_url_1.com/',
#                           'xscore': 0}),
#              2: defaultdict(int,
#                          {'p': 1,
#                           'pre': 5,
#                           'ul': 0,
#                           'url': 'http://www.example_url_2.com/',
#                           'xscore': 6}),
#                           ...


sorted_ranking = sorted(ranking.items(),key=lambda k_v: k_v[1]['xscore'],reverse=True)

def get_ranked_url(rank):
    return sorted_ranking[rank][1]['url']

#top_hit_url = "https://www.ablebits.com/office-addins-blog/2014/08/13/excel-index-match-function-vlookup/"
#top_hit_url = "http://www.deskbright.com/excel/using-index-match-match/"

nf = '' # new/scraped formula

# Parses data of input url and checks hits for each specified element
# If a paragraph contains every function in the original formula, set nf to p
def get_data(url):
    r = requests.get(url)
    soup = BeautifulSoup(r.text, "html.parser") # Parses the page
    for p in (soup.find_all()):
        if(all(functions in p.getText() for functions in found_functions)):
            nf_found_functions = []
            global nf
            nf = p.getText()
            nf_dl_formula = split_formula(nf)
            find_functions(nf_dl_formula, nf_found_functions)         
            if(set(found_functions) == set(nf_found_functions)) and len(found_functions) == len(nf_found_functions) and nf[0] == "=":         
                # Dictionary to keep track of which formula element is const, var, or sep
                nf_dl_type = collections.defaultdict(list)

                # Appends element type to dict with appropriate element type key
                for f in nf_dl_formula:
                    if(f in nf_found_functions):
                        nf_dl_type['const'].append(f)
                    elif(not re.search(variables, f)):
                        nf_dl_type['sep'].append(f)
                    else:
                        nf_dl_type['var'].append(f)
                
                try:
                    # Fetch the match +- 5 paragraphs
                    excerpt = (p.findPrevious(elements).findPrevious(elements).findPrevious(elements).findPrevious(elements).getText())
                    excerpt += ("\n\n")
                    excerpt += (p.findPrevious(elements).findPrevious(elements).findPrevious(elements).getText())
                    excerpt += ("\n\n")
                    excerpt += (p.findPrevious(elements).findPrevious(elements).getText())
                    excerpt += ("\n\n")                
                    excerpt += (p.findPrevious(elements).getText())
                    excerpt += ("\n\n")
                    excerpt += (p.getText())
                    excerpt += ("\n\n")
                    excerpt += (p.findNext(elements).getText())
                    excerpt += ("\n\n")
                    excerpt += (p.findNext(elements).findNext(elements).getText())
                    excerpt += ("\n\n")
                    excerpt += (p.findNext(elements).findNext(elements).findNext(elements).getText())
                    excerpt += ("\n\n")
                    excerpt += (p.findNext(elements).findNext(elements).findNext(elements).findNext(elements).getText())
                    excerpt += ("\n\n")
                    excerpt += (p.findNext(elements).findNext(elements).findNext(elements).findNext(elements).findNext(elements).getText())
                    excerpt += ("\n\n")
                    excerpt += (p.findNext(elements).findNext(elements).findNext(elements).findNext(elements).findNext(elements).findNext(elements).getText())
                
                    var_dict = {}  
                    
                    for i,v in enumerate(nf_dl_type['var']):
                        var_dict[v] = dl_type['var'][i]
                    
                    xpattern = re.compile('|'.join(var_dict.keys()))                
                    result = xpattern.sub(lambda x: var_dict[x.group()], excerpt)
                    
                    print("Excerpt from %s\n" % top_hit_url)
                    print(result)
     
                    return True                
                    break
                except:
                    pass
    global current_rank
    current_rank += 1
    return False

current_rank = 0
top_hit_url = get_ranked_url(current_rank)

look_for_text = False
while(look_for_text) == False:
    if(current_rank < 5):
        top_hit_url = get_ranked_url(current_rank)
        print('\nSelected web page: %s' % top_hit_url)
        print('\n')
        look_for_text = get_data(top_hit_url)
    else:
        print("Unable to find feasible match in top %s scrapes." % max_results)
        break


endTime = time.time()
duration = (endTime - startTime)
print(duration)
