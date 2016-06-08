#!/usr/bin/env python
# -*- coding: utf-8 -*-

__author__ = "Martin Rydén"
__copyright__ = "Copyright 2016, Martin Rydén"
__license__ = "GNU"
__version__ = "1.0.1"
__email__ = "pemryd@gmail.com"

### Reddit bot config ###
import praw
import re
import os
from config_bot import *
import sys

sys.stdout = open("out.txt", "w")

# Check that the file that contains our username exists
if not os.path.isfile("config_bot.py"):
    print ("You must create a config file with your username and password.")
    print ("Please see config_skel.py")
    exit(1)

# Create the Reddit instance
user_agent = ("xltldr 0.1")
r = praw.Reddit(user_agent=user_agent)

# and login
r.login(REDDIT_USERNAME, REDDIT_PASS, disable_warning=True)

# Have we run this code before? If not, create an empty list
if not os.path.isfile("posts_replied_to.txt"):
    print("post_replied_to does not exist")
    posts_replied_to = []

# If we have run the code before, load the list of posts we have replied to
else:
    # Read the file into a list and remove any empty values
    with open("posts_replied_to.txt", "r") as f:
        posts_replied_to = f.read()
        posts_replied_to = posts_replied_to.split("\n")
        posts_replied_to = list(filter(None, posts_replied_to))

subreddit = r.get_subreddit('pythonforengineers')
subreddit_comments = subreddit.get_comments(limit=100)

c_formula = ""
for comment in subreddit_comments:
    if comment.id not in posts_replied_to:
        print("Comment not replied to: %s" % comment, comment.id)
        if comment.body.startswith("!formulate"):
            c_formula = comment.body.replace("!formulate", "")
            c_formula = ''.join(list(filter(None, c_formula))).strip()
            break

if(len(c_formula) < 2):
    sys.exit("No comments found.")
    
### All set, let's begin ###

import requests
from bs4 import BeautifulSoup
from GoogleScraper import scrape_with_config, GoogleSearchError
import collections
import time
import logging
logging.basicConfig(level=logging.DEBUG)

tries = 0
maxchars = 250
# Input Excel formula
while True:
    formula = c_formula#input('\nFunction: ')
    if(formula[0] != '=') or (len(formula) > maxchars):
        if(tries == 0):
            print('Wrong input format. Try again.')
            tries += 1
        elif(tries == 1):
            print('The formula must begin with "=" and cannot exceed %s characters.' % maxchars)
            tries += 1
        else:
            break
    else:
        break

startTime = time.time()

# Split by and keep any non-alphanumeric delimiter, filter blanks
def split_formula(f):
    split = list(filter(None, re.split('([^\\w.":!$<>%&^/\s])', f)))
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
variables = re.compile('([\w\.":!$<>%&^/\s]+)',re.I)
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

# Sites / keywords to exclude
unwanted= ['youtube',
           'stackoverflow',
           '"stack exchange"',
           'site:knowexcel.com'
           #'site:social.technet.microsoft.com'
           ]
unwanted = ' -'.join(unwanted)

#### Scrape google for top hits ####
# sm = ['http-async', 'http']

config = {
    'use_own_ip': 'True',
    'keyword': ('excel formula -%s %s') % (unwanted, wcf),
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
numf = len(found_functions)


if(numf <= 2):
    max_results = (12 * numf)
elif(numf > 2 and numf <= 5):
    max_results = (10 * numf)
else:
    max_results = 50

r = 0
# Results - append URL for each hit to urls list
for serp in search.serps:
    for link in serp.links:
        if(r < max_results):
            urls.append(link.link)
            r += 1
        
#### Parse with BS4 ####

# This dict will be used to store each hit with an id as key
# and chosen web element, its URL, and a total score (sum of found elements)
ranking = collections.defaultdict(dict)

# Web elements to look for
elements = ['pre', 'p', 'ul', 'td', 'h1', 'h2', 'h3', 'h4', 'xl-formula']


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
        soup = BeautifulSoup(r.text.encode('utf-8'), "lxml") # Parses the page
    
        # Iterate through each chosen element, which are counted
        # using the find_elements function
        for e in elements:
            find_elements(e)
       
        try:
            stitle = soup.title.string.encode('utf-8')
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
        continue # Skips sites with SSL cert errors

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

nf = '' # new/scraped formula

# This function 
def getAttribute(attr, f, t):
    try:
        return getattr(attr, f)(t)
    except:
        pass

# Parses data of input url and checks hits for each specified element
# If a paragraph contains every function in the original formula, set nf to p
def get_data(url):
    r = requests.get(url)
    soup = BeautifulSoup(r.text.encode('utf-8'), "lxml") # Parses the page
    for p in (soup.find_all()): # Iterate through every page element
        if(all(functions in p.getText() for functions in found_functions)): # Check if current element contains function names
            nf_found_functions = [] # Here we create a new function list for the current text
            global nf
            nf = p.getText()
            nf_dl_formula = split_formula(nf.replace(' ', '')) # Split the formula and get rid of whitespaces
            find_functions(nf_dl_formula, nf_found_functions) # Add function names to list
#            print(url)
#            print(nf[0])
#            print(len(found_functions))
#            print(len(nf_found_functions))
#            print(set(found_functions))
#            print(set(nf_found_functions))  
#            print(nf.encode('utf-8'))   
            
            # The next requirements are that the new function list and the original function list need to contain the same
            # names, and be the same length. The text need to begin with an equal sign
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
                
                num_p = 5 # Set number of paragraphs before and after p to collect
                plist = [] # List to store paragraph +- num_p paragraphs
                
                prevtext = p
                nexttext = p
                
                # Find previous and next num_p paragraphs, add to list
                for i in range(0,num_p):
                    prevtext = getAttribute(prevtext, 'findPrevious', elements)
                    try:                    
                        plist.insert(0, prevtext.getText()).encode('utf-8')
                    except:
                        pass        
                try:
                    plist.append(p.getText()).encode('utf-8')
                except:
                    pass
                for i in range(0,num_p):
                    nexttext = getAttribute(nexttext, 'findNext', elements)
                    try:
                        plist.append(nexttext.getText()).encode('utf-8')
                    except:
                        pass
                # Join list of paragraphs, replace variables with original vars, if possible
                try:                
                    excerpt = '\n'.join(plist)
                    var_dict = {}  
                    global result                    
                    
                    if(len(dl_type['var']) == len(nf_dl_type['var'])):
                            
                        for i,v in enumerate(nf_dl_type['var']):
                            var_dict[v] = dl_type['var'][i]

                        xpattern = re.compile('|'.join(var_dict.keys()))
                        result = xpattern.sub(lambda x: var_dict[x.group()], excerpt)
                    else:
                        result = excerpt
              
                    # Post text to specified reddit comment
                    result = ("Excerpt from %s\n%s" % (top_hit_url, result))
                    
                    comment.reply(result)
                    print(result)
                    print(comment.id)

                    # Write comment id to file
                    with open("posts_replied_to.txt", "a") as f:
                        f.write(comment.id + "\n")
                    f.close()
     
                    return True                
                    break
                except Exception as e:
                    print(e)

    global current_rank
    current_rank += 1
    return False

# Get the top ranked url
current_rank = 0
if(len(sorted_ranking) > 0):
    top_hit_url = get_ranked_url(current_rank)
else:
    sys.exit("No matches found.")

# Get the top ranked url, iterate for max results
look_for_text = False
while(look_for_text) == False:
    if(current_rank < max_results):
        try:
            top_hit_url = get_ranked_url(current_rank)
            print('\nSelected web page: %s' % top_hit_url)
            print('\n')
            look_for_text = get_data(top_hit_url)
        except IndexError:
            print("Unable to find feasible match in top %s scrapes." % max_results)
            break
    else:
        print("Unable to find feasible match in top %s scrapes." % max_results)
        break

endTime = time.time()
duration = (endTime - startTime)
print("\nTime elapsed: %s seconds." %round(duration,0))
