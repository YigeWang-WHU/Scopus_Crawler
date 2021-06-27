from elsapy import elsclient
from elsapy import elssearch
from elsapy import elsdoc
from elsapy import elsprofile
import json
import xlsxwriter
import requests
from urllib.parse import quote_plus as url_encode


# If you dont want to search author names, leave it with "" for each varaible
author_last = ""
author_first = ""

# define the time range you want to search
year_start = 2009  # >year_start, start from (year_start + 1)
assert year_start is not None

# define the database
database = 'scopus'  # 'scopus'

# define the keywords you want to search
# if you dont want to search with certain keywords, make it equal to ""

# Precedence rule: OR > AND > AND NOT. (OR logic connect first and then AND and finally AND NOT)

# search articles contain "explosion" and "wall" but may appear separately : "explosion AND wall"

# search articles contain "explosion" or "wall" : "explosion OR wall"

# if you want to search phrases consisting of multiple words e.g. contact detonation
# instead of individual words, use {} to quote the phrase, e.g. {contact detonation} OR {contact explosion}

keywords = "{contact detonation} OR {contact explosion}"

# define the journal/conference you want to search in, you dont need to give the exact conference/journal name.
# if you dont want to search in certain journals or conferences, make it equal to ""
jour_conf = "{Advanced Materials Research}"

# Load API key
# 1. 'https://dev.elsevier.com/apikey/create': create an APIKey
# 2. save an .txt file as 'config.json' under the guidance of this webpage: \
# https://github.com/ElsevierDev/elsapy/blob/master/CONFIG.md
con_file = open("config.json")
config = json.load(con_file)
con_file.close()

# initialize search client
client = elsclient.ElsClient(config['apikey'])


# Retrieve author id by quering author name
author_id = None
if author_last and author_first: # Only if users provide valie author names
    auth_srch = elssearch.ElsSearch('AUTHLASTNAME(%s)'%author_last + ' AUTHFIRST(%s)'%author_first,'author')
    auth_srch.execute(client)
    print ("Found ", len(auth_srch.results), " authors \n")
    print('{:<6} {:<6} {:<6}'.format('Index |', 'First name |', 'Last name |'))
    print('-'*80)
    for i, author in enumerate(auth_srch.results):
        #let's look on every author and print the name and affiliaiton stored in Scopus  
        index = str(i)
        author_id = author['dc:identifier'].split(':')[1]
        first_name_scopus = author['preferred-name']['given-name']
        last_name_scopus = author['preferred-name']['surname']
        
        print('{:<12} {:<12} {:<11}'.format(index, first_name_scopus, last_name_scopus))

    if len(auth_srch.results) > 1:
        # If multiple authors are found, users are responsible for specifying one of them

        selection = input('Please input the index of the author you want to search for:')
        author_id = auth_srch.results[int(selection)]['dc:identifier'].split(':')[1]
    

# search dictionary
search_dictionary = {'year_start': year_start, 'author_id': author_id, 'jour_conf': jour_conf, 'keywords':keywords}
# define whether get all documents or not (only for the first 25 documents)
get_all_key = True


# initialize doc search object using Scopus or ScienceDirect, execute search, \
# and retrieve all results
# DOC:
# https://dev.elsevier.com/technical_documentation.html
# https://dev.elsevier.com/sd_article_meta_tips.html
# ScienceDirect Article Metadata Guide
# https://dev.elsevier.com/sd_article_meta_tips.html
# Scopus Search Guide
# https://dev.elsevier.com/sc_search_tips.html

if database == "scopus":
    search_language = ''
    for k,v in search_dictionary.items():
        if k == 'year_start' and v:
            search_language += "PUBYEAR > " + str(v) 
        elif k == 'author_id' and v:
            search_language += " AND AU-ID(" + v + ")"
        elif k == 'jour_conf' and v:
            search_language += " AND SRCTITLE(" + v + ")"
        elif k == 'keywords' and v:
            search_language += " AND Title(" + v + ")" \
                + " OR KEY(" + v + ")" \
                + " OR ABS(" + v + ")"

doc_srch = elssearch.ElsSearch(search_language, database)
doc_srch.execute(client, get_all=get_all_key)
print("\nThe number of papers found: {}".format(len(doc_srch.results)))

if len(doc_srch.results) > 500:  # for saving time
    print("Document search has", len(doc_srch.results), "results. \
    Please refine your searching strategy")
else:
    # record the results
    # open a new xlsx file
    workbook = xlsxwriter.Workbook('Test_elsevier.xlsx')
    worksheet = workbook.add_worksheet()
    row = 0
    col = 0

    # define the contents; could be updated later
    col_names = ['Title', 'Author(s)', 'Data', 'Journal', 'DOI']

    for name in col_names:
        worksheet.write(row, col, name)
        col += 1

    # move to the first observation row
    row = 1

    for count, paper in enumerate(doc_srch.results):  # each paper
        title = authors = journal = data = doi = 'No values'
        # take the results

        try:
            
            title = paper['dc:title']
            #authors = paper['dc:creator']
            uri = paper['prism:url']
            journal = paper['prism:publicationName']
            data = paper['prism:coverDate']
            doi = 'https://doi.org/' + paper['prism:doi']
        
            # Retrieve all authors for that document
            
            paper_id = paper['dc:identifier']
            url = ("http://api.elsevier.com/content/abstract/scopus_id/" \
                + paper_id \
                + "?field=authors")
            resp = requests.get(url, headers={'Accept':'application/json', 'X-ELS-APIKey': config['apikey']})
            results = json.loads(resp.text.encode('utf-8'))
            authors=', '.join([au['ce:indexed-name'] for au in results['abstracts-retrieval-response']['authors']['author']])
            
        except:
            pass

        related_section_text = None

        # define the contents
        if related_section_text:
            items = [title, authors, data, journal, doi, related_section_text]
        else:
            items = [title, authors, data, journal, doi]

        # record the information in the excel file
        col = 0
        for item in items:
            worksheet.write(row, col, item)
            col += 1
        row += 1

    # end the excel file recording
    workbook.close()


