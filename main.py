import json
from datetime import datetime

import pubmed_parser as pp
from Bio import Entrez
from openpyxl import load_workbook, Workbook

EMAIL = 'your.email@example.com'
FILE_NAME = 'input.xlsx'


def search(query):
    Entrez.email = EMAIL
    handle = Entrez.esearch(db='pubmed',
                            sort='pub+date',
                            retmax='200',
                            retmode='xml',
                            mindate='2020/01/01',
                            maxdate=datetime.today().strftime('%Y/%m/%d'),
                            datetype='pdat',
                            term=query)
    results = Entrez.read(handle)
    return results


def fetch_details(id_list):
    ids = ','.join(id_list)
    Entrez.email = EMAIL
    handle = Entrez.efetch(db='pubmed', retmode='xml', id=ids)
    results = Entrez.read(handle)
    return results


def read_authors(file_name):
    wb = load_workbook(file_name)

    author_ws = wb['List 1_by names']
    author_list = []
    for i in range(1, author_ws.max_row + 1):
        if isinstance(author_ws.cell(row=i, column=1).value, int):
            author_list.append(
                author_ws.cell(row=i, column=4).value + '[Author]')

    return author_list


def read_institutes(file_name):
    wb = load_workbook(file_name)

    institute_ws = wb['List 2_by Institute']
    institute_list = []

    for i in range(1, institute_ws.max_row + 1):
        if isinstance(institute_ws.cell(row=i, column=1).value, int):
            institute_list.append(
                institute_ws.cell(row=i, column=2).value + '[Affiliation]')

    return institute_list


if __name__ == '__main__':

    authors = read_authors(FILE_NAME)
    institutes = read_institutes(FILE_NAME)

    # Get Authors
    #results = search(' OR '.join(authors))

    # Get Organisations
    results = search(' OR '.join(institutes))

    if results['IdList']:
        id_list = results['IdList']
        papers = fetch_details(id_list)

        wb = Workbook()
        ws1 = wb.active
        ws1.title = 'KKH'

        for i, paper in enumerate(papers['PubmedArticle']):

            _ = ws1.cell(column=1, row=i+1, value="{0}".format(paper['MedlineCitation']['Article']['AuthorList']))
            _ = ws1.cell(column=2, row=i+1, value="{0}".format(paper['MedlineCitation']['Article']['ArticleTitle']))
            _ = ws1.cell(column=3, row=i+1, value="{0}".format(paper['MedlineCitation']['Article']['Journal']))
            _ = ws1.cell(column=4, row=i+1, value="{0}".format(paper['MedlineCitation']['Article']['ELocationID']))


            #print(json.dumps(paper, indent=2))

            #print(json.dumps(papers[0], indent=2, separators=(',', ':')))

        wb.save(filename='output.xlsx')

    else:
        print('Nothing found!')
