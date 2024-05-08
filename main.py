import logging
import re
import urllib
import urllib.request
import atexit

import openpyxl

headers = {
    'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) '
                 'AppleWebKit/537.11 (KHTML, like Gecko) '
                 'Chrome/23.0.1271.64 Safari/537.11',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Charset': 'ISO-8859-1,utf-8;q=0.7,*;q=0.3',
    'Accept-Encoding': 'none',
    'Accept-Language': 'en-US,en;q=0.8',
    'Connection': 'keep-alive'
}


def read_keywords_from_xlsx(filename):
    sheet1 = openpyxl.load_workbook(filename, data_only=True)["data"]
    keywords1 = []
    for row1 in sheet1.values:
        row_filtered = [value for value in row1[:3] if value is not None]  # Фильтрация None
        if len(row_filtered) >= 2:  # Проверяем, есть ли хотя бы 2 значения
            keywords1.append(row_filtered)
        if len(row_filtered) < 2:
            return keywords1
    return keywords1


def add_column_to_row(row, column):
    if column is not None:
        row.append(column)
    return row


def get_filter(shorthand):
    if shorthand == "line" or shorthand == "linedrawing":
        return "+filterui:photo-linedrawing"
    elif shorthand == "photo":
        return "+filterui:photo-photo"
    elif shorthand == "clipart":
        return "+filterui:photo-clipart"
    elif shorthand == "gif" or shorthand == "animatedgif":
        return "+filterui:photo-animatedgif"
    elif shorthand == "transparent":
        return "+filterui:photo-transparent"
    else:
        return ""


def run(keywords, looking_limit, adult, filter='', badsites=[]):
    linkResult = ""
    looking_count = 0
    seen = set()
    while looking_count < looking_limit:
        try:
            request_url = (
                    'https://www.bing.com/images/async?q='
                    + urllib.parse.quote_plus(keywords)
                    + '&count=' + str(looking_limit)
                    + '&adlt=' + adult
                    + '&qft=' + ('' if filter is None else get_filter(filter))
            )
            request = urllib.request.Request(request_url, None, headers=headers)
            response = urllib.request.urlopen(request)
            html = response.read().decode('utf8')

            if html == "":
                break
            links = re.findall('murl&quot;:&quot;(.*?)&quot;', html)

            for link in links:
                isbadsite = False
                for badsite in badsites:
                    isbadsite = badsite in link
                    if isbadsite:
                        break
                if isbadsite:
                    continue

                if looking_count < looking_limit and link not in seen:
                    linkResult = link
                    seen.add(link)
                    looking_count += 1

        except urllib.error.HTTPError as e:
            logging.error('URLError while making request to Bing: %s', e)
        except urllib.error.URLError as e:
            logging.error('URLError while making request to Bing: %s', e)
    return linkResult


keywords = read_keywords_from_xlsx("dataset200.xlsx")
looking_limit = 1
adult = "off"
badsites = []
workbook = openpyxl.Workbook()
sheet = workbook.active
len_keywords = len(keywords)

for i in range(len_keywords):
    row = keywords[i]
    print("process: " + str(i) + " / " + str(len_keywords))
    keyword_phrase = " ".join(str(value) for value in row)
    link_of_the_photo = run(keyword_phrase, looking_limit, adult, None, badsites)
    modified_row = add_column_to_row(row.copy(), link_of_the_photo)
    sheet.append(modified_row)


def close_file(some_workbook):
    some_workbook.save("result.xlsx")
    some_workbook.close()


atexit.register(close_file, workbook)