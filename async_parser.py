import asyncio
import logging
import re
import urllib.parse
import atexit
from multiprocessing import Pool

import aiohttp
import openpyxl


headers = {
    'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) '
                  'AppleWebKit/537.11 (KHTML, like Gecko) '
                  'Chrome/23.0.1271.64 Safari/537.11',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Charset': 'ISO-8859-1,utf-8;q=0.7,*;q=0.3',
    'Accept-Encoding': 'none',
    'Accept-Language': 'en-US,en;q=0.8',
    'Connection': 'close'
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


def get_filter(self, shorthand):
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


async def run(keywords, looking_limit, adult, filter='', badsites=[]):
    linkResult = ""
    looking_count = 0
    seen = set()

    async with aiohttp.ClientSession(trust_env=True) as session:
        while looking_count < looking_limit:
            try:
                request_url = (
                        'https://www.bing.com/images/async?q='
                        + urllib.parse.quote_plus(keywords)
                        + '&count=' + str(looking_limit)
                        + '&adlt=' + adult
                        + '&qft=' + ('' if filter is None else get_filter(filter))
                )

                async with session.get(request_url, headers=headers, timeout=3) as response:
                    html = await response.text()

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

            except aiohttp.ClientError as e:
                logging.error('Error while making request to Bing: %s', e)
    return linkResult


async def process_row(row, semaphore, looking_limit, adult, badsites):
    async with semaphore:
        try:
            print(f"Response row: {row}")
            keyword_phrase = " ".join(str(value) for value in row)
            link_of_the_photo = await run(keyword_phrase, looking_limit, adult, None, badsites)
            return add_column_to_row(row.copy(), link_of_the_photo)
        except Exception as e:
            logging.error(f"Error processing row {row}: {e}")


async def main():
    max_concurrent_requests = 50
    semaphore = asyncio.Semaphore(max_concurrent_requests)
    keywords = read_keywords_from_xlsx("dataset2000.xlsx")
    looking_limit = 1
    adult = "off"
    badsites = []
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    len_keywords = len(keywords)
    with Pool() as pool:
        results = await asyncio.gather(*[
            process_row(row, semaphore, looking_limit, adult, badsites)
            for row in keywords
        ])

    for i, modified_row in enumerate(results):
        try:
            print(f"Process: {i + 1} / {len_keywords}")
            sheet.append(modified_row)
        except Exception as e:
            logging.error('Error while try bind row: %s', e)

    def close_file(some_workbook):
        some_workbook.save("result.xlsx")
        some_workbook.close()

    atexit.register(close_file, workbook)


if __name__ == "__main__":
    asyncio.run(main())