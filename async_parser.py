import asyncio
import logging
import re
import urllib.parse
import atexit
import yaml
from multiprocessing import Pool

import aiohttp
import openpyxl


def load_config(filename):
    with open(filename, 'r') as f:
        return yaml.safe_load(f)


config = load_config("config.yaml")


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


async def run(row, looking_limit, adult, filter='', badsites=[]):
    link_result = ""
    looking_count = 0
    seen = set()
    headers = config['bing']['headers']
    keyword_phrase = " ".join(str(value) for value in row)

    async with aiohttp.ClientSession(trust_env=True) as session:
        while looking_count < looking_limit:
            try:
                request_url = (
                        'https://www.bing.com/images/async?q='
                        + urllib.parse.quote_plus(keyword_phrase)
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
                            link_result = link
                            seen.add(link)
                            looking_count += 1

            except aiohttp.ClientError as e:
                logging.error('Error while making request to Bing: %s', e)
    return link_result


async def process_row(row, semaphore, looking_limit, adult, badsites, filter):
    async with semaphore:
        try:
            print(f"Response row: {row}")
            link_of_the_photo = await run(row, looking_limit, adult, filter, badsites)
            return add_column_to_row(row.copy(), link_of_the_photo)
        except Exception as e:
            logging.error(f"Error processing row {row}: {e}")


async def main():
    max_concurrent_requests = config['concurrency']['max_concurrent_requests']
    keywords = read_keywords_from_xlsx(config['input_file'])
    looking_limit = config['bing']['looking_limit']
    adult = config['bing']['adult_filter']
    badsites = config['bing']['bad_sites']
    filter = config['bing']['filter']

    semaphore = asyncio.Semaphore(max_concurrent_requests)
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    with Pool() as pool:
        results = await asyncio.gather(*[
            process_row(row, semaphore, looking_limit, adult, badsites, filter)
            for row in keywords
        ])

    len_keywords = len(keywords)
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
