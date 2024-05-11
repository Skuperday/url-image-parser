if __name__ == '__main__':
    import multiprocessing

    multiprocessing.freeze_support()

import asyncio
import logging
import re
import urllib.parse
import atexit
from collections import OrderedDict
import plyer

import yaml
from multiprocessing import Pool

import aiohttp
import openpyxl
import itertools


def load_config(filename):
    with open(filename, 'r') as f:
        return yaml.safe_load(f)


config = load_config("config.yaml")


def combine_columns(keywords, columns):
    strict_columns = config['strict_columns']
    columns_offset = min(columns)
    columns_w_offset = [col - columns_offset for col in columns]
    strict_w_offset = [col - columns_offset for col in strict_columns]

    keywords_in_func = keywords.copy()
    for i in range(len(keywords_in_func)):
        if i in strict_w_offset:
            keywords_in_func[i] = f'"{keywords_in_func[i]}"'

    combinations = set()
    for combination_size in range(1, len(columns) + 1):
        for column_subset in itertools.combinations(columns_w_offset, combination_size):
            selected_keywords = [str(keywords_in_func[i]) for i in range(len(column_subset))]
            sorted_keywords = " ".join(sorted(selected_keywords))
            combinations.add(sorted_keywords)
    return list(combinations)


def read_keywords_from_xlsx(filename):
    sheet_name = config['sheet_name']
    cols_to_search = config['columns']

    sheet1 = openpyxl.load_workbook(filename, data_only=True)[sheet_name]
    keywords1 = []

    none_value_counter = 0
    for row1 in sheet1.values:
        if none_value_counter > 5:
            break
        row_filtered = []
        for col in cols_to_search:
            if row1[col] is not None:
                row_filtered.append(row1[col])
        if len(row_filtered) > 0:
            keywords1.append(row_filtered)
        else:
            none_value_counter += 1
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


async def run(row, looking_limit, adult, filter='', badsites=[], white_list=[]):
    seen = OrderedDict()
    headers = config['bing']['headers']
    columns = config['columns']

    async with aiohttp.ClientSession(trust_env=True) as session:
        combined_columns = combine_columns(row, columns)
        print(combined_columns)
        for keywords in combined_columns:
            try:
                request_url = (
                        'https://www.bing.com/images/async?q='
                        + urllib.parse.quote_plus(str(keywords))
                        + '&count=' + str(looking_limit)
                        + '&adlt=' + adult
                        + '&qft=' + ('' if filter is None else get_filter(filter))
                )
                async with session.get(request_url, headers=headers, timeout=3) as response:
                    html = await response.text()
                    for link in re.findall('murl&quot;:&quot;(.*?)&quot;', html):
                        if any(white_site in link for white_site in white_list):
                            seen[link] = seen.get(link, 0) + 10
                            break
                        if any(badsite in link for badsite in badsites):
                            continue
                        seen[link] = seen.get(link, 0) + 1  # Increment occurrence count

            except aiohttp.ClientError as e:
                logging.error('Error while making request to Bing: %s', e)
    best_link = max(seen, key=seen.get) if seen else ""
    return best_link


async def process_row(row, semaphore, looking_limit, adult, badsites, white_list, filter):
    async with semaphore:
        try:
            print(f"Response row: {row}")
            link_of_the_photo = await run(row, looking_limit, adult, filter, badsites, white_list)
            return add_column_to_row(row.copy(), link_of_the_photo)
        except Exception as e:
            logging.error(f"Error processing row {row}: {e}")


async def main():
    max_concurrent_requests = config['concurrency']['max_concurrent_requests']
    keywords = read_keywords_from_xlsx(config['input_file'])
    looking_limit = config['bing']['looking_limit']
    adult = config['bing']['adult_filter']
    black_list = config['bing']['black_list']
    white_list = config['bing']['white_list']
    filter = config['bing']['filter']

    semaphore = asyncio.Semaphore(max_concurrent_requests)
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    with Pool() as pool:
        results = await asyncio.gather(*[
            process_row(row, semaphore, looking_limit, adult, black_list, white_list, filter)
            for row in keywords
        ])

    len_keywords = len(keywords)
    errors = 0
    for i, modified_row in enumerate(results):
        try:
            print(f"Process: {i + 1} / {len_keywords}")
            sheet.append(modified_row)
        except Exception as e:
            errors += 1
            logging.error('Error while try bind row: %s', e)

    def close_file(some_workbook):
        some_workbook.save("result.xlsx")
        some_workbook.close()

    plyer.notification.notify(
        message=f'Завершено. Обработано {str(len(results))} из {str(len(keywords))}.'
                f'ошибки: {errors}',
        app_name='URL PARSER',
        title='Приложение обработало все записи', )

    print('Please wait for file...')
    atexit.register(close_file, workbook)


if __name__ == "__main__":
    asyncio.run(main())
