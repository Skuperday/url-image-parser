# Image URL Parser: Ваш асинхронный помощник по поиску изображений

**Надоело вручную искать изображения для таблиц?** 
<br>Image URL Parser приходит на помощь! Этот Python-проект позволяет автоматизировать поиск и извлечение URL-адресов изображений из Bing Image, экономя ваше время и силы.

***

### Особенности

*   **Молниеносная скорость:** Асинхронные запросы с помощью `aiohttp` для максимальной эффективности.
*   **Гибкость:** Настройка параметров поиска через YAML-файл:
    *   Количество результатов для поиска
    *   Фильтры по типу изображения (line, photo, clipart, gif, transparent)
    *   Чёрные и белые списки сайтов 
*   **Удобство:** Работа с Excel:
    *   Читает ключевые слова из таблицы
    *   Записывает URL-адреса найденных изображений в новый Excel-файл
*   **Простота:**
    *   **В IDE:** <br>1. Запустите `async_parser_v_2.py`.
    *   **Исполняемый файл:**
1. Установите PyInstaller: `pip install pyinstaller`
2. Соберите исполняемый файл:
    ```bash
    pyinstaller --onefile async_parser_v_2.py --hidden-import plyer.platforms.win.notification
    ``` 
3.  Поместите `config.yaml` и Excel-файл в папку с исполняемым файлом.
4.  Запустите исполняемый файл.


*** 

### Примеры использования

*   **SEO-специалисты и контент-менеджеры:** Быстро находите релевантные изображения для статей и веб-страниц. 
*   **Исследователи и аналитики данных:** Собирайте большие наборы данных изображений для машинного обучения или анализа. 
*   **Разработчики:** Интегрируйте поиск изображений в свои приложения.
*   **И многие другие!**

*** 

### Начало работы 

1.  **Скачайте код проекта:** Клонируйте репозиторий или загрузите архив.
2.  **Настройте config.yaml:** Укажите путь к Excel-файлу, параметры поиска и фильтры. 
3.  **Запустите:** (см. раздел "Простота" выше)

***

### Внесите свой вклад!

Image URL Parser - проект с открытым исходным кодом. Мы приветствуем участие сообщества! 

*   **Сообщайте об ошибках и предлагайте идеи:**  [Создайте Issue](https://github.com/Skuperday/urlParser/issues)
*   **Улучшайте код:** [Отправьте Pull Request](https://github.com/Skuperday/urlParser/pulls)
*   **Расскажите о проекте:** Поделитесь ссылкой на GitHub  и расскажите о своём опыте использования Image URL Parser.

***
***

## Image URL Parser: Your Asynchronous Image Search Companion

**Tired of manually searching for images to tables?** 
<br>Image URL Parser is here to help! This Python project automates the process of searching and extracting image URLs from Bing Image, saving you time and effort. 

*** 

### Features

*   **Blazing Fast:** Utilizes asynchronous requests with `aiohttp` for maximum efficiency. 
*   **Flexible:** Customize search parameters through a YAML file: 
    *   Number of looking images
    *   Filters by image type (line, photo, clipart, gif, transparent)
    *   Blacklists and whitelists
    *   And much more
*   **Convenient:** Works with Excel:
    *   Reads keywords from a spreadsheet
    *   Writes URLs of found images into a new Excel file 
*   **Easy to Use:**
    *   **In your IDE:**<br>1. Run `async_parser_v_2.py`. 
    *   **Executable file:** 
    <br> 1.  Install PyInstaller: `pip install pyinstaller`
    <br> 2.  Build the executable: ``` pyinstaller --onefile async_parser_v_2.py --hidden-import plyer.platforms.win.notification ``` 
    <br>3.  Place `config.yaml` and your Excel file in the folder with the executable.
    <br>4.  Run the executable. 

***

### Use Cases 

*   **SEO specialists and content managers:** Quickly find relevant images for articles and web pages.
*   **Data researchers and analysts:** Gather large datasets of images for machine learning or analysis.
*   **Developers:** Integrate image search functionality into your applications.
*   **And many more!**

***

### Getting Started

1.  **Download the project code:** Clone the repository or download the archive. 
2.  **Configure config.yaml:** Specify the path to your Excel file, search parameters, and filters.
3.  **Run:** (see the "Easy to Use" section above)

*** 

### Contribute!

Image URL Parser is an open-source project. We welcome contributions from the community!

*   **Report bugs and suggest ideas:** [Open an Issue](https://github.com/Skuperday/urlParser/issues)
*   **Improve the code:** [Submit a Pull Request](https://github.com/Skuperday/urlParser/pulls)
*   **Spread the word:** Share the GitHub repository link and talk about your experience using Image URL Parser.

***