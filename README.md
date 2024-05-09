# Image url parser

simple asynchronous parser for bing image:
### in Intellij
1. install dependencies
2. configurate your dataset and some parameters in config
3. run code under async_parser_v_2.py


### in explorer:
1. add config.yaml and your dataset into /dist
2. run executable

to make your own executable run:
> pyinstaller --onefile async_parser_v_2.py --hidden-import plyer.platforms.win.notification

all results will store at .xlsx file in same directory. check name of your "output_file" at config<br>
after exec wait for console closes, cause you might have little pause before file will create