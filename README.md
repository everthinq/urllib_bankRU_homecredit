### This project parses branch data from [Home Credit and Finance Bank Russia map](https://www.homecredit.ru/contacts/otdeleniya/offices/)

### Instructions:

0) Install Python3 -- https://www.python.org/
1) Install requirements.txt -- *pip install -r requirements.txt*
2) Run *main.py*
3) To get branches data you need to do API requests
4) *get_cities_id* function is using russian capital letters [А-Я] to get cities id and returns a JSON with city_id and city_name 
5) *get_branch_data* function is using the JSON from the step 4 and sends GET request with city_id to get branches data. Returns a JSON with all the branches data
6) *write_xlsx* function is using the JSON from the step 5 and writes the data in excel
