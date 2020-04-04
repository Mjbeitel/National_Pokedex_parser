
import requests
import xlsxwriter 
import pandas as pd
import os

from bs4 import BeautifulSoup

page = requests.get("https://pokemondb.net/pokedex/national")


################################################## INITALIZATIONS ######################################################
datas = []
numbers=[]
names = []
types = []
i = 0
j = 0
gens = []

error = 'Page is unable to be reached\n'
success = 'Page successfully reached\n'

if page.status_code != 200:
    print(error)

else:

    print(success)


soup = BeautifulSoup(page.content, 'html.parser') # parse url


poke_dex = soup.find_all(class_ = 'infocard') # find all including image SRC's 


nat_dex = soup.find_all(class_ = 'infocard-lg-data text-muted') # get specific text data


#print(nat_dex) # verification that works. 

x = range(len(poke_dex)) # capture length


#print(x) # works


print("\n Number:     Name:      Type:     ")

print(" ---------------------------------------------------------")


workbook = xlsxwriter.Workbook('National-Pokedex.xlsx')
workbook.close() # close excel if left open

def national_dex(i,j):
    for link in nat_dex: # get all data for later use
        i = i + 1

        data = link.get_text() # collect all data this is a list that contains everything

        print(data)

        datas.append(data)




    for link in soup.find_all('small'): # get number and type both of type small 

        j = j + 1 # incriment specific counter

        number = link.get_text()

        if j % 2 == 1: # odd number works

            numbers.append(number) # append all number in array
           

        if j % 2 == 0: # even number works
            typez = number
            types.append(typez) # append all type in array
            



    for link in soup.find_all(class_ = 'ent-name'): # get names

        name = link.get_text()
        names.append(name) # append all name in array




def print_excel(numbers,names,types,datas,i):
    workbook = xlsxwriter.Workbook('National-Pokedex.xlsx') 
  

    worksheet = workbook.add_worksheet('National Pokedex') # adds new worksheet
  

   

    merge_format = workbook.add_format({ # format of merge
        'bold': 1,
        'font_size': 18,
        'align': 'left',
        'valign': 'vcenter'})

    worksheet.merge_range('A1:C1', 'National Pokedex', merge_format)

    worksheet.set_column('A:A', 9.51) # set column A width
    worksheet.set_column('B:B', 12) # set column B-D width
    worksheet.set_column('C:C', 14) # set column B-D width

    

    row = 2
    col = 0

    # Iterate over the data and write it out row by row. 
    for item in numbers: 
   
        # write operation perform 
        worksheet.write(row, col, item) 
  
        # incrementing the value of row by one 
        # with each iteratons. 
        row += 1


    row = 2
    col = col + 1 # incriment column

    for item in names: #prices array
   
        # write operation perform 
        worksheet.write(row, col, item) 
  
        # incrementing the value of row by one 
        # with each iteratons. 
        row += 1


    row = 2
    col = col + 1 # incriment column

    for item in types: #prices array
   
        # write operation perform 
        worksheet.write(row, col, item) 
  
        # incrementing the value of row by one 
        # with each iteratons. 
        row += 1


    worksheet.add_table('A2:C892', {'style':'Table Style Dark 2','columns': [{'header': 'Number'},
                                           {'header': 'Name'},
                                           {'header': 'Type'},
                                           ]}) # add data to table or else everything breaks, I am an idiot
   
    workbook.close() # close excel file with all new information
    



def excel_opener():

     os.system('National-Pokedex.xlsx') # open excel file 





national_dex(i,j) # call function

print_excel(numbers,names,types,datas,i) # call function

excel_opener() # open excel file