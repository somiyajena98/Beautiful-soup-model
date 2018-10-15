import urllib.request 
from bs4 import BeautifulSoup
import requests
import xlsxwriter  

req = urllib.request.Request( "http://www.team4adventure.com/tours", data=None, 
             headers={ 'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_9_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/35.0.1916.47 Safari/537.36' } 
                                      ) 
f = urllib.request.urlopen(req) 
s=f.read()


# Create a workbook and add a worksheet.

workbook = xlsxwriter.Workbook('hello.xlsx') #create an excel file of name hello
worksheet = workbook.add_worksheet('SEO') #create an excel sheet of name SEO
row = 0
col = 0
soup=BeautifulSoup(s)
a=soup.find(class_='main-navigation')
link=a.find_all('a')
for l in link:
        names = l.contents[0]  #names present in the website 
        links1 = l.get('href') #links of the categories in the website
        print(names)
        print(links1)
        #place the links and names in the columns in excel sheet.
        worksheet.write(row, col, links1) 
        worksheet.write(row, col+1, names)
        x='Home'
        y='Tours'
        z='Departures'
        reqst=requests.get(links1)
        data=reqst.text
        b=BeautifulSoup(data)
        text=b.get_text()
        t=text.split()
        a=len(text)
        
    #Calculate the keywords in the text
        
        calkey1=text.count(x)
        calkey2=text.count(y)
        calkey3=text.count(z)

        worksheet.write(row, col+2, calkey1)
        worksheet.write(row, col+3, calkey2)
        worksheet.write(row, col+4, calkey3)

     #Calculate the density of the words

        density1=calkey1/a
        density2=calkey2/a
        density3=calkey3/a

        worksheet.write(row, col+5, density1)
        worksheet.write(row, col+6, density2)
        worksheet.write(row, col+7, density3)

    #create a chart to plot the calculated keywords

        chart = workbook.add_chart({'type': 'column'})
        chart.add_series({'values': '=SEO!$C$1:$C$19'}) #$ represents the cell
        chart.add_series({'values': '=SEO!$D$1:$D$19'})
        chart.add_series({'values': '=SEO!$E$1:$E$19'})

    #Place the chart in the cells speified 

        worksheet.insert_chart('I22', chart)
        
                           
    #worksheet.write(row, col + 1,links)
        row += 1
    

workbook.close()
