import xlsxwriter
import urllib.request
import requests
from requests import get
from bs4 import BeautifulSoup
url = "https://in.bookmyshow.com/bengaluru/movies/nowshowing"
response= get(url) 

html_soup=BeautifulSoup(response.text,"lxml")
movies=html_soup.find_all(class_='card-container wow fadeIn movie-card-container')

links=[]
name=[]
censor=[]
lang=[]

#xlsxwriter operation

workbook=xlsxwriter.Workbook('trialp.xlsx')
worksheet= workbook.add_worksheet('SEO')

row,col=0,0
display=workbook.add_format({'bold':True, 'font_color':'black'})

worksheet.write(row, 0,"name",display)
worksheet.write(row, 1,"links",display)


worksheet.write(0, 2,"Keyword1(movie)",display)
worksheet.write(0, 3,"Keyword2(review)",display)
worksheet.write(0, 4,"Density(keyword1)",display)
worksheet.write(0, 5,"Density(keyword2)",display)

row=1

print(len(movies))

for data in movies:
     name=data.h4.text
     
     if data.find('li',class_="__language") is not None:
          lang=data.li.text
     
     rate=data.find('div', class_="card-tag")
     if rate.span.span is not None:
          censor= rate.span.span.text
          
     link1= data.find('a')
     links="https://in.bookmyshow.com"+link1['href']
     
     
     worksheet.write(row, 0,name)
     worksheet.write(row, 1,links)

       
     req1=requests.get(links)
     data1=req1.text
     s1=BeautifulSoup(data1, "lxml")
     text=s1.get_text()
     t1=text.split()
        #print(t1)
     a=len(text)
     print(a)

     k1="movie"
     k2="review"
        

     calkeyword1=text.count(k1)
     calkeyword2=text.count(k2)
     worksheet.write(row, 2, calkeyword1)
     worksheet.write(row, 3, calkeyword2)
     

     print(calkeyword1)
     print(calkeyword2)

     density1= calkeyword1/a
     worksheet.write(row, col+4, density1)
     print(density1)
     density2= calkeyword2/a
     worksheet.write(row, col+5, density2)
     print(density2)

     row+=1

       
chart1=workbook.add_chart({'type':'column'})
chart1.set_x_axis({'name': 'Density of SEO Keywords'})
chart1.set_title({'name': 'SEARCH ENGINE OPTIMIZATION'})
heading=workbook.add_format({'bold':True,'font_color':'red'})
chart1.add_series({'values':'=SEO!$C2:$C20'})
chart1.add_series({'values':'=SEO!$D2:$D20'})

worksheet.insert_chart("L20",chart1)


workbook.close()

























