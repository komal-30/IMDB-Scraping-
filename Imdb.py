#import modules
from bs4 import BeautifulSoup
from matplotlib.pyplot import text
import requests,openpyxl
from sqlalchemy import true

excel = openpyxl.Workbook()
print(excel.sheetnames)
sheet = excel.active
sheet.title = "Top Rated Movies"
print(excel.sheetnames)
sheet.append(['Rank','Movie Name','Year Of Release','IMDB Rating'])

try:
    data = requests.get('https://www.imdb.com/chart/top/')
    data.raise_for_status()

    soup = BeautifulSoup(data.text,'html.parser')
    print(soup.prettify)

    movies = soup.find('tbody',class_="lister-list").find_all('tr')
    for movie in movies:

        name=movie.find('td',class_="titleColumn").a.text
        #print(name)
    
        #rank = movie.find('td',class_='titleColumn').text
        rank = movie.find('td',class_='titleColumn').get_text(strip=True).split(".")[0]
        #print(rank)

        year = movie.find('td',class_='titleColumn').span.text.strip("()")
        #print(year)

        #rating = movie.find('td',class_='ratingColumn imdbRating').get_text()
        rating = movie.find('td',class_='ratingColumn imdbRating').strong.text
        print(rank,name,year,rating)
        sheet.append([rank,name,year,rating])
      

except Exception as e:
    print(e)

excel.save("IMDB Data.xlsx")