from bs4 import BeautifulSoup
import requests, openpyxl

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = 'Top Rated Movies'
sheet.append(['Rank', 'Name', 'Year', 'Rating'])

try:
    headers = {
    'User-Agent': 'Chrome/58.0.3029.110'}
    source = requests.get('https://www.imdb.com/chart/top/', headers=headers)
    source.raise_for_status()

    soup = BeautifulSoup(source.text, 'html.parser')
    movies = soup.find("ul", class_="ipc-metadata-list ipc-metadata-list--dividers-between sc-71ed9118-0 kxsUNk compact-list-view ipc-metadata-list--base").find_all('li')
    
    for movie in movies:
        name=movie.find("a", class_="ipc-title-link-wrapper").h3.text
        ranking = name.split('. ')[0]
        # Remove the serial number from the movie title
        name = name.split('. ')[1]
        year=movie.find("div", class_="sc-c7e5f54-7 brlapf cli-title-metadata").span.text
        rating=movie.find('span', class_="ipc-rating-star ipc-rating-star--base ipc-rating-star--imdb ratingGroup--imdb-rating").text.strip('')
        rating = movie.find('span', class_="ipc-rating-star ipc-rating-star--base ipc-rating-star--imdb ratingGroup--imdb-rating").text.strip()
        # Extract the text within parentheses (e.g., (2.8M)) and strip it
        rating = rating.split('(')[0].strip()

        # print(f"{ranking} : {name} : {year} : {rating}")
        sheet.append([ranking, name, year, rating])
except Exception as e:
    print(e)

excel.save('IMDB Movie Ratings.xlsx')
    