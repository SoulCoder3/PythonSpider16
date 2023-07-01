import multiprocessing
import requests
from bs4 import BeautifulSoup
import xlwt
import threading

book = xlwt.Workbook(encoding='utf-8', style_compression=0)

sheet = book.add_sheet('The top movies in douban', cell_overwrite_ok=True)
sheet.write(0, 0, 'Name')
sheet.write(0, 1, 'Image')
sheet.write(0, 2, 'Rank')
sheet.write(0, 3, 'Score')
sheet.write(0, 4, 'Author')
sheet.write(0, 5, 'Introduction')

global_lock = multiprocessing.Lock()
def save_to_excel(items, page):
    global_lock.acquire()
    row = page
    for item in items:
        row += 1
        item_name = item.find(class_='title').string
        item_img = item.find('a').find('img').get('src')
        item_index = item.find(class_="").string
        item_score = item.find(class_='rating_num').string
        item_author = item.find('p').text
        item_intr = item.find(class_='inq')
        if item_intr == None:
            item_intr = ''
        else:
            item_intr = item_intr.string
        sheet.write(row, 0, item_name)
        sheet.write(row, 1, item_img)
        sheet.write(row, 2, item_index)
        sheet.write(row, 3, item_score)
        sheet.write(row, 4, item_author)
        sheet.write(row, 5, item_intr)
    book.save('Top 250 Movies in Douban.xls')
    global_lock.release()
    return None


def main(url, page):
    html = request_douban(url)
    soup = BeautifulSoup(html, "lxml")

    items = soup.find(class_='grid_view').find_all('li')
    save_to_excel(items, page)
    return None


def request_douban(url):
    headers = {
        "User-Agent": 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36'
    }
    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            return response.text
    except requests.RequestException:
        return None


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    urls = []
    pages = []
    pool = multiprocessing.Pool(4)

    for i in range(0, 10):
        page = i * 25
        url = 'https://movie.douban.com/top250?start=' + str(page) + '&filter='
        pages.append(page)
        urls.append(url)
    zip_argus = list(zip(urls, pages))
    pool.starmap(main, zip_argus)
    pool.close()
    pool.join()
