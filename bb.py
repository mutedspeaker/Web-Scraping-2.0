import requests
from bs4 import BeautifulSoup
import openpyxl
import datetime
import pandas as pd
import re as rrr


# an amazon tailored soup extracting function
def soupExtract(site, keyword, page, brand=0):
    if brand != 0:
        sheet.title = brand
    url = site + keyword + "&page=" + str(page)
    while True:
        try:
            req = requests.get(url)
            req.raise_for_status()
            print("\nSuccess\n")
            soup = BeautifulSoup(req.text, "html.parser")
            break

        # the exception is due, to high traffic in Amazon Website, it is quite possible to not get our html request in the first time
        except Exception:
            print("r...", end=" ")
    a = soup.find('span', class_='rush-component s-latency-cf-section').find_all('span')
    name = soup.find('span', class_='rush-component s-latency-cf-section').find_all('span',
                                                                                    class_='a-size-base-plus a-color-base a-text-normal')
    price = soup.find('span', class_='rush-component s-latency-cf-section').find_all('span', class_='a-price-whole')
    rating = soup.find('span', class_='rush-component s-latency-cf-section').find_all('span', class_='a-icon-alt')
    reviews = soup.find('span', class_='rush-component s-latency-cf-section').find_all('span',
                                                                                       class_='a-size-base s-underline-text')
    beforeprice = soup.find('span', class_='rush-component s-latency-cf-section').find_all('span',
                                                                                           class_='a-price a-text-price')
    weightPerGram = soup.find('span', class_='rush-component s-latency-cf-section').find_all('span',
                                                                                             class_='a-size-base a-color-secondary')
    delivery = soup.find('span', class_='rush-component s-latency-cf-section').find_all('span',
                                                                                        class_='a-color-base a-text-bold')
    print(url)
    rdt = []
    rdt.append(["Platform", "Date", "Category", "Brand name", "Product Name", "Rating", "Reviews", "Discount", "Size",
                "Price/g", "Keyword", "Actual price", "Offer price", "Delivery Date"])

    for i in range(len(name)):
        if page == 7 & i == 10:
            break
        try:
            pl = 'Amazon'
            dt = datetime.datetime.now().strftime("%x")
            ct = 'Biscuits'
            bn = name[i].text.split()[0]
            n = " ".join(name[i].text.split()[:-1])
            n = name[i].text.split(',')[0]
            # n = name[i].text
            r = float(rating[i].text.split()[0].replace(',', ''))
            re = int(reviews[i].text.replace(',', ''))
            # discount
            si = name[i].text.split()[-1]
            try:
                si = rrr.search("([0-9]{1,10})\.{0,1}[0-9]*\s{0,3}(g)", name[i].text).group()
            except:
                si = ""
            # si = rrr.find("[0-9]*/.{0,1}[0-9]*/s{0,3}(g)", name[i].text)
            wpg = weightPerGram[i].text
            k = keyword
            b = float(beforeprice[i].text.split('₹')[1].replace(',', ''))
            p = float(price[i].text.replace(',', ''))
            de = delivery[i].text
            di = b - p
            di = di / b
            di = di * 100
            di = round(di, 2)
            if di > 0 :
                di = str(di) + "%"
            else:
                di = 0
                b = 0
            # di = str(di) + "%"
            wbst = "aafaf"
            # keyword
            if k.lower() not in name[i].text.lower():
                continue
            if (brand != 0) & (bn != brand):
                continue
            rdt.append([pl, dt, ct, bn, n, r, re, di, si, wpg, k, b, p, de, page, i + 1, wbst])
            sheet.append([pl, dt, ct, bn, n, r, re, di, si, wpg, k, b, p, de, page, i + 1, wbst])

        except:
            pass


    pager = soup.find('span', {'class': 's-pagination-strip'})

    # store the total number of pages in this variable and return it at the end
    totalPages = int(pager.text[-5])

    #     print(pager.find('span', class_ = 's-pagination-item s-pagination-disabled'))
    #     if pager.find('span', class_ = 's-pagination-item s-pagination-disabled'):
    #         return 1
    #     else:
    #         return 0

    #     totalPages = soup.find('span', 'aria-disabled'=='true').text
    #     print(totalPages)
    #     print(rdt)

    return totalPages


# ---------------------------------------------------------------------------------------------------------------------

# reading excel file
# hardcoded path
# data = pd.read_excel(r'C:\Users\Ojas Tewari\Desktop\input.xlsx')

data = pd.read_excel(r'input.xlsx')

# # if the user wants to give the path of the excel file
# print("Do u want to give a separate path? ")
# ans = input()
# if ans.lower() == 'y' or ans.lower() == 'yes':
#     path = input()
#     path = path[1:-1]
#     data = pd.read_excel(path)

# print("Enter column name: \n")
# keywords = input()
keywords = "Keys"
df = pd.DataFrame(data, columns=[keywords])
# print(df)
# print(df.head(2))
# print(df.shape)
# print(df.describe())

# Converting data frame taken from the above keyword to a numpy array
arr = df.to_numpy()
keywords = []

# Numpy arrays are 2d so making it 1d
for i in arr:
    if type(i[0]) == str:
        keywords.append(i[0])

# A typical amazon website broken in 3 parts:
# https://www.amazon.in/s?k= biscuit &page=2

site = "https://www.amazon.in/s?k="

# Open excel sheet
excel = openpyxl.Workbook()
sheet = excel.active

# Add the first row, with the following requirements
sheet.append(
    ["Platform", "Date", "Category", "Brand name", "Product Name", "Rating", "Reviews", "Discount", "Size", "Price/g",
     "Keyword", "Actual price", "Offer price", "Delivery Date", "Page Number", "Item Number", "Website"])
# ---------------------------------------------------------------------------------------------------------------------

# totalPages = soupExtract(site,'cream biscuit')
# for i in range(2, int(total_pages) + 1):

for keyword in keywords:

    total_pages = soupExtract(site, keyword, 1)         # brand
    i = 1
    while i != total_pages:
        i = i+1
        try:
            soupExtract(site, keyword, i)               # brand
        except:
            pass

# saving the excel file
excel.save("Data.xlsx")
