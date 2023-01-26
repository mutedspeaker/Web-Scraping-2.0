import requests
from bs4 import BeautifulSoup
import openpyxl
import datetime
import pandas as pd
import re as rrr

# an amazon tailored soup extracting function
print("r stands for retrying as the amazon site has a huge traffic so multiple tries are required.")
print("r is being printed for the sole benefit of the user, as it keeps him notified as to what is happening.")


def soupExtract(site, keyword, page, filtertherev, brand=0):
    ftav = filtertherev
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
            print("r", end=" ")
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
    rdt = [["Platform", "Date", "Category", "Brand name", "Product Name", "Rating", "Reviews", "Discount", "Size",
            "Price/g", "Keyword", "Actual price", "Offer price", "Delivery Date"]]

    for _ in range(len(name)):
        if page == 7 & _ == 10:
            break
        try:
            pl = 'Amazon'
            dt = datetime.datetime.now().strftime("%x")
            bn = name[_].text.split()[0]
            n = " ".join(name[_].text.split()[:-1])
            n = name[_].text.split(',')[0]
            # n = name[i].text
            r = float(rating[_].text.split()[0].replace(',', ''))
            re = int(reviews[_].text.replace(',', '')[1:-2])
            # discount
            si = name[_].text.split()[-1]
            try:
                si = rrr.search("([0-9]{1,10})\.{0,1}[0-9]*\s{0,3}(g)", name[_].text).group()
            except:
                si = ""
            # si = rrr.find("[0-9]*/.{0,1}[0-9]*/s{0,3}(g)", name[i].text)
            wpg = weightPerGram[_].text
            k = keyword
            b = float(beforeprice[_].text.split('₹')[1].replace(',', ''))
            p = float(price[_].text.replace(',', ''))
            de = delivery[_].text
            di = b - p
            di = di / b
            di = di * 100
            di = round(di, 2)
            if di > 0:
                di = str(di) + "%"
            else:
                di = ""
                b = ""
            # di = str(di) + "%"
            # keyword
            if ftav:
                # print(k, name[i].text)
                # if k.lower() not in name[i].text.lower():
                # #     continue
                # print(brand.lower(), bn.lower())
                if brand.lower() not in bn.lower():
                    continue
            rdt.append([pl, dt, k, bn, n, r, re, di, si, wpg, b, p, de, page, _ + 1])
            sheet.append([pl, dt, k, bn, n, r, re, di, si, wpg, b, p, de, page, _ + 1])

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

data = pd.read_excel(r'input2.xlsx')

# # if the user wants to give the path of the excel file
# print("Do u want to give a separate path? ")
# ans = input()
# if ans.lower() == 'y' or ans.lower() == 'yes':
#     path = input()
#     path = path[1:-1]
#     data = pd.read_excel(path)

# print("Enter column name: \n")
# keywords = input()
keywords = "Keywords"
brandExist = "Brand"
filterThere = 'Filter'
df = pd.DataFrame(data, columns=[keywords])
be = pd.DataFrame(data, columns=[brandExist])
ft = pd.DataFrame(data, columns=[filterThere])
# print(df)
# print(df.head(2))
# print(df.shape)
# print(df.describe())

# Converting data frame taken from the above keyword to a numpy array
arr = df.to_numpy()
be = be.to_numpy()
ft = ft.to_numpy()
keywords = []
bea = []
fta = []
# Numpy arrays are 2d so making it 1d
for i in arr:
    if type(i[0]) == str:
        keywords.append(i[0])
for i in be:
    print(i)
    if type(i[0]) == str:
        bea.append(i[0])

for i in ft:
    if type(i[0]) == str:
        fta.append(i[0])
if fta[0].lower() == "on":
    filterThere = 1
else:
    filterThere = 0
try:
    brandExist = bea[0]
except:
    brandExist = 0

# A typical amazon website broken in 3 parts:
# https://www.amazon.in/s?k= biscuit &page=2

site = "https://www.amazon.in/s?k="

# Open excel sheet
excel = openpyxl.Workbook()
sheet = excel.active

# Add the first row, with the following requirements
sheet.append(
    ["Platform", "Date", "Category (Keyword)", "Brand name", "Product Name", "Rating", "Reviews", "Discount", "Size",
     "Price/g", "Actual price", "Offer price", "Delivery Date", "Page Number", "Item Number"])

# totalPages = soupExtract(site,'cream biscuit')
# for i in range(2, int(total_pages) + 1):

if not brandExist:
    for keyword in keywords:
        total_pages = soupExtract(site, keyword, 1, filterThere)  # brand
        i = 1
        while i != total_pages:
            i = i + 1
            try:
                soupExtract(site, keyword, i, filterThere)  # brand
            except:
                pass
    # saving the excel file
    excel.save(f"{keyword}.xlsx")
else:
    for keyword in keywords:
        total_pages = soupExtract(site, keyword, 1, filterThere, brandExist)  # brand
        i = 1
        while i != total_pages:
            i = i + 1
            try:
                soupExtract(site, keyword, i, filterThere, brandExist)  # brand
            except:
                pass
    # saving the excel file
    excel.save(f"{keyword}.xlsx")
