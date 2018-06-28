
from bs4 import BeautifulSoup
import requests
import xlwt



#Method that finds and stores all data into the excel document
def get_info(input_url):


    #Initiating the excel file
    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet("Sheet 1")


    #Adding the correct columns
    sheet1.write(0, 0, "Comapany Name")
    sheet1.write(0, 1, "Comapany Address")
    sheet1.write(0, 2, "Comapany Phone")
    sheet1.write(0, 3, "Comapany Website")
    sheet1.write(0, 4, "CEO")
    sheet1.write(0, 5, "Employees")
    sheet1.write(0, 6, "State of inc")
    sheet1.write(0, 7, "Fiscal Year End")
    sheet1.write(0, 8, "Date Priced")
    sheet1.write(0, 9, "Symbol")
    sheet1.write(0, 10, "Exchange")
    sheet1.write(0, 11, "Share Price")
    sheet1.write(0, 12, "Shares Offered")
    sheet1.write(0, 13, "Offer Amount")
    sheet1.write(0, 14, "Total Expenses")
    sheet1.write(0, 15, "Shares Over Alloted")
    sheet1.write(0, 16, "Shareholder Shares Offered")
    sheet1.write(0, 17, "Shares Outstanding")
    sheet1.write(0, 18, "Lockup Period (days)")
    sheet1.write(0, 19, "Lockup Expiration")
    sheet1.write(0, 20, "Quiet Period Expiration")
    sheet1.write(0, 21, "CIK")


    #Using Requests and BeautifulSoup
    r = requests.get(input_url)
    br = BeautifulSoup(r.content,"html.parser")

    #List to keep the links
    link_list = []

    #Find the tables
    for table in br.find_all("div", {"class": "genTable"}):
        for tr in table.find_all("tr"):
            for td in tr.find_all("td"):
                for a in td.find_all("a"):

                    #Find all the links for the detailed page
                    tempLink = (a.get("href"))

                    #The actual links to the detailed page are all at least 60 characters in length.
                    #Thus, we can use this very simple and effective way to differentiate between good links and bad links.
                    if len(tempLink) > 60:
                        link_list.append(tempLink)

        #Since we only want the first table, we can break out of this loop
        break

    #Now, we iterate through all the links.
    i = 1
    for link in link_list:

        #Do the same thing to initialize beautifulsoup
        r = requests.get(link)
        br = BeautifulSoup(r.content, "html.parser")

        #Counter to make sure we get the info we need
        counter = 0

        #We find the table
        for table in br.find_all("div", {"class": "genTable"}):

            #Here, the tr tags are the columns
            cols = table.find_all("tr")

            #for every element in the column
            for j in range(0, len(cols)):
                for td in cols[j].find_all("td"):
                    counter += 1

                    #Now, in this detailed page, everything is either a legend or the detail.
                    #We already have the legend, so we just want the detail.
                    #Conveniently, the details are on every other line.
                    if(counter %2 == 0):

                        #Write it in the respective row/column in the excel file.
                        sheet1.write(i, j, td.get_text().rstrip())
                        print(td.get_text().rstrip())

            #Break out of the loop since we only want the first table
            break

        #Increment i by 1 to go to the next row
        i+=1


    #save it to an excel file
    book.save("Result.xls")


get_info("https://www.nasdaq.com/markets/ipos/activity.aspx?tab=pricings&month=2018-02")