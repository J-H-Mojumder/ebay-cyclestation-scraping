from selenium import webdriver
import time
import requests
import xlsxwriter

chrome_path = r"E:\chromedriver.exe"
driver = webdriver.Chrome(chrome_path)
url = "https://www.ebay.com.au/str/cyclestation"
total_data = 0
row_number = 1
page_counter = 0
outputWorkbook = xlsxwriter.Workbook("Ebay.xlsx")
outputsheet = outputWorkbook.add_worksheet()
while 1:

    try:
        driver.get(str(url))
        driver.maximize_window()

        outputsheet.write("A1", "Product Name")
        outputsheet.write("B1", "Price")
        outputsheet.write("C1", "Image Link")
        outputsheet.write("D1", "Image Name")
    except Exception as e:
        print("Problem opening the web link!!")
    else:
        #scrolling page
        start = 0
        end = 200
        while end < 1600:
            driver.execute_script("window.scrollBy(" + str(start) + "," + str(end) + ")", "")
            time.sleep(1.5)
            start = end
            end += 100

        items = driver.find_elements_by_class_name("s-item")
        print(str(len(items)))
        if (str(len(items)) != 0):
            for item in items:
                # title
                title = item.find_elements_by_class_name("s-item__title")
                # print(title[0].text)
                links = item.find_elements_by_tag_name("a")
                link1 = links[0]

                # item price
                price = item.find_elements_by_class_name("s-item__price")
                # print(price[0].text)

                # images
                img = link1.find_elements_by_class_name("s-item__image-img")
                # print(img[0].get_property("src"))

                # getting response
                response = requests.get(img[0].get_property("src"))
                name = title[0].text

                # formatting image name according to title
                name = name.replace("/", "or")
                name = name.replace('"', '')

                # directory
                imageName = "C:/Users/Masum's Computer/Desktop/Ebay" + "/" + name + ".jpg"

                outputsheet.write(row_number, 0, title[0].text)
                outputsheet.write(row_number, 1, price[0].text)
                outputsheet.write(row_number, 2, img[0].get_property("src"))
                name = name+".jpg"
                outputsheet.write(row_number, 3, name)
                row_number += 1
                print(str(row_number) + " Row added!!")
                # print(imageName)
                with open(imageName, "wb") as files:
                    files.write(response.content)
                    print("image downloaded")
                total_data += 1

            # next page
            next_page = driver.find_elements_by_class_name("ebayui-pagination__control")
            print(next_page[-1].text)
            next_page[-1].click()

            url = driver.current_url
            page_counter+=1
            print("Total data " + str(total_data))
        else:
            print("End of page")
            break

outputWorkbook.close()
print("Excel sheet closed!!")