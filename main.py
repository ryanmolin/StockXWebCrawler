from bs4 import BeautifulSoup
import urllib.request
import xlsxwriter
import time
import re


# List dictionary
brand = 'Yeezy', 'Ultra Boost', 'NMD', 'Iniki', 'Air Jordan', 'Jordan',
edition = '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', '23' 'XXX', 'XXX1',
origin = 'original', 'Original', 'retro', 'Retro', 'og', 'OG', 'legacy', 'Legacy', 'sb', 'SB'
weight = 'low', 'Low', 'high', 'High', 'mid', 'Mid'
gs = '(gs)', '(GS)', '(ps)', '(PS)', '(td)', '(TD)', '(w)', '(W)'

# Create an excel workbook
outWorkbook = xlsxwriter.Workbook("StockX/data.xlsx") # ex) StockX\data.xlsx
# Create a sheet within the workbook
outSheet = outWorkbook.add_worksheet()

# Make the titles
outSheet.write('A1', 'Brand')
outSheet.write('A2', 'Edition')
outSheet.write('A3', 'Origin')
outSheet.write('A4', 'Weight')
outSheet.write('A5', 'Color 1')
outSheet.write('A6', 'Color 2')
outSheet.write('A7', 'Color 3')
outSheet.write('A8', 'Color 4')
outSheet.write('A9', 'Color 5')
outSheet.write('A10', 'String')
outSheet.write('A11', 'Style')
outSheet.write('A12', 'Retail Price')
outSheet.write('A13', 'Release Date')
outSheet.write('A14', 'N of Sales')
outSheet.write('A15', 'Price Premium')
outSheet.write('A16', 'Average Sale Price')
outSheet.write('A17', 'Size')

# Start from the SECOND cell. Rows and columns are zero indexed.
col = 1

with open("StockX/filtered.txt") as file: # ex) StockX\crawled.txt
    try:
        for url in file:
            # Reset variables
            row = 0 # Starts writing the data from the top
            # Establish a connection
            req = urllib.request.Request(
                url,
                data=None,
                headers={
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.108 Safari/537.36'
                }
            )
            content = urllib.request.urlopen(req)
            content = content.read().decode('utf-8').strip()
            # Get the soup html
            soup = BeautifulSoup(content,"html.parser")
            # Get the name string from webpage
            shoename = soup.find('h1', {"class": "name"})

            # Continue if the string is full
            if shoename is not None:
                string = shoename.text
                # print('string = shoename.text:')
                # print(string)
                # Split the string into an array to work with the different parts
                string = string.split()
                print('string.split():')
                print(string)

                # Set variables that need to be reset (do not get declared somewhere else)
                dataBrand = "None"
                dataEdition = "None"
                dataOrigin = "None"
                dataWeight = "None"
                dataColor1 = "None"
                dataColor2 = "None"
                dataColor3 = "None"
                dataColor4 = "None"
                dataColor5 = "None"
                dataDetails1 = "None"
                dataDetails3 = "None"
                dataDetails4 = "None"
                dataDetails5 = "None"
                dataDetails6 = "None"
                dataDetails7 = "None"
                dataDetails8 = "None"

                # Get the brand from name string
                for element in string:
                    if "air" and "jordan" in brand:
                        dataBrand = 'Air Jordan'
                        x = string.index(element)
                        string.pop(x)
                    elif element in brand:
                        dataBrand = element
                        x = string.index(element)
                        string.pop(x)
                        # print(dataBrand)

                # Get edition from name string
                for element in string:
                    if element in edition:
                        dataEdition = element
                        x = string.index(element)
                        string.pop(x)
                        # print(dataEdition)

                # Get origin from name string
                for element in string:
                    if element in origin:
                        dataOrigin = element
                        x = string.index(element)
                        string.pop(x)
                        # print(dataOrigin)

                # Get low/high from name string
                for element in string:
                    if element in weight:
                        dataWeight = element
                        x = string.index(element)
                        string.pop(x)
                        # print(dataWeight)

                # Gradeschool (gs), preschool (ps), toddler (td), or not
                for element in string:
                    if element in gs:
                        dataDetails8 = element
                        x = string.index(element)
                        string.pop(x)
                    else:
                        dataDetails8 = "Regular"
                    # print(dataDetails8)

                # Put together extra string elements
                z = ' '
                string = z.join(string)
                print(string)

                # Error handling for data: if element is nonexistant
                try:
                    dataBrand
                except NameError:
                    dataBrand = "None"
                try:
                    dataEdition
                except NameError:
                    dataEdition = "None"
                try:
                    dataOrigin
                except NameError:
                    dataOrigin = "None"
                try:
                    dataWeight
                except NameError:
                    dataWeight = "None"
                try:
                    dataColor1
                except NameError:
                    dataColor1 = "None"
                try:
                    dataColor2
                except NameError:
                    dataColor2 = "None"
                try:
                    dataColor3
                except NameError:
                    dataColor3 = "None"
                try:
                    dataColor4
                except NameError:
                    dataColor4 = "None"
                try:
                    dataColor5
                except NameError:
                    dataColor5 = "None"
                try:
                    string
                except NameError:
                    string = "None"



                # Get the Style, Colorway, Retail Price, & Release Date
                childArray = []
                for test in soup.find_all('div', {'class': 'detail'}):
                    children = test.findChildren("span", recursive=False)
                    for child in children:
                        child = child.text
                        child = child[1:]
                        childArray.append(child)
                print(childArray)
                try:
                    if childArray[0] is not None:
                        dataDetails1 = childArray[0] # style
                    if childArray[1] is not None:
                        dataDetails2 = childArray[1] # colorway
                    if childArray[2] is not None:
                        dataDetails3 = childArray[2] # retail price
                    if childArray[3] is not None:
                        dataDetails4 = childArray[3] # release date
                except IndexError:
                    print('Information not found from:')
                    print(url)
                    pass

                # Put colorway into color list
                dataDetails2 = re.split(r'[`\-=~!@#$%^&*()_+\[\]{};\'\\:"|<,./<>?]', dataDetails2)
                print('Colorway.split:')
                print(dataDetails2)
                # Put colorway colors into variables
                try:
                    for colors in dataDetails2:
                        if dataColor1 is "None":
                            dataColor1 = colors.strip()
                        elif dataColor2 is "None":
                            dataColor2 = colors.strip()
                        elif dataColor3 is "None":
                            dataColor3 = colors.strip()
                        elif dataColor4 is "None":
                            dataColor4 = colors.strip()
                        elif dataColor5 is "None":
                            dataColor5 = colors.strip()
                        else:
                            print('Missing colors!!')
                except IndexError:
                    print('Colors not found from:')
                    print(url)
                    pass
                # Get the N of Sale, Price Premium, & Average Sale Price
                childArray2 = []
                try:
                    for divTag in soup.find_all("div", {"class": "gauge-value"}):
                        divTag = divTag.text
                        childArray2.append(divTag)
                    print(childArray2)
                    dataDetails5 = childArray2[0].strip() # N of Sales
                    dataDetails6 = childArray2[1].strip() # Price Premium
                    dataDetails7 = childArray2[2].strip() # Average Sale Price
                except:
                    print('Information not found from:')
                    print(url)
                    pass
                # Set an array data to easily write to the xlsx worksheet
                data = (
                     [dataBrand],
                     [dataEdition],
                     [dataOrigin],
                     [dataWeight],
                     [dataColor1],
                     [dataColor2],
                     [dataColor3],
                     [dataColor4],
                     [dataColor5],
                     [string],
                     [dataDetails1],
                     [dataDetails3],
                     [dataDetails4],
                     [dataDetails5],
                     [dataDetails6],
                     [dataDetails7],
                     [dataDetails8],
                )

                print('data array:')
                print(data)

                # Iterate over the data and write it out row by row.
                try:
                    for theData in (data):
                        outSheet.write_row(row, col, theData)
                        row += 1
                except Exception as ex:
                    print('Error: failed writing data to xlsx')
                    print(url)
                    print(ex)
                    pass
            else:
                print('Error: string is nonexistant:')
                print(url)
                pass

            # Each new url writes to a new column in excel
            col = col + 1
    # Catch all other errors that may pop up
    except Exception as ex2:
        print(ex2)
        print(url)
        pass

outWorkbook.close()
