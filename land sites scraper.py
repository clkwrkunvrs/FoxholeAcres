#Progess:
'''Trying to get APN, but getting message 'loading tax data' when i pull the tax data from landsofamerica. Need to see if I can introduce a delay
or something to allow the http reqeust time to fetch that data'''


#Land and Farm and Lands of America Scraper
import sys
import requests
import csv
from bs4 import BeautifulSoup
def main() :
    counter = 0
    #landandfarm scrape
    #***********url builder***************
    #get the state
    state = raw_input("What is the 2-letter abbreviation of the state you want to search? \n\r Note:If searching by entire state, type the whole state name: \n")
    type(state)
    #get the location
    location = raw_input("What is the name of the location you want to search?: ")
    type(location)
    #format the location for the url
    location = location.replace(" ",'-')
    #show what the user typed
    print("location is " + str(location) + " " + str(state))
    #if searching only by state, modify the url.
    if(not(location)):
    	url = "https://www.landandfarm.com/search/" + state + "-land-for-sale/?CurrentPage=1&IsResidence=False&SortBy=Size&SortOrder=Desc"
    else:
    	url ="https://www.landandfarm.com/search/" + state + "/" + location + "-land-for-sale/?CurrentPage=1&IsResidence=False&SortBy=Size&SortOrder=Desc"
    '''
    print("****************LAND AND FARM DATA********************\n")
    r = requests.get(url)
    soup = BeautifulSoup(r.text,'html.parser')
    # get number of pages
    numPages = soup.find('span',attrs={'class':'results-heading'})
    index = numPages.text.find('of')
    if(index == -1):
        numPages = 1
    else:
        numPages = int(numPages.text[index+3:len(numPages.text)])
    print("number of pages is: " +str(numPages))
    currentPage = 1
    while currentPage <= numPages:
        print("**************************************************************************")
        print("\t\t\tLand And Farm Page #: " + str(currentPage) + " of " + str(numPages))
        print("**************************************************************************")
        if(location):
            url ="https://www.landandfarm.com/search/" + state + "/" + location + "-land-for-sale/?CurrentPage=" + str(currentPage) + "&IsResidence=False&SortBy=Size&SortOrder=Desc"
            #print("url is: " + url)
        else:
        	url = "https://www.landandfarm.com/search/" + state + "-land-for-sale/?CurrentPage=" + str(currentPage) + "&IsResidence=False&SortBy=Size&SortOrder=Desc"

        r = requests.get(url)
        soup = BeautifulSoup(r.text,'html.parser')

        #get price
        results = soup.find_all('div',attrs={'class':'property-card--price'})
        #get size
        results2 = soup.find_all('div',attrs={'class':'property-card--quick-stats'})


        #results3 = soup.find_all('div',attrs={'class':'listingPrice-container'})
        #get location
        results4 = soup.find_all('div',attrs={'class':'property-card--street-address'})

        while counter < len(results4):
            print(str(results2[counter].text.strip() + ".........." + results[counter].text.strip() + ".........." + str(results4[counter].text.strip()).replace('\n',' | ')))
            counter = counter + 1
        counter = 0
        currentPage = currentPage + 1
        print("current page is " + str(currentPage) + " of " + str(numPages))

        '''
    #************lands of america scrape*******************
    print("\n****************LANDS OF AMERICA DATA*******************\n")
    url = "https://www.landsofamerica.com/" + location;
    if(location): url = url + "-" + state + "/all-land/no-house/sort-acre-high/"
    else: url = url + state + "/all-land/no-house/sort-acre-high/"
    #get the html
    r = requests.get(url)
    #prep the html for parsing
    soup = BeautifulSoup(r.text,'html.parser')
    #get number of pages
    numPages = soup.find('h1',attrs={'class':'listResultsLabel'})
    #find location of number of pages on 1st page
    index = int(numPages.text.find('of'));
    #extract that string and cast as int to get #pages
    if(index == -1):
        numPages = 1
    else:
        numPages = int(numPages.text[index+3:len(numPages.text)])

    #start loop to get all pages here
    currentPage = 1;
    while currentPage <= numPages:
        print("**************************************************************************")
        print("\t\t\tLands of America Page #: " + str(currentPage) + " of " + str(numPages))
        print("**************************************************************************")
        #dynamic url. page # changes with each loop iteration
        url = "https://www.landsofamerica.com/" + location + "-" + state + "/all-land/no-house/sort-acre-high/page-" + str(currentPage) + '/'
        r = requests.get(url)
        soup = BeautifulSoup(r.text,'html.parser')
        #get price
        results2 = soup.find_all('span',attrs={'class':'price'})
        #get size
        results3 = soup.find_all('span',attrs={'class':'size'})
        #get location
        results4 = soup.find_all(attrs={"itemprop":"name"})
        #lotURL = soup.find_all('div',{"class": ["clearfix list-group-item list-property free"]})
        lotURL = soup.find_all('h3',attrs={'class':'panel-title'})
        i = 0

         #get the URL for each entry
        while(i < len(lotURL)):
          swap = str(lotURL[i])
          #print(swap)
          #find start of URL
          indexStart = swap.find('/property')
          #slice everything before that off
          swap = swap[indexStart:]
          #find end of URL
          indexEnd = swap.find('">')
          #slice everything after URL off
          swap = swap[0:indexEnd]
          #give it back to the original variable
          lotURL[i] = swap
          i = i + 1


        #while(i < len(lotURL)):
        #)

        #concatenate and print all of this information to terminal
        while counter < len(results2):
    	       #print(results3[counter].text + ".......... " + results2[counter].text  + ".........."+ results4[counter]['content'])
             #get APN
             url = "https://www.landsofamerica.com" + lotURL[counter]
             print(url)
             r = requests.get(url)
             soup = BeautifulSoup(r.text,'html.parser')
             #list(soup.children)
             #('p', class_='outer-text')
             apn = soup.find_all('div',{'class':'parcelTaxDetails'})
             #apn = BeautifulSoup(apn.text, 'html.parser')
             #apn = apn.find_all('td',{})
             print(str(apn[0]) + '\n')
             counter = counter + 1
        #reset the array iterator
        counter = 0
        #increment the page number
        currentPage = currentPage + 1
main()
