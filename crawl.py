from selenium import webdriver
from openpyxl import  load_workbook
from time import sleep
from Product import Product
from shutil import copyfile
import os.path

INPUT_PATH = './Input/template.xlsx'
OUTPUT_PATH = './Output/ouput.xlsx'

if os.path.exists(OUTPUT_PATH):
    os.remove(OUTPUT_PATH)
copyfile(INPUT_PATH, OUTPUT_PATH)

###### setup ##########
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--headless')
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')
driver = webdriver.Chrome('chromedriver', chrome_options=chrome_options)
wb = load_workbook(OUTPUT_PATH)
sheetRanges = wb.worksheets[1]

global currentRow
currentRow = 6
base_url = 'https://shopee.vn'

###################################

#### Your config ####
linkShop = "https://shopee.vn/quandinhvan"
diffPercent = 30 #percent
prefixOfName = '[ULJU Store]'
remainProductInStock = '1000'
descriptionHeader = 'Sản phẩm: '
descriptionFooter = 'Chi tiết liên hệ: 0982255743'
channelLogistics = {'JandTExpress'  : 'Mở',
                    'NowShip'       : 'Mở', 
                    'NinjaVan'      : 'Mở', 
                    'BestExpress'   : 'Mở', 
                    'GHN'           : 'Mở',
                    'VTPost'        : 'Mở', 
                    'GHTK'          : 'Mở', 
                    'VNPostSave'    : 'Mở', 
                    'GrabExpress'   : 'Mở',
                    'VNPostFast'    : 'Mở'}
weight = 200
#####################

def writeProductIntoExcel(data, current_row):
    rangeList = ['A',
                 'B',
                 'C',
                 'D',
                 'E',
                 'F',
                 'G',
                 'H',
                 'I',
                 'J',
                 'K',
                 'L',
                 'M',
                 'N',
                 'O',
                 'P',
                 'Q',
                 'R',
                 'S',
                 'T',
                 'U',
                 'V',
                 'W',
                 'X',
                 'Y',
                 'Z',
                 'AA',
                 'AB',
                 'AC',
                 'AD',
                 'AE',
                 'AF',
                 'AG',
                 'AH',
                 'AI',
                 'AJ',
                 'AK'
                 ]

    it = 0
    for cell in data.Info:
        col_current = rangeList[it] + str(current_row)
        sheetRanges[col_current] = data.Info[cell]
        it+=1

def getSpecificProduct(linkProduct, index):
    driver.get(linkProduct)
    sleep(5)

    breadcrumb  = '/html/body/div[1]/div/div[2]/div[2]/div[2]/div[1]/a[@href]'
    name = '/html/body/div[1]/div/div[2]/div[2]/div[2]/div[2]/div[3]/div/div[1]/span'
    description = '/html/body/div[1]/div/div[2]/div[2]/div[2]/div[3]/div[2]/div[1]/div[1]/div[2]/div[2]/div/span'
    variation1 = '/html/body/div[1]/div/div[2]/div[2]/div[2]/div[2]/div[3]/div/div[4]/div/div[3]/div/div[1]/label'
    title_option_for_variation1 = '/html/body/div[1]/div/div[2]/div[2]/div[2]/div[2]/div[3]/div/div[4]/div/div[3]/div/div[1]/div/button'
    variation2 = '/html/body/div[1]/div/div[2]/div[2]/div[2]/div[2]/div[3]/div/div[4]/div/div[3]/div/div[2]/label'
    title_option_for_variation2 = '/html/body/div[1]/div/div[2]/div[2]/div[2]/div[2]/div[3]/div/div[4]/div/div[3]/div/div[2]/div/button'
    

    mProduct = Product()
    ### process get image ###
    # print('Process get Image')
    buttonNextImage = None
    try:
        buttonNextImage = driver.find_element_by_xpath(
            '/html/body/div[1]/div/div[2]/div[2]/div[2]/div[2]/div[2]/div[1]/div[2]/button[2]')
    except:
        # print("Dont have any button next")
        pass
    

    # get 5 image
    listLinkImage = []
    listElementImage = driver.find_elements_by_xpath(
        '/html/body/div[1]/div/div[2]/div[2]/div[2]/div[2]/div[2]/div[1]/div[2]/div/div/div[1]/div[@style]')

    for elem in listElementImage:
        style = elem.get_attribute('style')
        s = style.find('(')+2
        e = style.find(')')-4
        listLinkImage.append(style[s:e])
        #print(style[s:e])
    
    stillHaveImage = True
    lastImageXpath = '/html/body/div[1]/div/div[2]/div[2]/div[2]/div[2]/div[2]/div[1]/div[2]/div[5]/div/div[1]/div[@style]'

    while stillHaveImage and buttonNextImage!= None:
        buttonNextImage.click()
        sleep(2)
        style = driver.find_element_by_xpath(lastImageXpath).get_attribute('style')
        s = style.find('(')+2
        e = style.find(')')-4
        newImageLink = style[s:e]
        if newImageLink not in listLinkImage:
            listLinkImage.append(newImageLink)
        else:
            stillHaveImage = False
        
    # print(listLinkImage)
    it = 0
    for it in range(len(listLinkImage)):
        if it == 0:
            mProduct.Info['ps_item_cover_image'] = listLinkImage[it]
        else:
            key_info = 'ps_item_image_'+str(it)
            mProduct.Info[key_info] = listLinkImage[it]
        it+=1
    
    
    #########################

    # get category of product
    breadcrumTagA = driver.find_elements_by_xpath(breadcrumb)
    lastTagA = breadcrumTagA[len(breadcrumTagA) - 1].get_attribute('href').split('.')
    categoryNumber = lastTagA[len(lastTagA)-1]
    mProduct.Info['ps_category'] = categoryNumber
    ########################

    #get product name
    name_span = driver.find_element_by_xpath(name)
    real_name = name_span.text
    if '[' in real_name or ']' in real_name:
        e = real_name.find(']')+1
        real_name = real_name[e:]

    mProduct.Info['ps_product_name'] = prefixOfName + ' ' + real_name
    #########################

    #get product description
    description_text = driver.find_element_by_xpath(description).text
    mProduct.Info['ps_product_description'] = descriptionHeader + name_span.text + '\n' + description_text + descriptionFooter
    ##########################

    #et_title_variation_integration_no
    mProduct.Info['et_title_variation_integration_no'] = index
    ####################

    #price 
    price = '/html/body/div[1]/div/div[2]/div[2]/div[2]/div[2]/div[3]/div/div[3]/div/div/div[1]/div/div[2]/div[1]'
    price_temp = '/html/body/div[1]/div/div[2]/div[2]/div[2]/div[2]/div[3]/div/div[3]/div/div/div/div/div/div'
    price_text = None

    try:
        price_text = driver.find_element_by_xpath(price).text
    except Exception as ex:
        pass

    try:
        if price_text == None:
            price_text = driver.find_element_by_xpath(price_temp).text
    except Exception as ex:
        pass

    price_text = price_text.split('-')
    it = len(price_text) - 1
    price_text = price_text[it].strip()
    price_text = price_text[1:].split('.')
    price_text = ''.join(price_text)
    mProduct.Info['ps_price'] = str(int(int(price_text)*((diffPercent+100)/100)))
    #remain product in stock
    mProduct.Info['ps_stock'] = remainProductInStock
    #weight
    mProduct.Info['ps_weight'] = weight
    ####################

    #logistics channel
    mProduct.Info['channel_id_50018'] = channelLogistics['JandTExpress'] # J&T Express
    mProduct.Info['channel_id_50022'] = channelLogistics['NowShip'] # NowShip
    mProduct.Info['channel_id_50023'] = channelLogistics['NinjaVan'] # Ninja Van
    mProduct.Info['channel_id_50024'] = channelLogistics['BestExpress'] # BEST Express
    mProduct.Info['channel_id_50011'] = channelLogistics['GHN'] # Giao Hàng Nhanh
    mProduct.Info['channel_id_50010'] = channelLogistics['VTPost'] # Viettel Post
    mProduct.Info['channel_id_50012'] = channelLogistics['GHTK'] # Giao Hàng Tiết Kiệm
    mProduct.Info['channel_id_50016'] = channelLogistics['VNPostSave'] # VNPost Tiết Kiệm
    mProduct.Info['channel_id_50020'] = channelLogistics['GrabExpress'] # GrabExpress
    mProduct.Info['channel_id_50015'] = channelLogistics['VNPostFast'] # VNPost Nhanh
    ###########################

    #copy new output excel file and open it

    # variation
    variation1_label = None
    variation2_label = None
    title_option_for_variation1_list = None
    title_option_for_variation2_list = None

    try:
        variation1_label = driver.find_element_by_xpath(variation1).text
        title_option_for_variation1_list = driver.find_elements_by_xpath(
            title_option_for_variation1)
    except Exception as ex:
        # print(ex)
        pass

    try:
        variation2_label = driver.find_element_by_xpath(variation2).text
        title_option_for_variation2_list = driver.find_elements_by_xpath(
            title_option_for_variation2)
    except Exception as ex:
        # print(ex)
        pass

    global currentRow
    if variation1_label != None and variation2_label != None:
        for var1 in title_option_for_variation1_list:
            for var2 in title_option_for_variation2_list:
                mProduct.Info['et_title_variation_1'] = variation1_label
                mProduct.Info['et_title_variation_2'] = variation2_label
                mProduct.Info['et_title_option_for_variation_1'] = var1.text
                mProduct.Info['et_title_option_for_variation_2'] = var2.text
                mProduct.Info['et_title_image_per_variation'] = mProduct.Info['ps_item_cover_image']
                # print(mProduct.Info)
                writeProductIntoExcel(mProduct, currentRow)
                currentRow += 1
    elif variation1_label != None and variation2_label == None:
        for var1 in title_option_for_variation1_list:
            mProduct.Info['et_title_variation_1'] = variation1_label
            mProduct.Info['et_title_option_for_variation_1'] = var1.text
            mProduct.Info['et_title_image_per_variation'] = mProduct.Info['ps_item_cover_image']
            # print(mProduct.Info)
            writeProductIntoExcel(mProduct, currentRow)
            currentRow += 1
    elif variation1_label == None and variation2_label != None:
        for var2 in title_option_for_variation2_list:
            mProduct.Info['et_title_variation_1'] = variation2_label
            mProduct.Info['et_title_option_for_variation_1'] = var2.text
            mProduct.Info['et_title_image_per_variation'] = mProduct.Info['ps_item_cover_image']
            # print(mProduct.Info)
            writeProductIntoExcel(mProduct, currentRow)
            currentRow += 1


    ################################
    #print(mProduct.Info)
    sleep(5)

# Main process
def main():
    driver.get(linkShop)
    sleep(10)
    driver.find_element_by_xpath(
        '/html/body/div[1]/div/div[2]/div[2]/div[2]/div/div[3]/div[5]/div[2]/div[1]/div[1]/div[3]').click()
    sleep(2)
    # lấy tất cả sản phẩm ở trang 1
    items = driver.find_elements_by_xpath(
        '/html/body/div[1]/div/div[2]/div[2]/div[2]/div/div[3]/div[5]/div[2]/div[2]/div/div/div/a[@href]')

    links = []
    print(len(items))

    for item in items:
        i = item.get_attribute('href')
        links.append(i)
        # print(i)

    index = 1
    
    # getSpecificProduct(
    #     'https://shopee.vn/Quần-Baggy-Tây-chun-sau-i.587221.5110409602', index)
    for link in links:
        print('Product ', index)
        getSpecificProduct(link, index)
        index+=1

try:
    main()
except Exception as ex:
    print(ex)
finally:
    wb.save(OUTPUT_PATH)
    driver.close()
