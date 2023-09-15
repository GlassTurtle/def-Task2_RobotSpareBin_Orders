from robocorp.tasks import task
from robocorp import browser, http, excel
from RPA.PDF import PDF
from RPA.Tables import Tables
from RPA.Archive import Archive
import time


@task
def Task2_RobotSpareBin_Orders():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
     
    browser.configure(
        #slowmo=500,
        slowmo=50,
    )

    open_robot_order_website()
    orders = get_orders()
    
    # test1 = orders.get_row(1)
    # print(test1['Head'])
    # print(test1['Body'])

    # fill_the_form(orders.get_row(1), browser.page())
    # print(len(orders._data))
    for i in range(len(orders._data)):
    # for i in range(0,3):
        # print(i)
        # any way for print outputs to show during execution? Doesn't seem to show in terminal while robot is running.
        close_annoying_modal()
        OrderGot = fill_the_form(orders.get_row(i))
        page = browser.page()

        if OrderGot == False:
            page.reload()
            continue
        
        
        ordernum = page.locator('[class="badge badge-success"]').inner_html()
        # print (ordernum)
        # either create a new receipts folder/run? or clear existing receipts? 
        # Are the order number based on time or just the input? Probably time/random generation.
        path_orderpdf = store_receipt_as_pdf(ordernum)
        path_screensht = screenshot_robot(ordernum)
        embed_screenshot_to_receipt(path_screensht, path_orderpdf)

        # print(path_orderpdf)
        # print(path_screensht)
        # time.sleep(2)

        page.click('[id="order-another"]')

    archive_receipts()
    # time.sleep(5)

def archive_receipts():
    """ zip the output into an archive """
    lib = Archive()
    lib.archive_folder_with_zip('output/receipts', 'receipts.zip')
    # anyway to change to location of the .zip?
    # does this delete the receipts folder?
    # the zips are cumulative? is it zipping non-permanently deleted files?
         # I was stupid... '/output/receipts' stores in drive root...

def embed_screenshot_to_receipt(screenshot, pdf_file):
    """ add the screenshot to the pdf doc  """
    pdf = PDF()
    # pdf.add_files_to_pdf([pdf_file, screenshot], target_document=pdf_file)
    pdf.add_files_to_pdf([pdf_file, screenshot], pdf_file)
    # pdf.add_files_to_pdf(
    #     files=list_of_files,
    #     target_document="output/output.pdf"
    # )

def screenshot_robot(order_number):
    """ saves screenshot of the receipts page """
    page = browser.page()
    page.screenshot(path="output/receipts/"+ order_number + ".png")
    return "output/receipts/"+ order_number + ".png"
    

def store_receipt_as_pdf(order_number):
    """ takes the order number, and return the file system path to the PDF file which includes the order number"""
    page = browser.page()
    pdf = PDF()
    # filename = "/output/receipts/" + order_number + ".pdf"
    HTML_print = page.locator('[id="order-completion"]').inner_html() 
    pdf.html_to_pdf(HTML_print, "output/receipts/" + order_number + ".pdf")
    return "output/receipts/" + order_number + ".pdf"
    #the pdf functions don't return the path right? - nope


def fill_the_form(row):
    """ for an inputted row, fill in the form data  """
    page = browser.page()

    # page.select_option("#head",int(row['Head']))
    page.select_option("#head", row['Head'])
    
    
    # page.click('css:input[name="body"][value="1"]')
    # below doesn't work, is it just 'id:' just a misinterpretation from the ReMark assistant?
    # print('id:id-body-' + row['Body'])
    page.click("#id-body-"+ row['Body'])
    # page.click('css:input[id="id-body-1"]')
    # '.radio form-check'

    #are these in Playwright, not Selenium?

    # page.fill("#1694322371920", row['Legs'])
    # placeholder="Enter the part number for the legs"
    # page.fill('[id="1694322371920"]', row['Legs']) 1694335019978
    # why does the id attribute change? (according to the reqs - is it intentially set up that way? or something about the form type itself?)
    page.fill('[placeholder="Enter the part number for the legs"]', row['Legs'])

    page.fill("#address", row['Address'])

    page.click('[id="preview"]')

    # try for ordering, skip if none within n tries
    # while success == False
    i = 0
    while i < 8:
        page.click('[id="order"]')
        # try: #nevermind, try waits for timeout, just detect order another button instead
        #     #breaks if the order another button can be found, tries to order again if not found
        #     #waits for timeout, though, can change? or just if statement to look for?
        #     time.sleep(3)
        #     page.click('[id="order-another"]')
        #     break
        # except:
        #     print('Attempting to order again')
        #     i = i+1

        # click the order again if it can be found, if not then try to click order again.
        if page.query_selector('[id="order-another"]'):
            # time.sleep(3)
            # page.click('[id="order-another"]')
            # click 'another order' after the pdf export
            # moved to main.
            # break
            return True

        print('Attempting to order again')
        i = i+1


    if i==8:
        print('Could not order')
        return False

    



def open_robot_order_website():
    """ opens the website to order """
    # TODO: Implement your function here
    try:
        browser.goto("https://robotsparebinindustries.com/#/robot-order")
    except:
        print("Could not connect")


def get_orders():
    """ download the orders CSV file, read then return"""
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    
    try:
        # orders1 = Tables()
        # orders1 = orders1.read_table_from_csv('orders.csv')
        #what's the difference between table type and orders type?
        table = Tables()
        orders1 = table.read_table_from_csv('orders.csv')
        
        # something about looping orders at this step? But not at input stage yet?? above funtion reads all data...

        # print(orders1.get_cell(0,3)) 
        # print(orders1.get_cell(1,3))
        # print(orders1.get_cell(2,3))
        return orders1
    except:
        print("Table read failed")
        return 0

    
def close_annoying_modal():
    """click on a button to close the pop up"""

    page = browser.page()
    try:
        # in case the Order button breaks more than N times...
        # annoying modal won't be found, so just skip, and re-enter values for next bot
        page.click("button:text('I guess so...')")
    except:
        return
    # time.sleep(5)

    


    

