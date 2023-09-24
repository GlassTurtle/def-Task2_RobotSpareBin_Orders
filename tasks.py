from robocorp.tasks import task
from robocorp import browser, http, excel
from RPA.PDF import PDF
from RPA.Tables import Tables
from RPA.Archive import Archive
import time
import os
import glob

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

    # clear existing receipts folder for local runs
    # path_receipts = "output/receipts/"
    # clear_prev_receipts(path_receipts)

    open_robot_order_website()
    orders = get_orders()
    
    for i in range(len(orders._data)):
        close_annoying_modal()
        fill_the_form(orders.get_row(i))
        page = browser.page()
        ordernum = page.locator('[class="badge badge-success"]').inner_html()
        
        path_orderpdf = store_receipt_as_pdf(ordernum)
        path_screensht = screenshot_robot(ordernum)
        embed_screenshot_to_receipt(path_screensht, path_orderpdf)


        page.click('[id="order-another"]')

    archive_receipts()

def archive_receipts():
    """ zip the output into an archive """
    lib = Archive()
    lib.archive_folder_with_zip('output/receipts', 'receipts.zip')

def embed_screenshot_to_receipt(screenshot, pdf_file):
    """ add the screenshot to the pdf doc  """
    pdf = PDF()
    pdf.add_watermark_image_to_pdf(
        image_path=screenshot,
        source_path=pdf_file,
        output_path=pdf_file,
        coverage=0.2
    )

def screenshot_robot(order_number):
    """ saves screenshot of the receipts page """
    filename = "output/receipts/"+ order_number + ".png"
    time.sleep(0.5)#brief pause to allow image preview to fully load
    page = browser.page()
    RobImg = page.locator('[id="robot-preview-image"]')
    RobImg.screenshot(path=filename)
    return filename
    
def store_receipt_as_pdf(order_number):
    """ takes the order number, and return the file system path to the PDF file which includes the order number"""
    filename = "output/receipts/" + order_number + ".pdf"
    page = browser.page()
    HTML_print = page.locator('[id="receipt"]').inner_html()
    HTML_print = f'<div style="margin-top: 20cm;margin-left: 50px;margin-right: 20px;">{HTML_print}</div>'
    # HTML_print = f'<div style="margin: 50px 100px 50px 100px;">{HTML_print}</div>'

    pdf = PDF()
    pdf.html_to_pdf(HTML_print, filename)
    
    return filename

def clear_prev_receipts(path_clearfiles):
    """ Remove existing files from previous runs for local runs """
    #number of pdf files should be the same as png files, but just in case.
    files1 = glob.glob(f"{path_clearfiles}*.pdf")
    files2 = glob.glob(f"{path_clearfiles}*.png")
    for file1 in files1:
        os.remove(file1)
    for file2 in files2:
        os.remove(file2)


def fill_the_form(row):
    """ for an inputted row, fill in the form data  """
    page = browser.page()

    page.select_option("#head", row['Head'])
    page.click("#id-body-"+ row['Body'])
    page.fill('[placeholder="Enter the part number for the legs"]', row['Legs'])
    page.fill("#address", row['Address'])
    page.click('[id="preview"]')

    # Loop click "Order" unitl "Order-another" button is found
    while not(page.query_selector('[id="order-another"]')):
        page.click('[id="order"]')


def open_robot_order_website():
    """ opens the website to order """
    try:
        browser.goto("https://robotsparebinindustries.com/#/robot-order")
    except:
        print("Could not connect")


def get_orders():
    """ download the orders CSV file, read then return"""
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    table = Tables()
    orders1 = table.read_table_from_csv('orders.csv')

    return orders1

    
def close_annoying_modal():
    """click on a button to close the pop up, return if not found"""
    page = browser.page()
    page.click("button:text('I guess so...')")


    


    

