import os
import zipfile
from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF
from RPA.Tables import Tables

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """

    open_robot_order_website()
    orders = get_orders()
    fill_form_with_excel_data(orders)
    archive_receipts()

def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def close_annoying_modal():
    page = browser.page()
    page.click("button:text('OK')")

def get_orders():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

    library = Tables()
    orders = library.read_table_from_csv(
        "orders.csv", columns=["Order number", "Head", "Body", "Legs", "Address"]
    )

    return orders

def fill_form_with_excel_data(orders):
    """Read data from excel and fill in the sales form"""

    for row in orders:
        close_annoying_modal()
        fill_the_form(row)

def fill_the_form(row):
    """Read data from excel and fill in the sales form"""
    page = browser.page()

    page.select_option("#head", index=int(row["Head"]))
    page.fill("#address", str(row["Address"]))
    page.click(f'#id-body-{row["Body"]}')
    page.fill("input[type='number'][placeholder='Enter the part number for the legs']", str(row["Legs"]))
    
    # TRY wait_for_selector
    while True:
        try:
            page.click("#order")
            page.hover('#receipt', timeout=100)
            break
        except:
            print("An exception occurred")

    pdf_file = store_receipt_as_pdf(row["Order number"])
    screenshot = screenshot_robot(row["Order number"])
    embed_screenshot_to_receipt(screenshot, pdf_file)
    page.click("#order-another")
    
def screenshot_robot(order_number):

    page = browser.page()
    file_name = f'output/img_{order_number}.png'
    page.locator("#robot-preview-image").screenshot(path=file_name)
    return file_name

def store_receipt_as_pdf(order_number):

    page = browser.page()
    receipt_html = page.locator("#receipt").inner_html()

    pdf = PDF()
    file_name = f'output/receipt_{order_number}.pdf'
    pdf.html_to_pdf(receipt_html, file_name)
    return file_name

def embed_screenshot_to_receipt(screenshot, pdf_file):

    pdf = PDF()
    pdf.add_files_to_pdf(
        files=[screenshot],
        target_document=pdf_file,
        append=True
    )

def archive_receipts():

    pdf_files = [file for file in os.listdir('output') if file.endswith('.pdf')]

    with zipfile.ZipFile('output/receipts.zip', 'w') as zip_file:
        for file in pdf_files:
            complete_file = os.path.join('output', file)
            zip_file.write(complete_file, os.path.basename(complete_file))
