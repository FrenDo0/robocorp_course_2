from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

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
    accept_tos()
    download_csv_file()
    read_csv_file()
    archive_receipts()


def open_robot_order_website():
    browser.configure(
        slowmo=100,
    )
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def download_csv_file():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download("https://robotsparebinindustries.com/orders.csv",overwrite=True)

def read_csv_file():
    """Read csv data table"""
    data = Tables()
    orders = data.read_table_from_csv(
        "orders.csv", columns=["Order number","Head","Body","Legs","Address"]
        )
    for row in orders:
        fill_order_form(row)

def fill_order_form(order):
    """Fill the order form with data from csv file"""

    page = browser.page()
    page.select_option("#head",str(order["Head"]))
    body_id = "#id-body-" + str(order["Body"])
    page.click(body_id)
    page.fill("//html/body/div/div/div[1]/div/div[1]/form/div[3]/input",order["Legs"])
    page.fill("#address",order["Address"])
    page.click("text=Preview")
    order_number = order["Order number"]
    new_window_avaiable = False
    while not new_window_avaiable:
        page.click("#order")
        new_window_avaiable = page.is_visible("#order-another",timeout=1000)
        if new_window_avaiable:
            screenshot = screenshot_robot(order_number)
            file_pdf = store_receipt_in_pdf(order_number)
            embed_screenshot_to_receipt(screenshot,file_pdf)
            page.click("#order-another")
    accept_tos()

def accept_tos():
    """Accept terms of service"""
    page = browser.page()
    page.click("text=Ok")

def store_receipt_in_pdf(order_number):
    """Stores the receipt in pdf by given order number"""
    page = browser.page()
    receipt_result_html = page.locator("#receipt").inner_html()
    pdf = PDF()
    pdf.html_to_pdf(receipt_result_html, "archives/receipt_result_"+order_number+".pdf")
    return "receipt_result_"+order_number+".pdf"

def screenshot_robot(order_number):
    """Function to screenshot the robot"""
    page = browser.page()
    page.screenshot(path="archives/robot_"+order_number+".png")
    return "robot_"+order_number+".png"

def embed_screenshot_to_receipt(screenshot, pdf_file):
    """Function to insert schreenshot of the robot into pdf file with the receipt"""
    pdf = PDF()
    list_of_files = ["archives/"+pdf_file,"archives/"+screenshot]
    pdf.add_files_to_pdf(
        files=list_of_files,
        target_document="full_receipt_archive/"+pdf_file+"_full_receipt.pdf"
    )

def archive_receipts():
    """Function to put receipts in ZIP"""
    lib = Archive()
    lib.archive_folder_with_zip('./full_receipt_archive', 'output/receipts.zip')

