from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
import shutil

@task
def order_robots_from_robot_spare_bin():
    """
    Places orders for robots from Robot Spare Bin Industries Inc.
    Saves the order confirmation as a PDF file.
    Captures a screenshot of the ordered robot.
    Embeds the screenshot into the PDF receipt.
    Creates a ZIP archive containing the receipts and screenshots.
    """
    # Setting up browser speed
    browser.configure(slowmo=200)
    navigate_to_order_website()
    download_orders_csv()
    populate_form_with_csv_data()
    create_zip_archive()
    cleanup_files()

def navigate_to_order_website():
    """Opens the specified URL and handles any pop-up dialogs."""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    page = browser.page()
    page.click('text=OK')

def download_orders_csv():
    """Fetches the orders file from the provided URL."""
    http = HTTP()
    http.download("https://robotsparebinindustries.com/orders.csv", overwrite=True)

def order_additional_robot():
    """Clicks the button to order another robot."""
    page = browser.page()
    page.click("#order-another")

def confirm_order():
    """Handles confirmation for new robot orders."""
    page = browser.page()
    page.click('text=OK')

def fill_and_submit_robot_order(order):
    """Inputs the robot order details and clicks the 'Order' button."""
    page = browser.page()
    head_options = {
        "1": "Roll-a-thor head",
        "2": "Peanut crusher head",
        "3": "D.A.V.E head",
        "4": "Andy Roid head",
        "5": "Spanner mate head",
        "6": "Drillbit 2000 head"
    }
    selected_head = order["Head"]
    page.select_option("#head", head_options.get(selected_head))
    page.click('//*[@id="root"]/div/div[1]/div/div[1]/form/div[2]/div/div[{0}]/label'.format(order["Body"]))
    page.fill("input[placeholder='Enter the part number for the legs']", order["Legs"])
    page.fill("#address", order["Address"])
    
    while True:
        page.click("#order")
        additional_order_button = page.query_selector("#order-another")
        if additional_order_button:
            pdf_file_path = save_receipt_as_pdf(int(order["Order number"]))
            robot_screenshot_path = capture_robot_screenshot(int(order["Order number"]))
            insert_screenshot_into_receipt(robot_screenshot_path, pdf_file_path)
            order_additional_robot()
            confirm_order()
            break

def save_receipt_as_pdf(order_number):
    """Saves the order receipt as a PDF file."""
    page = browser.page()
    receipt_html = page.locator("#receipt").inner_html()
    pdf_handler = PDF()
    pdf_file_path = "output/receipts/{0}.pdf".format(order_number)
    pdf_handler.html_to_pdf(receipt_html, pdf_file_path)
    return pdf_file_path

def populate_form_with_csv_data():
    """Reads orders from a CSV file and fills in the form for each order."""
    csv_handler = Tables()
    robot_orders = csv_handler.read_table_from_csv("orders.csv")
    for order in robot_orders:
        fill_and_submit_robot_order(order)

def capture_robot_screenshot(order_number):
    """Takes a screenshot of the ordered robot."""
    page = browser.page()
    screenshot_file_path = "output/screenshots/{0}.png".format(order_number)
    page.locator("#robot-preview-image").screenshot(path=screenshot_file_path)
    return screenshot_file_path

def insert_screenshot_into_receipt(screenshot_path, pdf_file_path):
    """Embeds the robot screenshot into the PDF receipt."""
    pdf_handler = PDF()
    pdf_handler.add_watermark_image_to_pdf(image_path=screenshot_path, 
                                            source_path=pdf_file_path, 
                                            output_path=pdf_file_path)

def create_zip_archive():
    """Archives the receipts into a ZIP file."""
    archive_lib = Archive()
    archive_lib.archive_folder_with_zip("./output/receipts", "./output/receipts.zip")

def cleanup_files():
    """Removes the output folders for receipts and screenshots."""
    shutil.rmtree("./output/receipts")
    shutil.rmtree("./output/screenshots")
