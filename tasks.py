# from robocorp.tasks import task
# from robocorp import browser

# from RPA.HTTP import HTTP
# from RPA.Excel.Files import Files
# from RPA.PDF import PDF

# @task
# def robot_spare_bin_python():
#     """Insert the sales data for the week and export it as a PDF"""
#     browser.configure(
#         slowmo=100,
#     )
#     open_the_intranet_website()
#     log_in()
#     download_excel_file()
#     fill_form_with_excel_data()
#     collect_results()
#     export_as_pdf()
#     log_out()

# def open_the_intranet_website():
#     """Navigates to the given URL"""
#     browser.goto("https://robotsparebinindustries.com/")

# def log_in():
#     """Fills in the login form and clicks the 'Log in' button"""
#     page = browser.page()
#     page.fill("#username", "maria")
#     page.fill("#password", "thoushallnotpass")
#     page.click("button:text('Log in')")

# def fill_and_submit_sales_form(sales_rep):
#     """Fills in the sales data and click the 'Submit' button"""
#     page = browser.page()

#     page.fill("#firstname", sales_rep["First Name"])
#     page.fill("#lastname", sales_rep["Last Name"])
#     page.select_option("#salestarget", str(sales_rep["Sales Target"]))
#     page.fill("#salesresult", str(sales_rep["Sales"]))
#     page.click("text=Submit")

# def download_excel_file():
#     """Downloads excel file from the given URL"""
#     http = HTTP()
#     http.download(url="https://robotsparebinindustries.com/SalesData.xlsx", overwrite=True)

# def fill_form_with_excel_data():
#     """Read data from excel and fill in the sales form"""
#     excel = Files()
#     excel.open_workbook("SalesData.xlsx")
#     worksheet = excel.read_worksheet_as_table("data", header=True)
#     excel.close_workbook()

#     for row in worksheet:
#         fill_and_submit_sales_form(row)

# def collect_results():
#     """Take a screenshot of the page"""
#     page = browser.page()
#     page.screenshot(path="output/sales_summary.png")

# def export_as_pdf():
#     """Export the data to a pdf file"""
#     page = browser.page()
#     sales_results_html = page.locator("#sales-results").inner_html()

#     pdf = PDF()
#     pdf.html_to_pdf(sales_results_html, "output/sales_results.pdf")

# def log_out():
#     """Presses the 'Log out' button"""
#     page = browser.page()
#     page.click("text=Log out")


from robocorp.tasks import task
from robocorp import browser
import requests
from RPA.Tables import Tables

from RPA.PDF import PDF
from pathlib import Path
import time
import os 
import zipfile
from RPA.FileSystem import FileSystem


@task
def order_robots_from_RobotSpareBin():
    csv_url='https://robotsparebinindustries.com/orders.csv'
    local_csv_file='order_downloaded.csv'
    folder_path = 'output/receipt'
    output_path = 'output/receipt.zip'
    os.makedirs(folder_path)

    browser.configure(
        slowmo=1000,
    )
    open_robot_order_website()
    close_annoying_modal()

    download_csv(csv_url, local_csv_file)
    orders = get_orders(local_csv_file)
    fill_the_form(orders)
    
    archive_receipts(folder_path , output_path)


def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")


def close_annoying_modal():
    page = browser.page()
    page.click("text=OK")


def download_csv(url: str, filename: str, overwrite: bool = True):
    if overwrite:
        response = requests.get(url)
        response.raise_for_status()  # Ensure we notice bad responses
        with open(filename, 'wb') as f:
            f.write(response.content)
        print(f"File downloaded and saved as {filename}")
    else:
        print("File exists and overwrite is set to False. Download skipped.")
    return filename

def get_orders(csv_file: str):
    table_library = Tables()
    orders_table = table_library.read_table_from_csv(csv_file)
    return orders_table
     
def store_receipt_as_pdf(order_number, page):
    page = browser.page()    
    order_results_html = page.locator("#receipt").inner_html()
    pdf = PDF()
    receipt_pdf_file_path =f"output/receipt/order_results_{order_number}.pdf"

    pdf.html_to_pdf(order_results_html,receipt_pdf_file_path )
    return receipt_pdf_file_path

def screenshot_robot(order_number):
    """Take a screenshot of the page"""
    page = browser.page()
    path = f"output/receipt/sales_summary_{order_number}.png"
    page.screenshot(path=path )
    return path

def embed_screenshot_to_receipt(screenshot, pdf_file):
    pdf = PDF()
    pdf.add_watermark_image_to_pdf(image_path=screenshot, output_path=pdf_file, source_path=pdf_file, coverage=0.2)


def archive_receipts(folder_path , output_path):
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        # Walk through the folder and add files to the zip file
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                # Create the full file path
                file_path = os.path.join(root, file)
                # Add the file to the zip file, preserving the directory structure
                zipf.write(file_path, os.path.relpath(file_path, folder_path))


def fill_the_form(orders, max_retries=3):
    unique_number = 0
    for order in orders:
        retries = 0
        while retries < max_retries:
            try:

                ###################### First page Filling Form ########################
                page = browser.page()
                unique_number += 1
                page.select_option("#head", value=str(order["Head"]))
                selector = f"input[name='body'][value='{order['Body']}']"
                page.click(selector)
                page.fill("input.form-control[type='number']", str(order["Legs"]))
                page.fill("#address", str(order["Address"]))
                page.click("#order")


                ###################### Second page saving pdf and Start next Order #################  
                page = browser.page()
                pdf_file_path     =   store_receipt_as_pdf(unique_number, page)
                screen_short_path =   screenshot_robot(unique_number)
                embed_screenshot_to_receipt(screen_short_path, pdf_file_path)
                page.click('#order-another')
                close_annoying_modal()


                break
            except Exception as e:
                print(f"Error encountered: {e}. Retrying...")
                retries += 1
                time.sleep(2)  # Wait for 2 seconds before retrying
                if retries >= max_retries:
                    print("Max retries reached. Moving to next order.")
                    break
    
