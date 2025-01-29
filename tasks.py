from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP  
from RPA.PDF import PDF
from RPA.Tables import Tables
from RPA.Archive import Archive
import os
import time

@task
def order_robots_from_RobotSpareBin():
    """Procesa todas las órdenes de robots desde RobotSpareBin Industries Inc."""
    browser.configure(slowmo=100, headless=False)
    open_robot_order_website()
    orders = get_orders()
    
    for order in orders:
        close_annoying_modal()
        fill_the_form(order)  
        preview_robot()
        submit_order(order["Order number"])
        
        # Manejo de comprobantes
        pdf_path = store_receipt_as_pdf(order["Order number"])
        screenshot_path = screenshot_robot(order["Order number"])
        embed_screenshot_to_receipt(screenshot_path, pdf_path)

        # Eliminar o comentar estas líneas de depuración
        # if order["Order number"] == "5":  
        #     breakpoint()
        
        go_to_order_another_robot()
    
    # Archivo final fuera del loop
    zip_path = archive_receipts()
    print(f"ZIP creado en: {zip_path}")

def open_robot_order_website():
    """Abre el sitio web de pedidos"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def get_orders():
    """Descarga y lee el CSV"""
    http = HTTP()  
    http.download(
        url="https://robotsparebinindustries.com/orders.csv",
        target_file="orders.csv",
        overwrite=True
    )
    return Tables().read_table_from_csv("orders.csv")

def close_annoying_modal():
    """Cierra el modal de alerta"""
    page = browser.page()
    page.click("button:text('OK')")

def fill_the_form(order):
    """Rellena el formulario con datos de la orden"""
    page = browser.page()
    page.select_option("#head", str(order["Head"]))
    page.click(f'xpath=//input[@name="body" and @value="{order["Body"]}"]')
    page.fill("input[type='number'][placeholder='Enter the part number for the legs']", str(order["Legs"]))
    page.fill("#address", order["Address"])

def preview_robot():
    """Genera la vista previa del robot"""
    page = browser.page()
    page.click("#preview")

def submit_order(order_number):
    """Envía la orden con manejo de errores"""
    page = browser.page()
    for _ in range(5):
        page.click("#order")
        if page.query_selector("#receipt"):
            return
        time.sleep(1)
    raise Exception(f"Fallo al enviar orden {order_number}")

def go_to_order_another_robot():
    """Reinicia el formulario para nueva orden"""
    page = browser.page()
    page.click("#order-another")

def store_receipt_as_pdf(order_number):
    """Guarda el recibo como PDF"""
    page = browser.page()
    receipt_html = page.locator("#receipt").inner_html()
    
    output_dir = "output/receipts"
    os.makedirs(output_dir, exist_ok=True)
    pdf_path = os.path.join(output_dir, f"receipt_{order_number}.pdf")
    
    PDF().html_to_pdf(receipt_html, pdf_path)
    return pdf_path

def screenshot_robot(order_number):
    """Captura la imagen del robot"""
    page = browser.page()
    output_dir = "output/screenshots"
    os.makedirs(output_dir, exist_ok=True)
    screenshot_path = os.path.join(output_dir, f"robot_{order_number}.png")
    
    page.locator("#robot-preview").screenshot(path=screenshot_path)
    return screenshot_path

def embed_screenshot_to_receipt(screenshot, pdf_file):
    """Incrusta la captura en el PDF"""
    PDF().add_files_to_pdf(
        files=[screenshot],
        target_document=pdf_file,
        append=True
    )

def archive_receipts():
    """Crea archivo ZIP con los comprobantes"""
    lib = Archive()
    receipts_dir = os.path.abspath("output/receipts")
    output_dir = os.path.abspath("output")
    
    # Asegurar que el directorio de salida existe
    os.makedirs(output_dir, exist_ok=True)
    
    if not os.path.exists(receipts_dir):
        raise FileNotFoundError(f"Directorio no encontrado: {receipts_dir}")
    
    # Verificar que hay archivos PDF para comprimir
    pdf_files = [f for f in os.listdir(receipts_dir) if f.endswith(".pdf")]
    if not pdf_files:
        raise Exception("No hay archivos PDF para comprimir")
    
    archive_name = os.path.join(output_dir, "receipts.zip")
    
    # Eliminar el archivo ZIP si ya existe
    if os.path.exists(archive_name):
        os.remove(archive_name)
    
    # Crear el archivo ZIP sin el parámetro overwrite
    lib.archive_folder_with_zip(
        folder=receipts_dir,
        archive_name=archive_name,
        include="*.pdf",
        recursive=True
    )
    
    # Verificar que el ZIP se creó correctamente
    if not os.path.exists(archive_name):
        raise Exception(f"Fallo al crear el archivo ZIP: {archive_name}")
    
    print(f"Archivo ZIP creado en: {archive_name}")
    return archive_name