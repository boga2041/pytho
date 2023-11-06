import shutil
import time
from datetime import date
import pymysql


from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from urllib.parse import urlparse
import zipfile
import time
import os
import glob
from email import encoders
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
import smtplib

from webdriver_manager.chrome import ChromeDriverManager
from win10toast import ToastNotifier
import email.message
from selenium import webdriver
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.chrome.webdriver import WebDriver
import threading

from conexion import miConexion


s = miConexion.cursor()

sql = "SELECT correo_pac,password_pac FROM `empresas` where estatus = 1"

s.execute(sql)

myresult = s.fetchall()

chrome_options = webdriver.ChromeOptions()

preferences = {"download.default_directory": r"C:\Test",
               "download.prompt_for_download": False,
               "download.directory_upgrade": True}

chrome_options.add_experimental_option("prefs", preferences)


WebDriver = webdriver.Chrome(ChromeDriverManager().install(),
                             chrome_options=chrome_options)

urls = ['https://facturacion.finkok.com/cuentas/ingresar/',
        'https://facturacion.finkok.com/cuentas/ingresar/?next=/app/',
        'https://facturacion.finkok.com/app/',
        'https://facturacion.finkok.com/cuentas/ingresar/?next=/app/rbilling/']


def minfun(user, password, url):
    WebDriver.get(url)

    try:

        input_user = WebDriverWait(WebDriver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//input[@name="username"]'))
        )
        input_user.send_keys(user)
    except WebDriverException as e:
        print("no se ejecuto1")
        print(e)

    try:
        input_pass = WebDriverWait(WebDriver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//input[@name="password"]'))
        )
        input_pass.send_keys(password)
    except WebDriverException as e:
        print("no se ejecuto2")
        print(e)
    try:
        boon = WebDriverWait(WebDriver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//button[normalize-space()="Ingresar"]'))
        )
        boon.click()
    except WebDriverException as e:
        print("no se ejecuto3")
        print(e)

    try:
        time.sleep(2)
        WebDriver.refresh()
        time.sleep(3)
        element = WebDriverWait(WebDriver, 15).until(
            EC.presence_of_element_located((By.ID, "link-creditcard"))
        )

        element.click()
    except WebDriverException as e:
        print("no se ejecuto4")
        print(e)

    try:
        time.sleep(2)
        WebDriver.refresh()
        time.sleep(3)
        dos = WebDriverWait(WebDriver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//img[@src="/static/icons/document-pdf-text.png"]'))
        )
        dos.click()
    except WebDriverException as e:
        print("no se ejecuto5")
        print(user)

        minfun(user, password, url)

        print(e)
    try:
        time.sleep(3)
        targetPattern = r"C:\Test\*.pdf"

        datos = glob.glob(targetPattern)
        if len(datos) > 0:
            characters = "[]"
            string = ''.join(b for b in datos[0] if b not in characters)
            if os.path.isfile(string):
                nombre2 = user.split("@")

                nombre_nuevo = r"C:\Test\\" + nombre2[0] + ".pdf"

                os.rename(string, nombre_nuevo)
                if os.path.isfile(r"C:\Test\compro\\" + nombre2[0] + ".pdf"):
                    os.remove(r"C:\Test\compro\\" + nombre2[0] + ".pdf")
                    shutil.move(nombre_nuevo, r"C:\Test\compro")
                else:
                    shutil.move(nombre_nuevo, r"C:\Test\compro")

    except WebDriverException as e:
        print("no se ejecuto6")
        print(user)
        print(e)

    try:
        wait = WebDriverWait(WebDriver, 10)
        sign_in = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Cerrar sesión")))
        sign_in.click()

    except WebDriverException as e:
        print("no se ejecuto7")
        print(user)
        minfun(user, password, url)
        print(e)

    return


for x in myresult:
    minfun(x[0], x[1], urls[0])
    print(x)
    time.sleep(3)
    continue

today = date.today()
mes = today.month

mes = mes - 1

if mes == 1:
    mes = 'ENERO'
elif mes == 2:
    mes = 'FEBRERO'
elif mes == 3:
    mes = 'MARZO'
elif mes == 4:
    mes = 'ABRIL'
elif mes == 5:
    mes = 'MAYO'
elif mes == 6:
    mes = 'JUNIO'
elif mes == 7:
    mes = 'JULIO'
elif mes == 8:
    mes = 'AGOSTO'
elif mes == 9:
    mes = 'SEPTIEMBRE'
elif mes == 10:
    mes = 'OCTUBRE'
elif mes == 11:
    mes = 'NOVIEMBRE'
elif mes == 12:
    mes = 'DICIEMBRE'
else:
    print('error')

fantasy_zip = zipfile.ZipFile(r'C:\Test\COMPROBANTES_' + mes + '.zip', 'w')

for folder, subfolders, files in os.walk(r'C:\Test\compro'):

    for file in files:
        if file.endswith('.pdf'):
            fantasy_zip.write(os.path.join(folder, file),
                              os.path.relpath(os.path.join(folder, file), r'C:\Test\compro'),
                              compress_type=zipfile.ZIP_DEFLATED)

fantasy_zip.close()
if os.path.isfile(r'C:\Test\COMPROBANTES_' + mes + '.zip'):
    today = date.today()
    mes2 = today.month

    if mes2 == 1:
        mes2 = 'Enero'
    elif mes2 == 2:
        mes2 = 'Febrero'
    elif mes2 == 3:
        mes2 = 'Marzo'
    elif mes2 == 4:
        mes2 = 'Abril'
    elif mes2 == 5:
        mes2 = 'Mayo'
    elif mes2 == 6:
        mes2 = 'Junio'
    elif mes2 == 7:
        mes2 = 'Julio'
    elif mes2 == 8:
        mes2 = 'Agosto'
    elif mes2 == 9:
        mes2 = 'Septiembre'
    elif mes2 == 10:
        mes2 = 'Octubre'
    elif mes2 == 11:
        mes2 = 'Noviembre'
    elif mes2 == 12:
        mes2 = 'Diciembre'
    else:
        print('error')

    datos2 = format(today.day)

    msg = MIMEMultipart()
    email_content = """
    <html xmlns="http://www.w3.org/1999/xhtml" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
    <head>
    <!--[if gte mso 9]>
    <xml>
      <o:OfficeDocumentSettings>
        <o:AllowPNG/>
        <o:PixelsPerInch>96</o:PixelsPerInch>
      </o:OfficeDocumentSettings>
    </xml>
    <![endif]-->
      <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <meta name="x-apple-disable-message-reformatting">
      <!--[if !mso]><!--><meta http-equiv="X-UA-Compatible" content="IE=edge"><!--<![endif]-->
      <title></title>

        <style type="text/css">
          table, td { color: #000000; } @media only screen and (min-width: 520px) {
      .u-row {
        width: 500px !important;
      }
      .u-row .u-col {
        vertical-align: top;
      }

      .u-row .u-col-33p33 {
        width: 166.65px !important;
      }

      .u-row .u-col-100 {
        width: 500px !important;
      }

    }

    @media (max-width: 520px) {
      .u-row-container {
        max-width: 100% !important;
        padding-left: 0px !important;
        padding-right: 0px !important;
      }
      .u-row .u-col {
        min-width: 320px !important;
        max-width: 100% !important;
        display: block !important;
      }
      .u-row {
        width: calc(100% - 40px) !important;
      }
      .u-col {
        width: 100% !important;
      }
      .u-col > div {
        margin: 0 auto;
      }
    }
    body {
      margin: 0;
      padding: 0;
    }

    table,
    tr,
    td {
      vertical-align: top;
      border-collapse: collapse;
    }

    p {
      margin: 0;
    }

    .ie-container table,
    .mso-container table {
      table-layout: fixed;
    }

    * {
      line-height: inherit;
    }

    a[x-apple-data-detectors='true'] {
      color: inherit !important;
      text-decoration: none !important;
    }

    </style>



    <!--[if !mso]><!--><link href="https://fonts.googleapis.com/css?family=Open+Sans:400,700&display=swap" rel="stylesheet" type="text/css"><!--<![endif]-->

    </head>

    <body class="clean-body" style="margin: 0;padding: 0;-webkit-text-size-adjust: 100%;background-color: #000000;color: #000000">
      <!--[if IE]><div class="ie-container"><![endif]-->
      <!--[if mso]><div class="mso-container"><![endif]-->
      <table style="border-collapse: collapse;table-layout: fixed;border-spacing: 0;mso-table-lspace: 0pt;mso-table-rspace: 0pt;vertical-align: top;min-width: 320px;Margin: 0 auto;background-color: #000000;width:100%" cellpadding="0" cellspacing="0">
      <tbody>
      <tr style="vertical-align: top">
        <td style="word-break: break-word;border-collapse: collapse !important;vertical-align: top">
        <!--[if (mso)|(IE)]><table width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td align="center" style="background-color: #000000;"><![endif]-->


    <div class="u-row-container" style="padding: 0px 10px 20px;background-color: #000000">
      <div class="u-row" style="Margin: 0 auto;min-width: 320px;max-width: 500px;overflow-wrap: break-word;word-wrap: break-word;word-break: break-word;background-color: transparent;">
        <div style="border-collapse: collapse;display: table;width: 100%;background-color: transparent;">
          <!--[if (mso)|(IE)]><table width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td style="padding: 0px 10px 20px;background-color: #000000;" align="center"><table cellpadding="0" cellspacing="0" border="0" style="width:500px;"><tr style="background-color: transparent;"><![endif]-->

    <!--[if (mso)|(IE)]><td align="center" width="500" style="width: 500px;padding: 0px;border-top: 0px solid transparent;border-left: 0px solid transparent;border-right: 0px solid transparent;border-bottom: 0px solid transparent;" valign="top"><![endif]-->
    <div class="u-col u-col-100" style="max-width: 320px;min-width: 500px;display: table-cell;vertical-align: top;">
      <div style="width: 100% !important;">
      <!--[if (!mso)&(!IE)]><!--><div style="padding: 0px;border-top: 0px solid transparent;border-left: 0px solid transparent;border-right: 0px solid transparent;border-bottom: 0px solid transparent;"><!--<![endif]-->

    <table style="font-family:helvetica,sans-serif;" role="presentation" cellpadding="0" cellspacing="0" width="100%" border="0">
      <tbody>
        <tr>
          <td style="overflow-wrap:break-word;word-break:break-word;padding:10px 50px;font-family:helvetica,sans-serif;" align="left">

      <div style="color: #ffffff; line-height: 130%; text-align: left; word-wrap: break-word;">
        <p style="font-size: 14px; line-height: 130%;"><strong><span style="font-size: 48px; line-height: 62.4px;">Hola buen dia,</span></strong></p>
    <p style="font-size: 14px; line-height: 130%;"><strong><span style="font-size: 48px; line-height: 62.4px;">Alan Padilla</span></strong></p>
    <p style="font-size: 14px; line-height: 130%;">&nbsp;</p>
      </div>

          </td>
        </tr>
      </tbody>
    </table>

    <table style="font-family:helvetica,sans-serif;" role="presentation" cellpadding="0" cellspacing="0" width="100%" border="0">
      <tbody>
        <tr>
          <td style="overflow-wrap:break-word;word-break:break-word;padding:10px 50px 20px;font-family:helvetica,sans-serif;" align="left">

      <div style="color: #ffffff; line-height: 160%; text-align: left; word-wrap: break-word;">
        <p style="line-height: 160%; font-size: 14px;"><em style="font-size: 14px;"><span style="font-family: georgia, palatino; font-size: 18px; line-height: 28.8px;">Por medio del presente le brindamos un saludo cordial esperando se encuentre bien.</span></em><br /><em><span style="line-height: 22.4px; font-size: 14px;"><span style="font-family: georgia, palatino;"><span style="font-size: 18px; line-height: 28.8px;">Adjunto a este correo encontrara los comprobantes de pago</span></span></span></em><em><span style="line-height: 22.4px; font-size: 14px;"><span style="font-family: georgia, palatino;"><span style="font-size: 18px; line-height: 28.8px;">&nbsp;del mes. </span></span></span></em></p>
      </div>

          </td>
        </tr>
      </tbody>
    </table>

      <!--[if (!mso)&(!IE)]><!--></div><!--<![endif]-->
      </div>
    </div>
    <!--[if (mso)|(IE)]></td><![endif]-->
          <!--[if (mso)|(IE)]></tr></table></td></tr></table><![endif]-->
        </div>
      </div>
    </div>



    <div class="u-row-container" style="padding: 40px 50px 5px;background-color: rgba(255,255,255,0)">
      <div class="u-row" style="Margin: 0 auto;min-width: 320px;max-width: 500px;overflow-wrap: break-word;word-wrap: break-word;word-break: break-word;background-color: transparent;">
        <div style="border-collapse: collapse;display: table;width: 100%;background-color: transparent;">
          <!--[if (mso)|(IE)]><table width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td style="padding: 40px 50px 5px;background-color: rgba(255,255,255,0);" align="center"><table cellpadding="0" cellspacing="0" border="0" style="width:500px;"><tr style="background-color: transparent;"><![endif]-->

    <!--[if (mso)|(IE)]><td align="center" width="167" style="width: 167px;padding: 0px;border-top: 0px solid transparent;border-left: 0px solid transparent;border-right: 0px solid transparent;border-bottom: 0px solid transparent;" valign="top"><![endif]-->
    <div class="u-col u-col-33p33" style="max-width: 320px;min-width: 167px;display: table-cell;vertical-align: top;">
      <div style="width: 100% !important;">
      <!--[if (!mso)&(!IE)]><!--><div style="padding: 0px;border-top: 0px solid transparent;border-left: 0px solid transparent;border-right: 0px solid transparent;border-bottom: 0px solid transparent;"><!--<![endif]-->

    <table style="font-family:helvetica,sans-serif;" role="presentation" cellpadding="0" cellspacing="0" width="100%" border="0">
      <tbody>
        <tr>
          <td style="overflow-wrap:break-word;word-break:break-word;padding:20px 0px;font-family:helvetica,sans-serif;" align="left">

      <div style="color: #ffffff; line-height: 140%; text-align: center; word-wrap: break-word;">
        <p style="font-size: 14px; line-height: 140%; text-align: center;"><span style="font-size: 24px; line-height: 33.6px;">Fecha</span></p>
      </div>

          </td>
        </tr>
      </tbody>
    </table>

      <!--[if (!mso)&(!IE)]><!--></div><!--<![endif]-->
      </div>
    </div>
    <!--[if (mso)|(IE)]></td><![endif]-->
    <!--[if (mso)|(IE)]><td align="center" width="167" style="width: 167px;padding: 0px;border-top: 0px solid transparent;border-left: 0px solid transparent;border-right: 0px solid transparent;border-bottom: 0px solid transparent;" valign="top"><![endif]-->
    <div class="u-col u-col-33p33" style="max-width: 320px;min-width: 167px;display: table-cell;vertical-align: top;">
      <div style="width: 100% !important;">
      <!--[if (!mso)&(!IE)]><!--><div style="padding: 0px;border-top: 0px solid transparent;border-left: 0px solid transparent;border-right: 0px solid transparent;border-bottom: 0px solid transparent;"><!--<![endif]-->

    <table style="font-family:helvetica,sans-serif;" role="presentation" cellpadding="0" cellspacing="0" width="100%" border="0">
      <tbody>
        <tr>
          <td style="overflow-wrap:break-word;word-break:break-word;padding:10px 5px;font-family:helvetica,sans-serif;" align="left">

    <table width="100%" cellpadding="0" cellspacing="0" border="0">
      <tr>
        <td style="padding-right: 0px;padding-left: 0px;" align="center">

          <img align="center" border="0" src="https://tglobally.mx/images/mail_imagen.png" alt="Image" title="Image" style="outline: none;text-decoration: none;-ms-interpolation-mode: bicubic;clear: both;display: inline-block !important;border: none;height: auto;float: none;width: 100%;max-width: 50px;" width="50"/>

        </td>
      </tr>
    </table>

          </td>
        </tr>
      </tbody>
    </table>

      <!--[if (!mso)&(!IE)]><!--></div><!--<![endif]-->
      </div>
    </div>
    <!--[if (mso)|(IE)]></td><![endif]-->
    <!--[if (mso)|(IE)]><td align="center" width="167" style="width: 167px;padding: 0px;border-top: 0px solid transparent;border-left: 0px solid transparent;border-right: 0px solid transparent;border-bottom: 0px solid transparent;" valign="top"><![endif]-->
    <div class="u-col u-col-33p33" style="max-width: 320px;min-width: 167px;display: table-cell;vertical-align: top;">
      <div style="width: 100% !important;">
      <!--[if (!mso)&(!IE)]><!--><div style="padding: 0px;border-top: 0px solid transparent;border-left: 0px solid transparent;border-right: 0px solid transparent;border-bottom: 0px solid transparent;"><!--<![endif]-->

    <table style="font-family:helvetica,sans-serif;" role="presentation" cellpadding="0" cellspacing="0" width="100%" border="0">
      <tbody>
        <tr>
          <td style="overflow-wrap:break-word;word-break:break-word;padding:20px 0px;font-family:helvetica,sans-serif;" align="left">

      <div style="color: #ffffff; line-height: 140%; text-align: left; word-wrap: break-word;">
        <p style="font-size: 14px; line-height: 140%; text-align: center;"><span style="font-size: 24px; line-height: 33.6px;">""" + mes2 + " """ + datos2 + """th<br /></span></p>
      </div>

          </td>
        </tr>
      </tbody>
    </table>
    
      <!--[if (!mso)&(!IE)]><!--></div><!--<![endif]-->
      </div>
    </div>
    <!--[if (mso)|(IE)]></td><![endif]-->
          <!--[if (mso)|(IE)]></tr></table></td></tr></table><![endif]-->
        </div>
      </div>
    </div>


        <!--[if (mso)|(IE)]></td></tr></table><![endif]-->
        </td>
      </tr>
      </tbody>
      </table>
      <!--[if mso]></div><![endif]-->
      <!--[if IE]></div><![endif]-->
    </body>

    </html>"
    """

    message = "Thank you"

    to = "administracion@tglobally.com"

    bcc = "mvillasenor@tglobally.com,rbogarin@tglobally.com"

    rcpt = bcc.split(",") + [to]
    msg1 = email.message.Message()
    password = "[FW?k7_^;_Y="
    msg['From'] = "avisos@tglobally.com"
    msg['To'] = to
    msg['Bcc'] = bcc

    msg['Subject'] = "Subscription"
    msg1.add_header('Content-Type', 'text/html')
    msg1.set_payload(email_content)
    # add in the message body
    each_zip = r'C:\Test'
    os.chdir(each_zip)
    da = os.listdir(each_zip)

    part2 = MIMEText(email_content, 'html')

    f = open(da[1], 'rb')
    # Establezca el MIME y el nombre de archivo del adjunto, aquí está el tipo de zip:
    mime = MIMEBase('rar', 'rar', filename=da[1])
    # Agregue la información de encabezado necesaria:
    mime.add_header('Content-Disposition', 'attachment', filename=da[1])
    mime.add_header('Content-ID', '<0>')
    mime.add_header('X-Attachment-Id', '0')
    # Lea el contenido del adjunto:
    mime.set_payload(f.read())
    f.close()

    # Codificar con Base64
    encoders.encode_base64(mime)
    # Añadir a MIMEMultipart
    msg.attach(mime)
    msg.attach(part2)

    # create server
    server = smtplib.SMTP(host='r03.iservidorweb.com',
                          port=587)

    server.starttls()

    # Login Credentials for sending the mail
    server.login(msg['From'], password)

    # send the message via the server.
    server.sendmail(msg['From'], rcpt, msg.as_string())

    server.quit()
    os.remove(r'C:\Test\COMPROBANTES_' + mes + '.zip')
    py_files = glob.glob(r"C:\Test\compro\*.pdf")

    for py_file in py_files:
        try:
            os.remove(py_file)
        except OSError as a:
            print(f"Error:{a.strerror}")

    if __name__ == "__main__":
        toaster = ToastNotifier()

        toaster.show_toast(
            "Hello World!!!",
            "successfully sent email to %s:" % msg['To'],
            icon_path="python_icon.ico",
            duration=10
        )
    print("successfully sent email to %s:" % msg['To'])
