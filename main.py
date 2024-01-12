import qrcode
from qrcode.image.styledpil import StyledPilImage
from email.message import EmailMessage
from email.utils import make_msgid
import mimetypes
import smtplib
from qrcode.image.styles.colormasks import SolidFillColorMask
from PIL import Image, ImageDraw
import openpyxl
import math
import requests
from pyppeteer import launch
from xhtml2pdf import pisa             # import python module
from dotenv import load_dotenv
import os

load_dotenv()
wb = openpyxl.load_workbook('GuestListTest.xlsx')
sheet = wb.active
api_url = "https://ruskokaaccess.azurewebsites.net/api/guests"
headers = {"Authorization":os.environ.get("api-key")}

#ToDo
    #Save the excel sheet
    #Add the guest to the database
    #Improve Email and PDF
    #Remove generic Table number

async def generate_pdf_from_html(html_content, pdf_path):
    browser = await launch()
    page = await browser.newPage()

    await page.setContent(html_content)

    await page.pdf({'path': pdf_path, 'format': 'A4'})

    await browser.close()

def convert_html_to_pdf(source_html, output_filename):
    # open output file for writing (truncated binary)
    result_file = open(output_filename, "w+b")

    # convert HTML to PDF
    pisa_status = pisa.CreatePDF(
            source_html,                # the HTML to convert
            dest=result_file)           # file handle to recieve result

    # close output file
    result_file.close()                 # close output file

    # return False on success and True on errors
    return pisa_status.err

def style_eyes(img):
    img_size = img.size[0]
    eye_size = 70  # default
    quiet_zone = 40  # default
    mask = Image.new('L', img.size, 0)
    draw = ImageDraw.Draw(mask)
    draw.rectangle((10, 10, 80, 80), fill=255)
    draw.rectangle((img_size - 80, 10, img_size - 10, 80), fill=255)
    draw.rectangle((10, img_size - 80, 80, img_size - 10), fill=255)
    return mask

for i in range(sheet.max_row-2):
    emailGroup = [[0,0,0,0,0]]
    guests = 1
    names = [0]
    namesUse = [0]
    first = sheet.cell(row=3+i, column=1)
    last = sheet.cell(row=3+i, column=2)
    emails = sheet.cell(row=3+i, column=4)
    status = sheet.cell(row=3+i, column=5)

    names[0] = first.value + last.value
    namesUse[0] = first.value + " " + last.value

    emailGroup[0][0] = first.value
    emailGroup[0][1] = last.value
    emailGroup[0][2] = emails.value
    emailGroup[0][3] = status.value

    if emailGroup[0][3] == False:
        sheet.cell(row=3+i, column=5).value = True
        for l in range ((sheet.max_row-3)-i):
            firstCheck = sheet.cell(row=4+i+l, column=1)
            lastCheck = sheet.cell(row=4+i+l, column=2)
            emailsCheck = sheet.cell(row=4+i+l, column=4)
            statusCheck = sheet.cell(row=4+i+l, column=5)
            if emailsCheck.value == emailGroup[0][2] and statusCheck.value == False:
                emailGroup.append([0,0,0,0,0])
                sheet.cell(row=4 + i + l, column=5).value = True
                names.append(firstCheck.value + lastCheck.value)
                namesUse.append(firstCheck.value + " " + lastCheck.value)
                emailGroup[l+1][0] = firstCheck.value
                emailGroup[l+1][1] = lastCheck.value
                emailGroup[l+1][2] = emailsCheck.value
                emailGroup[l+1][3] = statusCheck.value


        print(first.value, last.value, emails.value, emailGroup[0][3])
        print(emailGroup)

        for index, name in enumerate(names):
            # Generates the ticket code
            a = emailGroup[index][0][0]
            b = emailGroup[index][1][0]
            c = str((math.ceil((len(emailGroup[index][1] + emailGroup[index][0]))/3.141592)*2))
            d = emailGroup[index][0][-1]
            e = emailGroup[index][1][-1]
            code = ("WPB" + a+b+c+d+e + "23")
            emailGroup[index][4] = code

            #Generate a qr code based off the ticket code
            qr = qrcode.QRCode(version=5, error_correction=qrcode.constants.ERROR_CORRECT_H, border=1)
            qr.add_data(code)
            qr_eyes_img = qr.make_image(image_factory=StyledPilImage,
                            color_mask=SolidFillColorMask(back_color=(255, 255, 255), front_color=(158, 42, 43)))
            qr_img = qr.make_image(image_factory=StyledPilImage,
                       color_mask=SolidFillColorMask(front_color=(84, 11, 14)))


            mask = style_eyes(qr_img)
            final_img = Image.composite(qr_eyes_img, qr_img, mask)
            final_img.save(emailGroup[index][0] + emailGroup[index][1] + ".png")

        #Email Sending

        msg = EmailMessage()
        # generic email headers
        msg['Subject'] = 'Winter Palace Ball Tickets'
        msg['From'] = '<deema@ruskoka.com>'  # Change this
        msg['To'] = emails.value

        # set the plain text body
        msg.set_content('This is a plain text body.')

        # now create a Content-ID for the image
        image_cid = [make_msgid(domain="ruskoka.com")[1:-1],
                     make_msgid(domain="ruskoka.com")[1:-1],
                     make_msgid(domain="ruskoka.com")[1:-1],
                     make_msgid(domain="ruskoka.com")[1:-1],
                     make_msgid(domain="ruskoka.com")[1:-1],
                     make_msgid(domain="ruskoka.com")[1:-1],
                     make_msgid(domain="ruskoka.com")[1:-1],
                     make_msgid(domain="ruskoka.com")[1:-1],
                     make_msgid(domain="ruskoka.com")[1:-1],
                     make_msgid(domain="ruskoka.com")[1:-1]]
        # if `domain` argument isn't provided, it will
        # use your computer's name

        trs = []
        for index,name in enumerate(namesUse):
            trs.append(f'''\
                  <tr>
                    <td align="center">{name}</td>
                  </tr>
                  <tr>
                    <td align="center"><img src="cid:{image_cid[index+1]}" width="50%" align="center"></td>
                  </tr>''')
        name_table = '\n'.join(trs)

        # set an alternative html body
        msg.add_alternative(f"""\
        <html>
            <body>
                <div align="center">
                         <img src="cid:{image_cid[0]}" style="width:50%">
                         <p>Thank you for purchasing a ticket!<br></p>
                         <table width="50%" border="0" cellspacing="0" cellpadding="0">
                             {name_table}
                         </table>
                        </div>
            </body>
        </html>
        """.format(image_cid=image_cid,name_table=name_table), subtype='html')
        # image_cid looks like <long.random.number@xyz.com>
        # to use it as the img src, we don't need `<` or `>`
        # so we use [1:-1] to strip them off

        with open("WPB-removebg.png", 'rb') as img:
            # know the Content-Type of the image
            maintype, subtype = mimetypes.guess_type(img.name)[0].split('/')

            # attach it
            msg.get_payload()[1].add_related(img.read(),
                                             maintype=maintype,
                                             subtype=subtype,
                                             cid=f"<{image_cid[0]}>")

        # now open the image and attach it to the email
        for index,item in enumerate(names):
            with open(item + ".png", 'rb') as img:
                # know the Content-Type of the image
                maintype, subtype = mimetypes.guess_type(img.name)[0].split('/')

                # attach it
                msg.get_payload()[1].add_related(img.read(),
                                             maintype=maintype,
                                             subtype=subtype,
                                             cid=f"<{image_cid[index+1]}>")

        #Creating a pdf
        trs = []
        for index, name in enumerate(namesUse):
            trs.append(f'''\
                          <tr>
                            <td align="center" size="bigger">{name}</td>
                          </tr>
                          <tr>
                            <td align="center"><img src="{names[index]}.png" style="zoom:60%" align="middle"></td>
                          </tr>
                          <tr>
                          <td><br></td>
                          </tr>''')
        name_pdf_table = '\n'.join(trs)

        html = f"""\
        <html>
            <body style="font-size:20px">
                <div align="center">
                         <img src="WPB-removebg.png">
                         <p>Thank you for purchasing a ticket!<br></p>
                         <table border="0" cellspacing="0" cellpadding="0">
                             {name_pdf_table}
                         </table>
                </div>
            </body>
        </html>
        """.format(image_cid=image_cid,name_pdf_table=name_pdf_table)
        print(html)

        #asyncio.get_event_loop().run_until_complete(generate_pdf_from_html(html, names[0] + '.pdf'))
        convert_html_to_pdf(html, names[0] + '.pdf')

        # Send the email (this example assumes SMTP authentication is required)

        with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
            smtp.ehlo()  # send the extended hello to our server
            smtp.starttls()  # tell server we want to communicate with TLS encryption
            smtp.login("deema@ruskoka.com", os.environ.get("apppass"))
            #smtp.sendmail("deema@ruskoka.com", emails.value, msg.as_string())

        print("Message sent!")

        for index, name in enumerate(names):
            guest = {"name": emailGroup[index][0] + " " + emailGroup[index][1],"tables": 1 ,"ticket": emailGroup[index][4]}
            response = requests.post(api_url, json=guest, headers=headers)
            print(response)

        #smtp.quit()  # finally, don't forget to close the connection