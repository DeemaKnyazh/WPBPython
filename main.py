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
    #Improve Email and PDF
    #Remove generic Table number when adding to DB

#ToDo
    #When final Run
        #Uncomment the smtp emailer

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
                       color_mask=SolidFillColorMask(back_color=(207, 234, 250), front_color=(84, 11, 14)))


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
        msg.set_content('Thank you for purchasing a ticket and supporting Ruskoka Camp!'+
                        'If you are seeing this that means there was an error in loading the email'+
                        'Please check that the PDFs containing the tickets are attached, if not please reply all to this email')

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
                     make_msgid(domain="ruskoka.com")[1:-1],
                     make_msgid(domain="ruskoka.com")[1:-1],
                     make_msgid(domain="ruskoka.com")[1:-1]]
        # if `domain` argument isn't provided, it will
        # use your computer's name

        # trs = []
        # for index,name in enumerate(namesUse):
        #     trs.append(f'''\
        #           <tr>
        #             <td align="center"><p style="font-size:20px;text-decoration:underline;margin-bottom:1px;">{name}</p></td>
        #           </tr>
        #           <tr>
        #             <td align="center"><img src="cid:{image_cid[index+1]}" width="70%" align="center"></td>
        #           </tr>''')
        # name_table = '\n'.join(trs)

        # set an alternative html body
        msg.add_alternative(f"""\
        <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
        <html>
        <meta http-equiv="Content-Type" content="text/html charset=UTF-8" />
        <meta name="supported-color-schemes" content="light">
            <body>
                <div align="center" style="background-color: #cfeafa">
                         <img src="cid:{image_cid[0]}" style="width:95%">
                         <p style="font-size:30px;text-decoration:underline;">Thank you for purchasing a ticket and supporting Ruskoka Camp!<br></p>
                         <img src="cid:{image_cid[1]}" style="width:60%">
                         <table width="70%" border="0" cellspacing="0" cellpadding="0">
                         <span style="font-family:arial, helvetica neue, helvetica, sans-serif;text-align: center;width:90%;">
                                <p style="font-size:20px;margin-bottom: 20px;">Your generosity and kindness are greatly appreciated by us and the children who will benefit from your support.<br></p>
                                <br>
                                <p style="font-size:20px;margin-bottom: 20px;">Our camp has a one of a kind program that provides a fun and safe enviroment for children from all walks of life, giving them a chance to enjoy exciting outdoor activities, learn new skills, make new friends, and create lasting memories. Ruskoka also helps them develop their self-esteem, confidence, and resilience. </p>
                                <br>
                                <p style="font-size:20px;margin-bottom: 20px;">We are immensely grateful for you support and hope that you enjoy a magical evening at the Winter Palace Ball, we look forward to seeing you there!</p>
                                <br>
                                <p style="font-size:20px;margin-bottom: 20px;">Please do not forget to bring the attached tickets to the ball with you, as they will be required upon entry! If guests you purchased a ticket for will be arriving at differing times, please give each of them a copy of the Ticket PDF.</p>
                        </span>
                         </table>
                         <img src="cid:{image_cid[1]}" style="width:75%">
                         <p><br></p>
                         <p style="font-family: "Helvetica", sans-serif;">If there are any issues seeing the images that means there was an error in loading the email<br></p>
                         <p style="font-family: "Helvetica", sans-serif;">Please check that the PDFs containing the tickets are attached, if not please reply all to this email<br></p>
                         <p style="font-family: "Helvetica", sans-serif;">The content of this email is confidential and intended for the recipient specified in message only. If you received this message by mistake, please reply to this message and follow with its deletion, so that we can ensure such a mistake does not occur in the future.<br></p>
                        </div>
            </body>
        </html>
        """.format(image_cid=image_cid), subtype='html')
        # image_cid looks like <long.random.number@xyz.com>
        # to use it as the img src, we don't need `<` or `>`
        # so we use [1:-1] to strip them off

        with open("WPB-NewBG.png", 'rb') as img:
            # know the Content-Type of the image
            maintype, subtype = mimetypes.guess_type(img.name)[0].split('/')
            # attach it
            msg.get_payload()[1].add_related(img.read(),
                                             maintype=maintype,
                                             subtype=subtype,
                                             cid=f"<{image_cid[0]}>",
                                             filename='Ball Logo')
        with open("Divider.png", 'rb') as img:
            # know the Content-Type of the image
            maintype, subtype = mimetypes.guess_type(img.name)[0].split('/')
            # attach it
            msg.get_payload()[1].add_related(img.read(),
                                             maintype=maintype,
                                             subtype=subtype,
                                             cid=f"<{image_cid[1]}>",
                                             filename='Ball Logo')


        #Creating a pdf
        trs = []
        for index, name in enumerate(namesUse):
            trs.append(f'''\
                          <tr>
                            <td align="center" size="bigger">{name}</td>
                          </tr>
                          <tr>
                            <td align="center"><img src="{names[index]}.png" style="zoom:90%" align="middle"></td>
                          </tr>
                          <tr>
                            <td><br></td>
                          </tr>
                          <tr>
                            <td><br></td>
                          </tr>
                          <tr>
                            <td style="margin: 20px"><br></td>
                          </tr>
                          ''')
        name_pdf_table = '\n'.join(trs)

        html = f"""\
        <html>
            <body style="font-size:20px">
                <div align="center">
                         <img src="WPB-NewBG.png">
                         <p>Thank you for purchasing a ticket and supporting Ruskoka Camp!<br></p>
                         <table border="0" cellspacing="0" cellpadding="0">
                         <tr>
                            <td><br></td>
                          </tr>
                          <tr>
                            <td><br></td>
                          </tr>
                             {name_pdf_table}
                         </table>
                </div>
            </body>
        </html>
        """.format(image_cid=image_cid,name_pdf_table=name_pdf_table)
        print(html)

        convert_html_to_pdf(html, names[0] + '.pdf')

        # now open the pdf and attach it
        namess = [0]
        namess[0] = emailGroup[0][0]+emailGroup[0][1]
        print(namess)
        for index, name in enumerate(namess):
            with open(name + ".pdf", 'rb') as pdf:
                pdf_data = pdf.read()
            msg.add_attachment(pdf_data, maintype='application', subtype='pdf', filename='Tickets.pdf')

        # Send the email (this example assumes SMTP authentication is required)
        with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
            smtp.ehlo()  # send the extended hello to our server
            smtp.starttls()  # tell server we want to communicate with TLS encryption
            smtp.login("deema@ruskoka.com", os.environ.get("apppass"))
            smtp.sendmail("deema@ruskoka.com", emails.value, msg.as_string())

        print("Message sent!")

        for index, name in enumerate(names):
            guest = {"name": emailGroup[index][0] + " " + emailGroup[index][1],"tables": 1 ,"ticket": emailGroup[index][4]}
            response = requests.post(api_url, json=guest, headers=headers)
            print(response)

#Saves the workbook after each success is set to true
#wb.save('GuestListTest.xlsx')

        #smtp.quit()  # finally, don't forget to close the connection