# -*- coding: utf-8 -*-
"""
Created on Apr-11-2021

@author: nsriram
"""
import os
from docx import Document
import docx
from docx2pdf import convert
from flask import Flask,render_template,url_for,request
from datetime import date
from docx.shared import Cm, Pt
from docx.enum.style import WD_STYLE_TYPE

today = date.today()
sys_date = today.strftime("%d/%m/%Y")
#print("d1 =", d1)

UPLOAD_FOLDER = 'uploads'


app = Flask(__name__)
app.secret_key = "secret key"
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/')
def home():
	return render_template('index.html')


@app.route('/index.html')
def returnhome():
    
    return render_template('index.html')

@app.route('/formtodocx',methods=['POST'])
def formtodocx():
    if request.method == 'POST':
        
        client_name=request.form['clientname']
        print(client_name)
        community_name=request.form['commname']
        print(community_name)
        email=request.form['email']
        print(email)
        contact=request.form['contact']
        print(contact)
        address=request.form['address']
        print(address)
        has_b1=request.form['has_b1']
        print(has_b1)
        count_b1=request.form['count_b1']
        print(count_b1)
        has_b2=request.form['has_b2']
        count_b2=request.form['count_b2']
        has_b3=request.form['has_b3']
        count_b3=request.form['count_b3']
        watt_low=request.form['9w']
        watt_high=request.form['12w']
        
        client="Niroop"
        doc = docx.Document("D:\\Udak\\InstantSalesQuote\\downloads\\template.docx")
        
        quote_path="D:\\Udak\\InstantSalesQuote\\downloads\\"+'client'".docx"
        p=doc.add_paragraph("\n\n\n\n\n\n\n")
        p.add_run('To,\n')
        p.add_run('Botanika\n').bold = True
        #p.add_run("Botanika").bold=True
        p=doc.add_paragraph("#242 Raja Rajeshwari Nagar Colony\nKondapur, Hyderabad\nTelangana 500 084\nPhone: +91 97010 49715\nE-mail: business@udaktech.com,")
        p=doc.add_paragraph("\n\n")
        p=doc.add_paragraph("Dear "+client+" Garu,")
        p=doc.add_paragraph("\n")
        p=doc.add_paragraph("The below quote is based on the following details we obtained from you on "+sys_date+".The quote would be adjusted as and when new details/updates are observed during the course of discussion/consultation.")
        p=doc.add_paragraph("\n")
        p=doc.add_paragraph("\t""1. B1, B2 basements are common for everyone and they have 190 plus LED tube lights, which are turned on 24X7.\n"
                            "\t2. Based on our field study we are going install ELSA sensors to around 70% of lights.\n"
                            "\t3. We will leave 30% of tube lights to maintain…\n"
                            "\t\t\ta. Ambient lighting\n"
                            "\t\t\tb. Driveways\n"
                            "\t4. The number of sensors need to be installed may vary based on the ground truth. The invoice would be adjusted, to reflect the final number of sensors installed.\n"
                            "\t5. As discussed and suggested, installation would be done by UDAK Team.\n"
                            "\t6. Payment of One Time Model would be in 2 installments (50% on delivery and 50% on completion of installation).\n"
                            "\t7. Payment of Subscription Model would be post-paid and will start from the date of delivery.\n")
        
        
        doc.add_page_break()
        p=doc.add_paragraph("\n\n\n\n")
        p.add_run('Quote\n').bold = True
        p=doc.add_paragraph("The following equipment and charges would be part of the final quotation submitted.\n")
        ################ONE TIME PAYMENT########################
        p=doc.add_paragraph("OPTION 1: One Time Payment\n")
        table1 = doc.add_table(rows = 1, cols = 4, style='Table Grid')
        OTP_cells = table1.rows[0].cells
        OTP_cells[0].text = 'Item'
        OTP_cells[1].text = 'Description'
        OTP_cells[2].text = 'Quantity'
        OTP_cells[3].text = 'Price (Rs.)'
        cells = table1.add_row().cells
        cells[0].text = "1."
        cells[1].text = "ELSA sensor with inbuilt activity and lux detection, along with adjustable knobs for specific adjustments. (Rs. 799/Sensor)*"
        cells[2].text = "133"
        cells[3].text = "1,06,267"
        cells = table1.add_row().cells
        cells[0].text = "2."
        cells[1].text = "One Year Full Product Replacement Guarantee**"
        cells[2].text = " "
        cells[3].text = "Free"

        ##########################################################
        
        
       

        ################Subscription Model########################
        
        p=doc.add_paragraph("\nOPTION 2: Subscription Model\n")
        table2 = doc.add_table(rows = 1, cols = 4, style='Table Grid')
        SM_cells = table2.rows[0].cells
        SM_cells[0].text = 'Item'
        SM_cells[1].text = 'Description'
        SM_cells[2].text = 'Quantity'
        SM_cells[3].text = 'Price (Rs.)'
        cells = table2.add_row().cells
        cells[0].text = "1."
        cells[1].text = "Solution is subscription based, wherein company gives the product free and users pay nominal monthly fee of Rs 75/Sensor*"
        cells[2].text = "133"
        cells[3].text = "9,975/month"
        cells = table2.add_row().cells
        cells[0].text = "2."
        cells[1].text = "Full Product Maintenance & Replacement Guarantee**"
        cells[2].text = " "
        cells[3].text = "Lifetime Free***"
        ###########################################################
        p=doc.add_paragraph("\n\n* 18% GST is exclusive of the above details and would be part of the final invoice.\n"
                             "** Any and every defective sensor would be fully replaced, with no additional charges to the customer.\n"
                             "*** Product will be owned & maintained by UDAK under active subscription\n"
                             "\nIf you have any questions, please do not hesitate to contact us.\n"
                             "\n\nKind regards,\n"
                             "UDAK Technologies Pvt Ltd.")
        
        
        
        
        
        
        
        
        
        
        saved_doc=doc.save(quote_path)
        docxtopdf(quote_path)
        
       



def generate_smatrix(docx_file):
    convert(docx_file)


def docxtopdf(docx_file):
    convert(docx_file)
    
    #return render_template('result.html')
if __name__ == '__main__':
	app.run()
