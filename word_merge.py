
import sys
import os
import comtypes.client
import logging
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
from docx import Document
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import ntpath
import time


g_logger = logging.getLogger(__name__)

class CSVReader:

	def __init__(self, csv_file_name):
		self.__file_name = csv_file_name
	def __call__(self):
		g_logger.debug("File Name: {}".format(self.__file_name))

		df = pd.read_csv(self.__file_name)
		names = df['name']
		
		email = df['email']
		if len(names) != len(email):
			g_logger.error(" file %s is invalid/corrupted", self.__file_name)
		_details = {}
		for i in range(len(names)):
			_details[names[i]] = email[i]
		g_logger.error("_details:\n{}".format(_details))
		return _details
#c = CSVReader('test1.csv')()
class CreateOfferLetter:
	def __init__(self, template, name, email):
		self.__template = template
		self.__name = name
		self.__email = email
		self.__new_offer_letter = ""

	def __call__(self):

		g_logger.debug("Name: %s", self.__name)
		g_logger.debug("Address: %s", self.__email)
		g_logger.debug("Offer Letter Template Name: {}".format(self.__template))

		_file = open(self.__template, 'rb')
		_doc = Document(_file)

		for paragraph in _doc.paragraphs:

			if '[Name]' in paragraph.text:
				inline = paragraph.runs
			# Loop added to work with runs (strings with same style)
				for i in range(len(inline)):
					if '[Name]' in inline[i].text:
						text = inline[i].text.replace('[Name]', self.__name)
						inline[i].text = text


			if '[Email]' in paragraph.text:
				inline = paragraph.runs
			# Loop added to work with runs (strings with same style)
				for i in range(len(inline)):
					if '[Email]' in inline[i].text:
						text = inline[i].text.replace('[Email]', self.__email)
						inline[i].text = text

		self.__name = self.__name.replace(" ", "_")
		self.__new_offer_letter = self.__name.replace(" ", "_") + "_Offer_Letter.docx"
		g_logger.debug("New offer letter name: %s", self.__new_offer_letter)

		# If this file already Present, then remove it
		if os.path.exists(self.__new_offer_letter):
			g_logger.info("New offer letter file {} already exist. Deleting older one.".format(self.__new_offer_letter))
			os.remove(self.__new_offer_letter)

		_doc.save(self.__new_offer_letter)
		_file.close()

		return self.__new_offer_letter



#offer_letter = CreateOfferLetter("template.docx", "sania", "saniathobani@yahoo.co.in")()

class Word2PdfConverter:
    def __init__(self, input_file):
        self.__input_file = os.path.abspath(ntpath.basename(input_file))
        self.__output_file = os.path.splitext(ntpath.basename(input_file))[0] + ".pdf"
        self.__output_file = os.path.abspath(self.__output_file)
        self.__word_format_pdf = 17


    def __call__(self):

        g_logger.debug("input file name: {}".format(self.__input_file))
        g_logger.debug("output file name: {}".format(self.__output_file))

        # Error the log and return false if input file not present
        if not os.path.exists(self.__input_file):
            g_logger.error("input doc file {} not exist".format(self.__output_file))
            return False

        # Error the log and return false if output file already present
        if os.path.exists(self.__output_file):
            g_logger.info("output pdf file {} already exist. Deleting older one.".format(self.__output_file))
            os.remove(self.__output_file)

        # create COM object
        word = comtypes.client.CreateObject('Word.Application')
        # key point 1: make word visible before open a new document
        word.Visible = True
        # key point 2: wait for the COM Server to prepare well.
        time.sleep(3)

        # convert docx file 1 to pdf file 1
        doc = word.Documents.Open(self.__input_file)
        doc.SaveAs(self.__output_file, FileFormat=self.__word_format_pdf)
        doc.Close()
        word.Quit()

        # Delete newly created doc file
        # print(os.path.basename(self.__output_file))
        if os.path.exists(self.__input_file):
            os.remove(self.__input_file)
		
        return os.path.basename(self.__output_file)
#offer_letter_pdf = Word2PdfConverter(offer_letter)()

class SendMail:
    def __init__(self, to_name, to_email, attach_file_name):
        self.__to_name = to_name
        self.__to_email = to_email
        self.__attach_file_name = attach_file_name

        # change these according to your requirements
        self.__from_email = os.environ.get('EMAIL_CSV')
        self.__from_email_password = os.environ.get('PASS_CSV')
        self.__subject = "Congrats {} - Offer Letter!!".format(self.__to_name)
        self.__body =''' "Hi {},\n\n"
                "
                ""'''.format(self.__to_name)

    def __call__(self):
        g_logger.debug("self.__to_name: {}".format(self.__to_name))
        g_logger.debug("self.__to_email: {}".format(self.__to_email))
        g_logger.debug("self.__attach_file_name: {}".format(self.__attach_file_name))
        g_logger.debug("self.__from_email: {}".format(self.__from_email))
        g_logger.debug("self.__subject: {}".format(self.__subject))
        g_logger.debug("self.__body: {}".format(self.__body))
        msg = MIMEMultipart()
        msg['From'] = self.__from_email
        msg['To'] = self.__to_email
        msg['Subject'] = self.__subject

        msg.attach(MIMEText(self.__body, 'plain'))

        attachment = open(self.__attach_file_name, 'rb')

        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= " + self.__attach_file_name)

        msg.attach(part)
        text = msg.as_string()
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(self.__from_email, self.__from_email_password)
        server.sendmail(self.__from_email, self.__to_email, text)
        server.quit()

        attachment.close()

        # Delete pdf file
        if os.path.exists(self.__attach_file_name):
            os.remove(self.__attach_file_name)
# SendMail('sania', 'saniathobani@yahoo.co.in', 'sania_Offer_Letter.pdf')()

def main():
    global g_logger
    logging.basicConfig(filename="Send_Offer_Letter.log", filemode='w', level=logging.DEBUG, )
    g_logger.info("Staring the script")
    details = CSVReader('test1.csv')()
    g_logger.debug("main() - details: {}".format(details))
    for name in details.keys():
        offer_letter = CreateOfferLetter("tempLate.docx", name, details[name])()
        offer_letter = Word2PdfConverter(offer_letter)()
        SendMail(name, details[name], offer_letter)()

if __name__ == '__main__':
    main()