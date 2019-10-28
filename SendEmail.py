
# Send an HTML email with an embedded image and a plain text message for
# email clients that don't want to display the HTML.

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from bs4 import BeautifulSoup, SoupStrainer

# Define these once; use them twice!
strFrom = 'lexisnexispsl@gmail.com'
recipients = ['daniel.hutchings.1@lexisnexis.co.uk', 'stephen.leslie@lexisnexis.co.uk', 'danielmhutchings@gmail.com', 'emma.millington@lexisnexis.co.uk', 'lisa.moore@lexisnexis.co.uk', 'claire.hayes@lexisnexis.co.uk', 'Ruth.Newman@lexisnexis.co.uk']
emaildirectory = 'C:\\Users\\Hutchida\\Documents\\PSL\\AICER\\'
#directory = '\\\\atlas\\Knowhow\\ContentHub\\'

filename = emaildirectory + 'newcontentreport_email.html'
soup = BeautifulSoup(open(filename),'lxml') 
try: subject = soup.find('title').text
except: subject = 'Not present'

# Create the root message and fill in the from, to, and subject headers
msgRoot = MIMEMultipart('related')
msgRoot['Subject'] = subject
msgRoot['From'] = strFrom
msgRoot['To'] = ", ".join(recipients)
msgRoot.preamble = 'This is a multi-part message in MIME format.'

# Encapsulate the plain and HTML versions of the message body in an
# 'alternative' part, so message agents can decide which they want to display.
msgAlternative = MIMEMultipart('alternative')
msgRoot.attach(msgAlternative)

msgText = MIMEText('This is the alternative plain text message.')
msgAlternative.attach(msgText)

#send html from a file
f = open(filename)
msgText = MIMEText(f.read(),'html')
# We reference the image in the IMG SRC attribute by the ID we give it below
#msgText = MIMEText('<b>Some <i>HTML</i> text</b> and an image.<br><img src="cid:image1"><br><img src="cid:image2"><br>Nifty!', 'html')
msgAlternative.attach(msgText)

# This example assumes the image is in the current directory
fp = open('C:\\Users\\Hutchida\\Documents\\PSL\\AICER\\newcontentpie.png', 'rb')
msgImage = MIMEImage(fp.read())
fp.close()

# Define the image's ID as referenced above
msgImage.add_header('Content-ID', '<image1>')
msgRoot.attach(msgImage)

# This example assumes the image is in the current directory
fp = open('C:\\Users\\Hutchida\\Documents\\PSL\\AICER\\newcontentbar.png', 'rb')
msgImage = MIMEImage(fp.read())
fp.close()

# Define the image's ID as referenced above
msgImage.add_header('Content-ID', '<image2>')
msgRoot.attach(msgImage)

# Send the email (this example assumes SMTP authentication is required)
import smtplib
smtp = smtplib.SMTP()
smtp.connect('smtp.gmail.com', 587)
smtp.ehlo()
smtp.starttls()
smtp.login(strFrom, 'PlantP0t1')
smtp.sendmail(strFrom, recipients, msgRoot.as_string())
smtp.quit()