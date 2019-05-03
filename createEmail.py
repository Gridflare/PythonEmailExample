"""
Build an HTML type email containing multiple images and an attachment
Then save the email to a file and attempt to load it with an email
client for sending.
The script this example is derived from was tested on Win 7 with Outlook
Thunderbird also mostly works if the X-Unsent support addon is installed

Images from Pixabay.com
"""

# For creating the email body
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# For assigning recipients
from email.headerregistry import Address

# For  images and arbitrary attachments
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication

# For saving the email to be loaded in a client for sending
import os
import email.generator

subject = 'Cute Cats'
imageNames = ['treecat.jpg','rockcat.jpg','flowercat.jpg','catstretch.jpg']
attachmentFileName = 'azipfile.zip'

# Make email container, add recipients
message = MIMEMultipart('related')
message['Subject'] = subject
# message['From'] =  # Email client should fill this for you
message['To'] = 'thecatdude@example.com'
# You can add a CC field if you want, but not a BCC field.
# message['CC'] =

htmlContent = """
<strong>I want to share these cats with you!</strong>
<br>
<table>
    <tr>
        <td><img src="cid:{}"></td>
        <td><img src="cid:{}"></td>
    </tr>
    <tr>
        <td><img src="cid:{}"></td>
        <td><img src="cid:{}"></td>
    </tr>
</table>

""".format(*imageNames)

textContent = """
I want to share these cats with you!
"""

# Record the MIME types of both content types - text/plain and text/html
# Attach content into message container.
# Put HTML first for Thunderbird compatibility
message.attach(MIMEText(htmlContent, 'html'))
message.attach(MIMEText(textContent, 'plain'))

# Add the images to the email
for imgName in imageNames:
    with open(imgName, 'rb') as imgFile:
        imgData = MIMEImage(imgFile.read())
        imgData.add_header('Content-ID', f'<{imgName}>')
        message.attach(imgData)

# Attach a file to the email
# Thunderbird doesn't seem to like this in combo with X-Unsent = 1
with open(attachmentFileName, 'rb') as attFile:
    attData = MIMEApplication(attFile.read())
    attData.add_header('Content-Disposition', f'attachment; filename="{attachmentFileName}"')
    message.attach(attData)

# Mark message as unsent so that it can be sent with a mail client
message.add_header('X-Unsent', '1')

emailfilename = subject.replace(' ','_') + '.eml'

# Save the email in eml format
# This method does not support multiple recipients
with open(emailfilename, 'w') as f:
    gen = email.generator.Generator(f)
    gen.flatten(message)
    
# print(message) # In case you want to take a peak a the raw MIME

# Open the email with the local mail client
# Try both, hopefully the wrong one fails reasonably quietly
os.popen(emailfilename) # Works well on Windows with Outlook
os.popen('thunderbird  ' + emailfilename) #Assume Thunderbird
