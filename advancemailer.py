import os
import pandas as pd
from random import choices, randint
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import smtplib
import logging
import time
import sys
import string
import random
import binascii
from datetime import datetime
import imgkit
import uuid

path_wkhtmltoimage = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltoimage.exe'
imgkit.config(wkhtmltoimage=path_wkhtmltoimage)

logging.basicConfig(filename='mail.log', level=logging.DEBUG)

totalSend = 1
if len(sys.argv) > 1:
    totalSend = int(sys.argv[1])

emaildf = pd.read_csv('gmail.csv')
contactsData = pd.read_csv('test.txt')
subjects = pd.read_csv('subject.csv')
bodies = ['body.txt']

spam_subjects = []
spam_sender_names = []
additional_subject_words = []  # Define additional_subject_words if not defined
From = ["ORDER-Confirmation"]

include_emojis = True  # Global variable for emojis in subject and sender name
throttling_speed = 1.5  # Default throttling speed

def get_smtp_servers():
    # Extract unique SMTP domains from the gmail.csv file
    unique_domains = set(emaildf['email'].str.split('@').str[-1].str.lower())
    
    smtp_servers = {
        'gmail.com': ('smtp.gmail.com', 587),
        'yahoo.com': ('smtp.mail.yahoo.com', 587),
        'outlook.com': ('smtp.office365.com', 587),
        'icloud.com': ('smtp.mail.me.com', 587),
        # Add more domain mappings as needed
    }
    
    # Add dynamically detected SMTP servers
    for domain in unique_domains:
        if domain not in smtp_servers:
            smtp_servers[domain] = ('smtp.' + domain, 587)

    return smtp_servers

def Random_Name():
    Names_List = [" "]
    FIRST_NAME = random.choice(Names_List)
    return FIRST_NAME

def Random_Ticket_ID():
    UpperLetters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    OnlyNumbers = '0123456789'
    
    RAND_TICKET_ID = (
        "V" + random.choice(OnlyNumbers) + random.choice(OnlyNumbers) + "/" +
        random.choice(UpperLetters) + random.choice(UpperLetters) + random.choice(OnlyNumbers) + "/" +
        random.choice(OnlyNumbers) + random.choice(OnlyNumbers) + random.choice(OnlyNumbers) +
        random.choice(OnlyNumbers) + random.choice(OnlyNumbers) + random.choice(OnlyNumbers)
    )
    return RAND_TICKET_ID
    
def convert_html_to_png(html_content, output_path):
    # Use imgkit to convert HTML to PNG and save it to the specified output path
    options = {
        'format': 'png',
        'quiet': '',
        'encoding': 'utf-8',
        'quality': 50,  # Adjust the quality (compression level)
    }
    imgkit.from_string(html_content, output_path, options=options)

def generate_random_subject(use_emojis=None):
    adjectives = ["TX ID","TX Approved ID","TX Generated  ID"]
    verbs = [""]
    emojis = ["💰", "💵", "💴", "💶", "💷", "💸", "🤑", "🏦", "🌐", "📈", "💳", "📉", "💹", "💲", "💱", "🚀", "🌏", "🏧", "🏨",
              "🏦", "🏪", "🏫", "🏛"]

    if use_emojis is None:
        use_emojis = include_emojis

    if use_emojis:
        subject = f"{random.choice(adjectives)} {random.choice(verbs)} {random.choice(emojis)}"
    else:
        subject = f"{random.choice(adjectives)} {random.choice(verbs)}"

    if additional_subject_words:
        subject += f" {random.choice(additional_subject_words)}"

    spam_trigger_words = ["sale", "discount", "Renewal", "win", "free", "cash", "urgent", "limited time", "transaction",
                          "order", "click here"]
    for trigger_word in spam_trigger_words:
        if trigger_word in subject.lower():
            return generate_random_subject(use_emojis)

    if subject.lower() in spam_subjects:
        return generate_random_subject(use_emojis)

    return subject

def generate_random_sender_name(use_emojis=None):
    emojis = ["💰", "💵", "💴", "💶", "💷", "💸", "🤑", "🏦", "🌐", "💳", "📈", "💹", "💲", "💱"]
    adjectives = ["Service", "Update"]
    nouns =  ["Service", "Update"]


    if use_emojis is None:
        use_emojis = include_emojis

    if use_emojis:
        sender_name = f"{random.choice(emojis)} {random.choice(adjectives)} {random.choice(nouns)}"
    else:
        sender_name = f"{random.choice(adjectives)} {random.choice(nouns)}"

    for spam_word in spam_sender_names:
        if spam_word.lower() in sender_name.lower():
            return generate_random_sender_name(use_emojis)

    return sender_name

def Random_MD5():
    RAND_MD5 = binascii.hexlify(os.urandom(16))
    return RAND_MD5

RAND_MD5 = Random_MD5()

def Current_Time():
    now = datetime.now()
    formatted_time = now.strftime("%B %d, %Y %H:%M:%S")
    return formatted_time

CURR_TIME = Current_Time()

def delete_smtp_entry(emailId):
    global emaildf
    emaildf = emaildf[emaildf['email'] != emailId]
    emaildf.to_csv('gmail.csv', index=False)
    print(f"SMTP entry for {emailId} deleted from gmail.csv")

def get_extra_subjects(csv_file):
    with open(csv_file, 'r', encoding='utf-8') as file:
        reader = csv.DictReader(file)
        return [row['ExtraSubject'] for row in reader]
def generate_random_uuid():
    # Generate a random UUID and convert it to a string
    return str(uuid.uuid4()).upper()

def send_mail(email, emailId, password, bodyFile, subjectWord, fromName, smtp_server, smtp_port):
    global include_emojis
    time.sleep(throttling_speed)

    invoiceNo = randint(10000000, 99999999)
    randomString = ''.join(choices(string.ascii_lowercase, k=4))

    # Updated code to use the new random methods
    subject = generate_random_subject() + " # " + Random_Ticket_ID()  # Add ticket ID to the subject
    sender_name = Random_Name() + " " + fromName  # Add random name to the "From" name

    num = randint(1111, 9999)

    transaction_id = randint(10000000, 99999999)
    random_body_lines = [
        "Your Invoice Status: Immediate Confirmation Required - Your Cooperation is Valued. H:1(818) 805-2112",
        # Add more lines as needed
    ]

    # Directly include HTML content within the code
    body = """
          <html>
            <body>
                <p>Hello,
<br><br>
Please find attached the Payment receipt for your recent purchase.
<br>If you have any questions or concerns, you can contact  with us at $phone_number
<br><br>
<br>Invoice Number :$random_product_id
<br>Issue date : $current_date
<br><br>
Kindky review the attached invoice for detailed breakdown of the charged.
</p> 
            
            </body>
        </html>
    """
    

    # Append random MD5 and current time to the body
    rand_md5 = Random_MD5().decode('utf-8')
    curr_time = Current_Time()

    tempEmail = ''
    for i in email:
        if i == '@':
            break
        tempEmail += i

        
        # Replace placeholders in the body with fixed values
        body = body.replace('$customer', tempEmail)
        body = body.replace('$email', email)
        body = body.replace('$invoice_no', str(transaction_id))
        body = body.replace('$random_line', random.choice(random_body_lines))
        body = body.replace('$random_uuid', generate_random_uuid())
        body = body.replace('$current_time', curr_time)
        body = body.replace('$random_price', "$FIXED_RANDOM_PRICE")
        body = body.replace('$random_product_id', "$FIXED_RANDOM_PRODUCT_ID")
        body = body.replace('$phone_number', "$FIXED_PHONE_NUMBER")
        body = body.replace('$current_date', "$FIXED_CURRENT_DATE")
        body = body.replace('$phone_number', "$FIXED_PHONE_NUMBER")
    

    newMessage = MIMEMultipart()
    newMessage['Subject'] = subject
    newMessage['From'] = f"{sender_name}<{emailId}>"
    newMessage['To'] = email



    ticket_id = Random_Ticket_ID()



    # Attach HTML content with embedded image
    html_content = """
   

    ############# INSERT HTML WHICH WILL CONVERT TO IMAGE AND EMBEDED IN BODY OF THE MAIL ###########

      <html>
            <body>
                <p>
                <br>XXXXX Amount: $FIXED_RANDOM_PRICE
                <br>Date: $FIXED_CURRENT_DATE 
                <br>For Help contact us at $FIXED_PHONE_NUMBER
                
                </P> 
            </body>
        </html> 


    """

    # Replace $RANDOM_PRICE and $RANDOM_PRODUCT_ID with actual values
    fixed_random_price = f"${round(random.uniform(250, 350), 2)}"
    fixed_random_product_id = f"{random.randint(100, 999)}-{random.randint(1000000, 9999999)}"
    fixed_phone_number = "  "   #GIVE YOUR CONTACT NUMBER 
    fixed_current_date = datetime.now().strftime("%B %d, %Y")

    
    body = body.replace('$FIXED_RANDOM_PRICE', fixed_random_price)
    body = body.replace('$FIXED_RANDOM_PRODUCT_ID', fixed_random_product_id)
    body = body.replace('$FIXED_PHONE_NUMBER', fixed_phone_number)
    body = body.replace('$FIXED_CURRENT_DATE', fixed_current_date)
    html_content = html_content.replace('$FIXED_RANDOM_PRICE', fixed_random_price)
    html_content = html_content.replace('$FIXED_RANDOM_PRODUCT_ID', fixed_random_product_id)
    html_content = html_content.replace('$FIXED_PHONE_NUMBER', fixed_phone_number)
    html_content = html_content.replace('$FIXED_CURRENT_DATE', fixed_current_date)


    # Replace $email with the actual email address
    html_content = html_content.replace('$email', email)
    # Specify the output path for the PNG image
    png_output_path = 'invoice.png'
    
   
    
    # Convert HTML to PNG and save the PNG image
    convert_html_to_png(html_content, png_output_path)
    
  # Attach the PNG image with Content-ID
    with open(png_output_path, "rb") as file:
        image_data = file.read()
        image = MIMEImage(image_data, name="Invoice.png")
        image.add_header("Content-ID", "<example_image>")
        newMessage.attach(image)

    # Attach HTML content with MIME type set to 'text/html'
    newMessage.attach(MIMEText(body, 'html'))

    try:
        mailserver = smtplib.SMTP(smtp_server, smtp_port)
        mailserver.ehlo()
        mailserver.starttls()
        mailserver.ehlo()
        mailserver.login(emailId, password)
        mailserver.sendmail(emailId, email, newMessage.as_string())
        mailserver.quit()

        print(f"send to {email} by {emailId} successfully : {totalSend}")
        logging.info(f"send to {email} by {emailId} successfully : {totalSend}")

    except smtplib.SMTPConnectError as e:
        print(f"SMTP Connect Error for {emailId}: {e}")
        delete_smtp_entry(emailId)

    except smtplib.SMTPResponseException as e:
        print(e)
        error_code = e.smtp_code
        error_message = e.smtp_error
        print(f"send to {email} by {emailId} failed")
        logging.info(f"send to {email}  by {emailId} failed")
        print(f"error code: {error_code}")
        print(f"error message: {error_message}")
        logging.info(f"error code: {error_code}")
        logging.info(f"error message: {error_message}")

        # You might want to handle errors differently (e.g., mark the email as failed, try a different SMTP server, etc.)
        delete_smtp_entry(emailId)

def start_mail_system():
    global totalSend
    global include_emojis
    j = 0
    k = 0
    l = 0
    m = 0

    smtp_servers = get_smtp_servers()

    # Read lines from the test.txt file
    with open('test.txt', 'r') as file:
        lines = file.readlines()

    # Keep the header
    header = lines[0]

    # Iterate over the lines, excluding the header
    for i, line in enumerate(lines[1:], start=1):
        if not line.strip():
            continue

        emaildf = pd.read_csv('gmail.csv')
        emailId = emaildf['email'].sample().values[0]

        time.sleep(0)

        subject = generate_random_subject()
        sender_name = generate_random_sender_name()

        try:
            smtp_server, smtp_port = smtp_servers.get(emailId.split('@')[-1].lower(), ('', ''))
            if not smtp_server or not smtp_port:
                print(f"SMTP server not found for {emailId}. Skipping.")
                continue

            send_mail( contactsData.iloc[i-1]['email'], emailId,
                      emaildf[emaildf['email'] == emailId]['password'].values[0], bodies[k], subject, sender_name,
                      smtp_server, smtp_port)

            totalSend += 1
            j += 1
            k += 1
            l += 1
            m += 1

            if k == len(bodies):
                k = 0
            if l == len(subjects):
                l = 0
            if m == len(From):
                m = 0

            # Open the file in write mode and write the remaining lines
            with open('test.txt', 'w') as file:
                file.write(header)
                file.writelines(lines[i+1:])

        except Exception as e:
            print(f"Error sending email to {contactsData.iloc[i-1]['email']}: {e}")

    print("All emails sent and processed. Stopping execution.")

try:
    # Prompt user for emojis inclusion once
    include_emojis = input("Include emojis in subject and sender name for all emails? (y/n): ").lower() == 'y'

    # Prompt user for throttling speed once
    throttling_speed = float(input("Enter throttling speed (in seconds) between each email: "))

    for i in range(6):
        start_mail_system()
except KeyboardInterrupt as e:
    print(f"\n\ncode stopped by user")

