import os
import random
import string
import datetime
import base64
import logging
import hashlib
import concurrent.futures
import requests
import qrcode
import smtplib
import tempfile
import shutil
import imgkit
import configparser
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from exchangelib import Credentials, Account, Message, Mailbox, FileAttachment, HTMLBody, Configuration, DELEGATE
from colorama import Fore, Style, init

init()

def show_banner():
    banner = """
  ██████╗  ██╗  ██╗ ██████╗ ███████╗████████╗██████╗ ███████╗██╗      █████╗ ██╗   ██╗
 ██╔════╝  ██║  ██║██╔═══██╗██╔════╝╚══██╔══╝██╔══██╗██╔════╝██║     ██╔══██╗╚██╗ ██╔╝
 ██║  ███╗███████║██║   ██║███████╗   ██║   ██████╔╝█████╗  ██║     ███████║ ╚████╔╝ 
 ██║   ██║██╔══██║██║   ██║╚════██║   ██║   ██╔══██╗██╔══╝  ██║     ██╔══██║  ╚██╔╝  
 ╚██████╔╝██║  ██║╚██████╔╝███████║   ██║   ██║  ██║███████╗███████╗██║  ██║   ██║   
  ╚═════╝ ╚═╝  ╚═╝ ╚═════╝ ╚══════╝   ╚═╝   ╚═╝  ╚═╝╚══════╝╚══════╝╚═╝  ╚═╝   ╚═╝   
    """
    colored_banner = ""
    for char in banner:
        if char != ' ' and char != '\n':
            colored_banner += Fore.GREEN + char
        else:
            colored_banner += char
    print(colored_banner + Style.RESET_ALL)

def show_status(message, status="info"):
    if status == "success":
        print(f"{Fore.GREEN}✓ {message}{Style.RESET_ALL}")
    elif status == "error":
        print(f"{Fore.RED}✗ {message}{Style.RESET_ALL}")
    elif status == "warning":
        print(f"{Fore.YELLOW}⚠ {message}{Style.RESET_ALL}")
    else:
        print(f"{Fore.CYAN}ℹ {message}{Style.RESET_ALL}")

# Initialize configuration
config = configparser.ConfigParser()
config.read('config.ini')

# Setup logging and temp directory
os.makedirs(config['options']['temp_dir'], exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('email_sender.log'),
        logging.StreamHandler()
    ]
)

def generate_random_string(length=38):
    """Generate random hexadecimal string"""
    return ''.join(random.choices(string.hexdigits, k=length))

def generate_qr_code(data, file_path=None):
    """Generate QR code image"""
    qr = qrcode.QRCode(version=1, box_size=10, border=5)
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill='black', back_color='white')
    if file_path:
        img.save(file_path)
    return file_path if file_path else None

def get_company_logo(domain, file_path=None):
    """Fetch company logo from Clearbit"""
    try:
        response = requests.get(f"https://logo.clearbit.com/{domain}", stream=True)
        if response.status_code == 200 and file_path:
            with open(file_path, 'wb') as f:
                for chunk in response.iter_content(1024):
                    f.write(chunk)
            return file_path
    except Exception as e:
        logging.warning(f"Failed to fetch logo for {domain}: {e}")
    return None

def get_company_name(domain):
    """Extract company name from domain"""
    return domain.split(".")[0].capitalize()

def replace_placeholders(content, placeholders):
    """Replace template placeholders with actual values"""
    for key, value in placeholders.items():
        content = content.replace(f"##{key}##", str(value))
    return content

def convert_html_to_image(html_content, output_path, placeholders, logo_path=None, qr_path=None):
    """Convert HTML to image using imgkit"""
    try:
        temp_dir = tempfile.mkdtemp()
        temp_html_path = os.path.join(temp_dir, 'temp.html')
        
        with open(config['options']['conversion_template_path'], 'r', encoding='utf-8') as f:
            template = f.read()
        
        processed_content = replace_placeholders(html_content, placeholders)
        full_html = template.replace('<!-- CONTENT_PLACEHOLDER -->', processed_content)
        full_html = replace_placeholders(full_html, placeholders)
        
        if logo_path and os.path.exists(logo_path):
            logo_temp_path = os.path.join(temp_dir, 'logo.png')
            shutil.copy(logo_path, logo_temp_path)
            full_html = full_html.replace('src="cid:logo"', 'src="logo.png"')
        
        if qr_path and os.path.exists(qr_path):
            qr_temp_path = os.path.join(temp_dir, 'qrcode.png')
            shutil.copy(qr_path, qr_temp_path)
            full_html = full_html.replace('src="cid:qrcode"', 'src="qrcode.png"')
        
        with open(temp_html_path, 'w', encoding='utf-8') as f:
            f.write(full_html)
        
        imgkit_options = {
            'format': 'png',
            'encoding': "UTF-8",
            'quiet': '',
            'enable-local-file-access': '',
            'width': '800',
            'quality': '100'
        }
        
        imgkit.from_file(temp_html_path, output_path, options=imgkit_options)
        return True
    except Exception as e:
        logging.error(f"HTML to image conversion failed: {e}")
        return False
    finally:
        if 'temp_dir' in locals() and os.path.exists(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)

def send_via_smtp(recipient_address, processed_email_html, placeholders, logo_file, qr_file, html_image_file):
    """SMTP sender with flexible TLS/non-TLS handling"""
    try:
        # Prepare email message
        msg = MIMEMultipart('related')
        msg['From'] = f"{placeholders['SENDER_NAME']} <{placeholders['SENDER_EMAIL']}>"
        msg['To'] = recipient_address
        msg['Subject'] = placeholders['SUBJECT']

        # Create email body
        html_part = MIMEMultipart('alternative')
        html_part.attach(MIMEText("Please view this email in an HTML-compatible client", 'plain'))
        html_part.attach(MIMEText(processed_email_html, 'html'))
        msg.attach(html_part)

        # Attach images
        def attach_image(file_path, cid):
            if os.path.exists(file_path):
                with open(file_path, 'rb') as f:
                    img = MIMEImage(f.read())
                    img.add_header('Content-ID', f'<{cid}>')
                    msg.attach(img)

        attach_image(logo_file, 'logo')
        if config['options'].getboolean('include_qr_code'):
            attach_image(qr_file, 'qrcode')
        if config['options'].getboolean('convert_to_img'):
            attach_image(html_image_file, 'htmlimage')

        # Connection handling
        host = config['smtp']['server']
        port = config['smtp'].getint('port')
        use_tls = config['smtp'].getboolean('use_tls')
        use_auth = config['smtp'].getboolean('use_auth')

        # Create SMTP connection
        with smtplib.SMTP(host, port) as server:
            # Attempt STARTTLS if configured
            if use_tls:
                server.starttls()
            
            # Authenticate if required
            if use_auth:
                server.login(config['smtp']['username'], config['smtp']['password'])
            
            # Send message
            server.send_message(msg)
            return True

    except Exception as e:
        logging.error(f"SMTP failed for {recipient_address}: {str(e)}")
        return False

def send_via_owa(recipient_address, processed_email_html, placeholders, logo_file, qr_file, html_image_file):
    """Send email via Exchange OWA"""
    try:
        credentials = Credentials(
            username=config['owa']['email'],
            password=config['owa']['password']
        )
        config_ex = Configuration(
            server=config['owa']['server'],
            credentials=credentials
        )
        account = Account(
            primary_smtp_address=config['owa']['email'],
            config=config_ex,
            autodiscover=False,
            access_type=DELEGATE
        )
        
        # Create message
        m = Message(
            account=account,
            subject=placeholders['SUBJECT'],
            body=HTMLBody(processed_email_html),
            to_recipients=[Mailbox(email_address=recipient_address)]
        )
        
        # Attach images as inline
        def attach_inline(file_path, cid, name):
            if os.path.exists(file_path):
                with open(file_path, 'rb') as f:
                    m.attach(FileAttachment(
                        name=name,
                        content=f.read(),
                        content_id=cid,
                        is_inline=True
                    ))
        
        attach_inline(logo_file, 'logo', 'logo.png')
        if config['options'].getboolean('include_qr_code'):
            attach_inline(qr_file, 'qrcode', 'qrcode.png')
        if config['options'].getboolean('convert_to_img'):
            attach_inline(html_image_file, 'htmlimage', 'email_body.png')
        
        m.send_and_save()
        return True
    except Exception as e:
        logging.error(f"OWA send failed for {recipient_address}: {e}")
        return False

def send_email(recipient_address, email_count):
    """Main email sending function"""
    temp_dir = ""
    try:
        # Extract recipient info
        recipient_name = recipient_address.split("@")[0]
        recipient_domain = recipient_address.split("@")[1]
        
        # Create temp directory for this email
        temp_dir = os.path.join(config['options']['temp_dir'], f"email_{email_count}")
        os.makedirs(temp_dir, exist_ok=True)
        
        # Generate files
        logo_file = os.path.join(temp_dir, "logo.png")
        qr_file = os.path.join(temp_dir, "qrcode.png")
        html_image_file = os.path.join(temp_dir, "email_body.png")
        
        if config['options'].getboolean('include_qr_code'):
            generate_qr_code(f"http://{recipient_name}{recipient_address}", qr_file)
        
        logo_available = get_company_logo(recipient_domain, logo_file)

        # Prepare all placeholders
        placeholders = {
            "RECIPIENT_NAME": recipient_name,
            "RECIPIENT_EMAIL": recipient_address,
            "RECIPIENT_DOMAIN": recipient_domain,
            "CURRENT_DATE": datetime.datetime.now().strftime("%A, %B %d, %Y"),
            # [Add all other placeholders...]
            "SENDER_NAME": config['smtp']['sender_name'],
            "SENDER_EMAIL": config['smtp']['sender_email'],
            "SUBJECT": config['options']['subject_template']
        }

        # Process email template
        with open(config['options']['email_body_path'], 'r', encoding='utf-8') as f:
            email_html = f.read()
        processed_email_html = replace_placeholders(email_html, placeholders)

        # Convert to image if needed
        if config['options'].getboolean('convert_to_img'):
            if not convert_html_to_image(
                email_html,
                html_image_file,
                placeholders,
                logo_path=logo_file if logo_available else None,
                qr_path=qr_file if config['options'].getboolean('include_qr_code') else None
            ):
                logging.warning(f"Failed to generate image for {recipient_address}")

        # Send via appropriate method
        if config['owa'].getboolean('use_owa'):
            success = send_via_owa(recipient_address, processed_email_html, placeholders, 
                                 logo_file, qr_file, html_image_file)
        else:
            success = send_via_smtp(recipient_address, processed_email_html, placeholders,
                                  logo_file, qr_file, html_image_file)

        if success:
            with open(config['options']['log_success'], 'a') as f:
                f.write(f"{datetime.datetime.now()} - Email sent to {recipient_address}\n")
            show_status(f"Email sent to {recipient_address}", "success")
            return True
        else:
            raise Exception("Sending failed (check logs for details)")
            
    except Exception as ex:
        logging.error(f"Failed to send to {recipient_address}: {ex}")
        with open(config['options']['log_failure'], 'a') as f:
            f.write(f"{datetime.datetime.now()} - Failed to send to {recipient_address}: {ex}\n")
        show_status(f"Failed to send to {recipient_address}: {ex}", "error")
        return False
    finally:
        if temp_dir:
            shutil.rmtree(temp_dir, ignore_errors=True)

def start_sending_emails():
    """Main function to start email sending process"""
    try:
        show_banner()
        
        # Validate required files
        if not os.path.exists('list.txt'):
            show_status("list.txt file not found", "error")
            return
        
        with open('list.txt', 'r') as f:
            recipient_addresses = [line.strip() for line in f if line.strip()]
            if not recipient_addresses:
                show_status("list.txt is empty", "error")
                return

        total_emails = len(recipient_addresses)
        success_count = 0
        failure_count = 0
        
        show_status(f"Starting to send {total_emails} emails...")
        
        # Process emails with thread pool
        with concurrent.futures.ThreadPoolExecutor(
            max_workers=config['options'].getint('concurrency_limit')) as executor:
            
            future_to_email = {
                executor.submit(send_email, addr, idx): addr 
                for idx, addr in enumerate(recipient_addresses, 1)
            }
            
            for future in concurrent.futures.as_completed(future_to_email):
                email = future_to_email[future]
                try:
                    if future.result():
                        success_count += 1
                    else:
                        failure_count += 1
                except Exception as exc:
                    failure_count += 1
                    logging.error(f"Error processing {email}: {exc}")
        
        # Show final results
        show_status("Email sending completed.")
        show_status(f"Success: {success_count}/{total_emails} ({success_count/total_emails*100:.1f}%)", "success")
        if failure_count > 0:
            show_status(f"Failures: {failure_count}", "error")
        
    except Exception as ex:
        logging.error(f"Fatal error: {ex}")
        show_status(f"Fatal error: {ex}", "error")

if __name__ == "__main__":
    start_sending_emails()