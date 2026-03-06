import sys
import logging
from pathlib import Path
from smtplib import SMTP
from utils import create_msg, get_msg, validate, required

# 1. Setup Logging
logging.basicConfig(
    filename=r'..\log\send_mail.log',
    level=logging.INFO,
    format='%(levelname)s => %(message)s => %(asctime)s',
    filemode='a'
)

ADDRESS = "smtp.gmail.com"
PORT = 587

def main():
    if len(sys.argv) < 2:
        logging.error("Script triggered without a file path argument.")
        sys.exit(1)

    file_path = Path(sys.argv[1])
    
    if not file_path.is_file():
        logging.error(f"Invalid file path: {file_path}")
        sys.exit(1)

    if file_path.suffix.lower() != ".xlsx":
        logging.warning(f"Skipped: {file_path.name} is not an XLSX file.")
        sys.exit(1)
    
    # Get Environment Variables
    email = required("EMAIL")
    password = required("PASS")

    Fname = file_path.name
    if Fname.lower().startswith("estate"):
        recipients = required("ELI")
    elif Fname.lower().startswith("plant"):
        recipients = required("OLU")
    else:
        print ("Unresolved file trying to be sent, exiting", Fname)
        exit()
     
    cc = required("TEST-CC").split()
    
    try:
        msg_text = get_msg(r"..\message.txt")
        validate(email=email, password=password, cc=cc, to=recipients, msg_text=msg_text)
        
        msg = create_msg(from_addr=email, to_addrs=recipients, msg_content=msg_text, cc=cc)

        # Attach XLSX
        with open(file_path, 'rb') as f:
            msg.add_attachment(
                f.read(),
                maintype='application',
                subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                filename=file_path.name
            )

        # Send Email
        with SMTP(ADDRESS, PORT) as server:
            server.starttls()
            server.login(email, password)
            server.send_message(msg)
        
        logging.info(f"SUCCESS: Sent {file_path.name} to {recipients} and {cc}")
        print(f"Message sent with attachments: {file_path.name}")

    except Exception as e:
        logging.error(f"FAILED to send {file_path.name}: {str(e)}")
        print(f"Error encountered. Check email_log.txt for details.")

main()