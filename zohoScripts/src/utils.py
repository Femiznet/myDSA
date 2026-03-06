import os
from dotenv import load_dotenv
from email.message import EmailMessage
from pathlib import Path

load_dotenv()

def required(key:str):
    try:
        value = os.getenv(key.upper())
        if not value:
            raise ValueError (f"Environment variable returned {value}")
        return value
    except Exception as e:
        raise RuntimeError(f"{key} not set") from e

def create_msg(from_addr:str, to_addrs:str, msg_content:str=None, subject:str=None, cc:None=None):
    msg  = EmailMessage()
    msg["from"] = from_addr
    msg["to"] = to_addrs
    msg["cc"] = cc
    msg["subject"] = subject
    msg.set_content(msg_content)
    
    return msg

def get_msg(path:str):
    if not path:
        return
    
    try:
        filepath = Path(path)
    except Exception as Exc:
        print (Exc)
        return
    
    if not filepath.exists() or not filepath.is_file():
        print (f"{filepath} is not a file or Doesnt exist")
        return 

    with open(path) as readFile:
        return "\n".join(readFile.readlines())
    
def validate(email:str, password:str, cc:str, to:str, msg_text:str):
    if not email:
        raise ValueError("Missing input for Email")
    if not password:
        raise ValueError("Missing input for Password")
    if not cc:
        raise ValueError("Missing input for CC")
    if not to:
        raise ValueError("Missing input for Recipient Email")
    if not msg_text:
        print ("Message body is empty")