
import argparse
import sys
from win32com.client import Dispatch
import os
import tempfile


def main():
    parser = argparse.ArgumentParser(description="Sends an email using Outlook")
    parser.add_argument("-t", "--to", required=True, help="Recipient of the email")
    parser.add_argument(
        "-s", "--subject", default="Message subject", help="Subject of the email"
    )
    parser.add_argument(
        "-b", "--body", default="Message body", help="Body of the email", nargs="*"
    )
    parser.add_argument(
        "-hb", "--htmlbody", default=None, help="HTML Body of the email"
    )
    parser.add_argument(
        "-a",
        "--attachment",
        nargs="?",
        const=None,
        type=str,
        help="Path to attachment file (optional)",
    )

    parser.add_argument("-ds", "--dont-send", action="store_true", help="Don't send email, only show")

    args = parser.parse_args()

    outlook = Dispatch("outlook.application")
    mail = outlook.CreateItem(0)
    mail.To = args.to
    mail.Subject = args.subject
    if not args.body:
        body = sys.stdin.read()
        mail.Body = body
    else:
        mail.Body = " ".join(args.body)

    if args.htmlbody:
        mail.HTMLBody = args.htmlbody

    if args.attachment:
        attachment_path = args.attachment
        mail.Attachments.Add(attachment_path)

    if args.dont_send:
        temp_dir = tempfile.gettempdir()    
        msg_file_path = os.path.join(temp_dir, "email.msg")
        mail.SaveAs(msg_file_path, 3)
        os.startfile(msg_file_path)
    else:
        mail.Send()


if __name__ == "__main__":
    main()
