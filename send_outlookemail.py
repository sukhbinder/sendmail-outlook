
import argparse
import sys
from win32com.client import Dispatch


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

    mail.Send()


if __name__ == "__main__":
    main()
