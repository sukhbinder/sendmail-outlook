# sendmail-outlook
Send email using outlook application 

## Usage
```
usage: sendmail [-h] -t TO [-s SUBJECT] [-b [BODY ...]] [-hb HTMLBODY] [-a [ATTACHMENT]]

Sends an email using Outlook

options:
  -h, --help            show this help message and exit
  -t TO, --to TO        Recipient of the email
  -s SUBJECT, --subject SUBJECT
                        Subject of the email
  -b [BODY ...], --body [BODY ...]
                        Body of the email
  -hb HTMLBODY, --htmlbody HTMLBODY
                        HTML Body of the email
  -a [ATTACHMENT], --attachment [ATTACHMENT]
                        Path to attachment file (optional)
```

# Install
pip install sendmail-outlook
