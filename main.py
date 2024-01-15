import ExcelParser
import SendEmail


def main():
    """Doctype"""
    columns_config = {
        "name": "Name",
        "email_to": "Email_to",
        "email_cc": "Email_cc",
        "links": "links",
        "body": "body",
    }

    signature = (f"Best regards,\n"
                 f"Name\n"
                 f"something\n"
                 f":)")

    parser = ExcelParser.ExcelParser("test.xlsx",  columns_config)
    sender = SendEmail.EmailSender()
    for data in parser.parse_data():
        sender.send_email(data, signature=signature)


if __name__ == '__main__':
    main()
