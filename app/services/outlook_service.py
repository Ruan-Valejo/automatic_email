import win32com.client


def create_draft_email(to_email: str, cc_email: str, subject: str, body: str, attachments: list[str]) -> None:
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    mail.To = to_email
    mail.CC = cc_email
    mail.Subject = subject
    mail.Body = body

    for file_path in attachments:
        mail.Attachments.Add(file_path)

    mail.Display()