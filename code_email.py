import smtplib  # to deal with gmail mails
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


def Email_send_function(to, subject, message, uname, pasw, attachment_path=None):
    """
    Sends an email, optionally with a single document attached.

    attachment_path: full path to a file (pdf, docx, xlsx, image, etc.)
                      or None / "" if no attachment is needed.
    """
    try:
        s = smtplib.SMTP("smtp.gmail.com", 587)  # create session for gmail
        s.starttls()  # transport layer
        s.login(uname, pasw)

        # Build a proper MIME message so attachments work
        msg = MIMEMultipart()
        msg["From"] = uname
        msg["To"] = to
        msg["Subject"] = subject
        msg.attach(MIMEText(message, "plain"))

        if attachment_path:
            if os.path.isfile(attachment_path):
                with open(attachment_path, "rb") as f:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(f.read())
                encoders.encode_base64(part)
                filename = os.path.basename(attachment_path)
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename=\"{filename}\"",
                )
                msg.attach(part)
            else:
                # Attachment path was given but file doesn't exist - fail loudly
                s.close()
                return "f"

        s.sendmail(uname, to, msg.as_string())

        x = s.ehlo()
        s.close()

        if x[0] == 250:
            return "s"
        else:
            return "f"

    except Exception as e:
        print("Email send error:", e)
        return "f"