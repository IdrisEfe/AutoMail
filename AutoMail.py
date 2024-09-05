import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication

class autoMail:
    def __init__(self, *, text, dear_name, student_name):
        self.body_text = text.format(dear_name = dear_name, student_name = student_name)

    def getBody(self):
        return self.body_text
    
    def get_mime_type(self, file):
        ext = os.path.splitext(file)[1].lower()
        if ext in ['.jpeg', '.jpg']:
            return 'image/jpeg'
        elif ext == '.png':
            return 'image/png'
        elif ext == '.pdf':
            return 'application/pdf'
        elif ext == '.pptx':
            return 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        return 'application/octet-stream'

    def create_mail(self, *, from_email, subject, body, to_email, photo=None, presentation=None):

        photo_format = self.get_mime_type(photo)
        presentation_format = self.get_mime_type(presentation)

        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = ', '.join(to_email) if isinstance(to_email, list) else to_email
        msg['Subject'] = subject

        msg.attach(MIMEText(body, 'plain'))

        if photo and os.path.exists(photo): # in case if it is still None
            with open(photo, 'rb') as f:
                img_data = f.read()
                image = MIMEImage(img_data, _subtype = photo_format, name = os.path.basename(photo))
                msg.attach(image)
        elif photo: # in case if it is still None
            print(f'Error: Photo {photo} not found.')

        if presentation and os.path.exists(presentation): # in case if it is still None
            with open(presentation, 'rb') as f:
                pres_data = f.read()
                attachment = MIMEApplication(pres_data, _subtype = presentation_format, name = os.path.basename(presentation))
                attachment['Content-Disposition'] = f'attachment; filename="{os.path.basename(presentation)}"'
                msg.attach(attachment)
        elif presentation: # in case if it is still None
            print(f'Error: Presentation {presentation} not found.')

        return msg

    def send_mail(self, *, from_email, from_password, subject, body, to_email, photo=None, presentation=None):
        mail = self.create_mail(
                from_email=from_email,
                subject=subject,
                body=body,
                to_email=to_email,
                photo=photo,
                presentation=presentation
                         )
        
        try:

            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(from_email, from_password)
            txt = mail.as_string()
            server.sendmail(from_email, to_email, txt)
            server.quit()
            print('Mail is sent.')

        except Exception as e:
            print(f"Couldn't send the mail: {str(e)}")