import smtplib
from datetime import date, datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

from dotenv import dotenv_values

class EmailSender:
    def __init__(self):
        self.totalPass = 8 
        self.totalFail = 2
        self.totalError = 3
         # Load environment variables from .env.secrets file
        config = dotenv_values('.env.secrets')

        # Setup port number and server name
        self._smtpServer = 'mail.genmills.com'     # Google SMTP Server
        self._smtpPort = 587                       # 587 is the Standard secure SMTP Port
        self._smtpUserEmail = config['SMTPServer_Email']
        # self._smtpPassword = config['SMTPServer_Password']
        
    def send_email(self, EMAIL_TO,  FILE_PATH=None):

        if isinstance(EMAIL_TO, str):
            self.emailToList = [EMAIL_TO]  # Convert single email to a list
        elif isinstance(EMAIL_TO, list):
            self.emailToList = EMAIL_TO    # Use the provided list of emails
        else:
            raise ValueError("EMAIL_TO must be a string or a list of strings")

        subject = "Content Audit Summary - PET"

        for email in self.emailToList:
            personEmail = email
            personName = email.split('@')[0].replace('.', ' ')

            # Body of email
            body = f"""
                <html>
                    <head>
                        <style>
                            h2 {{
                                text-align: center;
                            }}
                            h2, b {{ color: darkblue; }}
                            th {{ color: #0ea5e9 }}
                            
                            th, td {{
                                border: 1px solid #96D4D4;
                                border-collapse: collapse;
                            }}
                            
                            table {{
                                width: 250px;
                            }}

                            th, td {{
                                padding: 10px;
                                /* border-color: #96D4D4; */
                            }}
                            
                            .value {{
                                text-align: center;
                            }}
                        </style>
                    </head>
                    <body>
                        <h2>Content Audit Summary - PET</h2>
                        <p>Below is a summary of products that you have requested to compare/Audit.<p/>
                        <p>
                            <b> Item Summary – {date.today().strftime("%d/%m/%Y")} | {datetime.now().strftime("%H:%M:%S")} </b>
                        </p>
                        <table>
                            <tr>
                                <th colspan="2">Walmart Network Items</th>
                            </tr>
                            <tr>
                                <td>SKUs Pass</td>
                                <td class="value">{self.totalPass}</td>
                            </tr>
                            <tr>
                                <td>SKUs Fail</td>
                                <td class="value">{self.totalFail}</td>
                            </tr>
                            <tr>
                                <td>SKUs Error</td>
                                <td class="value">{self.totalError}</td>
                            </tr>
                        </table>

                        <p>If you have any questions or received this in error, please contact sachin.bairi@genmills.com</p>
                    </body>
                </html>
                """
            
            # Make a MIME object to define parts of the email
            msg = MIMEMultipart()
            msg['From'] = self._smtpUserEmail
            msg['To'] = personEmail
            msg['Subject'] = subject
            msg.attach(MIMEText(body, 'html'))

            if FILE_PATH:
                # Define the file to attach
                fileName = self.get_file_name(FILE_PATH)

                try:
                    # Open the file in python as a binary
                    with open(FILE_PATH, 'rb') as attachment:
                        # Encode as Base64
                        attachmentPackage = MIMEBase('application', 'octet-stream')
                        attachmentPackage.set_payload(attachment.read())
                        encoders.encode_base64(attachmentPackage)
                        attachmentPackage.add_header('Content-Disposition', 'attachment', filename=fileName)
                        msg.attach(attachmentPackage)
                except Exception as e:
                    print(f'Error converting attachment: {str(e)}')

            try:
                # message = msg.as_string()  # Cast as string

                # Connect with the server
                print("Connecting to server...")
                # Create a secure SSL/TLS connection to the SMTP server
                server = smtplib.SMTP(self._smtpServer)
                print("Connected to server...")
                # server.starttls()
                print(f"Sending email to {personEmail}")
                server.send_message(msg)

            except Exception as e:
                print(f'Error sending email: {str(e)}')

            finally:
                # Close the SMTP server connection
                server.quit()
    
    def get_file_name(self, FILE_PATH):
        # Simple utility func to trim file_name
        file_name = FILE_PATH.split('/')[-1]
        return file_name
        

if __name__ == "__main__":
    emailToList = ["sachin.bairi@genmills.com"]
    email = EmailSender()
    
    email.send_email(EMAIL_TO=emailToList, FILE_PATH='./Destination/AuditSheet.xlsx')



