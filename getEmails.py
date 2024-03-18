from exchangelib import Credentials, Account, Configuration, DELEGATE, FileAttachment
import os
import configparser

# Erstelle ein ConfigParser-Objekt und lese die INI-Datei
config = configparser.ConfigParser()
config.read('config.ini')

# Hole die Anmeldeinformationen
email = config['credentials']['email']
password = config['credentials']['password']

# Erstelle Anmeldeinformationen und konfiguriere den Account
credentials = Credentials(email, password)
config = Configuration(server='outlook.office365.com', credentials=credentials)
account = Account(primary_smtp_address=email, config=config, autodiscover=False, access_type=DELEGATE)

# Ordner für die Speicherung der Anhänge
attachments_folder = r'C:\Users\User\git\FeGEschweiler\Attachments'
# Stelle sicher, dass der Zielordner existiert
if not os.path.exists(attachments_folder):
    os.makedirs(attachments_folder)

# Durchsuche die Inbox nach E-Mails und lade Anhänge herunter
for item in account.inbox.all():
    for attachment in item.attachments:
        if isinstance(attachment, FileAttachment):
            local_path = os.path.join(attachments_folder, attachment.name)
            with open(local_path, 'wb') as f:
                f.write(attachment.content)
            print(f'Anhang {attachment.name} wurde gespeichert.')

print('Alle Anhänge wurden heruntergeladen.')
