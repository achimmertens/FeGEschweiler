from exchangelib import Credentials, Account, Configuration, DELEGATE, FileAttachment
import os
import configparser
import re

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


# Funktion zum Extrahieren von Feldwerten aus dem E-Mail-Text
def extract_field(text, field_name):
    pattern = re.compile(fr"{field_name}: ([^\n\r]+)")
    match = pattern.search(text)
    return match.group(1).strip() if match else None

# Funktion zum Umwandeln des Datumsformats von DD.MM.YYYY zu YYYYMMDD
def convert_date_format(date_str):
    return "".join(reversed(date_str.split('.')))

# Durchsuche die Inbox nach E-Mails und verarbeite Anhänge
for item in account.inbox.all():
    # Extrahiere benötigte Informationen aus dem E-Mail-Text
    bereich = extract_field(item.text_body, "Bereich")
    kaufdatum = extract_field(item.text_body, "Kaufdatum")
    empfaenger = extract_field(item.text_body, "Empfänger")
    summe = extract_field(item.text_body, "Summe")
    print ('Bereich: ', bereich, '   Empfänger: ', empfaenger, '   Summe: ', summe, '   Kaufdatum: ', kaufdatum)
    if bereich and kaufdatum and empfaenger and summe:
        # Erstelle einen spezifischen Ordner basierend auf dem Bereich
        specific_folder = os.path.join(attachments_folder, bereich)
        if not os.path.exists(specific_folder):
            os.makedirs(specific_folder)

        kaufdatum_formatted = convert_date_format(kaufdatum)

        for attachment in item.attachments:
            if isinstance(attachment, FileAttachment):
                # Konstruiere den neuen Dateinamen
                new_filename = f"{kaufdatum_formatted}_{empfaenger}_{summe}_{attachment.name}"
                new_filename = new_filename.replace("€", "EUR")  # Ersetze das Euro-Symbol für die Dateinamenskompatibilität
                new_filename = f"{kaufdatum_formatted}_{empfaenger}_{summe.replace(',', '.').replace(' ', '')}_{attachment.name}"
                # Vollständiger Pfad für den Anhang
                local_path = os.path.join(specific_folder, new_filename)

                # Speichere den Anhang
                with open(local_path, 'wb') as f:
                    f.write(attachment.content)
                print(f'Anhang {attachment.name} wurde als {new_filename} in {specific_folder} gespeichert.')
    else:
        print("Nicht alle erforderlichen Felder wurden in der E-Mail gefunden.")

print('Verarbeitung abgeschlossen.')