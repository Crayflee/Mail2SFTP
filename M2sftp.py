import tkinter as tk
from tkinter import messagebox
from tkinter import scrolledtext
from threading import Thread
import imaplib
import email
import os
import paramiko
from datetime import datetime
import sys,traceback
from dotenv import load_dotenv


load_dotenv()



# E-Mail-Zugangsdaten
EMAIL = os.getenv("EMAIL")
PASSWORD = os.getenv('EMAIL_PASSWORD')
IMAP_SERVER = os.getenv('IMAP_SERVER')

# Benötigte mailboxen 
MAILBOX = 'INBOX'
ERROR_MAILBOX = 'INBOX.fehler'
PROCESSED_MAILBOX = 'INBOX.verarbeitet'

# SFTP-Zugangsdaten
SFTP_HOST = os.getenv('SFTP_HOST')
SFTP_PORT = os.getenv('SFTP_PORT')
SFTP_USERNAME = os.getenv('SFTP_USERNAME')
SFTP_PASSWORD = os.getenv('SFTP_PASSWORD')


# Dictionary für GUI
strings = {
    "win_title": "Mail2SFTP",
    "last_exec_label": "Die letzte Ausführung war am:",
    "progress_label": "Fortschritt:",
    "start_button_text": "Start",
    "error_title": "Fehler",
    "success_title": "Erfolg",
    "success_message": "Alle E-Mails und Anhänge wurden verarbeitet!",
    "no_exec_msg": "Das Skript wurde bisher noch nicht ausgeführt!",
    "con_imap": "Verbindung zum IMAP-Server wird hergestellt...",
    "con_sftp": "Verbindung zum SFTP-Server wird hergestellt...",
    "move_to_err": "E-Mail erfolgreich in den Fehler-Ordner verschoben.",
    "move_to_proc": "E-Mail erfolgreich in den Verarbeitet-Ordner verschoben."
}

class MailSftpGUI:
    def __init__(self, master):
        self.master = master
        master.title("Mail2SFTP")
        master.geometry("1200x400")  

        self.last_execution_label = tk.Label(master, text="Letzte Ausführung:")
        self.last_execution_label.pack()

        self.last_execution_time_label = tk.Label(master, text="")
        self.last_execution_time_label.pack()

        self.progress_label = tk.Label(master, text="Fortschritt:")
        self.progress_label.pack()

        self.progress_text = scrolledtext.ScrolledText(master, width=1200, height=20)  # Breite und Höhe des Textbereichs
        self.progress_text.pack()

        self.start_button = tk.Button(master, text="Start", command=self.start_process)
        self.start_button.pack()

        # letzte ausführung
        self.last_exec_time = None
        self.load_last_exec_time()
        
    def print_str(self, text):
        str_to_print = text + "\n"
        self.progress_text.insert(tk.END, str_to_print)

    def load_last_exec_time(self):
        try:
            with open('last_execution_time.txt', 'r') as file:
                self.last_exec_time = file.read()
                
        except FileNotFoundError:
            self.last_exec_time = None

        if self.last_exec_time:
            self.last_execution_time_label.config(text=f"Letzte Ausführung: {self.last_exec_time}")
        else:
            self.last_execution_time_label.config(text="Letzte Ausführung: Noch nicht ausgeführt")

    def save_last_exec_time(self):
        with open('last_execution_time.txt', 'w') as file:
            file.write(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
       
    def start_process(self):
        self.start_button.config(state="disabled")
        self.progress_text.delete(1.0, tk.END)
        self.progress_text.insert(tk.END, "Prozess gestartet...\n")
        
        Thread(target=self.run_script).start()
            

    def run_script(self):
        
        try:
            
            self.progress_text.insert(tk.END, "Verbindung zum IMAP-Server wird hergestellt...\n")
            mail = imaplib.IMAP4_SSL(IMAP_SERVER)
            mail.login(EMAIL, PASSWORD)
            mail.select(MAILBOX)
            # wirklich garantie für erfolgreiche Verbindung ?
            self.progress_text.insert(tk.END, "Verbindung zum IMAP-Server wurde hergestellt !\n")
            
            self.progress_text.insert(tk.END, "Verbindung zum SFTP-Server wird hergestellt...\n")
            transport = paramiko.Transport((SFTP_HOST, SFTP_PORT))
            transport.connect(username=SFTP_USERNAME, password=SFTP_PASSWORD)
            sftp = paramiko.SFTPClient.from_transport(transport)
            # wirklich garantie für erfolgreiche Verbindung ?
            self.progress_text.insert(tk.END, "Verbindung zum SFTP-Server wurde hergestellt !\n")
            # result wird benötigt, sonst fehler, wieso ? 
            all_result, all_data = mail.search(None, 'ALL')
            
            
            for email_id in all_data[0].split():
                
                print(email_id)
                
                result,data = mail.fetch(email_id, '(RFC822)') #änderung von uid zu direktem fetch hat auswirkungen auf email ID
                raw_email = data[0][1]
                msg = email.message_from_bytes(raw_email)
                if msg.get_content_maintype() == 'multipart':
                    self.move_mail_to_sftp(msg, sftp, email_id)

            sftp.close()
            transport.close()
            mail.close()
            mail.logout()
            
        except Exception as e:
            traceback.print_exc(None,sys.stdout,True)
            messagebox.showerror("Fehler", f"Ein Fehler ist aufgetreten: {e}") 
        else:
            messagebox.showinfo("Erfolg", "Prozess abgeschlossen!")
            self.save_last_exec_time()
            self.last_exec_time = None
            self.load_last_exec_time()

        self.start_button.config(state="normal")

    def move_mail_to_sftp(self, msg, sftp, email_id):
        attachment_found = False
        for part in msg.walk():
            filename = part.get_filename()
            if filename and filename.lower().endswith('.csv'):
                attachment_found = True
                with open(filename, 'wb') as f:
                    f.write(part.get_payload(decode=True))
                    self.progress_text.insert(tk.END, f"CSV-Anhang heruntergeladen: {filename}\n")
                    # f.close wird benötigt da sonst datei geöffnet bleibt ( Win Fehler) eigentlich auto, wieso nicht ?
                    f.close()

                    # SFTP Upload
                    try:
                        sftp.put(filename, filename)
                        self.progress_text.insert(tk.END, f"CSV-Datei auf SFTP-Server hochgeladen: {filename}\n")
                        os.remove(filename)
                        self.progress_text.insert(tk.END, f"Lokale Datei gelöscht: {filename}\n")
                    except Exception as e:
                        traceback.print_exc(None,sys.stdout,True)
                        self.progress_text.insert(tk.END, f"Fehler beim Upload der Datei auf den SFTP-Server {filename}: {e}\n")
                        self.move_to_mailbox(email_id, True)
                        return

        # Kein Anhang -> mail in fehler
        if not attachment_found:
            self.move_to_mailbox(email_id, True)
        else:
            self.move_to_mailbox(email_id, False)

    def move_to_mailbox(self, email_id, error):
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL, PASSWORD)
        mail.select()
        if error:
            try:
                mail.copy(email_id, ERROR_MAILBOX)
                mail.store(email_id , '+FLAGS', '(\\Deleted)')
                mail.expunge()
                self.progress_text.insert(tk.END, "E-Mail erfolgreich in den Fehler-Ordner verschoben.\n")
            except Exception as e:
                traceback.print_exc(None,sys.stdout,True)
                self.progress_text.insert(tk.END, f"Fehler beim Verschieben der E-Mail in den Fehler-Ordner: {e}\n")
        else:
            try:
                mail.copy(email_id, PROCESSED_MAILBOX)
                mail.store(email_id, '+FLAGS', '(\\Deleted)')
                mail.expunge()
                self.progress_text.insert(tk.END, "E-Mail erfolgreich in den Verarbeitet-Ordner verschoben.\n")
            except Exception as e:
                traceback.print_exc(None,sys.stdout,True)
                self.progress_text.insert(tk.END, f"Fehler beim Verschieben der E-Mail in den Verarbeitet-Ordner: {e}\n")       

        mail.close()
        mail.logout()

def main():
    root = tk.Tk()
    app = MailSftpGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
