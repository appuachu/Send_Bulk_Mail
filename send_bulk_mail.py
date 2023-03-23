import tkinter as tk
from tkinter import filedialog, messagebox
import smtplib
from email.mime.text import MIMEText

class BulkEmailSender:
    def __init__(self, master):
        self.master = master
        master.title("Bulk Email Sender")

        # sender email address
        self.label_sender = tk.Label(master, text="Sender Email (Outlook Mail Id):")
        self.label_sender.grid(row=0, column=0)
        self.entry_sender = tk.Entry(master)
        self.entry_sender.grid(row=0, column=1)

        # sender email password
        self.label_password = tk.Label(master, text="Sender Mail Password:")
        self.label_password.grid(row=1, column=0)
        self.entry_password = tk.Entry(master, show="*")
        self.entry_password.grid(row=1, column=1)

        # email subject
        self.label_subject = tk.Label(master, text="Subject:")
        self.label_subject.grid(row=2, column=0)
        self.entry_subject = tk.Entry(master)
        self.entry_subject.grid(row=2, column=1)

        # email message
        self.label_message = tk.Label(master, text="Message:")
        self.label_message.grid(row=3, column=0)
        self.entry_message = tk.Text(master, height=10)
        self.entry_message.grid(row=3, column=1)

        # recipients list file dialog button
        self.label_recipients = tk.Label(master, text="Recipients List:")
        self.label_recipients.grid(row=4, column=0)
        self.button_recipients = tk.Button(master, text="Select File", command=self.select_file)
        self.button_recipients.grid(row=4, column=1)

        # send email button
        self.button_send = tk.Button(master, text="Send", command=self.send_email)
        self.button_send.grid(row=5, column=0)

        # clear form button
        self.button_clear = tk.Button(master, text="Clear", command=self.clear_form)
        self.button_clear.grid(row=5, column=1)

    def select_file(self):
        self.filename = filedialog.askopenfilename()

    def send_email(self):
        # read email addresses from file
        with open(self.filename, 'r') as f:
            email_list = f.readlines()
        email_list = [email.strip() for email in email_list]

        # send email separately to each recipient
        all_sent = True  # flag to indicate whether all emails were sent successfully
        for email in email_list:
            # email content
            msg = MIMEText(self.entry_message.get("1.0", 'end-1c'))
            msg['Subject'] = self.entry_subject.get()
            msg['From'] = self.entry_sender.get()
            msg['To'] = email

            try:
                server = smtplib.SMTP('smtp.outlook.com', 587)
                server.starttls()
                server.login(self.entry_sender.get(), self.entry_password.get())
                server.sendmail(self.entry_sender.get(), email, msg.as_string())
                server.quit()
            except Exception as e:
                messagebox.showerror("Error", f"Error sending email to {email}: {str(e)}")
                all_sent = False  # set flag to False if any email fails to send


        if all_sent:
            messagebox.showinfo("Success", "All emails sent successfully!")

    def clear_form(self):
        self.entry_sender.delete(0, 'end')
        self.entry_password.delete(0, 'end')
        self.entry_subject.delete(0, 'end')
        self.entry_message.delete("1.0", 'end')

if __name__ == "__main__":
    root = tk.Tk()
    sender = BulkEmailSender(root)
    root.mainloop()
