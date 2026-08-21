from tkinter import *
from tkinter import messagebox, filedialog
import os
import smtplib
import time
import pandas as pd  # pandas and pyxl for reading the content from the excel file
from PIL import ImageTk

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

root = Tk()

# ---------------------------------------------------------------------------
# Color palette (kept in one place so the look is easy to tweak later)
# ---------------------------------------------------------------------------
COLOR_BG = "#F4F6F9"          # main window background (soft light grey)
COLOR_HEADER = "#1F2A44"      # deep navy header
COLOR_HEADER_SUB = "#2E3B5C"  # slightly lighter navy for the subtitle strip
COLOR_ACCENT = "#F2A33C"      # warm gold accent (buttons, highlights)
COLOR_ACCENT_DARK = "#D4881F"
COLOR_TEXT = "#1F2937"        # near-black text
COLOR_MUTED = "#6B7280"       # muted grey text
COLOR_CARD = "#FFFFFF"        # white "card" fields
COLOR_DANGER = "#E5484D"
COLOR_SUCCESS = "#16A34A"
COLOR_FOOTER = "#141B2E"      # near-black footer bar


####################################### EMAIL SENDING LOGIC #######################################
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


####################################### GUI APP #######################################
class Email:
    def __init__(self, root):
        self.root = root
        self.root.title("BulkPost — Bulk Email Sender")
        self.root.geometry("1000x620+200+40")
        self.root.resizable(False, False)
        self.root.config(bg=COLOR_BG)

        # path of the document to attach (None = no attachment)
        self.attachment_path = None

        ######################### ICON #######################################
        self.Email_icon = ImageTk.PhotoImage(file="email01.png")
        self.Setting_icon = ImageTk.PhotoImage(file="setting.png")

        ##################### HEADER ###########################
        header = Label(self.root, bg=COLOR_HEADER)
        header.place(x=0, y=0, relwidth=1, height=90)

        brand = Label(header, text="BulkPost", font=("Segoe UI", 34, "bold"),
                      bg=COLOR_HEADER, fg="white")
        brand.place(x=110, y=8)

        tagline = Label(header, text="Send Smarter. Reach Further.", font=("Segoe UI", 11, "italic"),
                         bg=COLOR_HEADER, fg=COLOR_ACCENT)
        tagline.place(x=113, y=55)

        btn_set1 = Button(header, image=self.Email_icon,
                           bg=COLOR_HEADER, bd='0', command=LEFT, height=65, width=90,
                           activebackground=COLOR_HEADER).place(x=5, y=10)

        btn_set = Button(header, image=self.Setting_icon,
                          bg=COLOR_HEADER, bd='0', command=self.setting_window, height=65, width=90,
                          activebackground=COLOR_HEADER).place(x=900, y=10)

        sub_strip = Label(self.root, text="Bulk mode reads recipients straight from an Excel file",
                           font=("Segoe UI", 11), bg=COLOR_HEADER_SUB, fg="#D7DEEC")
        sub_strip.place(x=0, y=90, relwidth=1, height=28)

        ##################### MODE SELECT #######################
        self.var_choice = StringVar()
        single = Radiobutton(root, text="Single", value="single", command=self.check_single_OR_bulk,
                              activebackground=COLOR_BG, variable=self.var_choice, font=("Segoe UI", 16, "bold"),
                              bg=COLOR_BG, fg=COLOR_TEXT, selectcolor=COLOR_CARD).place(x=50, y=138)

        multiple = Radiobutton(root, text="Multiple (Excel)", value="multiple", command=self.check_single_OR_bulk,
                                variable=self.var_choice, activebackground=COLOR_BG,
                                font=("Segoe UI", 16, "bold"), bg=COLOR_BG, fg=COLOR_TEXT,
                                selectcolor=COLOR_CARD).place(x=200, y=138)
        self.var_choice.set("single")

        ######################### FIELD LABELS #####################################
        To = Label(self.root, text="To (Email Address)", font=("Segoe UI", 13, "bold"), bg=COLOR_BG,
                   fg=COLOR_TEXT).place(x=50, y=200)

        Subject = Label(self.root, text="Subject", font=("Segoe UI", 13, "bold"), bg=COLOR_BG,
                         fg=COLOR_TEXT).place(x=50, y=250)

        Message = Label(self.root, text="Message", font=("Segoe UI", 13, "bold"), bg=COLOR_BG,
                         fg=COLOR_TEXT).place(x=50, y=300)

        ############################## ATTACHMENT #####################################
        Attachment = Label(self.root, text="Attachment", font=("Segoe UI", 13, "bold"), bg=COLOR_BG,
                            fg=COLOR_TEXT).place(x=50, y=505)

        self.attach_entry = Entry(self.root, font=("Segoe UI", 11), bg=COLOR_CARD, fg=COLOR_MUTED,
                                   relief=FLAT, highlightthickness=1, highlightbackground="#D1D5DB",
                                   highlightcolor=COLOR_ACCENT, state='readonly')
        self.attach_entry.place(x=280, y=508, width=400, height=32)

        self.btn_attach = Button(self.root, activebackground=COLOR_ACCENT_DARK, text="Attach File",
                                  font=("Segoe UI", 11, "bold"), bg=COLOR_ACCENT, relief=FLAT,
                                  command=self.Attach_button, cursor="hand2", fg="white")
        self.btn_attach.place(x=695, y=508, width=140, height=32)

        self.btn_attach_clear = Button(self.root, activebackground="#C13F44", text="✕",
                                        font=("Segoe UI", 11, "bold"), bg=COLOR_DANGER, relief=FLAT,
                                        command=self.clear_attachment, cursor="hand2", fg="white")
        self.btn_attach_clear.place(x=845, y=508, width=40, height=32)

        ############################################## STATUS ##################################################################
        self.Total = Label(self.root, font=("Segoe UI", 13, "bold"), bg=COLOR_BG, fg=COLOR_TEXT)
        self.Total.place(x=50, y=558)

        self.Sent = Label(self.root, font=("Segoe UI", 13, "bold"), bg=COLOR_BG, fg=COLOR_SUCCESS)
        self.Sent.place(x=280, y=558)

        self.Left = Label(self.root, font=("Segoe UI", 13, "bold"), bg=COLOR_BG, fg=COLOR_ACCENT_DARK)
        self.Left.place(x=430, y=558)

        self.Failed = Label(self.root, font=("Segoe UI", 13, "bold"), bg=COLOR_BG, fg=COLOR_DANGER)
        self.Failed.place(x=580, y=558)

        ############### ENTRY FIELDS ##########################################
        self.to_entry = Entry(self.root, font=("Segoe UI", 13), bg=COLOR_CARD, fg=COLOR_TEXT,
                               relief=FLAT, highlightthickness=1, highlightbackground="#D1D5DB",
                               highlightcolor=COLOR_ACCENT)
        self.to_entry.place(x=280, y=200, width=350, height=32)

        self.sub_entry = Entry(self.root, font=("Segoe UI", 13), bg=COLOR_CARD, fg=COLOR_TEXT,
                                relief=FLAT, highlightthickness=1, highlightbackground="#D1D5DB",
                                highlightcolor=COLOR_ACCENT)
        self.sub_entry.place(x=280, y=250, width=450, height=32)

        self.message_entry = Text(self.root, font=("Segoe UI", 13), bg=COLOR_CARD, fg=COLOR_TEXT,
                                   relief=FLAT, highlightthickness=1, highlightbackground="#D1D5DB",
                                   highlightcolor=COLOR_ACCENT, wrap=WORD)
        self.message_entry.place(x=280, y=300, width=700, height=190)

        btn1 = Button(root, activebackground=COLOR_ACCENT_DARK, command=self.send_email, text="SEND",
                      font=("Segoe UI", 14, "bold"), bg=COLOR_ACCENT, relief=FLAT,
                      fg="white", cursor="hand2").place(x=700, y=555, width=130, height=34)

        btn2 = Button(root, activebackground="#9CA3AF", command=self.clear1, text="CLEAR",
                      font=("Segoe UI", 14, "bold"), bg="#D1D5DB", relief=FLAT,
                      fg=COLOR_TEXT, cursor="hand2").place(x=850, y=555, width=130, height=34)

        self.btn3 = Button(root, activebackground=COLOR_HEADER, text="BROWSE", font=("Segoe UI", 13, "bold"),
                            bg=COLOR_HEADER_SUB, relief=FLAT, command=self.Browse_button, cursor="hand2",
                            state=DISABLED, fg="white", disabledforeground="#9AA5BD")
        self.btn3.place(x=650, y=200, width=150, height=32)

        ############################## DEVELOPER CREDIT (FOOTER) ##############################
        footer_strip = Label(self.root, bg=COLOR_FOOTER)
        footer_strip.place(x=0, y=598, relwidth=1, height=22)

        footer_bar = Frame(footer_strip, bg=COLOR_FOOTER)
        footer_bar.place(relx=0.5, rely=0.5, anchor="center")

        footer_left = Label(footer_bar, text="BulkPost  v1.0   |   Developed by ",
                             font=("Segoe UI", 10), bg=COLOR_FOOTER, fg="white")
        footer_left.pack(side=LEFT)

        footer_name = Label(footer_bar, text="Abhay Pande",
                             font=("Segoe UI", 10, "bold"), bg=COLOR_FOOTER, fg=COLOR_ACCENT)
        footer_name.pack(side=LEFT)

        ################################################ Browse ########################################################################
        self.check_file_exist()

    def Browse_button(self):
        op = filedialog.askopenfile(initialdir='/', title="Select Excel File for Emails",
                                     filetypes=(("All Files", "*.*"), ("Excel Files", ".xlsx")))
        if op != None:
            data = pd.read_excel(op.name)
            if 'Email' in data.columns:
                self.EMAIL = list(data['Email'])
                c = []
                for i in self.EMAIL:
                    if (pd.isnull(i)) == False:
                        c.append(i)
                self.EMAIL = c
                if len(self.EMAIL) > 0:
                    self.to_entry.config(state=NORMAL)
                    self.to_entry.delete(0, END)
                    self.to_entry.insert(0, str(op.name.split("/")[-1]))
                    self.to_entry.config(state='readonly')
                    self.Total.config(text="Total: " + str(len(self.EMAIL)))
                    self.Sent.config(text="Sent: ")
                    self.Left.config(text="Left: ")
                    self.Failed.config(text="Failed: ")
            else:
                messagebox.showinfo("Error", "Please Select A File Which Has Emails", parent=self.root)

    ############################## ATTACHMENT HANDLERS #####################################
    def Attach_button(self):
        """Let the user pick any document (pdf, docx, xlsx, image, etc.) to attach."""
        file_path = filedialog.askopenfilename(
            initialdir='/', title="Select a Document to Attach",
            filetypes=(("All Files", "*.*"),
                       ("PDF", "*.pdf"),
                       ("Word Documents", "*.docx *.doc"),
                       ("Excel Files", "*.xlsx *.xls"),
                       ("Images", "*.png *.jpg *.jpeg")))
        if file_path:
            self.attachment_path = file_path
            self.attach_entry.config(state=NORMAL)
            self.attach_entry.delete(0, END)
            self.attach_entry.insert(0, os.path.basename(file_path))
            self.attach_entry.config(state='readonly')

    def clear_attachment(self):
        self.attachment_path = None
        self.attach_entry.config(state=NORMAL)
        self.attach_entry.delete(0, END)
        self.attach_entry.config(state='readonly')

    #################################### SEND EMAIL #############################
    def send_email(self):
        x = len(self.message_entry.get('1.0', END))
        if self.to_entry.get() == "" or self.sub_entry.get() == "" or x == 1:
            messagebox.showerror("ERROR", "All feilds are required", parent=self.root)
        else:
            if self.var_choice.get() == "single":
                status = Email_send_function(
                    self.to_entry.get(), self.sub_entry.get(), self.message_entry.get('1.0', END),
                    self.uname, self.pasw, self.attachment_path)
                if status == "s":
                    messagebox.showinfo("SUCCESS", "Email Has Been Sent", parent=self.root)
                if status == "f":
                    messagebox.showerror("Failed", "Email Not Sent", parent=self.root)

            if self.var_choice.get() == "multiple":
                self.failed = []
                self.s_count = 0
                self.f_count = 0
                for x in self.EMAIL:
                    status = Email_send_function(
                        x, self.sub_entry.get(), self.message_entry.get('1.0', END),
                        self.uname, self.pasw, self.attachment_path)
                    if status == "s":
                        self.s_count += 1
                    if status == "f":
                        self.f_count += 1
                    self.status_bar()
                    time.sleep(1)
                messagebox.showinfo("Success", "Email Has Been Sent,Please Check Status....", parent=self.root)

    def clear1(self):
        self.to_entry.config(state=NORMAL)
        self.to_entry.delete(0, END)
        self.sub_entry.delete(0, END)
        self.message_entry.delete('1.0', END)
        self.var_choice.set("single")
        self.btn3.config(state=DISABLED)
        self.clear_attachment()
        self.Total.config(text="")
        self.Sent.config(text="")
        self.Left.config(text="")
        self.Failed.config(text="")

    def status_bar(self):
        self.Total.config(text="Status " + str(len(self.EMAIL)) + ":-")
        self.Sent.config(text="Sent: " + str(self.s_count))
        self.Left.config(text="Left: " + str(len(self.EMAIL) - (self.f_count + self.s_count)))
        self.Failed.config(text="Failed: " + str(self.f_count))
        self.Total.update()
        self.Sent.update()
        self.Left.update()
        self.Failed.update()

    def check_single_OR_bulk(self):
        if self.var_choice.get() == "single":
            messagebox.showinfo("single", "Setted To Single", parent=self.root)
            self.btn3.config(state=DISABLED)
            self.to_entry.config(state=NORMAL)
            self.to_entry.delete(0, END)
            self.clear1()
        if self.var_choice.get() == "multiple":
            messagebox.showinfo("multiple", "Setted To Bulk", parent=self.root)
            self.btn3.config(state=NORMAL)
            self.to_entry.delete(0, END)
            self.to_entry.config(state='readonly')

    ####################################### SETTING FUNCTION ########################
    def setting_clear(self):
        self.uname_entry.delete(0, END)
        self.pasw_entry.delete(0, END)

    def setting_window(self):
        self.check_file_exist()
        self.root2 = Toplevel()
        self.root2.title("BulkPost — Settings")
        self.root2.resizable(False, False)
        self.root2.geometry("700x460+350+90")
        self.root2.focus_force()
        self.root2.grab_set()
        self.root2.config(bg=COLOR_BG)

        header2 = Label(self.root2, bg=COLOR_HEADER)
        header2.place(x=0, y=0, relwidth=1, height=90)

        title2 = Label(header2, text="BulkPost", font=("Segoe UI", 30, "bold"),
                       bg=COLOR_HEADER, fg="white")
        title2.place(x=30, y=8)

        REF2 = Label(header2, text="Account Settings", font=("Segoe UI", 11, "italic"),
                     bg=COLOR_HEADER, fg=COLOR_ACCENT)
        REF2.place(x=33, y=55)

        REF3 = Label(self.root2, text="Enter your Gmail address and App Password",
                     font=("Segoe UI", 11), bg=COLOR_HEADER_SUB, fg="#D7DEEC")
        REF3.place(x=0, y=90, relwidth=1, height=28)

        uname = Label(self.root2, text="Email Address", font=("Segoe UI", 13, "bold"), bg=COLOR_BG,
                      fg=COLOR_TEXT).place(x=50, y=160)

        pasw = Label(self.root2, text="App Password", font=("Segoe UI", 13, "bold"), bg=COLOR_BG,
                     fg=COLOR_TEXT).place(x=50, y=210)

        self.uname_entry = Entry(self.root2, font=("Segoe UI", 13), bg=COLOR_CARD, fg=COLOR_TEXT,
                                  relief=FLAT, highlightthickness=1, highlightbackground="#D1D5DB",
                                  highlightcolor=COLOR_ACCENT)
        self.uname_entry.place(x=230, y=160, width=350, height=32)

        self.pasw_entry = Entry(self.root2, font=("Segoe UI", 13), bg=COLOR_CARD, fg=COLOR_TEXT, show="*",
                                 relief=FLAT, highlightthickness=1, highlightbackground="#D1D5DB",
                                 highlightcolor=COLOR_ACCENT)
        self.pasw_entry.place(x=230, y=210, width=350, height=32)

        ################################################# BUTTON OF SETTING ############################################
        btn1 = Button(self.root2, activebackground=COLOR_ACCENT_DARK, text="SAVE", font=("Segoe UI", 13, "bold"),
                      bg=COLOR_ACCENT, relief=FLAT,
                      fg="white", cursor="hand2", command=self.save_setting).place(x=230, y=265, width=140, height=34)

        btn2 = Button(self.root2, activebackground="#9CA3AF", text="CLEAR", font=("Segoe UI", 13, "bold"),
                      bg="#D1D5DB", relief=FLAT, command=self.setting_clear, cursor="hand2",
                      fg=COLOR_TEXT).place(x=390, y=265, width=140, height=34)

        self.uname_entry.insert(0, self.uname)
        self.pasw_entry.insert(0, self.pasw)

        ################################################# ABOUT / CREDIT ############################################
        about_strip = Label(self.root2, bg=COLOR_FOOTER)
        about_strip.place(x=0, y=438, relwidth=1, height=22)

        about_bar = Frame(about_strip, bg=COLOR_FOOTER)
        about_bar.place(relx=0.5, rely=0.5, anchor="center")

        about_left = Label(about_bar, text="BulkPost  v1.0   |   Developed by ",
                            font=("Segoe UI", 10), bg=COLOR_FOOTER, fg="white")
        about_left.pack(side=LEFT)

        about_name = Label(about_bar, text="Abhay Pande",
                            font=("Segoe UI", 10, "bold"), bg=COLOR_FOOTER, fg=COLOR_ACCENT)
        about_name.pack(side=LEFT)

    ####################################### FOR EMAIL AND PASS IN SETTING ##############################################
    def check_file_exist(self):
        if os.path.exists("important.txt") == False:
            f = open('important.txt', 'w')
            f.write(",")
            f.close()
        f2 = open('important.txt', 'r')
        self.credentials = []
        for i in f2:
            self.credentials.append([i.split(",")[0], i.split(",")[1]])
        self.uname = self.credentials[0][0]
        self.pasw = self.credentials[0][1]

    def save_setting(self):
        if self.uname_entry.get() == "" or self.pasw_entry.get() == "":
            messagebox.showinfo("ERROR", "All feilds are required", parent=self.root2)
        else:
            f = open('important.txt', 'w')
            f.write(self.uname_entry.get() + "," + self.pasw_entry.get())
            f.close()
            messagebox.showinfo("Sent", "Email and password are saved Successfully", parent=self.root2)
            self.check_file_exist()


if __name__ == "__main__":
    obj = Email(root)
    root.mainloop()