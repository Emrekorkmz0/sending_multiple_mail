import customtkinter
from tkinter import *
from tkintermapview import TkinterMapView
from PIL import Image
import threading
import requests
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from tkinter.filedialog import askopenfilename
import pandas as pd
customtkinter.set_default_color_theme("blue")
smtp_ssl_host = 'smtp.gmail.com'  # smtp.mail.yahoo.com
smtp_ssl_port = 465
defaultimg = customtkinter.CTkImage(Image.open("images.jpg"), size=(500, 300))


class App(customtkinter.CTk):
    APP_NAME = "EMRE KORKMAZ"
    WIDTH = 250
    HEIGHT = 250

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.title(App.APP_NAME)
        self.minsize(App.WIDTH, App.HEIGHT)

        self.marker_list = []
        self.attachment_path = None 
        self.username = None
        self.password = None

        self.show_login_popup()
        # ============ create two CTkFrames ============

        self.frame_left = customtkinter.CTkFrame(self, corner_radius=0, fg_color=None)
        self.frame_left.pack(side=LEFT, fill=Y)
        
        #------------Left Frame-----------#
        

        self.l_frame = customtkinter.CTkFrame(self.frame_left)
        self.l_frame.pack(fill=X)


        self.v_lable = customtkinter.CTkLabel(self.l_frame, text="", image=defaultimg, width=250, height=250)
        self.v_lable.pack(padx=5, pady=5)


        self.textBox = customtkinter.CTkTextbox(self.l_frame, height=25, width=300,)
        self.textBox.pack(padx=5, pady=5)

        self.textBox.bind("<KeyRelease>", self.adjust_textbox_height)


        self.attach_btn = customtkinter.CTkButton(self.l_frame,
                                                  text="Dosya Ekle",
                                                  command=self.add_attachment)
        self.attach_btn.pack(padx=5, pady=(5, 10))


        # Add placeholder text
        self.placeholder_text = "Mail içeriğini Yaz"
        self.add_placeholder()

        #------------Button Frame-----------#
        self.buttons_frame = customtkinter.CTkFrame(self.frame_left)
        self.buttons_frame.pack(side=BOTTOM, fill=BOTH)

        self.buttons_frame.grid_columnconfigure(0, weight=1)
        
        # Use lambda to call send_mail() when the button is clicked
        self.v_btn = customtkinter.CTkButton(self.buttons_frame,
                                             text="Toplu Mail gönder", height=40,
                                             command=lambda: threading.Thread(target=self.send_mail, daemon=True).start(),
                                             corner_radius=0)
        self.v_btn.grid(pady=(10, 0), padx=(5, 5), row=0, column=0, sticky="we")

    def show_login_popup(self):
        """Kullanıcı girişini almak için popup pencere."""
        self.login_popup = customtkinter.CTkToplevel(self)
        self.login_popup.title("Giriş")
        self.login_popup.geometry("300x250")
        self.login_popup.resizable(False, False)

        # Kullanıcı adı
        username_label = customtkinter.CTkLabel(self.login_popup, text="E-posta:")
        username_label.pack(pady=(20, 5))
        self.username_entry = customtkinter.CTkEntry(self.login_popup, placeholder_text="E-posta adresinizi girin")
        self.username_entry.pack(padx=20, pady=5)

        # Şifre
        password_label = customtkinter.CTkLabel(self.login_popup, text="Şifre:")
        password_label.pack(pady=(10, 5))
        self.password_entry = customtkinter.CTkEntry(self.login_popup, placeholder_text="Şifrenizi girin", show="*")
        self.password_entry.pack(padx=20, pady=5)

        login_button = customtkinter.CTkButton(self.login_popup, text="Giriş Yap", command=self.save_login_credentials)
        login_button.pack(pady=(10, 20))  

    def save_login_credentials(self):
        """Popup'tan kullanıcı adı ve şifreyi al."""
        self.username = self.username_entry.get().strip()
        self.password = self.password_entry.get().strip()
        if self.username and self.password:
            self.login_popup.destroy()  
        else:
            print("Lütfen tüm alanları doldurun!") 




    def adjust_textbox_height(self, event=None):
        """Metin kutusunun yüksekliğini içeriğe göre ayarlar."""
        num_lines = int(self.textBox.index("end-1c").split('.')[0])  
        line_height = 20  
        new_height = max(50, num_lines * line_height)  
        self.textBox.configure(height=new_height)  


    def add_attachment(self):
        # Kullanıcıdan dosya seçmesini isteyin
        file_path = askopenfilename(title="Dosya Seç", 
                                    filetypes=[("Tüm Dosyalar", "*.*"), 
                                               ("PDF Dosyaları", "*.pdf"),
                                               ("Resim Dosyaları", "*.jpg;*.png")])
        if file_path:
            self.attachment_path = file_path
            print(f"Eklenti seçildi: {file_path}")  # to control
            self.v_lable = customtkinter.CTkLabel(self.l_frame, text=f"Eklenti:{file_path}", width=15, height=15)
            self.v_lable.pack(padx=5, pady=5)

    def get_mails_from_exel(self):
        
        data=pd.read_excel("a.xlsx")
        
        first_column=data.iloc[:,0]
        first_column=first_column.tolist()
        
        return first_column
    def get_input(self):
        input_text = self.textBox.get("1.0", 'end-1c').strip()  #clear spaces
        if input_text == self.placeholder_text: 
            return ""  
        return input_text

    def start(self):
        self.mainloop()

    def add_placeholder(self):
        # Insert placeholder text
        self.textBox.insert("1.0", self.placeholder_text)
        self.textBox.configure(fg_color="white")  # Optional: Change text color for placeholder

        # Bind events to handle focus and blur
        self.textBox.bind("<FocusIn>", self.remove_placeholder)
        self.textBox.bind("<FocusOut>", self.add_placeholder_on_blur)

    def remove_placeholder(self, event):
        if self.textBox.get("1.0", "end-1c") == self.placeholder_text:
            self.textBox.delete("1.0", "end")
            self.textBox.configure(fg_color="white")  # Optional: Change text color for normal text

    def add_placeholder_on_blur(self, event):
        if not self.textBox.get("1.0", "end-1c"):
            self.textBox.insert("1.0", self.placeholder_text)
            self.textBox.configure(fg_color="lightgray")


    def send_mail(self):
        username=self.username
        
        password=self.password
        sender = username
        targets = self.get_mails_from_exel()

        message_content = self.get_input() if self.get_input() else "Mail içeriği boş"

        # creating multiple mail
        
        msg = MIMEMultipart()
        msg['Subject'] = 'Topic'
        msg['From'] = sender
        # add attachment
        if self.attachment_path:
            try:
                with open(self.attachment_path, "rb") as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                encoders.encode_base64(part)  # Base64 encoding

                # define attachment name
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename={self.attachment_path.split('/')[-1]}",
                )
                msg.attach(part)
                print("Eklenti başarıyla eklendi.")
            except Exception as e:
                print(f"Eklenti eklenirken hata oluştu: {e}")

        # sending mail operation
        try:
            server = smtplib.SMTP_SSL(smtp_ssl_host, smtp_ssl_port)
            server.login(username, password)
            for target in targets:
                msg['To'] = ', '.join(target)
                msg.attach(MIMEText(message_content, 'plain'))
                server.sendmail(sender, target, msg.as_string())
            print("Email sent successfully!")
        except Exception as e:
            print(f"Failed to send email: {e}")
        finally:
            server.quit()


if __name__ == "__main__":
    app = App()
    app.start()
    print('Finished')
