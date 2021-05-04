from tkinter import *
from tkinter import ttk
from ttkthemes import ThemedTk
from tkinter import messagebox
from tkinter import filedialog
import smtplib
from email.message import EmailMessage
from threading import Thread

def show_hide():
    if btn_switch['text'] == 'show':
        btn_switch['text'] = 'hide'
        entryPass['show'] = ''
    else:
        btn_switch['text'] = 'show'
        entryPass['show'] = '*'

def clear(event=None):
    global files
    password.set('')
    subject.set('')
    entryBody.delete('1.0', END)
    listbox1.delete(0, END)
    files = []

def thread_attach():
    Thread(target=attachFile).start()

def attachFile():
    global files
    selFiles = list(filedialog.askopenfilenames(initialdir='./', title='Select files',
                                                filetypes=(('All files', '*.*'), )))
    for file in selFiles:
        files.append(file)

    listbox1.delete(0, END)
    for file in files:
        f_name = file.split('/')[-1]
        listbox1.insert(END, f_name)

def thread_send():
    Thread(target=sendMail).start()

def sendMail():
    global files
    fromAddr = sender.get()
    passkey = password.get()
    toAddr = receiver.get()
    cc = CC.get()
    bcc = BCC.get()
    subject_of_mail = subject.get()
    body = entryBody.get('1.0', 'end-1c')

    conditions = [fromAddr != '',
                  passkey != '',
                  toAddr != '']

    try:
        if all(conditions):
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(fromAddr, passkey)

            email = EmailMessage()
            email['From'] = fromAddr
            email['To'] = toAddr
            try:
                if cc != '':
                    email['Cc'] = cc
                if bcc != '':
                    email['Bcc'] = bcc
            except:
                messagebox.showwarning('Warning',
                                       'Cc or Bcc not specified properly')
            email['Subject'] = subject_of_mail
            email.set_content(body)

            for file in files:
                with open(file, 'rb') as f:
                    f_data = f.read()
                    f_name = f.name
                    f_name = f_name[::-1]
                    f_name = f_name[:f_name.index('/')]
                    f_name = f_name[::-1]
                try:
                    email.add_attachment(f_data, maintype='image', subtype=file[file.index('.') + 1:],
                                         filename=f_name)
                except:
                    email.add_attachment(f_data, maintype='application', subtype=file[file.index('.') + 1:],
                                         filename=f_name)

                f.close()

            win = messagebox.askyesno('Confirmation',
                                      'Are you sure you want to send this mail?')
            if win > 0:
                server.send_message(email)
                messagebox.showinfo('Success',
                                    'Mail has been sent successfully')
                server.quit()
                del email
                files = []
                clear()
            else:
                listbox1.delete(0, END)
        else:
            messagebox.showerror('Error',
                                 'Something has not been specified properly')
    except smtplib.SMTPAuthenticationError:
        messagebox.showerror('Error',
                             'This error may be occured due to following reasons:' + '\n' +
                             '\t1. Password does not match' + '\n' +
                             '\t2. You may not have allowed less secure apps for\n\t   your google account')
    except smtplib.SMTPSenderRefused:
        messagebox.showerror('Error',
                             'The size of attached files has been exceeded the limit')


root = ThemedTk(theme='arc')
root.title('E-Mail Messenger')
root.geometry('800x600+230+70')
root.iconbitmap('./res/icon.ico')
root.config(bg='white')
root.minsize(800, 600)

# Variables
sender = StringVar()
password = StringVar()
receiver = StringVar()
subject = StringVar()
CC = StringVar()
BCC = StringVar()
files = []

heading = Label(root, text='E-Mail Messenger', bg='white',
                font=('arial', 20, 'bold'))
heading.pack(fill='x', padx=50, pady=20)

mainFrame = Frame(root, bg='white')
mainFrame.pack(expand='yes', padx=50, pady=10)

lblFrom = Label(mainFrame, text='From:', anchor='ne', bg='white', font=('arial', 10, 'bold'))
lblFrom.grid(row=0, column=0, pady=5, padx=10, sticky='e')

entryFrom = ttk.Entry(mainFrame, textvar=sender, width=30, font=('arial', 10, 'bold'))
entryFrom.grid(row=0, column=1, pady=5, padx=10, sticky='w')

lblPass = Label(mainFrame, text='Password:', bg='white', font=('arial', 10, 'bold'))
lblPass.grid(row=0, column=2, pady=5, padx=10, sticky='e')

entryPass = ttk.Entry(mainFrame, textvar=password, show='*', width=23, font=('calibri', 10, 'bold'))
entryPass.grid(row=0, column=3, pady=5, padx=10, sticky='w')

btn_switch = ttk.Button(mainFrame, text='show', width=5, command=show_hide)
btn_switch.grid(row=0, column=3, pady=5, sticky='e')

lblTo = Label(mainFrame, text='To:', bg='white', font=('arial', 10, 'bold'))
lblTo.grid(row=1, column=0, pady=5, padx=10, sticky='e')

entryTo = ttk.Entry(mainFrame, textvar=receiver, width=30, font=('arial', 10, 'bold'))
entryTo.grid(row=1, column=1, pady=5, padx=10, sticky='w')

lblCc = Label(mainFrame, text='Cc:', bg='white', font=('arial', 10, 'bold'))
lblCc.grid(row=1, column=2, pady=5, padx=10, sticky='e')

entryCc = ttk.Entry(mainFrame, textvar=CC, width=30, font=('arial', 10, 'bold'))
entryCc.grid(row=1, column=3, pady=5, padx=10, sticky='w')

lblBcc = Label(mainFrame, text='Bcc:', bg='white', font=('arial', 10, 'bold'))
lblBcc.grid(row=2, column=2, pady=5, padx=10, sticky='e')

entryBcc = ttk.Entry(mainFrame, textvar=BCC, width=30, font=('arial', 10, 'bold'))
entryBcc.grid(row=2, column=3, pady=5, padx=10, sticky='w')

lblSub = Label(mainFrame, text='Subject:', bg='white', font=('arial', 10, 'bold'))
lblSub.grid(row=2, column=0, pady=5, padx=10, sticky='e')

entrySub = ttk.Entry(mainFrame, textvar=subject, width=30, font=('arial', 10, 'bold'))
entrySub.grid(row=2, column=1, pady=5, padx=10, sticky='w')

lblBody = Label(mainFrame, text='Content:', bg='white', font=('arial', 10, 'bold'))
lblBody.grid(row=3, column=0, pady=5, padx=10, sticky='ne')

bodyFrame = Frame(mainFrame, bg='white')
bodyFrame.grid(row=3, column=1, columnspan=3, pady=5, padx=10, sticky='w')

entryBody = Text(bodyFrame, width=70, height=10, borderwidth=1, font=('calibri', 12))
entryBody.pack(side=LEFT, fill='both')
yScroll = ttk.Scrollbar(bodyFrame, orient='vertical', command=entryBody.yview)
yScroll.pack(side=RIGHT, fill='y')
entryBody.configure(yscroll=yScroll.set)

btnAdd = ttk.Button(mainFrame, text='Add files', command=thread_attach)
btnAdd.grid(row=5, column=0, pady=10)

listFrame1 = Frame(mainFrame, bg='white')
listFrame1.grid(row=5, column=1, columnspan=2, pady=5, padx=10, sticky='w')
listFrame = Frame(listFrame1, bg='white')
listFrame.pack()

listbox1 = Listbox(listFrame, height=5, width=40)
listbox1.pack(side=LEFT, fill='both')
yScroll1 = ttk.Scrollbar(listFrame, orient='vertical', command=listbox1.yview)
yScroll1.pack(side=RIGHT, fill='y')
listbox1.configure(yscroll=yScroll1.set)
xScroll1 = ttk.Scrollbar(listFrame1, orient='horizontal', command=listbox1.xview)
xScroll1.pack(fill='x')
listbox1.configure(xscroll=xScroll1.set)

btnFrame = Frame(root, bg='white')
btnFrame.pack(side=BOTTOM, padx=10, pady=25)

btnSend = ttk.Button(btnFrame, text='Clear', command=clear)
btnSend.grid(row=0, column=0, padx=10)

btnSend = ttk.Button(btnFrame, text='Send', command=thread_send)
btnSend.grid(row=0, column=1, padx=10)

btnQuit = ttk.Button(btnFrame, text='Exit', command=root.destroy)
btnQuit.grid(row=0, column=2, padx=10)

root.bind('<Escape>', clear)
root.mainloop()
