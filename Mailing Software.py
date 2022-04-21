# Mailing Software source code
# Importing the required modules
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase
from email import encoders 
import smtplib
import webbrowser
import re
import os.path

# Created path variable to find the location of the file
path = os.path.dirname(__file__)
if '\\' in path:
    path = path.replace('\\','/')
else:
    pass
# Created a global files variable of type list which is empty initially
global files
files = []
# Created to check whether user is logged in or logged out
global check
# setting check to false indicating that user has not logged in
check = False

global email
email = None

# Create refresh_login function to delete the login page entry box content
def refresh_login():
    email_entry.delete(0,END)
    password_entry.delete(0,END)
    
# Created forgot function to deal with forgot password
def forgot(event):
    webbrowser.open_new("https://accounts.google.com/")

# Created setup function to enable less secure apps access
def setup(event):
    webbrowser.open_new("https://www.google.com/settings/security/lesssecureapps")

# Used to login and send an email message
def login():
    # Stores the senders's email id
    sender = str(account.get())
    global check
    # Check if email id , password is entered and user is logged out
    if account.get() and pswrd.get()and check == False:
        # Run the try block 
        try:
            # Setting check variable to true to show user has logged in
            check = True
            # Stores the senders' password
            password = str(pswrd.get())
            # Creating an SMTP Session with name email
            global email
            email = smtplib.SMTP('smtp.gmail.com',587)
            # Starting the TLS for security reasons
            email.starttls()
            # Authentication
            email.login(sender,password)
        # Authentication exception
        except smtplib.SMTPAuthenticationError:
            error = messagebox.showerror("Error","Authentication Error")
            return None
    # Indicate that user has not filled  email id and password
    elif (not(account.get())and not(pswrd.get())):
        warning = messagebox.showwarning("Warning","Enter Email Id And Password")
        return None
    elif (not(account.get())):
        warning = messagebox.showwarning("Warning","Enter Email Id")
        return None
    elif (not(pswrd.get())):
        warning = messagebox.showwarning("Warning","Enter Password")
        return None
    else:
        pass
    
    # select_file function is used for selecting a file that is to be attached with an email message    
    def select_file():
        # Getting the location of file to be attached
        file_loc = filedialog.askopenfilename(initialdir = path,title = "Select A File",filetypes = (("All files","*.*"),))
        # Checking if a file_loc contains a location or not
        if file_loc:
            item = []
            loc = file_loc.split('/')
            # Get the name of the file to be attached
            file_name = str(loc[-1])
            # Append file name and file_loc to item
            item.append(file_name)
            item.append(file_loc)
            # Append item to the files which contain all files that are to be attached with email message
            global files
            files.append(item)
            # Add the selected file to the attach_box
            def add_file():
                # Insert the name of file in the attach box
                attach_box.insert(END,file_name)
            # Calling the add_file function
            add_file()
            # Assingning the anchor with index 0 and making it as default selection
            attach_box.selection_anchor(0)
            attach_box.selection_set(ANCHOR)
    
    # select_files function used for selecting files that are to be attached with email message    
    def select_files():
        # Getting the location of all files that are to be attached
        file_locs = filedialog.askopenfilenames(initialdir = path,title = "Select The Files",filetypes = (("All files","*.*"),))
        # Checking if a file_locs contains any location or not
        if file_locs:
            # Run for loop for each value of file location
            for file_loc in file_locs:
                item = []
                loc = file_loc.split('/')
                # Get the name of the file to be attached
                file_name = str(loc[-1])
                # Append file name and file_loc to item
                item.append(file_name)
                item.append(file_loc)
                # Append item to the files which contain all files that are to be attached with email message
                global files
                files.append(item)
                # Add the selected file to the attach_box
                def add_file():
                    # Insert the name of file in the attach box
                    attach_box.insert(END,file_name)
                # Calling the add_file function
                add_file()
            # Assingning the anchor with index 0 and making it as default selection
            attach_box.selection_anchor(0)
            attach_box.selection_set(ANCHOR)
        
    # remove_file function is used to remove file that are attached to the email message
    def remove_file():
        global files
        # Used to remove the file from the attach_box
        def delete():
            # Use to get the line number of the selected element in a tuple form
            item = attach_box.curselection()
            # Check if item contains a non-empty tuple
            if item:
                # Delete the selected element from attach box
                attach_box.delete(ANCHOR)
                # Delete the file of the selected element from attachments
                del files[item[0]]
                # Assigning the anchor with index of previous anchor and setting it as a default selection
                attach_box.selection_anchor(ANCHOR)
                attach_box.selection_set(ANCHOR)
            # Run if item contains an empty tuple
            else:
                # Assingning the anchor with index 0 and making it as default selection
                attach_box.selection_anchor(0)
                attach_box.selection_set(ANCHOR)
                # Delete the selected element from attach box
                attach_box.delete(ANCHOR)
                # Delete the file of the selected element from attachments
                del files[ANCHOR]
                # Assigning the anchor with index of previous anchor and setting it as a default selection
                attach_box.selection_anchor(ANCHOR)
                attach_box.selection_set(ANCHOR)
        # Calling the delete file function
        delete()

    # remove_all_files function is used to remove all files attached to an email message
    def remove_all_files():
        global files
        # Used to remove the file from the attach_box
        attach_box.delete(0,END)
        # Used to empty the list of files so that all the attached files get removed from the email attachments
        del files[:]
        
    # Used to refresh the contents of top  window
    def refresh_email():
        global files
        del files[:]
        # Destroy the top
        top.destroy()
        # Calling the login function
        login()
        
    # Used to log out of the top window
    def logout():
        # Quit the SMTP session
        global email
        email.quit()
        # Setting the check variable to false to represent that user has logged out
        global check
        check = False
        # Deleting the files
        global files
        del files[:]
        # View the root window
        root.deiconify()
        # Destroy the top
        top.destroy()
        # Call the refresh login function
        refresh_login()
        
    # Create sendemail function for performint the functionality of sending an email message with attachments
    def sendemail():
        # Run the try block
        try:
            # Storing the recepient email address
            recepient = str(receiver.get())
            # Stroring the subject
            sub = str(subject.get())
            # Storing the text part of the message body
            body = str(message.get('1.0','end'))
            # Used for sending multiple email of type cc or bcc or both
            Bcc = str(bcc.get())
            Cc = str(cc.get())
            # Create an empty list for storing bcc, cc and recepient email addresses
            Bcc_list = []
            Cc_list = []
            # Created an empty string for storing bcc, cc and recepient email addresses
            global bcc_str
            bcc_str = ""
            global cc_str
            cc_str = ""
            global recepient_str
            recepient_str = ""
            global verify_cc
            verify_cc = False
            global verify_bcc
            verify_bcc = False
            print(1)
            # Function to verify an email address on the basis of it's name
            def verify_email():
                # Regular expression for name of an email address
                regex = "[\w._%+-]{1,20}@[\w.-]{2,20}.[A-Za-z]{2,3}"
                # Created an empty string for storing bcc, cc and recepient email addresses
                global bcc_str
                bcc_str = ""
                global cc_str
                cc_str = ""
                global recepient_str
                recepient_str = ""
                global verify_cc
                global verify_bcc
                # Run the try block
                try:
                    # Check if recepient is empty or not
                    if recepient:
                        test = recepient + " "
                        email = test.split()
                        # Check the for valid email id's and add them to recepient_str and recepient_list
                        if re.search(regex,email[0]):
                            recepient_str = str(email[0])

                    # Check if Cc is empty or not              
                    if Cc:
                        test = Cc + " "
                        # Check whether email id's are seperated by spaces
                        if (" " in test and "," in test) or "," in test:
                            messagebox.showerror("Error in Cc","Use Spaces To Seperate Email Id's",parent = mainframe1)
                        # Run only when email id's are seperated by spaces
                        else:
                            # Make a list of unique email id's named as email_list
                            Cc_ = Cc.split()
                            email_list = list(set(Cc_))
                            # Check the for valid email id's and add them to cc_string and Cc_list
                            for email in email_list:
                                if re.search(regex,email):
                                    Cc_list.append(email)
                                    cc_str += str(email) + ","
                            verify_cc = True
                        
                    # Check if Bcc is empty or not
                    if Bcc:
                        test = Bcc + " "
                        # Check whether email id's are seperated by spaces
                        if (" " in test and "," in test) or "," in test:
                            messagebox.showerror("Error in Bcc","Use Only Spaces To Seperate Email Id's",parent = mainframe1)
                        # Run only when email id's are seperated by spaces
                        else:
                            # Make a list of unique email id's named as email_list
                            Bcc_ = Bcc.split()
                            email_list = list(set(Bcc_))
                            # Check the for valid email id's and add them to bcc_string and Bcc_list
                            for email in email_list:
                                if re.search(regex,email):
                                    Bcc_list.append(email)
                                    bcc_str += str(email)+","
                            verify_bcc = True
                # Just pass in case of any exception
                except:
                    messagebox.showerror("Error","Something Went Wrong1",parent = mainframe1)
                    return None
            # Calling the verify_email function and assigning it's return value to stop
            verify_email()
            
            if Cc and verify_cc == True:
                verify_cc = False
            elif not(Cc) and verify_cc == False:
                pass
            else:
                return None
            
            if Bcc and verify_bcc == True:
                verify_bcc = False
            elif not(Bcc) and verify_bcc == False:
                pass
            else:
                return None
            print(2)
            # Storing the receivers's email address in list format
            toaddrs = [recepient_str] + Cc_list + Bcc_list
            # Create an instance of MIMEMultipart named it as msg
            msg = MIMEMultipart()
            # Storing the required msg parameters in dictionary format
            msg['From'] = sender
            msg['To'] = recepient_str
            if Cc:
                msg['Cc'] = cc_str.rstrip(",")
            if Bcc:
                msg['Bcc'] = bcc_str.rstrip(",")
            msg['Subject'] = sub
            # Attact body part to the msg instance
            msg.attach(MIMEText(body,"plain"))
            print(3)
            # Used to attach files to the msg instance
            for item in files:
                # Store the name of file attached
                filename = item[0]
                # Open the file to be sent
                attachment = open(item[1],'rb')
                # Create an instance of MIMEBase and name it as part
                part = MIMEBase('application', 'octet-stream')
                # Used to set the payload of the part with the attacged file
                part.set_payload(attachment.read())
                # Encode part in base64 format
                encoders.encode_base64(part)
                # Add header to the part
                part.add_header("Content-Disposition", f"attachment; filename = {filename}",) 
                # Attach part to the msg instance
                msg.attach(part)
            # Convert the Multipart msg into a string
            text = msg.as_string()
            # Sending the email
            email.sendmail(sender,toaddrs,text)
            # Label to confirm the sending of email for the user
            send = messagebox.showinfo("Success","Email Send Successfuly",parent = mainframe1)
        # Server Discconected exception
            print(4)
        except smtplib.SMTPServerDisconnected:
            error = messagebox.showerror("Error","Server Disconnected",parent = mainframe1)
            return None
        #For any unexpected error
        except:
            error = messagebox.showerror("Error","Something went wrong2",parent = mainframe1)
            return None
 
    # Create a top level window and name it as top  
    top = Toplevel()
    # Title of top
    top.title("Mailing Software")
    # Used to adjust geometry of top 
    windowWidth = 600 
    windowHeight = 650
    positionRight = int(root.winfo_screenwidth()/2 - windowWidth/2)
    positionDown = int(root.winfo_screenheight()/2 - windowHeight/2)
    top.geometry("+{}+{}".format(positionRight, positionDown))
    # Fixing the size of top
    top.resizable(0,0)
    # Creating a frame for the top
    mainframe1 = LabelFrame(top,text = "COMPOSE EMAIL", padx = 10, pady = 10,bg = '#ff9a00', fg = "black",font = ('arial',10,'bold','italic'))
    mainframe1.pack(pady = 10)
    # Hiding the root window
    root.withdraw()
    # Creating  string variables to store the basic information used in sending an email
    sender_var = StringVar()
    sender_var.set(str(account.get()))
    receiver = StringVar()
    subject = StringVar()
    message = StringVar()
    cc = StringVar()
    bcc = StringVar()
    # GUI for composing our email message
    # Creating sender's label and entry widget
    sender_label = Label(mainframe1,text = "FROM :",bg = '#ff9a00', fg = "black",font = ('arial',10,'bold','italic')).grid(row = 1,column = 0, sticky = W, padx = 5)
    sender_entry = Entry(mainframe1, width = 50, textvariable = sender_var,state = DISABLED, font = ('arial',10,'bold','italic'), borderwidth = 3).grid(row = 1,column = 1, sticky = W, padx = 5, pady = 10)
    # Creating receiver's label and entry widget
    receiver_label = Label(mainframe1,text = "TO :",bg = '#ff9a00', fg = "black",font = ('arial',10,'bold','italic')).grid(row = 2,column = 0, sticky = W, padx = 5)
    receiver_entry = Entry(mainframe1, width = 50, textvariable = receiver,font = ('arial',10,'bold','italic'), borderwidth = 3).grid(row = 2,column = 1, sticky = W, padx = 5, pady = 10)
    # Creating label and entry widget used incase of sending cc and bcc enabled email messages
    cc_label = Label(mainframe1,text = "CC :",bg = '#ff9a00', fg = "black",font = ('arial',10,'bold','italic')).grid(row = 3,column = 0, sticky = W, padx = 5)
    cc_entry = Entry(mainframe1, width = 50, textvariable = cc,font = ('arial',10,'bold','italic'), borderwidth = 3).grid(row = 3,column = 1, sticky = W, padx = 5, pady = 10)   
    bcc_label = Label(mainframe1,text = "BCC :",bg = '#ff9a00', fg = "black",font = ('arial',10,'bold','italic')).grid(row = 4,column = 0, sticky = W, padx = 5)
    bcc_entry = Entry(mainframe1, width = 50, textvariable = bcc,font = ('arial',10,'bold','italic'), borderwidth = 3).grid(row = 4,column = 1, sticky = W, padx = 5, pady = 10)
    # Creating subject label and entry widget
    subject_label = Label(mainframe1,text = "SUBJECT :",bg = '#ff9a00', fg = "black",font = ('arial',10,'bold','italic')).grid(row = 5,column = 0, sticky = W, padx = 5)
    subject_entry = Entry(mainframe1, width = 50, textvariable = subject,font = ('arial',10,'bold','italic'), borderwidth = 3).grid(row = 5,column = 1, sticky = W, padx = 5, pady = 10)
    # Creating message label and textbox widget
    message_label = Label(mainframe1,text = "MESSAGE :",bg = '#ff9a00', fg = "black",font = ('arial',10,'bold','italic')).grid(row = 6,column = 0, sticky = W, padx = 5)
    message = Text(mainframe1, width = 50, height = 8,font = ('arial',10,'bold','italic'), borderwidth = 3)
    message.grid(row = 6,column = 1, sticky = W, padx = 5, pady = 10)
    # Creating label and listbox for attachment purposes
    # Creating the attach box label
    attach_label = Label(mainframe1,text = "ATTACHMENTS : ",bg = '#ff9a00', fg = "black",font = ('arial',10,'bold','italic')).grid(row = 7,column = 0, sticky = W, padx = 5)
    # Creating the attach box
    attach_box = Listbox(mainframe1, width = 52, height = 5, font = ('arial',10,'bold','italic'), borderwidth = 3)
    attach_box.grid(row = 7,column = 1, sticky = W,columnspan = 3)
    # Creating the scrollbar
    scrollbar = Scrollbar(mainframe1)
    scrollbar.grid(row = 7,column = 2, sticky = W+N+S)
    # Configuring attach box with scrollbar and setting exportselection to false
    attach_box.config(yscrollcommand = scrollbar.set,exportselection = False)
    scrollbar.config(command = attach_box.yview)
    # Create Add File Menu
    email_menu = Menu(top)
    top.config(menu = email_menu)
    # Add File Menu
    add_file_menu = Menu(email_menu)
    email_menu.add_cascade(label = "Add Files", menu = add_file_menu)
    add_file_menu.add_command(label = "Add One File To Attachments", command = select_file)
    # Add Many Files to attach-box
    add_file_menu.add_command(label = "Add Many File To Attachments", command = select_files)
    # Create Delete File Menu
    remove_file_menu = Menu(email_menu)
    email_menu.add_cascade(label = "Remove Files", menu = remove_file_menu)
    remove_file_menu.add_command(label = "Remove A File From Attachments", command = remove_file)
    remove_file_menu.add_command(label = "Remove All File From Attachments", command = remove_all_files)
    # Create Developers Menu
    developer_menu = Menu(email_menu)
    email_menu.add_cascade(label = "Developer",menu = developer_menu)
    developer_menu.add_command(label = "Aman Negi")
    developer_menu.add_command(label = "CSE - T1")
    developer_menu.add_command(label = "00615602718")
    
    # Creating all the basic buttons of the top
    send_ = Button(mainframe1, text = "Send Email",command = sendemail, padx = 11, bg = "#ff2e63",font = ('arial',10,'bold'), fg = "white").grid(row = 8,column = 1,sticky = E, padx = 10, pady = 5)
    refresh_ = Button(mainframe1, text = "Refresh",command = refresh_email, padx = 23, bg = "#ff2e63",font = ('arial',10,'bold'), fg = "white").grid(row = 9, column = 1,sticky = E, padx = 10, pady = 5)
    logout_ = Button(mainframe1, text = "Log Out", command = logout,  padx = 23, bg = "#ff2e63",font = ('arial',10,'bold'), fg = "white").grid(row = 10, column = 1,sticky = E, padx = 10, pady = 5)


# Creating a root window
root = Tk()
# Title of root window
root.title("Mailing Software")
# Used to adjust geometry of root window
windowWidth = 300 
windowHeight = 300
positionRight = int(root.winfo_screenwidth()/2 - windowWidth/2)
positionDown = int(root.winfo_screenheight()/2 - windowHeight/2)
root.geometry("+{}+{}".format(positionRight, positionDown))
# Fix the size of root window
root.resizable(0,0)
# Creating a frame for root window
mainframe = LabelFrame(root, text = "Mailing Software", padx = 10, pady = 10,bg = '#ff9a00', font = ('arial',11,'bold'), fg = "#222831")
mainframe.pack()
# Creating variables to store account and password of sender's
account = StringVar()
global pswrd
pswrd = StringVar()
# Defining  all the labels of the root window
email_id = Label(mainframe,text = "Email Id :",bg = '#ff9a00', fg = "black",font = ('arial',12,'bold','italic'))
password = Label(mainframe,text = "Password :",bg = '#ff9a00', fg = "black",font = ('arial',12,'bold','italic'))
f_password = Label(mainframe, text = "Forgot Password?",fg = "blue",cursor = "hand2", bg = "white",font = ('arial',10,'bold','italic','underline'))
secure_apps = Label(mainframe, text = "Turn this setting on before login",cursor = "hand2", bg = "white",font = ('arial',10,'bold',"italic","underline"), fg = "blue")
# Creating entry widget to get sender's email id and password for login
email_entry = Entry(mainframe, width = 35,font = ('arial',15,'bold','italic'), textvariable = account, borderwidth = 3)
password_entry = Entry(mainframe,width = 35,font = ('arial',15,'bold'), show = "*",fg = "#222831", textvariable = pswrd, borderwidth = 3)
password_entry.grid(row = 3, column = 0, pady = 10, ipady = 3)
# Define password display button images
global show
show = PhotoImage(file = 'images/show.png')
global hide
hide = PhotoImage(file = 'images/hide.png')
# Function to view/hide the password
def view():
    global pswrd
    # Used to hide the password
    def hide_():
        # Display the show image on button
        show.image = PhotoImage(file = 'images/show.png')
        if pswrd.get():
            # Allow to view and enter password in encrypted form
            password_entry = Entry(mainframe,width = 35,font = ('arial',15,'bold'), textvariable = pswrd, show = "*", fg = "#222831", borderwidth = 3)
            password_entry.grid(row = 3, column = 0, pady = 10, ipady = 3)     
        else:
            # Allow to view and enter password in encrypted form
            password_entry = Entry(mainframe,width = 35,font = ('arial',15,'bold'),show = "*" ,fg = "#222831", textvariable = pswrd, borderwidth = 3)
            password_entry.grid(row = 3, column = 0, pady = 10, ipady = 3)
        password_view = Button(mainframe,image = show,command = show_, borderwidth = 3).grid(row = 3,column = 0,sticky = E)
    # Used to show the password
    def show_():
        # Display the hide image on button
        hide.image = PhotoImage(file = 'images/hide.png')
        if pswrd.get():
            # Allow to view and enter password in unencrypted form
            password_entry = Entry(mainframe,width = 35,font = ('arial',15,'bold'), textvariable = pswrd, fg = "#222831", borderwidth = 3)
            password_entry.grid(row = 3, column = 0, pady = 10, ipady = 3)  
        else:
            # Allow to view and enter password in unencrypted form
            password_entry = Entry(mainframe,width = 35,font = ('arial',15,'bold'),fg = "#222831", textvariable = pswrd, borderwidth = 3)
            password_entry.grid(row = 3, column = 0, pady = 10, ipady = 3)
        password_view = Button(mainframe, image = hide , command = hide_, borderwidth = 3).grid(row = 3,column = 0,sticky = E)   
    show_()
# Button to hide or show the  sender's password 
password_view = Button(mainframe,image = show, command = view, borderwidth = 3).grid(row = 3,column = 0,sticky = E)
# Button to login into the sender's account
log_in = Button(mainframe, text = "Log In", bg = "#ff2e63",font = ('arial',10,'bold'), fg = "white", command = login)

# Displaying all the defined widgets on the root window
email_id.grid(row = 0, column = 0,sticky = W)
email_entry.grid(row = 1, column = 0, pady = 10, ipady = 3)
password.grid(row = 2, column = 0, sticky = W)
log_in.grid(row = 4,column = 0, ipadx = 60, ipady = 5, pady = 10)
f_password.grid(row = 5, column = 0, pady = 15, sticky = W)
secure_apps.grid(row = 5, column = 0,pady = 13, sticky = E)

# Using the binding function to bind f_password with forgot function and secure_apps with setup function
f_password.bind("<Button-1>",forgot)
secure_apps.bind("<Button-1>",setup)
# Closing the mainloop
mainloop()
