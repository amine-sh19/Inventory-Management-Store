from tkinter import *
import pandas as pd
from tkinter import messagebox
import tkinter.messagebox as tkMessageBox
import sqlite3
import sqlite3 as sql
import os
import csv
from datetime import date
import xlsxwriter
from xlsxwriter.workbook import Workbook
from PIL import Image, ImageTk
import tkinter.ttk as ttk
#AMINE SHABI#
root = Tk()
root.title("Gestion Des Stocks")

#Email: aminesh@inbox.lv #

canvas=Canvas(root,width=950,height=500)
image=ImageTk.PhotoImage(Image.open("am.jpg"))
canvas.create_image(0,0,anchor=NW,image=image)
canvas.pack()
width = 1100
height = 600
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x = (screen_width/2) - (width/2)
y = (screen_height/2) - (height/2)
root.geometry("%dx%d+%d+%d" % (width, height, x, y))
root.resizable(0, 0)
root.config(bg="brown")

# VARIABLES #
USERNAME = StringVar()
PASSWORD = StringVar()
PRODUCT_NAME = StringVar()
PRODUCT_PRICE = IntVar()
PRODUCT_QTY = IntVar()
SEARCH = StringVar()

# METHODS #

def Database():
    global conn, cursor
    conn = sqlite3.connect("mydata.db")
    cursor = conn.cursor()
    cursor.execute("CREATE TABLE IF NOT EXISTS `amine` (admin_id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, username TEXT, password TEXT)")
    cursor.execute("CREATE TABLE IF NOT EXISTS `product` (product_id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, product_name TEXT, product_qty TEXT, product_price TEXT)")
    cursor.execute("SELECT * FROM `amine` WHERE `username` = 'amine' AND `password` = 'aminesh00'")
    if cursor.fetchone() is None:
        cursor.execute("INSERT INTO `amine` (username, password) VALUES('amine', 'aminesh00')")
        conn.commit()

def Exit():
    result = tkMessageBox.askquestion('Gestion Des Stocks', 'Êtes-vous sûr de vouloir quitter ?', icon="warning")
    if result == 'yes':
        root.destroy()
        exit()

def Exit2():
    result = tkMessageBox.askquestion('Gestion Des Stocks', 'Êtes-vous sûr de vouloir quitter ?', icon="warning")
    if result == 'yes':
        Home.destroy()
        exit()

def ShowLoginForm():
    global loginform
    loginform = Toplevel()
    loginform.title("Gestion Des Stocks - Account Login")
    width = 700
    height = 600
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width/2) - (width/2)
    y = (screen_height/2) - (height/2)
    loginform.resizable(0, 0)
    loginform.geometry("%dx%d+%d+%d" % (width, height, x, y))
    LoginForm()
    
def LoginForm():
    global lbl_result
    TopLoginForm = Frame(loginform, width=600, height=100, bd=1, relief=SOLID)
    TopLoginForm.pack(side=TOP, pady=20)
    lbl_text = Label(TopLoginForm, text="Connectez-vous ici", font=('Helvetica', 18), width=600)
    lbl_text.pack(fill=X)
    MidLoginForm = Frame(loginform, width=600)
    MidLoginForm.pack(side=TOP, pady=50)
    lbl_username = Label(MidLoginForm, text="Nom d'utilisateur:", font=('Helvetica', 25), bd=18)
    lbl_username.grid(row=0)
    lbl_password = Label(MidLoginForm, text="Mot de passe:", font=('Helvetica', 25), bd=18)
    lbl_password.grid(row=1)
    lbl_result = Label(MidLoginForm, text="", font=('Helvetica', 18))
    lbl_result.grid(row=3, columnspan=2)
    username = Entry(MidLoginForm, textvariable=USERNAME, font=('Helvetica', 25), width=15)
    username.grid(row=0, column=1)
    password = Entry(MidLoginForm, textvariable=PASSWORD, font=('Helvetica', 25), width=15, show="*")
    password.grid(row=1, column=1)
    btn_login = Button(MidLoginForm, text="Se Connecter", font=('Helvetica', 18), width=30, command=Login)
    btn_login.grid(row=2, columnspan=2, pady=20)
    btn_login.bind('<Return>', Login)
    
def Home():
    global Home
    Home = Tk()
    Home.title("Gestion Des Stocks")
    width = 1024
    height = 520
    screen_width = Home.winfo_screenwidth()
    screen_height = Home.winfo_screenheight()
    x = (screen_width/2) - (width/2)
    y = (screen_height/2) - (height/2)
    Home.geometry("%dx%d+%d+%d" % (width, height, x, y))
    Home.resizable(0, 0)
    Title = Frame(Home, bd=1, relief=SOLID)
    Title.pack(pady=10)
    lbl_display = Label(Title, text="Gestion Des Stocks", font=('Helvetica', 45))
    lbl_display.pack()
    menubar = Menu(Home)
    filemenu = Menu(menubar, tearoff=0)
    filemenu2 = Menu(menubar, tearoff=0)
    filemenu.add_command(label="Se déconnecter", command=Logout)
    filemenu.add_command(label="Quitter l'application", command=Exit2)
    filemenu2.add_command(label="Ajouter un produit", command=ShowAddNew)
    filemenu2.add_command(label="Voir les produits", command=ShowView)
    menubar.add_cascade(label="Mon Compte", menu=filemenu)
    menubar.add_cascade(label="Stockage", menu=filemenu2)
    Home.config(menu=menubar)
    Home.config(bg="brown")

def ShowAddNew():
    global addnewform
    addnewform = Toplevel()
    addnewform.title(" - Gestion Des Stocks - ")
    width = 600
    height = 500
    screen_width = Home.winfo_screenwidth()
    screen_height = Home.winfo_screenheight()
    x = (screen_width/2) - (width/2)
    y = (screen_height/2) - (height/2)
    addnewform.geometry("%dx%d+%d+%d" % (width, height, x, y))
    addnewform.resizable(0, 0)
    AddNewForm()

def AddNewForm():
    TopAddNew = Frame(addnewform, width=600, height=100, bd=1, relief=SOLID)
    TopAddNew.pack(side=TOP, pady=20)
    lbl_text = Label(TopAddNew, text="Ajouter un nouveau produit", font=('Helvetica', 18), width=600)
    lbl_text.pack(fill=X)
    MidAddNew = Frame(addnewform, width=600)
    MidAddNew.pack(side=TOP, pady=50)
    lbl_productname = Label(MidAddNew, text="Nom du produit:", font=('Helvetica', 25), bd=10)
    lbl_productname.grid(row=0, sticky=W)
    lbl_qty = Label(MidAddNew, text="La quantité de produit:", font=('Helvetica', 25), bd=10)
    lbl_qty.grid(row=1, sticky=W)
    lbl_price = Label(MidAddNew, text="Prix du produit:", font=('Helvetica', 25), bd=10)
    lbl_price.grid(row=2, sticky=W)
    productname = Entry(MidAddNew, textvariable=PRODUCT_NAME, font=('Helvetica', 25), width=15)
    productname.grid(row=0, column=1)
    productqty = Entry(MidAddNew, textvariable=PRODUCT_QTY, font=('Helvetica', 25), width=15)
    productqty.grid(row=1, column=1)
    productprice = Entry(MidAddNew, textvariable=PRODUCT_PRICE, font=('Helvetica', 25), width=15)
    productprice.grid(row=2, column=1)
    btn_add = Button(MidAddNew, text="Enregistrer un nouveau produit", font=('Helvetica', 18), width=30, bg="#cd0000", command=AddNew)
    btn_add.grid(row=3, columnspan=2, pady=20)

def AddNew():
    Database()
    cursor.execute("INSERT INTO `product` (product_name, product_qty, product_price) VALUES(?, ?, ?)", (str(PRODUCT_NAME.get()), int(PRODUCT_QTY.get()), int(PRODUCT_PRICE.get())))
    conn.commit()
    PRODUCT_NAME.set("")
    PRODUCT_PRICE.set("")
    PRODUCT_QTY.set("")
    cursor.close()
    conn.close()

def ViewForm():
    global tree
    TopViewForm = Frame(viewform, width=600, bd=1, relief=SOLID)
    TopViewForm.pack(side=TOP, fill=X)
    LeftViewForm = Frame(viewform, width=600)
    LeftViewForm.pack(side=LEFT, fill=Y)
    MidViewForm = Frame(viewform, width=600)
    MidViewForm.pack(side=RIGHT)
    lbl_text = Label(TopViewForm, text="Voir les produits", font=('Helvetica', 18), width=600)
    lbl_text.pack(fill=X)
    lbl_txtsearch = Label(LeftViewForm, text="MES PRODUITS", font=('Helvetica', 15))
    lbl_txtsearch.pack(side=TOP, anchor=W)
    search = Entry(LeftViewForm, textvariable=SEARCH, font=('Helvetica', 15), width=10)
    search.pack(side=TOP,  padx=10, fill=X)
    btn_search = Button(LeftViewForm, text="Rechercher un produit", command=Search)
    btn_search.pack(side=TOP, padx=10, pady=10, fill=X)
    btn_reset = Button(LeftViewForm, text="Annuler", command=Reset)
    btn_reset.pack(side=TOP, padx=10, pady=10, fill=X)
    btn_delete = Button(LeftViewForm, text="Retirer", command=Delete)
    btn_delete.pack(side=TOP, padx=10, pady=10, fill=X)
    btn_save = Button(LeftViewForm, text="Enregistrer-ici", command=Database)
    btn_save.pack(side=TOP, padx=10, pady=10, fill=X)
    btn_export = Button(LeftViewForm, text="Exporter tous les produits", command=Export)
    btn_export.pack(side=TOP, padx=10, pady=10, fill=X)
    scrollbarx = Scrollbar(MidViewForm, orient=HORIZONTAL)
    scrollbary = Scrollbar(MidViewForm, orient=VERTICAL)
    tree = ttk.Treeview(MidViewForm, columns=("ProductID", "Product Name", "Product Qty", "Product Price"), selectmode="extended", height=100, yscrollcommand=scrollbary.set, xscrollcommand=scrollbarx.set)
    scrollbary.config(command=tree.yview)
    scrollbary.pack(side=RIGHT, fill=Y)
    scrollbarx.config(command=tree.xview)
    scrollbarx.pack(side=BOTTOM, fill=X)
    tree.heading('ProductID', text="ProductID",anchor=W)
    tree.heading('Product Name', text="Nom du produit",anchor=W)
    tree.heading('Product Qty', text="Quantité de produit",anchor=W)
    tree.heading('Product Price', text="Prix du produit",anchor=W)
    tree.column('#0', stretch=NO, minwidth=0, width=0)
    tree.column('#1', stretch=NO, minwidth=0, width=0)
    tree.column('#2', stretch=NO, minwidth=0, width=200)
    tree.column('#3', stretch=NO, minwidth=0, width=120)
    tree.column('#4', stretch=NO, minwidth=0, width=120)
    tree.pack()
    DisplayData()

def DisplayData():
    Database()
    cursor.execute("SELECT * FROM `product`")
    fetch = cursor.fetchall()
    for data in fetch:
        tree.insert('', 'end', values=(data))
    cursor.close()
    conn.close()

def Search():
    if SEARCH.get() != "":
        tree.delete(*tree.get_children())
        Database()
        cursor.execute("SELECT * FROM `product` WHERE `product_name` LIKE ?", ('%'+str(SEARCH.get())+'%',))
        fetch = cursor.fetchall()
        for data in fetch:
            tree.insert('', 'end', values=(data))
        cursor.close()
        conn.close()

def Reset():
    tree.delete(*tree.get_children())
    DisplayData()
    SEARCH.set("")

def Delete():
    if not tree.selection():
       print("ERROR")
    else:
        result = tkMessageBox.askquestion('Gestion Des Stocks', 'Voulez-vous vraiment supprimer cet enregistrement ?', icon="warning")
        if result == 'yes':
            curItem = tree.focus()
            contents =(tree.item(curItem))
            selecteditem = contents['values']
            tree.delete(curItem)
            Database()
            cursor.execute("DELETE FROM `product` WHERE `product_id` = %d" % selecteditem[0])
            conn.commit()
            cursor.close()
            conn.close()


def Export():
    conn=sql.connect('mydata.db')
    cursor=conn.cursor()
    cursor.execute("SELECT * FROM product")
    with open("product.csv","w")as csv_file:
        csv_writer=csv.writer(csv_file,delimiter="\t")
        csv_writer.writerow([i[0]for i in cursor.description])
        csv_writer.writerows(cursor)
    dir_path =os.getcwd() + "/product.cv"
    messagebox.showinfo('1 File','Exported Successfully')    


def Save():
    conn=sql.connect('mydata.db')
    cursor=conn.cursor()
    cursor.execute("SELECT * FROM product")
    with open("product.csv","w")as csv_file:
        csv_writer=csv.writer(csv_file,delimiter="\t")
        csv_writer.writerow([i[0]for i in cursor.description])
        csv_writer.writerows(cursor)
    dir_path =os.getcwd() + "/product.cv"
    messagebox.showinfo('1 File','Exported Successfully')  

def Export2():
    conn=sql.connect('mydata.db')
    cursor=conn.cursor()
    cursor.execute("SELECT * FROM product")
    with open("product.csv","w")as csv_file:
        csv_writer=csv.writer(csv_file,delimiter="\t")
        csv_writer.writerow([i[0]for i in cursor.description])
        csv_writer.writerows(cursor)
    dir_path =os.getcwd() + "/product.cv"
    messagebox.showinfo('1 File','Exported Successfully')

#====================Exporting To Excel=================================#

def Export():
    if not os.path.exists('./Excel_import'):
        os.makedirs('./Excel_import')
    conn=sqlite3.connect("mydata.db")
    c=conn.cursor()
    c.execute("SELECT * FROM product")
    data = c.fetchall()
    time=str(date.today())
    df=pd.DataFrame(data, columns=['ID', 'Nom', 'Qty','Prix'])
    datatoexcel = pd.ExcelWriter("./Excel_import/My Work"+time+".xlsx", engine='xlsxwriter')
    df.to_excel(datatoexcel, index= False, sheet_name = "Sheet1")
    worksheet = datatoexcel.sheets['Sheet1']
    worksheet.set_column('A:A', 5)
    worksheet.set_column('B:B', 10)
    worksheet.set_column('C:C', 13)
    worksheet.set_column('D:D', 10)
    worksheet.set_column('E:E', 10)
    worksheet.set_column('F:F', 10)
    worksheet.set_column('G:G', 15)
    worksheet.set_column('H:H', 26)
    datatoexcel.save()
    messagebox.showinfo("1 Fichier","Fichier Excel généré avec succès")

#====================Exporting To CSV=================================#

def ShowView():
    global viewform
    viewform = Toplevel()
    viewform.title("TOUS LES PRODUITS")
    width = 600
    height = 400
    screen_width = Home.winfo_screenwidth()
    screen_height = Home.winfo_screenheight()
    x = (screen_width/2) - (width/2)
    y = (screen_height/2) - (height/2)
    viewform.geometry("%dx%d+%d+%d" % (width, height, x, y))
    viewform.resizable(0, 0)
    ViewForm()

def Logout():
    result = tkMessageBox.askquestion('Panel Area', 'Êtes-vous sûr de vouloir vous déconnecter ?', icon="warning")
    if result == 'yes': 
        admin_id = ""
        root.deiconify()
        Home.destroy()
  
def Login(event=None):
    global admin_id
    Database()
    if USERNAME.get == "" or PASSWORD.get() == "":
        lbl_result.config(text="Veuillez remplir le champ requis !", fg="red")
    else:
        cursor.execute("SELECT * FROM `amine` WHERE `username` = ? AND `password` = ?", (USERNAME.get(), PASSWORD.get()))
        if cursor.fetchone() is not None:
            cursor.execute("SELECT * FROM `amine` WHERE `username` = ? AND `password` = ?", (USERNAME.get(), PASSWORD.get()))
            data = cursor.fetchone()
            admin_id = data[0]
            USERNAME.set("")
            PASSWORD.set("")
            lbl_result.config(text="")
            ShowHome()
        else:
            lbl_result.config(text="Invalid username or password", fg="red")
            USERNAME.set("")
            PASSWORD.set("")
    cursor.close()
    conn.close() 

def ShowHome():
    root.withdraw()
    Home()
    loginform.destroy()
 

#MENUBAR WIDGETS
menubar = Menu(root)
filemenu = Menu(menubar, tearoff=0)
menubar.add_cascade(label="Accueil", menu=Exit)
menubar.add_cascade(label="Gérer Mon Compte", command=ShowLoginForm)
root.config(menu=menubar)

#FRAME
Title = Frame(root, bd=1, relief=SOLID)
Title.pack(pady=10)

#LABEL WIDGET
lbl_display = Label(Title, text="WELCOME TO YOUR STORE", font=('Helvetica', 45))
lbl_display.pack()

#INITIALIZATION
if __name__ == '__main__':
    root.mainloop()
