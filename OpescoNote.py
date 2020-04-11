# import a tkinter package for gui and set the window object
from tkinter import *
from tkinter.filedialog import *
import re
import os
import sys
from tkinter.messagebox import *
import win32ui
import tkinter.ttk as ttk
from tkinter import colorchooser, font 
class OpescoNote:
    def __init__(self):
        self.mainBody = self.bodytxt = self.in_path = self.finded = self.infile = self.val =self.val2= self.numberlist=''
        self.count = 1
        self.val3 =0
        self.wcount = self.findw = self.replace = self.lb= None
        self.mainWindow()
        self.container.mainloop()

    def dothis(self):
        pass
    def printfile(self):
        #with open(self.in_path, 'r') as self.infile:
            #textfile= self.infile.read()
        textfile = self.mainBody.get(1.0, END)
        dc = win32ui.CreateDC()
        dc.CreatePrinterDC()
        dc.SetMapMode(4)
        font = win32ui.CreateFont({'name':'Arial','height':16})
        dc.SelectObject(font)
        dc.StartDoc(textfile)
        dc.StartPage()
        dc.TextOut(40,40, textfile)
        dc.EndPage()
        dc.EndDoc()
    
    def exitpage(self):
        msg = askyesnocancel("warning", "Do you want to save this page?")
        if msg == True:
            self.savefile()
            if self.in_path == None:
                self.in_path = ''
            else:
                self.container.destroy()
        elif msg == False:
            self.container.destroy()

    def newPage(self, event=None):
        msg = askyesnocancel("warning", "Do you want to save this page?")
        if msg == True:
            self.savefile()
            if self.in_path == None:
                self.in_path = ''
            else:
                self.count +=1
                self.container.title("Untitled"+ str(self.count) +" -OpescoNote")
                self.in_path = ''
                self.mainBody.delete(1.0,END)
        elif msg == False:
            self.count +=1
            self.container.title("Untitled "+ str(self.count) +"-OpescoNote")
            self.in_path = ''
            self.mainBody.delete(1.0,END)

    def openfile(self):
        try:
            self.in_path = askopenfilename(defaultextension=".txt", filetypes=[("All Files","*.*"), ("Text Documents","*.txt")])
            if self.in_path != " ":
                with open(self.in_path, 'r') as self.infile:
                    self.container.title(os.path.basename(self.in_path) + " - OpescoNote")
                    self.mainBody.delete(1.0, END) 
                    myfile = self.infile.read()
                    self.mainBody.insert(1.0, myfile, END)
        except:
            pass 
    def savefile(self):
        if self.in_path == '':
            #Save as new file
            try:
                self.in_path = asksaveasfilename(initialfile='Untitled.txt', defaultextension=".txt", filetypes=[("All Files","*.*"), \
                                          ("Text Documents","*.txt")])
                #Try to save the file
                with open(self.in_path,"w") as self.infile:
                    self.infile.write(self.mainBody.get(1.0, END))
                    # Change the window title
                    self.container.title(os.path.basename(self.in_path) + " - OpescoNote")
            except:
                pass
        else:
            with open(self.in_path,"w") as self.infile:
                self.infile.write(self.mainBody.get(1.0, END))
                # Change the window title
                self.container.title(os.path.basename(self.in_path) + " - OpescoNote")

    def saveasfile(self):
        #Save as new file
        self.in_path = asksaveasfilename(initialfile='Untitled.txt', defaultextension=".txt", filetypes=[("All Files","*.*"), \
                               ("Text Documents","*.txt")])
        try:
            #Try to save the file
            with open(self.in_path,"w") as self.infile:
                self.infile.write(self.mainBody.get(1.0, END))
                # Change the window title
                self.container.title(os.path.basename(self.in_path) + " - OpescoNote")
        except:
            pass
    def undoevent(self):
        try:
            self.mainBody.edit_undo()
        except:
            showinfo("Massage", "Nothing to Undo")
    def redoevent(self):
        try:
            self.mainBody.edit_redo()
        except:
            showinfo("Massage", "You need to undo before you can redo")
    def cutfile(self):
        self.mainBody.event_generate("<<Cut>>")
    def copyfile(self):
        self.mainBody.event_generate("<<Copy>>")
    def pastefile(self):
        self.mainBody.event_generate("<<Paste>>")
    def delete_event(self):
        try:
            self.mainBody.delete(SEL_FIRST, SEL_LAST)
        except:
            showinfo("Massage", "Select portion of text to delete")
    def selectall(self):
        try:
            self.mainBody.tag_add(SEL, '1.0', END)
            self.mainBody.mark_set(INSERT, '1.0')
            self.mainBody.see(INSERT)
            return 'break'
        except:
            showinfo("Message", "Nothing to select")

    def aboutus(self):
       showinfo("About OpescoNote","Simple text editor like notepad using Python. It was created by Yemi lead instructor for \
           Pythonclass at SQI college of ICT. Copyright of SQI")
    def contactus(self):
       showinfo("Contact Us", "SQI college of ICT, Opposit Yoako filling station Ilorin road, Ogbomosho, Oyo State, Nigeria.")
   
    # the below code is use to set the status 
    def setStatus_bar(self, event):
        list1=self.mainBody.index(INSERT).split('.')
        statusbar = "Line= "+str(self.mainBody.count('1.0', END, 'lines'))+",   Cursor Position = row: "+list1[0]+", col: "+list1[1]+",    wordcount= "+str(len(self.mainBody.get('1.0', 'end-1c').split()))
        self.showtext.set(statusbar)
    
    def findtext(self):

        def find_next():
            start = "1.0"
            end = "end"
            start = self.mainBody.index(start)
            end = self.mainBody.index(end)
            self.mainBody.mark_set("matchStart", start)
            self.mainBody.mark_set("matchEnd", start)
            self.mainBody.mark_set("searchLimit", end)
            targetfind = self.findw.get()
            if targetfind:
                state = True
                while state:
                    where = self.mainBody.search(targetfind, "matchEnd", "searchLimit", count=count)
                    if where == "":
                            break
                    else:
                            pastit = where + ('+%dc' % len(targetfind)) 
                            self.mainBody.mark_set("matchStart", where)
                            self.mainBody.mark_set("matchEnd", "%s" % (where))
                            self.mainBody.tag_add(SEL, where, pastit)
                            self.mainBody.see(INSERT)
                            self.mainBody.focus()
                            print(where)
                            state = False
                            start = where       
        self.mainBody.tag_remove(SEL, '1.0', END)
      
        def find_all():
            start = "1.0"
            end = "end"
            start = self.mainBody.index(start)
            end = self.mainBody.index(end)
            count= IntVar() 
            count=count
            self.mainBody.mark_set("matchStart", start)
            self.mainBody.mark_set("matchEnd", start)
            self.mainBody.mark_set("searchLimit", end)
            targetfind = self.findw.get()
            if targetfind:
                while True:
                    where = self.mainBody.search(targetfind, "matchEnd", "searchLimit", count=count)
                    if where == "":
                        break
                    elif where:
                        pastit = where + ('+%dc' % len(targetfind)) 
                        self.mainBody.mark_set("matchStart", where)
                        self.mainBody.mark_set("matchEnd", "%s+%sc" % (where, count.get()))
                        self.mainBody.tag_add(SEL, where, pastit)
                        self.mainBody.see(INSERT)
                        self.mainBody.focus()
        self.mainBody.tag_remove(SEL, '1.0', END)

        def replaceit():
            self.bodytxt = self.mainBody.get(1.0, END)
            self.finded = self.findw.get()
            self.replacew = self.replace.get()
            # self.bodytxt2 =self.bodytxt.replace(self.finded, self.replacew)
            # self.mainBody.replace(1.0, 'end', self.bodytxt2)
            self.mainBody.replace(1.0, 'end', self.mainBody.get(1.0, 'end').replace(self.finded, self.replacew))

        find = Toplevel()
        find.grid_propagate(0)
        frm = Frame(find)
        frm.pack(side=TOP, pady=6)
        frm.grid_propagate(0)
        Label(frm, text="Find what?   ").pack(side=LEFT, expand=YES, fill=BOTH)
        self.findw = Entry(frm, width=50)
        self.findw.focus_set()
        self.findw.pack(side=LEFT, expand=YES, fill=BOTH, padx=3)
        frm2 = Frame(find)
        frm2.pack(side=TOP, pady=6)
        frm2.grid_propagate(0)
        Label(frm2, text="replace with").pack(side=LEFT, expand=YES, fill=BOTH)
        self.replace = Entry(frm2, width=50)
        self.replace.pack(side=LEFT, expand=YES, fill=BOTH, padx=3)
        frm3 = Frame(find)
        frm3.pack(side=TOP, pady=6)
        frm3.grid_propagate(0)
        Button(frm3, text='Find Next', width=10, command=find_next).pack(side=LEFT, expand=YES, fill=BOTH, padx=3)
        Button(frm3, text='Find All', width=10, command=find_all).pack(side=LEFT, expand=YES, fill=BOTH, padx=3)
        Button(frm3, text='replace', width=10, command=replaceit).pack(side=LEFT, expand=YES, fill=BOTH, padx=3)
        
    def goTo(self):

        def findline():
            try:
                goline = gotow.get()
                self.mainBody.mark_set(INSERT, goline +'.0')
                self.mainBody.see(INSERT)
                self.mainBody.focus()
            except:
                pass
        
        go = Toplevel()
        go.grid_propagate(0)
        go.title('Go To')
        Label(go, text="Enter Line Number?  ").pack(side=TOP, anchor=W, padx=5, pady=3)
        gotow = Entry(go, width=30, text='goto')
        gotow.pack(side=TOP, padx=5, pady=3)
        Button(go, text='Go To', width=10, command=findline).pack(side=TOP, padx=5, pady=3)
    
    def getColor(self):
        self.mainBody.tag_remove('color1', SEL_FIRST, SEL_LAST)
        (rgb, color) = colorchooser.askcolor()
        self.mainBody.tag_add('color1', SEL_FIRST, SEL_LAST)
        self.mainBody.tag_configure('color1', foreground=color)

    def style_changer(self, style):
        style_select = style
        try:
            self.mainBody.tag_add('Style1', SEL_FIRST, SEL_LAST)
            self.mainBody.tag_config('Style1', font=(style_select))
        except:
            pass
    def number_changer(self, num):
        number = num
        try:
            self.mainBody.tag_add('Size1', SEL_FIRST, SEL_LAST)
            self.mainBody.tag_config('Size1', font=(number))
        except:
            pass
    def boldText(self):
        try:
            if bold.get() ==True:
                self.mainBody.tag_add('bold1', SEL_FIRST, SEL_LAST)
                self.mainBody.tag_config('bold1', font=('bold'))
            else:
                self.mainBody.tag_remove('bold1', SEL_FIRST, SEL_LAST)
        except:
             pass
    def italicText(self):
        try:
            if italic.get() ==True:
                self.mainBody.tag_add('italic1', SEL_FIRST, SEL_LAST)
                self.mainBody.tag_config('italic1', font=('italic'))
            else:
                self.mainBody.tag_remove('italic1', SEL_FIRST, SEL_LAST)
        except:
            pass
    def underlineText(self):
        try:
            if underline.get() ==True:
                self.mainBody.tag_add('underline1', SEL_FIRST, SEL_LAST)
                self.mainBody.tag_config('underline1', underline=True)
            else:
                self.mainBody.tag_remove('underline1', SEL_FIRST, SEL_LAST)
        except:
            pass
    def wrapText(self):
        try:
            if wrapT.get()==True:
                self.mainBody.config(wrap=WORD)
            else:
                self.mainBody.config(wrap=NONE)
        except:
            pass
    def change_font(self):
        
        def style_select(event):
            stent.set(stylebox.get(stylebox.curselection()))
            self.val =stent.get()
            self.lb.config(font=(self.val, self.val3, self.val2))
            
        def fsty_select(even):
            stent2.set(stylebox2.get(stylebox2.curselection()))
            self.val2 =stent2.get()
            self.lb.config(font=(self.val, self.val3, self.val2))
            
        def numb_select(eve):
            stent3.set(stylebox3.get(stylebox3.curselection()))
            self.val3 =stent3.get()
            self.lb.config(font=(self.val, self.val3, self.val2))
            
        def underline_Text():
            pass

        def apply_Effect():
            #try:
                
                if 'fonts' in self.mainBody.tag_names():
                    list(self.mainBody.tag_names()).remove('fonts')
                self.val =stent.get()
                self.val2 =stent2.get()
                self.val3 =stent3.get() 
                self.mainBody.tag_add('fonts', SEL_FIRST, SEL_LAST)
                self.mainBody.tag_config('fonts', font=(self.val, self.val3, self.val2))
                print(self.mainBody.tag_names())
                font.destroy()
            # except:
            #     pass

        def cancle_Effect():
            font.destroy()

        self.numberlist = [8, 9, 10,11,12,14,16,18,20,22,24,26,28,32,36,42,36,48,72]
        stylelist = ('AcadEref','Agency FB','AIGDT','AIGDT','Algerian','AmdtSymbols','AMGDT','Arial','Arial Black','Arial Narrow',\
    'Arial Rounded MT Bold','Arial Unicode MS','Bahnschrift','Bahnschrift Condensed','Bahnschrift Light','Bahnschrift Light Condensed',\
    'Bahnschrift Light SemiCondensed','Bahnschrift SemiBold','Bahnschrift SemiBold Condensed','Bahnschrift SemiBold SemiConden')
        fontstyle =['normal', 'italic', 'bold', 'roman', 'bold italic']
        font = Toplevel(height=200)
        Label(font, text="Fonts").grid(row=0, column=0, columnspan=2)
        stent =StringVar()
        stent.set('Times')
        stentry =Entry(font, width=20, textvariable=stent)
        stentry.grid(row=1, column=0, columnspan=2)
        yscrol = Scrollbar(font, orient='vertical')
        stylebox = Listbox(font, yscrollcommand=yscrol.set, exportselection=0)
        yscrol.config(command = stylebox.yview)
        for i in stylelist: 
            stylebox.insert(END, str(i))
        stylebox.grid(row=2, column=0, columnspan=2)
        yscrol.grid(row=2, column=2, sticky='ns')
        stylebox.bind('<<ListboxSelect>>', style_select)
        
        Label(font, text="Style").grid(row=0, column=3,columnspan=2, padx=7)
        stent2 =StringVar()
        stent2.set('normal')
        stentry2 =Entry(font, width=12, textvariable=stent2)
        stentry2.grid(row=1, column=3, columnspan=2,  padx=7)
        stylebox2 = Listbox(font, width=12, exportselection=0)
        for i in fontstyle: 
            stylebox2.insert(END, str(i))
        stylebox2.grid(row=2, column=3, columnspan=2,  padx=7)
        stylebox2.bind('<<ListboxSelect>>', fsty_select)
        
        Label(font, text="Size").grid(row=0, column=5)
        stent3 =IntVar()
        stent3.set(8)
        stentry3 =Entry(font, width=6, textvariable=stent3)
        stentry3.grid(row=1, column=5)
        yscrol3 = Scrollbar(font, orient='vertical')
        stylebox3 = Listbox(font, yscrollcommand=yscrol3.set, width=6, exportselection=0)
        yscrol3.config(command = stylebox3.yview)
        for i in self.numberlist: 
            stylebox3.insert(END, str(i))
        stylebox3.grid(row=2, column=5)
        yscrol3.grid(row=2, column=6, sticky='ns')
        stylebox3.bind('<<ListboxSelect>>', numb_select)
        Label(font, text=" ").grid(row=2, column=7)
        
        underline = Checkbutton(font, text="Underline", command=underline_Text)
        underline.grid(row=3, column=0, pady=10)
        self.lbframe =LabelFrame(font, text="Text Preview", width=100)
        self.lbframe.grid(row=3, column=1, columnspan=8, pady=10, sticky='ne')
        self.lb =Label(self.lbframe, text="Preview Text", font=("Times New Roman", '16'))
        self.lb.pack(side=LEFT)
        Button(font, text="OK", command=apply_Effect, width=7).grid(row=4, column=3, columnspan=1, pady=10)
        Button(font, text="Cancle", command=cancle_Effect, width=7).grid(row=4, column=4, columnspan=3, pady=10)
       
    #the main window starts from here.
    def mainWindow(self):
        self.container = Tk()
        self.container.title("Untitled "+ str(self.count) +" -OpescoNote")

        # create all menubar objects
        menubar = Menu(self.container)

        #create file and its submenus
        fileMenu = Menu(menubar, tearoff=0)
        fileMenu.add_command(label='New Page', command=self.newPage, accelerator="Ctrl+N")
        fileMenu.add_command(accelerator="Ctr+O", label='Open', command=self.openfile)
        fileMenu.add_command(accelerator="Ctr+S", label='Save', command=self.savefile)
        fileMenu.add_command(label='Save As', command=self.saveasfile)
        fileMenu.add_separator()
        fileMenu.add_command(label='Page Setup', command=self.dothis)
        fileMenu.add_command( accelerator="Ctr+P", label='Print', command=self.printfile)
        fileMenu.add_separator()
        fileMenu.add_command(label='Exit', command=self.exitpage)
        menubar.add_cascade(label='File', menu=fileMenu, underline=0)

        #create edit and its submenus
        editMenu = Menu(menubar, tearoff=0)
        editMenu.add_command(label='Undo', command=self.undoevent)
        editMenu.add_command(label='Redo', command=self.redoevent)
        editMenu.add_separator()
        editMenu.add_command(label='Copy', command=self.copyfile)
        editMenu.add_command(label='Cut', command=self.cutfile)
        editMenu.add_command(label='Paste', command=self.pastefile)
        editMenu.add_separator()
        editMenu.add_command(label='Find & replace', command=self.findtext)
        editMenu.add_command(label='Go To', command=self.goTo)
        editMenu.add_command(label='Select All', command=self.selectall)
        editMenu.add_command(label='Delete', command=self.delete_event)
        menubar.add_cascade(label='Edit', menu=editMenu)

        #create the format and its submenus
        formatMenu = Menu(menubar, tearoff=0)
        formatMenu.add_command(label='Font', command=self.change_font)

        numberMenu = Menu(menubar, tearoff=0)
        for num in self.numberlist:
            numberMenu.add_command(label=num, command=lambda i=num: self.number_changer(i))
        formatMenu.add_cascade(label='Size', menu=numberMenu)
           
        styleMenu = Menu(menubar, tearoff=0)
        styleMenu.add_command(label='Agency FB', command=lambda:self.style_changer('Agency FB'))
        styleMenu.add_command(label='AcadEref', command=lambda:self.style_changer('AcadEref'))
        styleMenu.add_command(label='Arial', command=lambda:self.style_changer('Arial'))
        styleMenu.add_command(label='Arial Black', command=lambda:self.style_changer('Arial Black'))
        styleMenu.add_command(label='Bahnschrift', command=lambda:self.style_changer('Bahnschrift'))
        styleMenu.add_command(label='Arial Unicode MS', command=lambda:self.style_changer('Arial Unicode MS'))
        styleMenu.add_command(label='Bahnschrift SemiBold Condensed', command=lambda:self.style_changer('Bahnschrift SemiBold Condensed'))
        styleMenu.add_command(label='Bahnschrift Light', command=lambda:self.style_changer('Bahnschrift Light'))
        styleMenu.add_command(label='Bahnschrift Condensed', command=lambda:self.style_changer('Bahnschrift Condensed'))
        styleMenu.add_command(label='Algerian', command=lambda:self.style_changer('Algerian'))
        styleMenu.add_command(label='AmdtSymbols', command=lambda:self.style_changer('AmdtSymbols'))
        formatMenu.add_cascade(label='Style', menu=styleMenu)

        formatMenu.add_command(label='Color', command= self.getColor)
        formatMenu.add_separator()
        bold =BooleanVar()
        italic =BooleanVar()
        underline =BooleanVar()
        formatMenu.add_checkbutton(label='Bold', variable=bold, command=self.boldText)
        formatMenu.add_checkbutton(label='Italic', variable=italic, command=self.italicText)
        formatMenu.add_checkbutton(label='Underline', variable=underline, underline=0, command=self.underlineText)
        menubar.add_cascade(label='Format', menu=formatMenu)

        # create the view submenu
        wrapT = BooleanVar()
        view = Menu(menubar, tearoff=0)
        view.add_checkbutton(label='Word wrap', variable=wrapT, command=self.wrapText)
        view.add_checkbutton(label='Show Status', variable=bold, command=self.boldText)
        menubar.add_cascade(label='View', menu=view) 

        #create the help and its submenus
        helpMenu = Menu(menubar, tearoff=0)
        helpMenu.add_command(label='About OpescoNote', command=self.aboutus)
        helpMenu.add_command(label='Contact Us', command=self.contactus)
        menubar.add_cascade(label='help', menu=helpMenu)
        self.container.config(menu=menubar)

        #create a frame to hold the text area
        bodyfrm = Frame(self.container)
        bodyfrm.pack(side=TOP, expand=YES, fill=BOTH)
        bodyfrm.grid_propagate(0)
        self.mainBody = Text(bodyfrm, height=30, width=125, font=('Times New Roman', 12), wrap=NONE)
        self.mainBody.focus_set()
        self.mainBody.pack(side=LEFT, expand=YES, fill=BOTH)
        vscroll = Scrollbar(bodyfrm, orient = "vertical", command = self.mainBody.yview)
        vscroll.pack(side= LEFT, anchor=E, fill=BOTH)
        hscroll = Scrollbar(self.container, orient = "horizontal", command = self.mainBody.xview)
        hscroll.pack(side=TOP, anchor=W, fill=BOTH)

        self.mainBody.configure(yscrollcommand = vscroll.set)
        self.mainBody.configure(xscrollcommand=hscroll.set)
        #Adding a binding to the menu
        self.container.bind_all("<Control-n>", self.newPage)
        self.mainBody.bind_all("<Control-Key-N>", self.newPage)
        self.mainBody.bind("<KeyPress>", self.setStatus_bar)
        self.mainBody.bind("<ButtonPress-1>", self.setStatus_bar)
        #create a status bar
        self.showtext = StringVar()
        statusbarfrm = Frame(self.container)
        statusbarfrm.pack(side=TOP, expand=YES, fill=BOTH, anchor=W)
        statusbarfrm.grid_propagate(0)
        linecount = Label(statusbarfrm, textvariable= self.showtext ).pack(side=LEFT)
        
if __name__ == "__main__":
    opnote = OpescoNote()

   
