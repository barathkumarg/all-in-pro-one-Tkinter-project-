

from tkinter import *
from tkinter.ttk import *
import tkinter as tk
from tkinter import ttk,Canvas
from docx import Document
from docxcompose.composer import Composer
from docx import Document as Document_compose
from gtts import gTTS
from tkinter import messagebox
import os
import cv2
from PIL import ImageTk,Image
from PyPDF2 import PdfFileReader,PdfFileMerger
from tkinter.filedialog import askopenfilename,asksaveasfile,askopenfilenames,askdirectory



LARGEFONT = ("Verdana", 35)


class tkinterApp(tk.Tk):

    # __init__ function for class tkinterApp
    def __init__(self, *args, **kwargs):
        # __init__ function for class Tk
        tk.Tk.__init__(self, *args, **kwargs)

        # creating a container
        container = tk.Frame(self)
        container.pack(side="top", fill="both", expand=True)

        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.title('All In One Pro')



        # initializing frames to an empty array
        self.frames = {}
        self.geometry('550x410')

        # iterating through a tuple consisting
        # of the different page layouts
        for F in (StartPage, Page1, Page2, Page3, Page4):
            frame = F(container,self)


            # initializing frame of that object from
            # startpage, page1, page2 respectively with
            # for loop
            self.frames[F] = frame

            frame.grid(row=0, column=0, sticky="nsew")


        self.show_frame(StartPage)

    # to display the current frame passed as
    # parameter
    def show_frame(self, cont):
        frame = self.frames[cont]
        frame.tkraise()

    # first window frame startpage


class StartPage(tk.Frame):
    def __init__(self, parent,controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller



        # Setting icon of master window


        #arrow1 = tk.Label(self, image=self.controller.arrow)

        l1=Label(self,text="All in one Pro",font=('VERDANA',22))
        l1.place(x=310, y=0)

        #arrow1.grid(row=1, column=0)
        bard = Image.open("download.png")
        newsize = (300, 410)
        bard = bard.resize(newsize)
        bardejov = ImageTk.PhotoImage(bard)
        label1 = Label(self, image=bardejov)
        label1.image = bardejov
        label1.place(x=0, y=0)

        button1 = tk.Button(self, text="Word Document Merger",fg="black", background="white",width=25,
                             command=lambda: controller.show_frame(Page1))

        # putting the button in its place by
        # using grid
        button1.place(x=340, y=60)

        ## button to show frame 2 with text layout2
        button2 = tk.Button(self, text="Image Convertor",fg="black", background="white",width=25,
                             command=lambda: controller.show_frame(Page2))

        button2.place(x=340, y=125)

        button3 = tk.Button(self, text="Word to speech",fg="black", background="white",width=25,
                             command=lambda: controller.show_frame(Page3))

        button3.place(x=340, y=190)

        button3 = tk.Button(self, text="Pdf Merger",fg="black", background="white",width=25,
                             command=lambda: controller.show_frame(Page4))

        button3.place(x=340, y=255)

        label2=Label(self,text="Created by Barathkumar G")
        label2.place(x=305,y=360)

    # second window frame page1


class Page1(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        label = ttk.Label(self, text="Word Document Merger", font=('courier',22))
        label.place(x=200,y=0)

        labelx=ttk.Label(self,text=" Note: hold ctrl and select to make multiselect")
        labelx.place(x=10, y=200)

        bard = Image.open("word.jpg")
        newsize = (360, 410)
        bard = bard.resize(newsize)
        bardejov = ImageTk.PhotoImage(bard)
        label1 = Label(self, image=bardejov)
        label1.image = bardejov
        label1.place(x=280, y=90)


        def openFile():
            global var, var1
            var = StringVar()
            var1 = StringVar()
            var1 = ""
            var = askopenfilenames(defaultextension="docx",
                                   filetypes=[("Document files", "*.docx")])
            var=list(var)
            var1=var



            l3.config(text="{0} files selected".format(len(var)))
            if var == "":
                file = None




        def convert():
            base = os.path.basename(var[0])






            if var!='' or var1!="":

                number_of_sections = len(var1)
                master = Document_compose(var1[0])
                composer = Composer(master)
                for i in range(0, number_of_sections):
                    doc_temp = Document_compose(var1[i])
                    composer.append(doc_temp)



                composer.save("{0}\Combined_document.docx".format(os. path. dirname(var1[0])))
                l3.config(text="no file choosen")


                messagebox.showinfo('Success', "Converted successfully Check in same folder of file !!!")



        l2 = ttk.Label(self, text='Choose file: ')
        l2.place(x=0, y=50)

        l3= ttk.Label(self,text="no file choosen")
        l3.place(x=150, y=48)

        openFileButton = ttk.Button(self, text=" Open ", width=10,command=openFile)
        openFileButton.place(x=80, y=47)

        l4 = ttk.Button(self, text='Merge',command=convert)
        l4.place(x=78, y=95)






        button1 = tk.Button(self, text="Back",width=10,
                             command=lambda: controller.show_frame(StartPage))
        button1.place(x=130,y=350)


class Page2(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        label = ttk.Label(self, text="Image Format Convertor", font=('courier', 22))
        label.place(x=150, y=0)

        def openFile():
            global var, var1
            var = StringVar()
            var1 = StringVar()
            var1 = ""

            text_file_extensions = ["*.png","*.jpg","*.jpeg","*.jfif","*.gif","*.svg","*bmp","webp"]
            ftypes = [
                ('test files', text_file_extensions),

            ]



            var = askopenfilename(defaultextension=".png", filetypes=ftypes)


            l3.config(text=var)
            if var == "":
                file = None
        def convert():
            im = Image.open(var)
            print(im)
            base= os.path.basename(var)
            print(base)
            file= os.path.splitext(os.path.basename(base))[0]
            print(file)
            text = val.get()
            if text=='black':
                originalImage = cv2.imread(str(var),1)
                print(var)

                grayImage = cv2.cvtColor(originalImage, cv2.COLOR_BGR2GRAY)
                cv2.imwrite("{0}/{1}_grayscale.png".format(os. path. dirname(var),file),grayImage)
            else:

                rgb_im = im.convert("RGB")



                rgb_im.save("{0}/{1}.{2}".format(os. path. dirname(var),file,text))


            messagebox.showinfo('Success',"Converted successfully Check in same folder of image !!!")






        l2 = ttk.Label(self, text='Choose file: ')
        l2.place(x=0, y=50)

        l3 = ttk.Label(self, text="no file choosen")
        l3.place(x=150, y=48)

        openFileButton = ttk.Button(self, text=" Open ", width=10, command=openFile)
        openFileButton.place(x=80, y=47)

        l6=ttk.Label(self,text="Choose file to converted").place(x=0,y=80)

        val=StringVar()

        rb1=Radiobutton(self,text='jpg',value='jpg',var=val)
        rb2 = Radiobutton(self, text='jpeg', value='jpeg', var=val)
        rb3 = Radiobutton(self, text='png', value='png', var=val)
        rb4 = Radiobutton(self, text='jfif', value='jfif', var=val)
        rb5 = Radiobutton(self, text='bmp', value='bmp', var=val)
        rb6 = Radiobutton(self, text='gif', value='gif', var=val)
        rb7 = Radiobutton(self, text='tif', value='tif', var=val)
        rb8 = Radiobutton(self, text='Convert Gray scale', value='black', var=val)
        rb1.place(x=10,y=100)
        rb2.place(x=65, y=100)
        rb3.place(x=130, y=100)
        rb4.place(x=195, y=100)
        rb5.place(x=250, y=100)
        rb6.place(x=315, y=100)
        rb7.place(x=380, y=100)
        rb8.place(x=425, y=100)


        l4 = ttk.Button(self, text='Convert',command=convert)
        l4.place(x=120, y=135)

        bard = Image.open("reverse1.jpg")
        newsize = (350, 290)
        bard = bard.resize(newsize)
        bardejov = ImageTk.PhotoImage(bard)
        label1 = Label(self, image=bardejov)
        label1.image = bardejov
        label1.place(x=200, y=140)




        button2 = tk.Button(self, text="BACK",width=10,
                             command=lambda: controller.show_frame(StartPage))

        # putting the button in its place by
        # using grid
        button2.place(x=120,y=350)

class Page3(tk.Frame):
    def __init__(self, parent,controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller






        def text():
            text = str(text1)
            wordfile = asksaveasfile(mode='w', defaultextension=".doc",
                                     filetypes=[("word file", "*.doc"),
                                                ("text file", "*.txt"),
                                                ("PDF", "*.pdf")])

            if wordfile is None:
                return
            wordfile.write(text)
            wordfile.close()
            messagebox.showinfo('Success', "Written successfully !!!")


        def openFile():
            global var, var1
            var = StringVar()
            var1 = StringVar()
            var1 = ""

            text_file_extensions = ["*.txt"]
            ftypes = [
                ('test files', text_file_extensions),

            ]



            var = askopenfilename(defaultextension=".txt", filetypes=ftypes)


            l6.config(text=var)
            if var == "":
                file = None
        def convert():

            base= os.path.basename(var)
            file= os.path.splitext(os.path.basename(base))[0]
            file1 = open(var, "r",errors="ignore").read().replace("\n", " ")
            #print(file1)
            language = 'en'
            speech = gTTS(text=str(file1), lang=language, slow=False)
            speech.save("{0}\{1}_audio_book.mp3".format(os. path. dirname(var),file))

            messagebox.showinfo('Success', "Converted successfully Check in same folder of text file !!!")




        label = ttk.Label(self, text="Text to speech", font=('courier', 22))
        label.place(x=145, y=0)

        openFileButton = ttk.Button(self, text=" Choose ", width=10, command=openFile)
        openFileButton.place(x=85, y=40)

        l5 = ttk.Label(self, text="Select the file: ")
        l5.place(x=0, y=40)

        l6=ttk.Label(self,text="")
        l6.place(x=165,y=70)

        bt2= ttk.Button(self, text='Convert to mp3', command=convert)
        bt2.place(x=87, y=100)

        bard = Image.open("speechto.jpg")
        newsize = (270,200)
        bard = bard.resize(newsize)
        bardejov = ImageTk.PhotoImage(bard)
        label1 = Label(self, image=bardejov)
        label1.image = bardejov
        label1.place(x=280, y=260)















        button2 = tk.Button(self, text="BACK",width=10,
                             command=lambda: controller.show_frame(StartPage))

        button2.place(x=200, y=350)


class Page4(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)


        def open():
            text_file_extensions = ["*.pdf"]
            ftypes = [
                ('test files', text_file_extensions),

            ]
            global filez


            filez = askopenfilenames(defaultextension=".txt", filetypes=ftypes)
            filez=list(filez)

            l2.configure(text="{0} files selected".format(len(filez)))

        def convert():
            if filez :

                merger=PdfFileMerger()

                for pdf in filez:
                    merger.append(pdf)

                folder_selected = askdirectory()


                merger.write('{0}\Mergered_pdf.pdf'.format(folder_selected))
                merger.close()

                l2.config(text="no file choosen")
                messagebox.showinfo('Success', "Merged successfully!!!")
            else:
                messagebox.showinfo('Error', "No files selected")













        label = ttk.Label(self, text="Pdf Merger", font=('courier', 22))
        label.place(x=145, y=0)

        l1=ttk.Label(self,text="Select the files:")
        l1.place(x=0,y=50)
        labelx = ttk.Label(self, text=" Note: hold ctrl and select to make multiselect")
        labelx.place(x=10, y=200)

        bt1=ttk.Button(self,text="Open",command=open)
        bt1.place(x=90,y=48)

        l2=ttk.Label(self,text="No file selected")
        l2.place(x=180,y=49)

        bt2=ttk.Button(self,text="Convert",command=convert)
        bt2.place(x=91,y=90)

        bard = Image.open("pdfmerge.jpg")
        newsize = (350, 290)
        bard = bard.resize(newsize)
        bardejov = ImageTk.PhotoImage(bard)
        label1 = Label(self, image=bardejov)
        label1.image = bardejov
        label1.place(x=280, y=100)











        button2 = tk.Button(self, text="BACK",width=10,
                             command=lambda: controller.show_frame(StartPage))

        button2.place(x=120, y=350)




app = tkinterApp()
app.mainloop()
