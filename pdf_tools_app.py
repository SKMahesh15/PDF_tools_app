from tkinter import *
import os 
from pathlib import Path
from tkinter import filedialog
import win32com.client
from PyPDF2 import PdfWriter
import img2pdf
from PIL import Image
import pdfshrink
import fitz

class PDF_tools_app(Tk):

    def __init__(self):
        super().__init__()

        self.title("PDF Tools")
        self.app_size()

        self.show_main_page()
    
    def show_main_page(self):
        self.app_size()
        self.clear_window()
        bttn_pdf_merge = Button(self, text="Merger Pdf Files", command=self.bttn_merge_pdf)
        bttn_pdf_merge.place(x=0, y=0)
        bttn_pdf_merge.pack(pady=20)

        bttn_wrd_2_pdf = Button(self, text="Word to PDF", command=self.word_to_pdf)
        bttn_wrd_2_pdf.pack(pady=20)

        bttn_jpg_2_pdf = Button(self, text="JPEG to PDF", command=self.jpg_to_pdf)
        bttn_jpg_2_pdf.pack(pady=20)

        bttn_pdf_2_word = Button(self, text="PDF to Word", command=self.pdf_to_word)
        bttn_pdf_2_word.pack(pady=20)

        bttn_pdf_2_jpg = Button(self, text="PDF to PNG", command=self.pdf_to_jpg)
        bttn_pdf_2_jpg.pack(pady=20)

        bttn_2_compress = Button(self, text="Compress Pdf files", command=self.compress_files)
        bttn_2_compress.pack(pady=20)

        bttn_exit = Button(self, text="Exit", command=exit)
        bttn_exit.pack()

    def app_size(self):
        self.geometry("700x450")
    
    def output_path(self):
        output_path = filedialog.askdirectory(initialdir="/home/user/Documents")
        return output_path

    def clear_window(self):  
        for widget in self.winfo_children():
            widget.destroy()
    
    def browse_files(self, text, filetypes):
        file_path = filedialog.askopenfilenames(title=text, filetypes=filetypes)
        return file_path
    
    def back_main_screen(self):
        back_bttn = Button(self, text="Back", command=self.show_main_page)
        back_bttn.pack()
    
    def bttn_merge_pdf(self):
        self.clear_window()
        self.app_size()

        label = Label(self, text="Merge PDF files")
        label.pack(pady=10)

        self.list_box = Listbox(self, width=65, height=15)
        self.list_box.pack(pady=20)

        save_files = Button(self, text="Select Files", command=lambda:self.select_show_files(text="Select PDF files", filetypes=(("PDF Files", "*.pdf"), ("All Files", "*.*"))))
        save_files.pack()

        bttn_merge_pdf = Button(self, text="Merge PDF" , command=lambda:self.merge_pdfs(self.selected_files))
        bttn_merge_pdf.pack(pady=25)

        self.back_main_screen()

    def select_show_files(self, text, filetypes):
        self.filepaths = self.browse_files(text=text, filetypes=filetypes)

        self.list_box.delete(0, END)

        for self.file in self.filepaths:
            self.list_box.insert(END, self.file)

        self.selected_files = self.filepaths

        return self.filepaths

    def select_show_file_single(self, text, filetypes):
        self.filepaths = filedialog.askopenfilename(text=text, filetypes=filetypes)

        self.list_box.delete(0, END)

        for self.file in self.filepaths:
            self.list_box.insert(END, self.file)

        self.selected_files = self.filepaths

        return self.filepaths

    def merge_pdfs(self, filepaths):
        if self.filepaths is None:
            self.list_box.insert(END, "No Files selected")
        
        print(self.filepaths)
        
        merger = PdfWriter()
        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=(("PDF Files", "*.pdf"), ("All Files", "*.*")))
        for filepath in filepaths:
            merger.append(filepath)

        if output_path:
            merger.write(output_path)
            merger.close()
            self.list_box.delete(0, END)
            self.list_box.insert(END, f'Merged PDF saved at: {output_path}')
    
    def word_to_pdf(self):
        self.clear_window()
        self.app_size()

        label = Label(self, text="Word to PDF Converter")
        label.pack(pady=10)

        self.list_box = Listbox(self, width=65, height=10)
        self.list_box.pack(pady=20)

        select_files = Button(self, text="Select Word Files", command=lambda:self.select_show_files(text="Select Word Files", 
                                                                                                    filetypes=(("Doc Files", "*.docx"), 
                                                                                                               ("Word'.doc'", "*.doc"), 
                                                                                                               ("All Files", "*.*"))))
        select_files.pack(pady=25)

        convert_to_pdf = Button(self, text="Convert to PDF", command=self.word_to_pdf_cnvrt)
        convert_to_pdf.pack(pady=30)

        self.back_main_screen()
    
    def word_to_pdf_cnvrt(self):

        word_app = win32com.client.Dispatch('Word.Application')
        pdf_format = 17

        output_path = self.output_path()
        print(self.filepaths)
        
        for word in self.filepaths:
            new_word_file = os.path.normpath(word)
            file_name = str(os.path.basename(word)).split('.')[0]
            new_pdf_file = os.path.normpath(os.path.join(output_path, file_name + ".pdf"))
            
            try:
                in_file = word_app.Documents.Open(str(new_word_file))
                in_file.SaveAs(new_pdf_file, FileFormat=pdf_format)
                in_file.Close()
            except Exception as e:
                print(f'Error converting {word}: {e}')
        
        word_app.Quit()

        self.list_box.delete(0, END)
        self.list_box.insert(END, 'Converted all word files to PDF')
    
    def jpg_to_pdf(self):
        self.clear_window()
        self.app_size()

        label = Label(self, text="JPG to PDF Converter")
        label.pack(pady=10)

        self.list_box = Listbox(self, width=65, height=10)
        self.list_box.pack(pady=20)

        select_files = Button(self, text="Select Jpg Files", command=lambda:self.select_show_files(text="Select Images Files", 
                                                                                                    filetypes=(("JPG Files", "*.jpg"),
                                                                                                        ("All Files", "*.*"))))
        select_files.pack(pady=25)

        convert_to_jpg = Button(self, text="Convert to PDF", command=self.jpg_to_pdf_cnvrt)
        convert_to_jpg.pack(pady=30)

        self.back_main_screen()
    
    def jpg_to_pdf_cnvrt(self):

        output_path = self.output_path()
        print(self.filepaths)

        for file in self.filepaths:
            new_file_name = os.path.basename(file).split('.')[0]
            new_jpg_file = os.path.normpath(os.path.join(output_path, new_file_name + ".pdf"))

            try:
                image = Image.open(file)
                image.convert("RGB")
                image.save(new_jpg_file, "PDF", resolution=100.0)
                self.list_box.delete(0, END)
                self.list_box.insert(END, f'Converted to PDF')
            except Exception as e:
                print(f'Error converting {file}: {e}')
    
    def pdf_to_jpg(self):
        self.clear_window()
        self.app_size()

        label = Label(self, text="JPG to PDF Converter")
        label.pack(pady=10)

        self.list_box = Listbox(self, width=65, height=10)
        self.list_box.insert(END, "Please Insert Merged file only")
        self.list_box.pack(pady=20)

        select_files = Button(self, text="Select Pdf File", command=lambda:self.select_show_files(text="Select PDF File", 
                                                                                                    filetypes=(("PDF File", "*.pdf"),
                                                                                                        ("All Files", "*.*"))))
        select_files.pack(pady=25)

        convert_to_jpg = Button(self, text="Convert to PDF", command=self.pdf_to_jpg_cnvrt)
        convert_to_jpg.pack(pady=30)

        self.back_main_screen()
    
    def pdf_to_jpg_cnvrt(self):

        output_path = self.output_path()
        a = self.filepaths
        a = list(a)
        a = str(a[0])

        pdf_doc = fitz.open(a)
        
        for file in range(len(a)):
            try:
                page = pdf_doc.load_page(file)

                pix = page.get_pixmap()

                final_path = f"{output_path}/{(os.path.basename(a)).split('.')[0]} {file + 1}.png"
                final_path = os.path.normpath(final_path)
                print(final_path)
                pix.save(final_path)
            except Exception as e:
                print(f'Error converting {file}: {e}')

    def pdf_to_word(self):
        self.clear_window()
        self.app_size()

        label = Label(self, text="PDF to WORD Converter")
        label.pack(pady=10)

        self.list_box = Listbox(self, width=65, height=10)
        self.list_box.pack(pady=20)

        select_files = Button(self, text="Select Pdf File", command=lambda:self.select_show_files(text="Select PDF File", 
                                                                                                    filetypes=(("PDF File", "*.pdf"),
                                                                                                        ("All Files", "*.*"))))
        select_files.pack(pady=25)

        convert_to_jpg = Button(self, text="Convert to PDF", command=self.pdf_to_word_cnvrt)
        convert_to_jpg.pack(pady=30)

        self.back_main_screen()
    
    def pdf_to_word_cnvrt(self):
        
        output_path = self.output_path()
        print(self.filepaths)
        
        word_app = win32com.client.Dispatch('Word.Application')
        word_format = 16

        for word in self.filepaths:
            new_pdf = os.path.normpath(word)
            file_name = str(os.path.basename(word)).split('.')[0]
            new_pdf_file = os.path.normpath(os.path.join(output_path, file_name + ".docx"))
            
            try:
                in_file = word_app.Documents.Open(str(new_pdf))
                in_file.SaveAs(new_pdf_file, FileFormat=word_format)
                in_file.Close()
            except Exception as e:
                print(f'Error converting {word}: {e}')

        word_app.Quit()
    
    def compress_files(self):
        self.clear_window()
        self.app_size()

        self.clear_window()
        self.app_size()

        label = Label(self, text="PDF Compress")
        label.pack(pady=10)

        self.list_box = Listbox(self, width=65, height=10)
        self.list_box.pack(pady=20)

        select_files = Button(self, text="Select Pdf File", command=lambda:self.select_show_files(text="Select PDF File", 
                                                                                                    filetypes=(("PDF File", "*.pdf"),
                                                                                                        ("All Files", "*.*"))))
        select_files.pack(pady=25)

        convert_to_jpg = Button(self, text="PDF Compress", command=self.pdf_compresser)
        convert_to_jpg.pack(pady=30)

        self.back_main_screen()

    
    def pdf_compresser(self):

        output_path = self.output_path()
        print(self.filepaths)
        
        for file in self.filepaths:
            pdfshrink.compress(file, f"{output_path}.pdf")


if __name__ == "__main__":
    app = PDF_tools_app()
    app.mainloop()