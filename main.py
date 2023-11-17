from tkinter import filedialog, Tk, Label, Entry, Button, Canvas
from pdf2docx import Converter
from PIL import Image, ImageTk

class PDFtoWordConverter:
    def __init__(self, master):
        self.master = master
        self.master.title("PDF to Word Converter")

        # Load logo and resize it to 200x200
        self.logo_path = "D:\Leaning programing\python\PDFtoWord\logo.jpg"
        logo_image = Image.open(self.logo_path)
        logo_image = logo_image.resize((200, 200), Image.ANTIALIAS)
        self.logo_image = ImageTk.PhotoImage(logo_image)

        # Display logo
        canvas = Canvas(self.master, width=200, height=200)
        canvas.create_image(0, 0, anchor='nw', image=self.logo_image)
        canvas.grid(row=0, column=0, columnspan=2, pady=10)

        # Input entry
        self.pdf_path_entry = Entry(self.master, width=40)
        self.pdf_path_entry.grid(row=1, column=0, padx=10, pady=10)

        # Browse button
        browse_button = Button(self.master, text="Browse", command=self.browse_pdf)
        browse_button.grid(row=1, column=1, pady=10)

        # Convert button
        convert_button = Button(self.master, text="Convert", command=self.convert_pdf_to_word)
        convert_button.grid(row=2, column=0, columnspan=2, pady=10)

    def browse_pdf(self):
        pdf_path = filedialog.askopenfilename(title="Select PDF file", filetypes=[("PDF files", "*.pdf")])
        self.pdf_path_entry.delete(0, "end")
        self.pdf_path_entry.insert(0, pdf_path)

    def convert_pdf_to_word(self):
        pdf_path = self.pdf_path_entry.get()

        if pdf_path:
            # Convert PDF to Word
            word_path = filedialog.asksaveasfilename(title="Save Word file", defaultextension=".docx", filetypes=[("Word files", "*.docx")])

            if word_path:
                self.convert_pdf_to_docx(pdf_path, word_path)
                print("Conversion successful. Word file saved at", word_path)
            else:
                print("You canceled saving the Word file.")
        else:
            print("You canceled selecting the PDF file.")

    @staticmethod
    def convert_pdf_to_docx(pdf_path, output_path):
        cv = Converter(pdf_path)
        cv.convert(output_path)
        cv.close()

if __name__ == "__main__":
    root = Tk()
    app = PDFtoWordConverter(root)
    root.mainloop()
