from tkinter import *
from tkinter.filedialog import askopenfilename, askdirectory
from tkinter import messagebox
import webbrowser
import time
import asyncio
from WalmartScraper import ExcelScraper

# Define the functions to be called on button clicks
class App:
    def __init__(self, root):
        self.filePath = ""
        self.folderPath = ""
        self.selected_retailer = ""
        self.emailTo = ""
        self.tempCnt = 0

    def checkValid(self):
        
        self.emailTo = mail_txtbox.get("1.0", "end-1c").lower()
        mail_txtbox.configure(state='disabled')

        # if self.emailTo != "" and self.filePath != "" and self.folderPath != "" :
        if self.filePath != "" and self.folderPath != "" :
            
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            loop.run_until_complete(self.compare_generate_excel())

        # elif self.emailTo == "":
        #     messagebox.showwarning(title="Warning", message="Please add a email address of the receiver")
        elif self.filePath == "" and self.folderPath == "":
            messagebox.showwarning(title="Warning", message="Please select a source file and destination folder")
        elif self.filePath == "":
            messagebox.showwarning(title="Warning", message="Please select a source file")
        elif self.folderPath == "":
            messagebox.showwarning(title="Warning", message="Please select a destination folder")

    def Reset(self):
        # Reset variables
        self.filePath = ""
        self.folderPath = ""
        self.selected_retailer = ""
        self.emailTo = ""

        # Reset the labels and textboxes
        mail_txtbox.delete("1.0","end")
        file_label.configure(text="No file selected.", fg="#FF0000")
        dest_label.configure(text="No destination folder selected", fg="#FF0000")
        progress_label.configure(text="")
        elapsedTime_label.configure(text="Time Taken : ???")

        root.update()

    def upload_file(self):
        self.filePath = askopenfilename(filetypes=[("Excel files", "*.xlsx"),("All files","*.*")])

        lbl_text = "Source file path : " + self.filePath
        file_label.configure(text=lbl_text, wraplength=800,fg="#000")

    def choose_destination(self):
        self.folderPath = askdirectory()

        lbl_text = "Destination Path : " + self.folderPath
        dest_label.configure(text=lbl_text, wraplength=800, fg="#000")
        self.folderPath = f"{self.folderPath}"

    def open_file(self):
        if messagebox.askyesno("Open File", "Do you want to open the file?"):
            webbrowser.open(f"{self.folderPath}/AuditSheet.xlsx")

    async def compare_generate_excel(self):
        progress_label.configure(text="Starting...")
        root.update()

        # Process start time
        #st = time.process_time()
        start_time = time.time()
        
        gen_button.configure(state='disabled')

        es = ExcelScraper(emailTo=self.emailTo,sourceFilePath=self.filePath, destinationFilePath=self.folderPath, newFileName="AuditSheet")
        es.get_url()
        
        await asyncio.gather(
            es.scrape_product_data(),
            self.update_counter(es)
        )
        # Assign the result of scrape_product_data() to es.dataList
        # es.dataList = await es.scrape_product_data()

        es.main()
        
        print("Completed!!!")
        progress_label.configure(text="Completed Inserting Data")
        root.update()

        # process end time
        # e = time.process_time()
        end_time = time.time()

        elapsed_time = end_time - start_time
        
        # lbl_ET = f'Time taken: {elapsed_time:.2f} seconds'
        lbl_ET = f'Time taken: {elapsed_time/60:.2f} minutes'
        elapsedTime_label.configure(text=lbl_ET)

        messagebox.showinfo("Excel Generated", "Excel generated successfully")
        progress_label.configure(text="Excel Generated Succesfully!")
        root.update()
        
        # Enable the gen_button and update the window title
        gen_button.configure(state='normal')
        root.title("Scrapematch Tool - Complete")

        open_button.configure(state="normal")

    async def update_counter(self, es):
        # print(es.stopFlag, es.counter)
        while es.stopFlag:
            cnt = es.counter
            totalProducts = es.total_products
            # print(f"cnt : {cnt}/{totalProducts}")
            progress_label.configure(text=f"Scraping & Inserting Data... {cnt}/{totalProducts}")
            root.update()
            await asyncio.sleep(0.1)

    def start_generate(self):
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        loop.run_until_complete(self.compare_generate_excel())

    def Close(self):
        root.destroy()

root = Tk()
app = App(root)
root.title("Scrapematch Tool")
# set window size
root.geometry("850x530")

# Set the background color and font for the root window
root.configure(bg="#f5f5f5")
root.option_add("*Font", "Verdana 10")

# Create a label to describe the purpose of the GUI
description_label = Label(root, text="This program scrapes data, compares the original with scraped sheet, and creates a sheet with differences")
description_label.configure(bg="#f5f5f5", fg="#333", font="Helvetica 12 bold", padx=20, pady=20)
description_label.grid(row=0, column=0, columnspan=2)

mail_label = Label(root, text="Enter Your Email: ")
mail_label.configure(bg="#f5f5f5", fg="#333", font="Helvetica 12 bold", padx=20, pady=20)
mail_label.grid(row=1, column=0, padx=20, pady=10)

mail_txtbox = Text(root, height=1, width=50, bg="light cyan")
mail_txtbox.grid(row=1, column=1, padx=20, pady=10, sticky=W, columnspan=2)

# Button to upload a file
upload_button = Button(root, text="Upload File", command=app.upload_file)
upload_button.configure(bg="#007bff", fg="#fff", font="Helvetica 10 bold", padx=10, pady=5)
upload_button.grid(row=2, column=0, padx=10, pady=10)
#tooltip.ToolTip(upload_button, "Upload a file to compare")

# Button to choose the destination for the generated Excel file
dest_button = Button(root, text="Choose Destination", command=app.choose_destination)
dest_button.configure(bg="#007bff", fg="#fff", font="Helvetica 10 bold", padx=10, pady=5)
dest_button.grid(row=3, column=0, padx=10, pady=10)
#tooltip.ToolTip(dest_button, "Choose a destination folder")

# label to show the uploaded file name
file_label = Label(root, text="No file selected.",fg="#FF0000")
file_label.grid(row=4, column=0,columnspan=2,padx=20,  pady=(20,10), sticky=W)

# label to show the selected destination folder
dest_label = Label(root, text ="No destination folder selected", fg="#FF0000")
dest_label.grid(row=5, column=0,columnspan=2,padx=20, pady=10, sticky=W)

# label to show the status
status_label = Label(root, text="Status : ")
status_label.grid(row=6, column=0, padx=20, pady=10, sticky=W)

# label to show the progress
progress_label = Label(root, text="???", fg="#059669")
progress_label.grid(row=6, column=0, padx=20, pady=10)

# label to show out of products
#outOf_label = Label(root, text="", fg="#f43f5e")
#outOf_label.grid(row=5, column=1, padx=20, pady=10, sticky=W)

# label to show time take to complete process
elapsedTime_label = Label(root, text="Time Taken : ???", fg="#000")
elapsedTime_label.grid(row=7, column=0, padx=20, pady=10, sticky=W)

# button to open file
open_button = Button(root, text="Open File", command=app.open_file, bg="#4CAF50", fg="white", pady=10, padx=20)
open_button.grid(row=8, column=0,padx=20, pady=10, sticky=W)
open_button.configure(state="disabled")

# reset button
reset_button = Button(root, text="Reset", command=app.Reset, bg="#a5f3fc", pady=10, padx=20, font=("Helvetica",10, "bold"))
reset_button.grid(row=8, column=1, padx=20, pady=10, sticky=W)

# exit button
exit_button = Button(root, text="Exit", command=app.Close, bg="#ef4444", fg="white", pady=10, padx=20, font=("Helvetica",10, "bold"))
exit_button.grid(row=8, column=1, padx=20, pady=10, sticky=E)

# dropdown menu with retailer options
options = ["Walmart", "Amazon", "Kroger", "Target"]
app.selected_retailer = StringVar(root)
app.selected_retailer.set(options[0]) # Set the default option
retailer_dropdown = OptionMenu(root, app.selected_retailer, *options)
retailer_dropdown.configure(font=("Helvetica", 10, "bold"), bg="white", fg="black", padx=10, pady=5)
retailer_dropdown.grid(row=2, column=1, padx=(20,0), pady=(0,20))

# button to generate the Excel file
# gen_button = Button(root, text="Compare & Generate Excel", command=app.compare_generate_excel, bg="#4CAF50", fg="white", pady=10, padx=20, font=("Helvetica", 10, "bold"))
gen_button = Button(root, text="Compare & Generate Excel", command=app.checkValid, bg="#4CAF50", fg="white", pady=10, padx=20, font=("Helvetica", 10, "bold"))
gen_button.grid(row=3, column=1, padx=(20,0), pady=(0,20))

# # label to credit the developer
# credit_label = Label(root, text="Developed by SACHIN", font=("Helvetica", 8))
# credit_label.grid(row=8, column=0, columnspan=2, pady=(20,10))

root.mainloop()