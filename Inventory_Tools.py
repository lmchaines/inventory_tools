
#Uploading Tickets with Buttons Widgets

from tkinter import *
from tkinter import filedialog
from PIL import ImageTk, Image
import easygui
import pandas as pd
import pygetwindow
import pyautogui
from tkinter import ttk
import time
import os
import shutil
import tkinter as tk
from tkinter import ttk
import PyPDF2

##Test 


#Accelaration python.
pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0.01

#Initiation of root.
root=Tk()


root.title('Inventory Tools')
root.iconbitmap("D:\Media\Icons\Xylem.ico")
root.geometry("600x850") #To set the box size


style = ttk.Style()
style.theme_use("alt")  # You can experiment with different themes

my_notebook= ttk.Notebook(root)
my_notebook.grid(row=7, column=0, pady=10,sticky=EW)


#Xylem image
img = PhotoImage(file="C:/Users/Lawrence.McHaines/OneDrive - Xylem, Inc/Documents/Inventory Management/190 Python/Xylem2.png")
img1 =Label(root, image=img, anchor= CENTER).grid(row=0,column=0,sticky=N)


#Heather and Packing
label_1 =Label(root,text="STEEB INVENTORY TOOLS",bg = "black",fg="#00EE00",font=("Calibri", 13))
label_1.grid(row=1, column=0, sticky=EW)
label_2 =Label(root,text="Version 1.0",font=("Calibri", 8))
label_2.grid(row=2, column=0, pady=5,sticky=EW)


#Frame and Packing
my_frame1=Frame(my_notebook,width=600, height=700, pady=20)
my_frame1.grid()
my_frame2=Frame(my_notebook,width=600, height=700, pady=20)
my_frame2.grid()
my_frame3=Frame(my_notebook,width=600, height=700,pady=20)
my_frame3.grid()


#Frame Heather
my_notebook.add(my_frame1, text="PI Upload  ")
my_notebook.add(my_frame2, text="FSS Upload  ")
my_notebook.add(my_frame3, text="Bin Location Maintenance  ")




#Functions assing to widgets

#############Frame PI upload#########

####Batches Tickets Upload ###### 
def opt2():

    #set the path for Pre-Printed Tickets option 2
  
    path = 'C:\\Temp\\ticket2.png'

    window = pygetwindow.getWindowsWithTitle('STEEB')[0]

    #Difine the FSS windows size
    x1 = window.left
    y1 = window.top
    height = window.height
    width = window.width

    x2 = x1 + width
    y2 = y1 + height
    
    # Captures a screenshot of the screen or a specified region of the screen.
    pyautogui.screenshot(path)
    
    
    # This code opens an image, crops a specified region from it, and then saves the cropped version back to the same 
    # file path, effectively replacing the original image with the cropped one. 
    # The cropping region is defined by the (x1, y1, x2, y2) coordinates, and the code assumes 
    # that the image at the path location exists and can be opened and saved successfully.
    
    im=Image.open(path)
    im=im.crop((x1, y1, x2, y2))
    im.save(path)
  
    #to locate the image on the screen
    im=pyautogui.locateOnScreen('C:\\Temp\\ticket2.png')

    #Initial first contact
    pyautogui.leftClick(x1+287, y1+300)

    #Get the data from an Excel File
    df4=[]
   
    myexcelfile=easygui.fileopenbox(msg="Getting your data",title='Batches Tickets Upload',default='C:/Users/Lawrence.McHaines/Desktop/*.*')
    temp=pd.ExcelFile(myexcelfile)
    df4=temp.parse('Ticket2',dtype = str)



    #Start uploading FSS data & Looping through the data
    for index,row in df4.iterrows():
     
        flnb05=str(row['WHSE'])
        flnb06=str(row['FROM_LOC'])
        flnb07=str(row['TO_LOC'])
    
        pyautogui.typewrite('2')
        pyautogui.typewrite(['enter'])
    
    
        pyautogui.write(flnb05)  #Send WHSE
        pyautogui.press(['tab'])
    
        pyautogui.write(flnb06) #Send From Loc
        pyautogui.press(['tab'])
    
        pyautogui.write(flnb07) #Send to Loc
        pyautogui.press(['tab'])
    
        pyautogui.typewrite(['f5']) #Update the record


####Ticket Option 6 ######   

def opt6():

    #Get the data from an Excel File
    
    
    #set the path for Ticket Option 6

    path ='C:\\Temp\\ticket6.png'


    window = pygetwindow.getWindowsWithTitle('STEEB')[0]

    #Difine the FSS windows size
    x1 = window.left
    y1 = window.top
    height = window.height
    width = window.width

    x2 = x1 + width
    y2 = y1 + height

    #pyautogui.screenshot(path)
    pyautogui.screenshot(path)
    im=Image.open(path)
    im=im.crop((x1, y1, x2, y2))
    im.save(path)
    
    im=pyautogui.locateOnScreen('C:\\Temp\\ticket6.png')

    #Initial first contact with FSS
    pyautogui.leftClick(x1+334, y1+219)
    
    df2=[]
    myexcelfile=easygui.fileopenbox(msg="Getting your data",title='Ticket Option 6',default="C:/Users/Lawrence.McHaines/Desktop/*.*")
    temp=pd.ExcelFile(myexcelfile)
    df2=temp.parse('Ticket6',dtype = str)
    


    
    #Start uploading FSS data & Looping through the data
    for index,row in df2.iterrows():
        
        
        flnb02=str(row['QUANTITY'])
        pyautogui.press(['end']) # Clear line
        pyautogui.write(flnb02) #Send Item
        pyautogui.press(['tab'])
        pyautogui.press(['pagedown'])
        
        #my_label.config(text=my_progress['value'])
        #my_progress['value']+=1
        #root.update_idletasks() #To see has it is hapenning.
        
        
    
####Ticket Option 5 ######    
    
def opt5():

    #set the path for Ticket Option 5
 
    path ='C:\\Temp\\ticket5.png'
  
    window = pygetwindow.getWindowsWithTitle('STEEB')[0]

    #Difine the FSS windows size
    x1 = window.left
    y1 = window.top
    height = window.height
    width = window.width

    x2 = x1 + width
    y2 = y1 + height

    pyautogui.screenshot(path)
    im=Image.open(path)
    im=im.crop((x1, y1, x2, y2))
    im.save(path)


    im=pyautogui.locateOnScreen('C:\\Temp\\ticket5.png')

    #Initial first contact with FSS
    pyautogui.leftClick(x1+372, y1+576)
    
    #Get the data from an Excel File
    df3=[]
    myexcelfile=easygui.fileopenbox(msg="Getting your data",title='Ticket Option 5 Upload',default="C:/Users/Lawrence.McHaines/Desktop/*.*")
    temp=pd.ExcelFile(myexcelfile)
    df3=temp.parse('Ticket5',dtype = str)

    #Start uploading FSS data & Looping through the data
    for index,row in df3.iterrows():
     
       
        flnb02=str(row['TICKET_NO'])
        flnb03=str(row['ITEM_NUMBER'])
        flnb04=str(row['QUANTITY'])
    
        pyautogui.write(flnb02)  #Send Ticket No.
        pyautogui.press(['tab'])
    
        pyautogui.write(flnb03) #Send Item No.
        pyautogui.press(['tab'])
    
        pyautogui.write(flnb04) #Send Qty.
        pyautogui.press(['tab'])
        pyautogui.typewrite(['enter'])
        
        
####Files rename  ######

def filerename():
    global root1
    root1=Toplevel()
    root1.title('Inventory Tools')
    root1.iconbitmap("D:\Media\Icons\Xylem.ico")
    root1.geometry("603x850") #To set the box size
    
    # Main Image
    img8 = tk.PhotoImage(file="C:/Users/Lawrence.McHaines/Desktop/Python_Projects/Xylem10.png")
    img9 = tk.Label(root1, image=img8, anchor="center")
    img9.grid(row=0, column=0, columnspan=3)
    
    
    # Progress for the file renaming
#     def update_progress_cc():
#         global current_value2
#         if current_value2 <= 100:
#             progress2["value"] = current_value2
#             progress_label_cc.config(text=f"Progress: {current_value2}%")
#             current_value2 += 1
#             root1.after(2, update_progress_cc)  # Update every 50 milliseconds
#         else:
#             progress_label_cc.config(text="Files rename complete!")


#     def start_progress_cc():
#         global current_value2
#         current_value2 = 0
#         update_progress_cc()


    #Setup the title and frame.
#     global progress2
#     global progress_label_cc


    style = ttk.Style()
    style.theme_use("alt")  # You can experiment with different themes

#     progress_color2 = "green"
#     style.configure("TProgressbar", thickness=38, background=progress_color2)

    label1 = tk.Label(root1, text="FILES RENAMING FOR PI", bg="#31658A", fg="white", font=("Calibri", 13))
    label1.grid(row=1, column=0, sticky="ew", columnspan=3) 


    def rename_and_move_files(directory):
        target_strings = {                
            "PIRPG21": "Whse {} {} {} Warehouse Summary Report.txt",
            "PIRPG90": "Whse {} {} {} Ticket Cost Verification.txt",
            "PIRPG50": "Whse {} {} {} +- $100 or $500 Report.txt",
            "PIRPG15": "Whse {} {} {} Count by Item.txt",
            "PIRPG20": "Whse {} {} {} Variance Report.txt",
            "PIRPG22": "Whse {} {} {} Company summary Report.txt",
            "PIRPG28": "Whse {} {} {} Accuracy Report.txt"
        }

        # Variables to store warehouse and status, initially unknown
        warehouse = "Unknown"
        status = "Unknown"
        renamed_files = []

        for filename in os.listdir(directory):
            full_path = os.path.join(directory, filename)

            if os.path.isfile(full_path):
                try:
                    with open(full_path, 'r', encoding='utf-16-le') as f:
                        contents = f.read()

                    for target_string, new_format in target_strings.items():
                        if target_string in contents:

                            # Extract warehouse number and status from file contents
                            warehouse = extract_warehouse(contents)
                            status = determine_status(contents)
                            status1 =uscan_status(contents)

                            # Construct the new filename
                            new_filename = new_format.format(status1, warehouse, status)
                            new_full_path = os.path.join(directory, new_filename)

                            if not os.path.exists(new_full_path):
                                os.rename(full_path, new_full_path)
                                renamed_files.append(new_full_path)
                                print(f"File {filename} renamed to {new_filename}")
                            else:
                                print(f"Cannot rename {filename} to {new_filename} because a file with the new name already exists.")
                            break
                except UnicodeDecodeError as e:
                    print(f"Error reading {filename}: {e}")

        # After renaming all files, create the directory with the name of the warehouse and status
        if warehouse != "Unknown" and status != "Unknown":
            new_directory = f"{warehouse}_{status}"
            new_directory_path = os.path.join(directory, new_directory)
            os.makedirs(new_directory_path, exist_ok=True)  # Create the new directory if it doesn't exist

            # Move the renamed files into the new directory
            for file_path in renamed_files:
                shutil.move(file_path, new_directory_path)

        # Reset the check box to zero
        ch5var.set(0)
        ch6var.set(0)
        ch7var.set(0)
        
#          #start_progress_cc()
#         label_frame_cc(text="Files rename complete!")


    def extract_warehouse(contents):
        whsget=whs.get()
        return whsget
    
    # Define status1 as a global variable at the beginning of your code
    status1 = "Unknown"

    def determine_status(contents):
        global status1
        state_map = {
            ch5var.get(): 'Temp',  #Temporary
            ch6var.get(): 'Perm',  #Permanen
            ch7var.get(): 'All',  #All
        }

        for state, value in state_map.items():
            if state == 1:
                status1 = value            
                break
        wshstatus = status1
        return status1


    def uscan_status(contents):
        global status2
        state1_map = {
            ch1var.get(): 'CA',  
            ch2var.get(): 'US',  
        }

        for state1, value in state1_map.items():
            if state1 == 1:
                status2 = value            
                break
        uscan = status2
        return status2


    def selectfile():
        # Use easygui to select the directory
        selected_directory = easygui.diropenbox(title="Select a directory to process files", default="C:/Users/Lawrence.McHaines/Desktop/")

        if selected_directory:
            rename_and_move_files(selected_directory)
        else:
            print("No directory selected. Exiting...")

       

    #Selection check box
    global ch1var
    ch1var=tk.IntVar()
    ch1=Checkbutton(root1, text="Co.13",variable=ch1var)
    ch1.deselect()
    ch1.grid(row=2,column=0, padx=50,pady=20,sticky="w")

    global ch2var
    ch2var=tk.IntVar()
    ch2=Checkbutton(root1, text="Co.14",variable=ch2var)
    ch2.deselect()
    ch2.grid(row=2,column=1, padx=1,pady=20,sticky="w")

    ch3=Checkbutton(root1, text="Text Format")
    ch3.deselect()
    ch3.grid(row=3,column=0, padx=50,pady=20,sticky="w")

    ch4=Checkbutton(root1, text="PDF Format")
    ch4.deselect()
    ch4.grid(row=3,column=1, padx=1,pady=20,sticky="w")

    global ch5var
    ch5var=tk.IntVar()
    ch5=Checkbutton(root1, text="Temporary",variable=ch5var)
    ch5.deselect()
    ch5.grid(row=4,column=0, padx=50,pady=20,sticky="w")

    global ch6var
    ch6var=tk.IntVar()
    ch6=Checkbutton(root1, text="Permanent",variable=ch6var)
    ch6.deselect()
    ch6.grid(row=4,column=1, padx=1,pady=20,sticky="w")

    global ch7var
    ch7var=tk.IntVar()
    ch7=Checkbutton(root1, text="All",variable=ch7var)
    ch7.deselect()
    ch7.grid(row=4,column=2, padx=60,pady=20,sticky="w")

    labe1 = tk.Label(root1, text="SELECT THE WAREHOUSE :", font=("Calibri", 12))
    labe1.grid(row=5, column=0, padx=50, pady=10,sticky="w") 

    # Insert the warehouse
    whs=Entry(root1, width=5)
    whs.get()
    whs.grid(row=5, column=1, padx=1, pady=10, sticky="w")


    btn10 = tk.Button(root1, text="Rename Files",command=selectfile)
    btn10.config(width=20, height=2, bg="#44759E", fg="white", font=("Calibri", 10))
    btn10.grid(row=7, column=0, padx=50, pady=50, sticky="w")  

#     progress2 = ttk.Progressbar(root1, orient="horizontal", length=200, mode="determinate",style="TProgressbar")
#     progress2.grid(row=7, column=1, sticky="w", columnspan=2, padx=10, pady=1)

#     label_frame_cc = tk.Frame(root1)
#     label_frame_cc.grid(row=6, column=1, columnspan=2,padx=10,pady=1,sticky="w")

    # Overlay the progress label on the frame
#     progress_label_cc = tk.Label(label_frame_cc, font=("Calibri", 10))
#     progress_label_cc.grid(row=6, column=1, columnspan=2,padx=30,pady=1,sticky="w")

    #Exit and Packing
    btn11=Button(root1, text="Exit Program",command=root1.destroy)
    btn11.grid(row=8, column=2, padx=50, pady = 55,sticky="e")

    root1.mainloop()


#############Frame FSS Upload#########       

def fssup():
    #Get the data from an Excel File
    df1=[]
    myexcelfile=easygui.fileopenbox(msg="Getting your data",title='FSS Upload',default='C:/Users/Lawrence.McHaines/Desktop/*.*')
    temp=pd.ExcelFile(myexcelfile)
    df1=temp.parse('FSS',dtype = str)


    #Find the FSS Screen
    path ='C:\\Temp\\fss1.png'

    #window = pygetwindow.getWindowsWithTitle('Session A-[24x80]')[0]
    window = pygetwindow.getWindowsWithTitle('STEEB')[0]

    #Difine the FSS windows size
    x1 = window.left
    y1 = window.top
    height = window.height
    width = window.width

    x2 = x1 + width
    y2 = y1 + height

    pyautogui.screenshot(path)
    im=Image.open(path)
    im=im.crop((x1, y1, x2, y2))
    im.save(path)


    im=pyautogui.locateOnScreen('C:\\Temp\\fss1.png')

    #Initial first contact with FSS
    pyautogui.rightClick(x1+584, y1+417)


    #Start uploading FSS data & Looping through the data
    for index,row in df1.iterrows():
     
        flnb02=str(row['FLNB02']) #Item
        flnb03=str(row['FLNB03']) #Branch
        flnb16=str(row['FLNB16']) #Supplier
        flnb24=str(row['FLNB24']) #Transport mode       
        flnb38=str(row['FLNB38']) #Transit Time
        flnb39=str(row['FLNB39']) #Total Lead Time

        pyautogui.press(['end']) # Clear line
        pyautogui.write(flnb02) #Send Item
        pyautogui.typewrite(['tab']) #Jump to second field
        pyautogui.press(['end'])  #Clear line
        pyautogui.typewrite(flnb03) #Send branch
        pyautogui.typewrite(['f5']) #Go to second screen in update mode


        # Posisionate to supplier
        pyautogui.press('tab', presses=11) #Move on second screen
        pyautogui.press(['end']) #Clear line
        pyautogui.write(flnb16) #Send supplier
        pyautogui.press('tab', presses=4) #Move on second screen


        pyautogui.press(['end']) #Clear line
        pyautogui.write(flnb24) #Send transport mode
        pyautogui.write(flnb38) #Send transit time
        pyautogui.write(flnb39) #Send total lead time

      
        pyautogui.press(['end']) #Clear second line transport mode

        pyautogui.typewrite(['tab']) #Move to other field
        pyautogui.press(['end']) #Clear second line transit time

        pyautogui.typewrite(['tab']) #Move to other field
        pyautogui.press(['end'])  #Clear second line total transit time

        pyautogui.press(['enter', 'enter']) #update field and return to the 1st screen.  
        
 

def discon():
    #Get the data from an Excel File
    df6=[]
    myexcelfile=easygui.fileopenbox(msg="Getting your data",title='FSS Discontinued',default='C:/Users/Lawrence.McHaines/Desktop/*.*')
    temp=pd.ExcelFile(myexcelfile)
    df6=temp.parse('Discontinued',dtype = str)


    #Find the FSS Screen
    path ='C:\\Temp\\Discontinued.png'

    window = pygetwindow.getWindowsWithTitle('STEEB')[0]
        
    #Difine the FSS windows size
    x1 = window.left
    y1 = window.top
    height = window.height
    width = window.width

    x2 = x1 + width
    y2 = y1 + height

    pyautogui.screenshot(path)
    im=Image.open(path)
    im=im.crop((x1, y1, x2, y2))
    im.save(path)


    im=pyautogui.locateOnScreen('C:\\Temp\\Discontinued.png')

    #Initial first contact with FSS
    pyautogui.rightClick(x1+584, y1+417)

    
    im1=pyautogui.locateOnScreen('C:\\Temp\\error053.png')
    
    #Start uploading FSS data & Looping through the data
    for index, row in df6.iterrows():
     
        flnb02=str(row['FLNB02']) #Item
        flnb03=str(row['FLNB03']) #Branch
        date01=str(row['DATE01']) #Date discontinued       
 

        pyautogui.press(['end']) # Clear line
        pyautogui.write(flnb02) #Send Item
        pyautogui.typewrite(['tab']) #Jump to second field
        pyautogui.press(['end'])  #Clear line
        pyautogui.typewrite(flnb03) #Send branch
        pyautogui.typewrite(['f5']) #Go to second screen in update mode

        
        
        if im1==None:
            # Posision to the end date field
            pyautogui.press('tab') #Move on second field
            #pyautogui.press(['end']) #Clear line
            pyautogui.write(date01) #Send discontinued date
            #pyautogui.press('tab',presses=9) #Move on second field
            pyautogui.press(['enter', 'enter']) #update field and return to the 1st screen.
        else:
            with pyautogui.hold('shift'):
                pyautogui.typewrite(['f12']) #Go back
                #pyautogui.write('1') #Move on second field
                #pyautogui.press(['enter'])
        
        

#############Frame Bin Location Maintenance ######### 

def add_item_loc():
    #Get the data from an Excel File
    df7=[]
    myexcelfile=easygui.fileopenbox(msg="Getting your data",title='Add New Bin Locastion',default='C:/Users/Lawrence.McHaines/Desktop/*.*')
    temp=pd.ExcelFile(myexcelfile)
    df7=temp.parse('Binloc',dtype = str)


    #Find the FSS Screen
    path ='C:\\Temp\\Binloc.png'

    window = pygetwindow.getWindowsWithTitle('STEEB')[0]
        
    #Difine the FSS windows size
    x1 = window.left
    y1 = window.top
    height = window.height
    width = window.width

    x2 = x1 + width
    y2 = y1 + height

    pyautogui.screenshot(path)
    im=Image.open(path)
    im=im.crop((x1, y1, x2, y2))
    im.save(path)


    im=pyautogui.locateOnScreen('C:\\Temp\\Binloc.png')

    #Initial first contact with stock location maintenance
    pyautogui.rightClick(x1+529, y1+289)

    
    im1=pyautogui.locateOnScreen('C:\\Temp\\error053.png')
    
    #Start uploading the bin location data & Looping through the data
    for index, row in df7.iterrows():
     
        ln08a=str(row['LN08']) #Branch
        lps6a=str(row['LPS6']) #Bin location
        ln05a=str(row['LN05']) #Item to add      
 
        #First Screen
        pyautogui.press(['end']) # Clear line
        pyautogui.write(ln08a) #Send Branch
        pyautogui.press(['end'])  #Clear line
        pyautogui.typewrite(lps6a) #Bin location
        pyautogui.press(['f5']) #Go to second screen in update mode
        pyautogui.press(['enter']) #update field and go to sencond field

        
        #Adding new items second screen
        pyautogui.press('down',presses=10)
        pyautogui.press('tab') #Move on second field
        pyautogui.write(ln05a) #input item
        pyautogui.press(['f6']) #Adding new items
        pyautogui.press(['f1']) #Return to first screen



def tranf_item_loc():
    return

def del_item_loc():
    return


#Action of widget and functions and packing

#Frame PI Upload and Packing
btn1=Button(my_frame1,text="PRE-PRINTED TICKET OPT 2",command=opt2, padx=31, pady=20)
btn2=Button(my_frame1,text="TICKET MAINTENACE OPT 6",command=opt6, padx=30, pady=20)
btn3=Button(my_frame1,text="TICKET DATA ENTRY OPT 5 ",command=opt5, padx=30, pady=20)
btn3a=Button(my_frame1,text="FILES RENAME FOR PI          ", command=filerename, padx=30, pady=20)
btn1.grid(row=4, column=0,pady = 5, padx=20)
btn2.grid(row=5, column=0,pady = 5, padx=20)
btn3.grid(row=6, column=0,pady = 5, padx=20)
btn3a.grid(row=7, column=0,pady = 5, padx=20)

#Frame FSS Upload and Packing
btn5=Button(my_frame2,text="        FSS Upload       ",command=fssup, padx=40, pady=20)
btn6=Button(my_frame2,text="Discontined Upload",command=discon, padx=40, pady=20)
btn5.grid(row=4, column=0,pady = 5, padx=20)
btn6.grid(row=5, column=0,pady = 5, padx=20)

#Frame Bin Location Maintenance and Packing
btn7=Button(my_frame3,text="Add New Item In BinLoc",command=add_item_loc, padx=31, pady=20)
btn8=Button(my_frame3,text="Transfer Stock to BinLoc",command=tranf_item_loc, padx=30, pady=20)
btn9=Button(my_frame3,text="Delete Item In BinLoc     ",command=del_item_loc, padx=30, pady=20)
btn7.grid(row=4, column=0,pady = 5,padx=20)
btn8.grid(row=5, column=0,pady = 5, padx=20)
btn9.grid(row=6, column=0,pady = 5, padx=20)


#Exit and Packing
btn4=Button(root, text="Exit Program",command=root.destroy)
btn4.grid(row=8, column=0, padx=30, pady = 10, sticky="e")


root.mainloop()

