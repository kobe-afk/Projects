#!/usr/bin/env python
# coding: utf-8

# In[1]:


from IPython.core.display import display, HTML
display(HTML("<style>.container { width:95% !important; }</style>"))


# In[2]:


import win32print
import win32ui
import PIL

import os 
import pandas as pd
import tkinter as tk
from tkinter import *   
import tkinter.font as font
from tkinter import filedialog     # for the browse files
from tkinter import messagebox     # for warning messages
from tkinter import ttk            # for the treeview


import pyautogui                   # for the screenshot
import math

# from pyautogui import screenshotUtil      # uncomment this for the clean room 

###### for printing out

import win32print 
import win32ui
from PIL import Image, ImageWin

##### For image

from datetime import date
from datetime import datetime
import datetime


from collections import Counter  # check for which trays have duplicates


# In[3]:


# ensure to change printing points and check print out pictures in setup folder

# change current directory to clean room current directory


# In[4]:



x_labelframe_pass_fail_start_pos = 20

y_labelframe_height_pos_pass_Fail_labelframe = 130


########### browse csv frame 

x_labelframe_start_pos = 220                     # place where all the labelframes will be positioned according to the x axis
y_labelframe_height_pos = 15                     # height the labels and the entry boxes will be positioned at for the y axis


x_label_pos = 20
y_pos = 15
y_pos_browse_csv = 10

x_entry_pos = 250                           # x position of the Lot ID entry box 
x_browse_file_button_pos = 1090

browse_csv_button_width = '20'

entry_file_width = '80'

print_out_button_x_position = 820
print_out_button_y_position = 929


#################### Label frame widths

Browse_csv_cognex_labelframe_height = "90"

## if check_map_height is changed , check_map_canvas_height must also be changed but must be less than the check map height

check_map_height = "800"             # height of the pass fail indicator labelframe
check_map_canvas_height = "753"
all_width = "1500"

pass_fail_indicator_frame_width = "1860"
check_map_canvas_width = "1840"


########## product_check_map

frame_x_pos = 30
frame_y_pos = 17




###############  legend_for_pass_fail

place_legend_frame_x_pos = 500
legend_frame_y_pos = 670


legend_frame_height = '79'
legend_frame_width = '800'

## button positions in the legend labelframe (y)

y_labels_legend_pos = 0
y_pass_fail_func_widget_pos = -2
y_no_unit_widget_pos = -2

## button positions in the legend labelframe (x)

x_legend_label_pos_Bin_A = 55
x_legend_label_pos_Bin_C = 250
x_legend_label_pos_no_unit = 650


x_legend_button_pos_Bin_A = 185 
x_legend_button_pos_Bin_C = 584
x_legend_button_no_unit = 730


######## Wafer_Creation global variables

total_cols = 15 
total_rows = 6 

check_map_button_width = 36
check_map_button_height = 26

wafer_config_file = r"C:\Users\zkob\Desktop\NEW PROJECT (GUI)\book1.csv"     # configuration file for indication map orientation

####### check_pass_fail global variables

number_of_units = 90  # enter total number of units per tray 

####### create shape 

start_x_point = 45
start_y_point = 15

next_x_point = 765


# making the "\"  

slanted_x_point = 18   
slanted_y_point = 232
slanted_y_point_2 = 68

final_y_point = 310


line_width = 3  # line width of tray shape


####################### display pass fail indicator function global variables

####### shape of tray placement 

x_vals = 0   # where to place the shape of the tray
y_vals = 30  # height placement of the shape of the tray 


tray_sort_label_x_position = 300

tray_sort_label_y_position = 50

#### change placement of the pass fail map ( x & y position)

pass_fail_map_x_position = 50

y_indicator_frame_position = 80     # placing the    (was 115)


####### screenshot_printout
 
crop_points_for_print = {"left":0, "top": 30, "right": 1590 , "bottom":808}        

given_save_location = r'\\10.132.110.1\Products\SES_PRODUCT\06_MARUMO\6.0 Process data\8.0 Tray 2DBC\Setup Folder'   # screenshot pictures of the GUI

mphase_share_file_path = r'\\sgamkpdisscl01.heptagon.local\Testing\Data\_Marumo\mPhase01'

_2dbc_share_file_path = r'\\10.132.110.1\Products\SES_PRODUCT\06_MARUMO\6.0 Process data\8.0 Tray 2DBC'

current_dir = r'C:\Users\zkob\Desktop\NEW PROJECT (GUI)'    # ensure to change





###########

needed_parameters = ['SerialNumber','Bin']   

root = Tk()  


def gui():    
    
    os.chdir(current_dir)   # change directory
    
    root.state('zoomed')          # opens window in full screen
    
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    
    
    initial_frame = Frame(root, width = screen_width ,height = screen_height)
    
    start_scan_btn_Font = font.Font(family='Helvetica',size=30)
    
    # Create a initial_frame
    start_scan_btn = Button(initial_frame, text = 'Start Scanning' ,bd = '2',command = lambda: create_main_frame(root,screen_width,screen_height,initial_frame), bg='#0052cc' , fg = '#ffffff',font = start_scan_btn_Font)
    
    # Set the position of button on the top of window.  
    start_scan_btn.pack(side=TOP, expand=YES)   
    
    initial_frame.pack(fill=BOTH, expand=YES)            # use .pack to allow the button to appear because we initialized the button earlier but have not told the machine where to place it in the GUI

    root.mainloop()


# In[5]:


############################ Frame after pressing the start scan button 

def create_main_frame(window_name,given_screen_width,given_screen_height,given_intial_frame):      
    
    given_intial_frame.destroy()                                                                         # destroy the initial frame that has the start scan button 
     
    main_frame = Frame(window_name, width = given_screen_width ,height = given_screen_height)            # Create a new frame from the root window and specify the width and height of the screen   
    
    
    ########################### Placing of widgets on the main frame
    
    browse_Lot_Id(main_frame)                                                          # call the browse csv function and pass in the main frame which will be where we place the widgets 
    

    
    ############### stop scanning button 
    
    
    main_frame.pack(fill=BOTH, expand=YES)


# In[6]:


###################### Browse for csv widgets placement

def browse_Lot_Id(main_frame):
    
    labelframe_fonts = font.Font(family='Helvetica',size=13)         # font for the labelframe text 
    
    
    ########################## make the Lot ID labelframe
    
    lot_id_labelframe = LabelFrame(main_frame, text="Enter Lot ID " , height = Browse_csv_cognex_labelframe_height , width = all_width, font = labelframe_fonts)
    lot_id_labelframe.place(x = x_labelframe_start_pos , y= y_labelframe_height_pos)
    
    
    
    ######################### Place the csv label , entrybox and button inside the entrybox
    
    lot_number = Label(lot_id_labelframe,text = "Enter A Lot ID Number :")
    lot_number.place(x = x_label_pos, y= y_pos)   

    
    lot_num = StringVar()
    
    Lot_id_entry = Entry(lot_id_labelframe,textvariable = lot_num, width = entry_file_width ,bd = '3' , name = "csv_file_path" )
    Lot_id_entry.place(x = x_entry_pos, y= y_pos)   
    
    
    button_get_lot_num = Button(lot_id_labelframe,text = "Search Files With Lot ID ",command = lambda: get_info_and_run_wafer_map(main_frame,Lot_id_entry) , width = browse_csv_button_width ,bg = 'Orange')
    button_get_lot_num.place(x = x_browse_file_button_pos, y= y_pos_browse_csv)   


# In[7]:


################## button_get_lot_num function

def get_info_and_run_wafer_map(main_frame,lot_number_entry):   # gets the info for the wafer maps and displays the wafer maps
    
    
    
    ################################# if the user has not selected a csv file ,then warn the user , if not then allow them to choose a text file 
    
    if len(lot_number_entry.get()) == 0 :                                          
        
        messagebox.showwarning("showwarning", "Warning ! Please Enter a valid Lot ID Number !" )

    else:
        
#         browse_text_widget(main_frame,filepath)
#         global needed_df
        
        ### call the getdf function to get the lot id (1)  
        
        Lot_id_number = lot_number_entry.get()
        
        Lot_id_number = Lot_id_number.upper().strip()
        
        needed_df = getdf(Lot_id_number)
        
#         print('that is the ' , Lot_id_number)
        
        ### scan through a directory (for now set as cwd) for files that start with lot id  and return a list (2)
        
        cognex_files_per_lot_id = find_cognex_files(Lot_id_number,_2dbc_share_file_path)
        
        ### place into the cleaning cognex files (3)  [ call this for however many files are found ]
        
        serial_number_list,correct_2dbc_position = cleaning_cognex_file(cognex_files_per_lot_id)     # generate the serial number list 
        
#         correct_tray_id_list.sort()
        
#         print('correct_tray_id_list',correct_tray_id_list)
        
        check_map_labelframe = product_check_map(main_frame)
        
        check_map_canvas = Canvas(check_map_labelframe, width=check_map_canvas_width, height=check_map_canvas_height)
        check_map_canvas.place(x=0,y=0)      # canvas where we have all the shapes
        
        
        display_pass_fail_indicator(needed_df,serial_number_list,correct_2dbc_position,check_map_canvas,check_map_labelframe,main_frame,Lot_id_number)
        


# In[8]:


############ reads in the chosen file from the browse csv function 

def getdf(chosen_lot_id):
    
#     global combined_df
    
    os.chdir(mphase_share_file_path)
    
    print('chosen_lot_id',chosen_lot_id)

#     chosen_lot_id.split('-',1)[1]

    correct_chosen_lot_id = chosen_lot_id.split('-',1)[0]

    csv_files = [x for x in os.listdir() if correct_chosen_lot_id in x]   # gets the csv file according to the Lot Id in the shared folder
    print('that is the csv file list ',csv_files)
    
    ####### check if after splitting , the string is still the same as the chosen lot number 
    
    if correct_chosen_lot_id == chosen_lot_id:   # if user has not entered the tray number 
        
        print('same correct_chosen_lot_id',correct_chosen_lot_id)
        print('same chosen_lot_id',chosen_lot_id)
        
        messagebox.showwarning("showwarning", "Warning ! Incomplete ID & Lot number ! Please Enter a Tray number ")
        
    else:
        
        print('correct_chosen_lot_id',correct_chosen_lot_id)
        print('chosen_lot_id',chosen_lot_id)

        if len(csv_files) < 1:

            messagebox.showwarning("showwarning", "Warning ! No Mphase Files Found in directory for this Lot ID Number ! Check if there is a file with Lot ID number in mphase shared folder" )

            os.chdir(current_dir)

        ############### when there are duplicate csv files

        elif len(csv_files) > 1:

            csv_files.sort()
            latest_csv_file = csv_files[-1]
            print('this is the csv file list ',csv_files)
            print('this is the latest_csv_file',latest_csv_file)

            df_in = pd.read_csv( latest_csv_file , skiprows=[0,2,3,4] , index_col=False ) 

            combined_df = pd.DataFrame( df_in , columns=needed_parameters ) 

            os.chdir(current_dir)
    #         print(combined_df)
            return combined_df

        #################### when there are no duplicate csv files

        else:              

            df_in = pd.read_csv( csv_files[0] , skiprows=[0,2,3,4] , index_col=False ) 

            combined_df = pd.DataFrame( df_in , columns=needed_parameters ) 

            os.chdir(current_dir)
    #         print(combined_df)
            return combined_df
#     except IndexError:
        
#         os.chdir(current_dir)
        
#         messagebox.showwarning("showwarning", "Warning ! Incomplete ID & Lot number ! Please Enter a Tray number " )


# In[9]:


######## we can replace with ivans function if needed

def find_cognex_files(given_Lot_id_number,given_cognex_file_path):
    
    try:
        cognex_files = [x for x in os.listdir(given_cognex_file_path) if given_Lot_id_number in x]  
        
    except OSError as e:
        
        messagebox.showwarning("showwarning", "Warning ! Check if 2DBC shared folder is connected !" )
        
    
    if len(cognex_files) < 1:
        
        messagebox.showwarning("showwarning", "Warning ! No 2DBC Files Found in directory for this Lot ID Number ! Check if there is a file with Lot ID number in 2DBC shared folder" )
        
    print('cognex_files',cognex_files)
    
#     tray_lists = [x.split(given_Lot_id_number)[1].split("_",1)[0].split("-")[1] for x in cognex_files]
#     list_of_trays = Counter(tray_lists)
#     list_of_duplicate_trays = [key for key in list_of_trays.keys() if list_of_trays[key]>1]
    
    
    
    
    if len(cognex_files) > 1:
        
        cognex_files.sort()
        
        latest_cognex_files = [cognex_files[-1]]
#         latest_cognex_files.sort()
        print('these are the latest_cognex_files', latest_cognex_files)

        return latest_cognex_files
                        
    else:
        return cognex_files


# In[10]:


def cleaning_cognex_file(given_cognex_files):
    
    os.chdir(_2dbc_share_file_path)    
    
    List_of_list_of_2dbc_positions = []
    
    List_of_list_of_serial_numbers = []
    
    List_of_list_of_tray_ids = []
    
#     print(given_cognex_files)
    
    for filenum in range(len(given_cognex_files)):
        
#         print('filenum',filenum)
#         print('given_cognex_files',given_cognex_files)
#         print('given_cognex_files[filenum]',given_cognex_files[filenum])

        file_path = os.path.abspath(given_cognex_files[filenum]) 
#         print('file_path',file_path)

        with open(file_path) as f:
            lines = f.readlines()
            lines = [x.strip() for x in lines]

            checking = '('

            tray_id = 'sLotNo'

            getting_serial_numbers_and_tray_id = [_2dbc_num for _2dbc_num in lines if checking in _2dbc_num or tray_id in _2dbc_num]    # get the tray id and the 2dbc numbers
            
#             print('getting_serial_numbers_and_tray_id',getting_serial_numbers_and_tray_id)
            
#             found_tray_id = getting_serial_numbers_and_tray_id[0].split('=')[1].split('-')[1].split(',',1)[0]         # get the tray id
            
#             print('found_tray_id',found_tray_id)

            getting_serial_numbers_and_tray_id.pop(0)   # remove the tray id from the first element in the list 

            getting_serial_numbers = [_2dbc_num.split(',')[0] for _2dbc_num in getting_serial_numbers_and_tray_id]

            getting_serial_numbers = [_2dbc_num.split(')')[1] for _2dbc_num in getting_serial_numbers]
            
            getting_all_2dbc_position = [_2dbc_pos.split('(',1)[1].split(')')[0] for _2dbc_pos in getting_serial_numbers_and_tray_id ]
            
            getting_correct_2dbc_position = [position for serial_num , position in zip(getting_serial_numbers,getting_all_2dbc_position) if serial_num != 'NO READ']
#             print('this is the getting_correct_2dbc_position', getting_correct_2dbc_position)
            
            getting_correct_2dbc_position_and_serial_number = [(int(position),serial_num) for serial_num , position in zip(getting_serial_numbers,getting_all_2dbc_position) if serial_num != 'NO READ']
#             print('this is the getting_correct_2dbc_position_and_serial_number', getting_correct_2dbc_position_and_serial_number)
            
#             List_of_list_of_serial_numbers.append(getting_correct_2dbc_position_and_serial_number)

#             List_of_list_of_tray_ids.append(found_tray_id)
            
#             List_of_list_of_2dbc_positions.append(getting_correct_2dbc_position)

    os.chdir(current_dir)

    return(getting_correct_2dbc_position_and_serial_number,getting_correct_2dbc_position)
        
        


# In[11]:


def product_check_map(main_frame):
        
    
    labelframe_fonts = font.Font(family='Helvetica',size=13)
    
    check_map_labelframe = LabelFrame(main_frame, text="Pass / Fail Indicator " , height = check_map_height , width = pass_fail_indicator_frame_width , font = labelframe_fonts ,name = 'check_map_labelframe')
    check_map_labelframe.place(x = x_labelframe_pass_fail_start_pos , y= y_labelframe_height_pos_pass_Fail_labelframe)
    
    return check_map_labelframe
    


# In[12]:


def display_pass_fail_indicator(needed_df,serial_number_list,correct_2dbc_position,check_map_canvas,check_map_labelframe,main_frame,Lot_id_number):
    
    print_out_button_fonts = font.Font(family='Helvetica',size=13) 
    
    
    twodbc_label_fonts = font.Font(family='Helvetica',size=12)
    
     ###### checking the serial number for whether the product is in the csv file , or passes or fails then generate a color list 
        

    bins_list = check_pass_fail(needed_df,serial_number_list,correct_2dbc_position)    # get the color list after checking whether product exists , then passes or fails 


    cognexlabel = Label(check_map_canvas,text = Lot_id_number )
#                 print(correct_tray_id_list[i])
    cognexlabel.place(x = tray_sort_label_x_position, y= tray_sort_label_y_position) 

    create_shape(check_map_canvas,y_vals,x_vals)

    wafer_creation(check_map_canvas,bins_list,y_indicator_frame_position,pass_fail_map_x_position)    # 30 was initially 17 


    legend_for_pass_fail(check_map_labelframe)

    print_out_button = Button(main_frame, text = 'Print screen' ,command = lambda: screenshot_printout(Lot_id_number) , bg= 'red' , fg = '#ffffff',font = print_out_button_fonts)
    print_out_button.place(x = print_out_button_x_position , y  = print_out_button_y_position)  
    


# In[13]:


def check_pass_fail(given_df,given_list_of_serial_numbers,correct_2dbc_position):
    

    
    Bin_number_list = [-1.0 for i in range(1,number_of_units+1)]

    for i in range(len(given_list_of_serial_numbers)):
        
        serial_num = given_list_of_serial_numbers[i][1]
#         print('serial_num',serial_num)
    
        checking = given_df.SerialNumber.isin([serial_num]).any()           # checks if the serialnumber is in the dataframe
    
#         print('serial_num',serial_num)
#         print('checking',checking)

    ########################################### if the serial number is in the dataframe , then check if product is pass or fail , if pass then grey color , if fail then red color

        if checking == True:            
            

            bin_number = given_df['Bin'].where(given_df['SerialNumber'] == serial_num).dropna().values[-1]       # when 
            
#             print('SerialNumber',serial_num)
#             print('bin_number',bin_number,type(bin_number))
#             print('given_list_of_serial_numbers',given_list_of_serial_numbers[i][0]-1)
        
            Bin_number_list[given_list_of_serial_numbers[i][0]-1] = bin_number
            

            
    
    ###### For those with missing position number or 'NO READ'
    
    for i in range(1,number_of_units+1):
        if str(i) not in correct_2dbc_position:
            Bin_number_list[i-1] = 'No Unit'
    
#     print('Bin_number_list::',Bin_number_list)
    
    return(Bin_number_list)


# In[14]:


def create_shape(w,y_vals,x_vals):
    w.create_line(start_x_point + x_vals, start_y_point + y_vals , next_x_point + x_vals, start_y_point + y_vals, width = line_width)       # straight line (x1,y1,x2,y2) above the lot and tray name
    w.create_line(next_x_point + x_vals, start_y_point + y_vals, next_x_point + x_vals, final_y_point + y_vals , width = line_width)        # straight line downwards
    w.create_line(next_x_point + x_vals, final_y_point + y_vals, start_x_point + x_vals, final_y_point + y_vals , width = line_width)       # bottom line
    w.create_line(start_x_point + x_vals, final_y_point + y_vals, slanted_x_point + x_vals, slanted_y_point + y_vals , width = line_width)       # slant 1 
    w.create_line(slanted_x_point + x_vals, slanted_y_point + y_vals, slanted_x_point + x_vals, slanted_y_point_2 + y_vals , width = line_width) # straight line between slant
    w.create_line(slanted_x_point + x_vals, slanted_y_point_2 + y_vals, start_x_point + x_vals, start_y_point + y_vals , width = line_width)     # slant 2 


# In[15]:


def wafer_creation(check_map_labelframe,given_list_of_bin_numbers,given_y_position_frame, given_x_pos):
    
    check_map_frame = Frame(check_map_labelframe,height = check_map_height , width = all_width)
    check_map_frame.place(x = given_x_pos ,y = given_y_position_frame)
    
    pixelVirtual = tk.PhotoImage(width=1, height=1)
    
    ###################### Legend
        
    legend(pixelVirtual,check_map_frame)
    
    val = pd.DataFrame(given_list_of_bin_numbers)

    pos = val.index.tolist()
    pos =[x+1 for x in pos]
    pos = pd.DataFrame(pos)

    final = pd.concat([pos,val], axis = 1)
    final.columns = ['Pos','Val']
    
    
    
    config_df = pd.read_csv(wafer_config_file)
    config_df = config_df.loc[:, ~config_df.columns.str.contains('^Unnamed')]
    
    for i,val in enumerate(final['Pos']):
        
        d = dict(zip(config_df.columns, range(len(config_df.columns))))
        config_df = config_df.rename(columns=d)
        s = config_df.rename(columns=d).stack()
        f = (s == val).idxmax()  # generates the rc number
        
        if final['Val'][i] == 'No Unit':
            
            myFont = font.Font(family='Helvetica', size=20, weight='bold')

            color = None

            label_text = "X"

            fg= 'black'

            product_position = Button(check_map_frame,width = check_map_button_width , height = check_map_button_height , bg = color ,fg = fg , text = label_text , font= myFont, image  = pixelVirtual , compound = 'c')
            product_position.grid(row=f[0]+1, column=f[1]+1)

            
        elif final['Val'][i] == -1.0:
            
            fg= 'black'

            product_position = Button(check_map_frame,width = check_map_button_width , height = check_map_button_height , bg = 'red' ,fg = fg , font='sans 9 bold', image  = pixelVirtual , compound = 'c')
            product_position.grid(row=f[0]+1, column=f[1]+1)
            
        
        else:

            fg = '#006600'      # green color

            text = int(final['Val'][i])

            myFont = font.Font(family='Helvetica', size=17, weight='bold')

            product_position = Button(check_map_frame,width = check_map_button_width , height = check_map_button_height ,fg = fg , text = text , font=myFont, image  = pixelVirtual , compound = 'c')
            product_position.grid(row=f[0]+1, column=f[1]+1)
        


# In[16]:


##### legend in the map 

def legend(pixelVirtual,check_map_frame):
    
    label_text_number_legend_cols = 1
    
    label_text_number_legend_rows = 1
    
    ############ make the length of the legend (col number)
        
    for i in range(1):
        for j in range(total_cols):
            
            myFont = font.Font(family='Helvetica', size=10)
            
            product_position = Button(check_map_frame,width = check_map_button_width , height = check_map_button_height , bg = 'light blue' ,text = label_text_number_legend_cols , font=myFont, image  = pixelVirtual , compound = 'c')
            product_position.grid(row=i, column=j + 1)
            
            label_text_number_legend_cols += 1
            
    
    ########### make the height of the legend (rows number)
    
    for i in range(total_rows+1):
        for j in range(1):
            
            if i == 0:
                
                product_position = Button(check_map_frame,width = check_map_button_width , height = check_map_button_height , bg = 'light blue',image  = pixelVirtual,compound = 'c')
                product_position.grid(row=i, column=0)
                
            else:
                
                myFont = font.Font(family='Helvetica', size=10)
            
                product_position = Button(check_map_frame,width = check_map_button_width , height = check_map_button_height , bg = 'light blue' ,text = label_text_number_legend_rows , font=myFont, image  = pixelVirtual , compound = 'c')
                product_position.grid(row=i, column=0)
                

                label_text_number_legend_rows += 1
    


# In[17]:


########## legend for the buttons


def legend_for_pass_fail(check_map_labelframe):
    
    pixelVirtual = tk.PhotoImage(width=1, height=1)
    
    labelframe_fonts = font.Font(family='Helvetica',size=13)
    
    making_legend_frame = LabelFrame(check_map_labelframe , height = legend_frame_height , width = legend_frame_width, text = "Legend" , font = labelframe_fonts )
    making_legend_frame.place(x = place_legend_frame_x_pos ,y = legend_frame_y_pos)
    
    ################# Bin Numbers 
    
    Legend_Bin_Numbers_labels = Label(making_legend_frame,text = "Bin Numbers :")
    Legend_Bin_Numbers_labels.place(x = x_legend_label_pos_Bin_A , y = y_labels_legend_pos)
    
    fg = '#006600'      # green color
    
    letter_Font = font.Font(family='Helvetica', size=17, weight='bold')
    
    Legend_Bin_Numbers_button_example = Button(making_legend_frame, relief=RIDGE,width = check_map_button_width , height = check_map_button_height ,fg =fg , text = 'n' , image = pixelVirtual,compound = "c",font= letter_Font)
    Legend_Bin_Numbers_button_example.place(x = x_legend_button_pos_Bin_A , y = y_pass_fail_func_widget_pos)
    
    
    ################## serial number not found 
    
    Number_not_found_labels = Label(making_legend_frame,text = "Serial Number Not Found in 2DBC File:")
    Number_not_found_labels.place(x = x_legend_label_pos_Bin_C , y = y_labels_legend_pos)
    
    
    Number_not_found_button_example = Button(making_legend_frame, relief=RIDGE,width = check_map_button_width , height = check_map_button_height , bg = 'red'  , image = pixelVirtual)
    Number_not_found_button_example.place(x = x_legend_button_pos_Bin_C , y = y_pass_fail_func_widget_pos )
    
    
    ################### no unit
    
    myFont = font.Font(family='Helvetica', size=20, weight='bold')
    
    Legend_no_unit_labels = Label(making_legend_frame,text = "No Unit :")
    Legend_no_unit_labels.place(x = x_legend_label_pos_no_unit , y = y_labels_legend_pos)
    
    
    Legend_button_no_unit_example = Button(making_legend_frame, relief=RIDGE,width = check_map_button_width , height = check_map_button_height , bg = None , font= myFont, image = pixelVirtual, text = "X", compound = "c")
    Legend_button_no_unit_example.place(x = x_legend_button_no_unit , y = y_no_unit_widget_pos)
    
    


# In[18]:


def screenshot_printout(Lot_id_number):
    
    global new_img
    
    myScreenshot = screenshotUtil.screenshot()
    
    new_img = myScreenshot.crop(( crop_points_for_print["left"], crop_points_for_print["top"], crop_points_for_print["right"], crop_points_for_print["bottom"] ))
    
    date_string = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    
    if not os.path.exists(given_save_location):    # change directory
        os.mkdir(given_save_location)
    
#     save_img_directory = os.path.join(given_save_location, given_directory)
    
    new_img.save(os.path.abspath(given_save_location)  + "/" + date_string +  "__"+  Lot_id_number + ".png" )
    
    # Constants for GetDeviceCaps
    #
    #
    # HORZRES / VERTRES = printable area
    #
    HORZRES = 8
    VERTRES = 10
    #
    # LOGPIXELS = dots per inch
    #
    LOGPIXELSX = 88
    LOGPIXELSY = 90
    #
    # PHYSICALWIDTH/HEIGHT = total area
    #
    PHYSICALWIDTH = 110
    PHYSICALHEIGHT = 111
    #
    # PHYSICALOFFSETX/Y = left / top margin
    #
    PHYSICALOFFSETX = 112
    PHYSICALOFFSETY = 113

    printer_name = win32print.GetDefaultPrinter ()
    file_name = os.path.abspath(given_save_location)  + "/" + date_string +  "__"+  Lot_id_number + ".png"

    #
    # You can only write a Device-independent bitmap
    #  directly to a Windows device context; therefore
    #  we need (for ease) to use the Python Imaging
    #  Library to manipulate the image.
    #
    # Create a device context from a named printer
    #  and assess the printable size of the paper.
    #
    hDC = win32ui.CreateDC ()
    hDC.CreatePrinterDC (printer_name)
    printable_area = hDC.GetDeviceCaps (HORZRES), hDC.GetDeviceCaps (VERTRES)
    printer_size = hDC.GetDeviceCaps (PHYSICALWIDTH), hDC.GetDeviceCaps (PHYSICALHEIGHT)
    printer_margins = hDC.GetDeviceCaps (PHYSICALOFFSETX), hDC.GetDeviceCaps (PHYSICALOFFSETY)
#     print(printable_area,printer_size,printer_margins)
    #
    # Open the image, rotate it if it's wider than
    #  it is high, and work out how much to multiply
    #  each pixel by to get it as big as possible on
    #  the page without distorting.
    #
    bmp = Image.open (file_name)
    if bmp.size[0] > bmp.size[1]:
        bmp = bmp.rotate(90, expand=True)

    ratios = [1.0 * printable_area[0] / bmp.size[0], 1.0 * printable_area[1] / bmp.size[1]]
    scale = min (ratios)

    #
    # Start the print job, and draw the bitmap to
    #  the printer device at the scaled size.
    #
    hDC.StartDoc (file_name)
    hDC.StartPage ()

    dib = ImageWin.Dib (bmp)
    scaled_width, scaled_height = [int (scale * i) for i in bmp.size]
    x1 = int ((printer_size[0] - scaled_width) / 2)
    y1 = int ((printer_size[1] - scaled_height) / 2)
    x2 = x1 + scaled_width
    y2 = y1 + scaled_height
    dib.draw (hDC.GetHandleOutput (), (x1, y1, x2, y2))

    hDC.EndPage ()
    hDC.EndDoc ()
    hDC.DeleteDC ()

    return new_img


# In[19]:


gui()

