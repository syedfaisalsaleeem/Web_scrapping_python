import tkinter as tk
import xlwt 
from xlwt import Workbook 
import requests
 
def call_back():
    first_session = requests.Session()
    l4=tk.Label(root,text='Please wait',background='bisque',font=("Helvetica", 16)).place(x=30,y=300)
    global entry1
    global entry2
    global entry3
    entry1=e1.get()
    entry2=e2.get()
    entry3=e3.get()
    directory=e4.get()
    entry1=int(entry1)
    entry2=int(entry2)
    print(entry1,entry2,entry3)
    print(type(entry1))

    excel_list=[]#this is the list in which i stored all the values to be sent to excel file

    for line in range(entry1,entry2+1):
                
                r = first_session.get('https://nfs.punjab.gov.pk/?page='+str(line))  
                page=r.text
                
                list1=[]# 1stpage list1
                list2=[]# 1st page list2
                for x in range (len(page)):
                #print(page[x])
                         if page[x:x+4]=='<tr>': # finding and tr in page1
                         #print('<tr> found')
                                list1.append(x)

                for y in range (len(page)):
    #print(page[x])
                        if page[y:y+5]=='</tr>': # finding tr in page1 this makes rows and of data to be enterd 
        #print('<tr> found')
                                list2.append(y)
#print(list1)
#print("xxxxxx")
#print(list2)
                data_list=[] #sending the rows of data into data_list
    
                for i in range (len(list1)):
        
                         r=page[list1[i]:list2[i]]
                         data_list.append(r)
        #print(r)
        
    
                for new in range(len(data_list)): # now iteratively separting rows of excel usage and converting into string with index1
                        data_1=str(data_list[new])
                        list_for_data_1=[]#again creating alist for creating columns in excel
                        list_for_data_1_2=[]
                        for x in range (len(data_1)):
    #print(page[x])
                                 if data_1[x:x+3]=='<td':
        #print('<tr> found')
                                      list_for_data_1.append(x+1)

                        for y in range (len(data_1)):
    #print(page[x])
                                 if data_1[y:y+4]=='/td>':
        #print('<tr> found')
                                      list_for_data_1_2.append(y)
            
            
    
#print(list_for_data_1)
#print("xxxxxx")
#print(list_for_data_1_2)

                        entry_list=[]#sending columns to list for excel usuage # this have all columns of row by row
                        for i in range (len(list_for_data_1)):
        
                                  r=data_1[list_for_data_1[i]:list_for_data_1_2[i]] #columns storing then sending to entry list
                                  entry_list.append(r)
        #print(r)
#print(entry_list[0])
                        for t in range (len(entry_list)):# entry list is not pure with columns puring it for excel list 
                                  m=entry_list[t].find('>')
                                  n=entry_list[t].find('<')
                                  l=entry_list[t][m+1:n] # pehley wali index sai > and < kai darmian ka chie
                                  excel_list.append(l)
                #time.sleep(0.5)
                        
    #print(excel_list)

  
    #Workbook is created 
    wb = Workbook() 

    # add_sheet is used to create sheet. 
    sheet1 = wb.add_sheet('Sheet 1')
    element=0 # work book of excel
    
    sheet1.write(0,0,'Name')
    sheet1.write(0,1,'Father Name')
    sheet1.write(0,2,'CNIC')
    sheet1.write(0,3,'Province')
    sheet1.write(0,4,'District')
    for row in range (1,len(excel_list)):
  
        for col in range(0,6):
            element=element+1
            if element>=len(excel_list):# if element becomes greater or equal to len of excel list it will break loop
            
                break
         
            else:
                sheet1.write(row,col,excel_list[element])

        
    wb.save(directory+':/'+entry3+'.xls') 
    l4=tk.Label(root,text='Now You May Close',background='bisque',font=("Helvetica", 16)).place(x=30,y=340)
    
    
root = tk.Tk()#creating an object named root
root.title("Faisal Application")#name of application
#def retrieve_input():
#    inputValue=textBox.get("1.0","end-1c")
#    print(inputValue)

#textBox=Text(root, height=1, width=10,wrap='none')
#textBox.config(wrap='none')


#buttonCommit=Button(root, height=1, width=10, text="Commit", command=lambda: retrieve_input())
frame1 = tk.Frame(root, width=500, height=500,background='bisque') # creating the frame of defined size
root.resizable(0,0) # finish the full screen mode
l1=tk.Label(root,text='enter the starting page number',background='bisque',font=("Helvetica", 16)).place(x=10,y=50)
l2=tk.Label(root,text='enter the ending page number',background='bisque',font=("Helvetica", 16)).place(x=10,y=100)
l5=tk.Label(root,text='enter Directory name except C',background='bisque',font=("Helvetica", 16)).place(x=10,y=160)
l3=tk.Label(root,text='enter the name of excel file',background='bisque',font=("Helvetica", 16)).place(x=10,y=210)

e1 = tk.Entry(root)

e1.place(x=320,y=50,height=30)

#e1.grid(padx=5,pady=10,ipady=3)
e2 = tk.Entry(root)
e2.place(x=320,y=100,height=30)
e4 = tk.Entry(root)
e4.place(x=320,y=150,height=30)
e3 = tk.Entry(root)
e3.place(x=320,y=200,height=30)

#for single line we use entry and for double line we use textbox
#e.place(bordermode=INSIDE, anchor=S,y=60)
#frame2 = tk.Frame(root, width=50, height = 50, background="#b22222")
#textBox.pack()
#buttonCommit.pack()

Btn = tk.Button(root, text="OKAY", width=15, height=2, command=call_back).place(x=200, y=270)

#e.pack()
frame1.pack()# packing the frame
root.mainloop()
