from tkinter import *
from main import check_data

root=Tk()
root.geometry('600x200')
root.title('Certificate Check')

loc_cert=Label(root, text='Enter the location of Certificate folder:').pack()
loc_cert=Entry(root,width=80).pack()
loc_cert_eg=Label(root,text='c:\\User\\Nk\\Desktop').pack()

loc_excel=Label(root, text='Enter the location of Certificate folder:').pack()
loc_excel=Entry(root,width=80).pack()
loc_excel_eg=Label(root,text='c:\\User\\Nk\\Desktop').pack()

check=Button(root, text='Check', command=check_data).pack()

root.mainloop()
