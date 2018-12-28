#Get rid of flickring when maxmize || gone with py installer
#Make buttons nicer || cosmetic
#make variable window sizes and textbox || Minor
#**dict || add
#Written By: Roman Kleyn || V1 Completed Nov 20, 2018
sub_text = '>>>Subject Text Goes Here, Leave as is if Subject not needed to exclude Subject column.'
txt = '''>>>Instructions/Temp Text:\n \
1. Download CSV of Contacts from LinkedIn or other CSV.\n \
2. First Row of CSV is header, name appropreietly & Import CSV.\n \
3. Write or Paste Message, header buttons paste where cursor is.\n \
4. Modify text/move header attributes; Copy, Cut & Paste/PC:ctrl-c, ctrl-x & ctrl-v/MAC:cmd-c, cmd-x & cmd-v.\n \
5. Export Msg to CSV, Saved in same folder as Imported CSV.\n \
	NOTE: When copying messages from the CSV copy content within cell not the cell.
EXAMPLE LinkedIn Text(LinkedIn CSV for it to work): 
Hi {First_Name} {Last_Name},
Hope all you're still {Position} at {Company}, we should catch up {First_Name}!
Give me a call at (123) 456 7891
Best Regards,
PPP\n \
6.DELETE Text when Done Reading and Start Writing!'''

import csv
import time
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import tkinter.scrolledtext as scrolledtext

class CSV_MSG:
	def __init__(self):
		self.button_list = []
		self.sub_list = []
		self.GUI()
		
	def CSVread(self): 
		global attributes, master_list, contacts
		#Check if csv file
		err = 0
		csvname = filename.split('.')
		csvn = len(csvname)-1
		csv_test = csvname[csvn].lower()
		if csv_test != "csv":
			messagebox.showinfo("Import Error!", "Check a .CSV file is being used.\nIf so check filename is properly formatted.\nie. listname.csv")
			err = 1

		if err == 0:
			raw_list, master_list = [], []
			csv_file = open(filename)
			data = csv.reader(csv_file)
			for row in data:
				raw_list.append(row)
			attributes, contacts = raw_list[0], raw_list[1:]
			attributes = [att.replace(" ", "_") for att in attributes]
			
			#Contacts List to Dict 
			for contact_info in contacts:
				n, contact_dic = 0, dict()
				for attribute in attributes:
					contact_dic[attribute] = contact_info[n]
					n += 1
				master_list.append(contact_dic)
		return err
			
	def UploadCSV(self):
		global filename
		filename = filedialog.askopenfilename()
		err = self.CSVread()
		#button gen + windows
		if err == 0: self.button_gen()
		
	def MsgGen(self, msg, sub):
		used_msg_attr = list(set([x.split('}')[0] for x in msg.split('{') if '}' in x]))
		used_sub_attr = list(set([x.split('}')[0] for x in sub.split('{') if '}' in x]))
		#Sub check if needed
		sub_act = True
		if sub == sub_text:
			sub = ''
			sub_act = False
		
		#check if attrubte exist(msg + sub) | Finish Error
		err = 0
		for atr in used_msg_attr:
			if atr not in attributes:
				messagebox.showinfo("Message Box Error!", "Check Message Textbox, variable `{}` does not exist in the .CSV file header".format(atr))
				err = 1
				break

		for atr in used_sub_attr:
			if atr not in attributes:
				messagebox.showinfo("Subject Box Error!", "Check Subject Textbox, variable `{}` does not exist in the .CSV file header".format(atr))
				err = 1
				break
				
		#Check if used variable exist(msg + sub)
		n = 0
		for usr in master_list:
			v_missing = False
			for a in used_msg_attr:
				if err == 0 and usr[a] == '':
					v_missing = True
			for a in used_sub_attr:
				if err == 0 and usr[a] == '':
					v_missing = True
			master_list[n]['Missing_Vals'] = v_missing
			n += 1
					
		n = 0
		msg = "'''" + msg + "'''"
		sub = "'''" + sub + "'''"
		if err == 0:
			for contact in master_list:
				#create variable
				l_eval = ['{0}=contact["{0}"]'.format(x) for x in attributes]
				exec(';'.join(l_eval))
				#format msg
				l_fmt = [x+" = "+x for x in attributes]
				fmat = '.format({})'.format(', '.join(l_fmt))
				txt = eval(msg+fmat)
				subj = eval(sub+fmat)
				#adds msg to Master_List dict
				master_list[n]['Msg'] = txt 
				master_list[n]['Sub'] = subj 
				n += 1
				
			#ML dicts to List
			mll, t_l = [], []
			for u_info in master_list: 
				for i in u_info:
					t_l.append(u_info[i])
				mll.append(t_l)
				t_l = []
				
			#add missing info to \/
			att = attributes+['Missg Used Val', 'Msg', 'Subject']
			if not sub_act: att = attributes+['Missg Used Val', 'Msg']
			rename = filename.split('/')
			fname = 'MassMsgGen_'+rename[-1]
			fname = rename[:-1]+[fname]
			fname = '/'.join(fname)
			mmg_csv = open(fname,'w')
			wrt = csv.writer(mmg_csv)
			wrt.writerow(att)
			wrt.writerows(mll)

	def button_gen(self):

		#Create Button list and forget buttons old buttons
		if len(self.button_list) > 0:
			for i in range(len(self.button_list)):
				n = self.button_list[i]+".grid_forget()"
				exec(n)
			self.button_list = []
		
		#Subject Area
		sframe = Frame(win)
		sframe.grid(row=2, column=0, columnspan=2, pady = 10, sticky = W)
		self.subject_label = Label(sframe, text="Subject Buttons")
		self.subject_label.grid(row = 1, column = 0, sticky = W)
		self.button_list.append('self.subject_label')
		nr = 1
		for a in range(len(attributes)):
			att, n = attributes[a], a+1
			b = 'self.bn'+str(n)+' = Button(sframe,text="'+att+'",command=lambda: e_sub.insert("insert","{'+att+'}"))'
			bp = 'self.bn'+str(n)+'.grid(row = 1, column = '+str(nr)+', sticky = W)'
			nr += 1
			exec(b)
			exec(bp)
			self.button_list.append('self.bn'+str(n))

		#Text Area
		tframe = Frame(win)
		tframe.grid(row=4, column=0, columnspan=2, pady = 10, sticky = W)
		self.text_label = Label(tframe, text="Message Buttons")
		self.text_label.grid(row = 1, column = 0, sticky = W)
		self.button_list.append('self.text_label')
		nr = 1
		for a in range(len(attributes)):
			att, n = attributes[a], a+1
			b = 'self.bt'+str(n)+' = Button(tframe,text="'+att+'",command=lambda: e_text.insert("insert","{'+att+'}"))'
			bp = 'self.bt'+str(n)+'.grid(row = 1, column = '+str(nr)+', sticky = W)'
			nr += 1
			exec(b)
			exec(bp)
			self.button_list.append('self.bt'+str(n))
			
	def get_text(self):
		msg = e_text.get("1.0", "end-1c")
		sub = e_sub.get("1.0", "end-1c")
		self.MsgGen(msg,sub)
		return
		
	def GUI(self):
		global win, e_text, e_sub
		win = Tk()
		win.title("MassMsgGen V1.0")
		win.geometry('700x700')

		#Import Button
		bimport = Button(win, text='Import CSV Contacts', command=self.UploadCSV)
		bimport.grid(row = 0, column=0)

		#Subject Button || future note: e_sub = Entry(win,width=400)
		e_sub = scrolledtext.ScrolledText(win,width=180,height=1)
		e_sub.insert(INSERT, sub_text)
		e_sub.grid(row = 1, column=0)

		#Message Button
		e_text = scrolledtext.ScrolledText(win,width=180,height=14)
		e_text.insert(INSERT, txt)
		e_text.grid(row = 3, column=0)
		
		#Submit Msg Button
		b_Submit = Button(win, text="Export Msg to CSV", command=lambda: self.get_text())
		b_Submit.grid(row = 5, column=0)

		win.mainloop()
CSV_MSG()
