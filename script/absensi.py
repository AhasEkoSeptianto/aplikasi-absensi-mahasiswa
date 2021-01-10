import tkinter as tk
from tkinter import *
from tkinter import ttk ,messagebox
from PIL import ImageTk, Image
import zipfile,pandas as pd,csv,os
from . import PandasFile

class absensi():

	# init__
	def __init__(self):
		# root
		self.root 		= Tk()
		
		try :
			# auto maximaze windows
			self.root.attributes('-zoomed', True)
		except :
			# auto maximisze linux because iam use linux 
			self.root.attributes('-fullscreen', True)

		self.root.title("Absensi Mahasiswa")
		
		# panggil fungsi CSV local
		self.myFilesCsv = myCsv()
		
		# mengatur posisi layout
		self.position()
		
		# ketika user quit akan menjalankan fungsi dan save data
		self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
		
		# update data ketika user berinteraksi
		self.SortBy.bind('<<ComboboxSelected>>',self.UpdateTable)
		self.ChooseSemester.bind('<<ComboboxSelected>>',self.UpdateTable)
		
		# posisi untuk db dan menu
		self.PositionMenuAndData()
		self.root.mainloop() #looping applications

	def position(self):
		# img
		self.img 		= Image.open("../img/logo.png")
		self.img 		= self.img.resize((80,80), Image.ANTIALIAS)
		self.myLogoStikom = ImageTk.PhotoImage(self.img)
		self.Logo 		= Label(self.root, image=self.myLogoStikom)
		self.Logo.image = self.myLogoStikom

		# Label
		self.header		= Label(self.root, text="ABSENSI MAHASISWA\nStikomCki.D", font="Arial 15 bold")
		self.Semester 	= Label(self.root, text="Select Semester : ")
		
		# combobox Semester
		self.posSemester = tk.StringVar()
		self.ChooseSemester = ttk.Combobox(self.root, width= 25, textvariable=self.posSemester)

		# ambil data jika ada Semester
		self.ChooseSemester['values'] = tuple(self.myFilesCsv.OpenFile())
		Label(self.root, text='Info (!)  : Berhasil membuatkan lembar kerja baru').place(x=1160, y=60, anchor=CENTER)
		self.ChooseSemester.current(0)
		
		# pengurutan data
		self.LabelSortBy 	= Label(self.root, text='Urutkan Berdasarkan : ')
		self.SortBy 		= ttk.Combobox(self.root, width= 25)
		self.SortBy['values'] = ('Nim','Nama','Jurusan')
		self.SortBy.current(0)

		# button
		self.tambahSiswa 	= Button(self.root, text='tambah mahasiswa'	, command=self.tambahSiswabtn)
		self.hapusMahasiswa = Button(self.root, text='hapus mahasiswa'	, command=self.hapusMahasiswa)
		self.rekapSiswa 	= Button(self.root, text='rekap mahasiswa'	, command=self.rekupSiswabtn)
		self.ubahPassword 	= Button(self.root, text='ubah password'	, command=self.ubahpasswordbtn)
		self.ubahKehadiranMahasiswa = Button(self.root, text='ubah kehadiran mahasiswa', command=self.ubahKehadiranMahasiswaBtn)
		self.clearFormWindow = Button(self.root, text='clear form'		, command=self.clearAllForm)
		self.exit 			= Button(self.root, text='Exit',padx=30	 	, command=self.on_closing)

		# FormCanvas
		self.CanvasForm 	= Canvas(self.root, height=180, width=500)
		self.CanvasFormRight = Canvas(self.root, height=180, width=600)
		self.CanvasFormLeft = Canvas(self.root, height=250, width=600)
		


	def PositionMenuAndData(self):
		# position
		self.Logo.place(x=100, y=50, anchor=CENTER)
		self.header.place(relx=0.5, y=50, anchor=CENTER)
		self.Semester.place(x=1057, y=20, anchor=CENTER)
		self.ChooseSemester.place(relx=0.9, y=20, anchor=CENTER)

		self.LabelSortBy.place(x=1042, y=60, anchor=CENTER)
		self.SortBy.place(relx=0.9, y=60, anchor=CENTER)
		self.ListDataInsert()

		# scrool bar
		self.Scroll 	= ttk.Scrollbar(self.root, orient="vertical",command=self.listBox.yview)
		self.Scroll.place(x=1350, y=120, height=300)
		self.listBox.configure(yscrollcommand=self.Scroll.set)

		self.listBox.place(y=120)
		self.tambahSiswa.place(y=450, x=80, anchor=CENTER)
		self.hapusMahasiswa.place(y=450, x=240, anchor=CENTER)
		self.rekapSiswa.place(y=450, x=395, anchor=CENTER)
		self.ubahKehadiranMahasiswa.place(y=450, x=585, anchor=CENTER)
		self.ubahPassword.place(y=450, x=765, anchor=CENTER)

		self.clearFormWindow.place(y=450, x=1180, anchor=CENTER)
		self.exit.place(y=450, x=1300, anchor=CENTER)

	# membuat table menu untuk penempatan db
	def ListDataInsert(self,path=None):
		# header table
		self.listHeaderTable = ('No','Nim','Nama','Jurusan','1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16')
		# pembuatan listbox table
		self.listBox = ttk.Treeview(self.root, selectmode="extended", columns=self.listHeaderTable, show='headings',height=14)
		
		# memberikan heading ke table dan pengaturan panjang table perrow nya
		listmid = ["Nim"]
		listspecial = ["Jurusan"]
		listshort = ['No','1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16']

		for i in listmid :
			self.listBox.heading(i)
			self.listBox.column(i,minwidth=130, width=130, stretch=YES)

		for i in listspecial :
			self.listBox.heading(i)
			self.listBox.column(i,minwidth=80,width=80, stretch=YES)
		
		for i in listshort :
			self.listBox.heading(i)
			self.listBox.column(i,minwidth=50,width=50, stretch=YES)

		self.listBox.heading('Nama')
		self.listBox.column('Nama',minwidth=298,width=298, stretch=YES)

		for col in self.listHeaderTable:
			self.listBox.heading(col, text=col)

		# mengupdate table jika data telah terpacu
		self.UpdateTable()


	#  <<<    FORM USER    >>>>>

	def UpdateTable(self,path=None):
			# path = posisi semester saat ini
			path =  str(self.ChooseSemester.get())
			Sort = self.SortBy.get()
			
			# pengurutan data berdasarkan pemilihan pengurutan dari user, contoh => urut berdasarkan nama
			PandasFile.PandasCsv(path).SoringCsv(sortby=Sort)
			# menghapus data sebelumnya dan digantikan dengan yang baru
			for i in self.listBox.get_children():
				self.listBox.delete(i)

			with open(path) as f:
				reader 	= csv.DictReader(f, delimiter=',')
				no 		= 1

				for row in reader:
					Nim 		= row['Nim']
					Nama 		= row['Nama']
					Jurusan 	= row['Jurusan']
					pertemuan1 	= row['1']
					pertemuan2 	= row['2']
					pertemuan3 	= row['3']
					pertemuan4 	= row['4']
					pertemuan5 	= row['5']
					pertemuan6 	= row['6']
					pertemuan7 	= row['7']
					pertemuan8 	= row['8']
					pertemuan9 	= row['9']
					pertemuan10 = row['10']
					pertemuan11 = row['11']
					pertemuan12 = row['12']
					pertemuan13 = row['13']
					pertemuan14 = row['14']
					pertemuan15 = row['15']
					pertemuan16 = row['16']

					self.listBox.insert("" , "end", values=(no,Nim,Nama,Jurusan,pertemuan1,pertemuan2,pertemuan3,pertemuan4,pertemuan5,pertemuan6,pertemuan7,pertemuan8,pertemuan9,pertemuan10,pertemuan11,pertemuan12,pertemuan13,pertemuan14,pertemuan15,pertemuan16))
					
					no += 1


	def tambahSiswabtn(self):
		# membersihkan form sebelumnya jika ada
		self.clearAllForm()

		# function untuk generated table mahasiswa jika ada data di semester 1 dan user ingin menggenerated data ke 
		# semester 6 maka tidak perlu input ulang
		def Generated():
			Semester 		= self.ChooseSemester.get()
			semuaNama 		= []
			semuaNim 		= []
			semuaJurusan	= []

			path = str('Semester 1 Ganjil.csv')
			with open(path) as file:
				reader = csv.DictReader(file, delimiter=',')

				for row in reader:
					nama 	= row['Nama']
					semuaNama.append(nama)
					nim 	= row['Nim']
					semuaNim.append(nim)
					jurusan = row['Jurusan']
					semuaJurusan.append(jurusan)

			# statement eksekusi generated
			if path == self.ChooseSemester.get():
				pass # jika posisi semester sama maka fungsi ini hanya bypass saja 
			else:
				for i in range(len(semuaNama)):
					PandasFile.PandasCsv(str(self.posSemester.get())).saveCsv(semuaNama[i],semuaNim[i],semuaJurusan[i])
			
			# selalu update jika ada perubahan data
			self.UpdateTable()
		

		# function tambah siswa button
		def tambahsiswa():
			nama 	= self.InputNama.get()
			jurusan = self.InputJurusan.get()
			
			# check jika nim integer
			try :
				nim 	= int(self.InputNim.get())
				
				if nama != "" and nim != "" and jurusan != "":
					# save form
					PandasFile.PandasCsv(str(self.posSemester.get())).saveCsv(nama,nim,jurusan)			
					# increnemet jumlah tambah mahasiswa
					self.JumlahLabelTambah += 1
					# label pesan
					pesan = 'Info (!)  : ' + str(self.JumlahLabelTambah) + ' mahasiswa telah ditambahkan'
					self.pesanTambahSiswa.set(pesan)
			
			except :
				nim = self.InputNim.get()
				
				if nim != "" :
					self.pesanTambahSiswa.set('Info (!)  :  Nim harus berupa angka')

			self.UpdateTable()	

		# Canvas Left
		# canvas form adding siswa
		self.CanvasForm 	= Canvas(self.root, height=280, width=800)
		self.CanvasForm.place(y=600, x=270, anchor=CENTER)

		self.LabelNama 		= Label(self.CanvasForm, text='Nama Mahasiswa  ')
		self.LabelNama.grid(row=0,column = 0, sticky='W', pady=10, padx=40)

		self.InputNama 		= Entry(self.CanvasForm,)
		self.InputNama.grid(row=0,column = 1, sticky='W', pady=10, padx=40)
		
		self.LabelNim 		= Label(self.CanvasForm, text= 'Nim Mahasiswa    ')
		self.LabelNim.grid(row=1,column = 0, sticky='W', pady=10, padx=40)
		
		self.InputNim 		= Entry(self.CanvasForm,)
		self.InputNim.grid(row=1,column = 1, sticky='W', pady=10, padx=40)
		
		self.LabelJurusan 	= Label(self.CanvasForm, text= 'Jurusan Mahasiswa    ')
		self.LabelJurusan.grid(row=2,column = 0, sticky='W', pady=10, padx=40)
		
		self.InputJurusan 	= Entry(self.CanvasForm,)
		self.InputJurusan.grid(row=2,column = 1, sticky='W', pady=10, padx=40)

		# Canvas Right
		self.CanvasFormRight = Canvas(self.root, height=180, width=500)
		self.CanvasFormRight.place(y=580, x=1000, anchor=CENTER)

		self.LabelInfoGeneratedMahasiswa = Label(self.CanvasFormRight, text='Info (!)  : Generated hanya berlaku jika\n lembar kerja "Semester 1 Ganjil.csv" telah tersedia')
		self.LabelInfoGeneratedMahasiswa.place(rely=0.3, x=235, anchor=CENTER)

		self.LabelGenerated = Button(self.CanvasFormRight, text='Generated Form Files', command=Generated)
		self.LabelGenerated.place(rely=0.6, x=235, anchor=CENTER)


		# label info 
		self.pesanTambahSiswa 		= tk.StringVar()
		self.labelInfoTambahSiswa 	= Label(self.CanvasForm, text="", textvariable=self.pesanTambahSiswa).place(y=155 , x=300, anchor=CENTER)
		self.JumlahLabelTambah 		= 0

		# save button
		self.SubmitTambahMahasiswa 	= Button(self.CanvasForm, text='Submit', command= tambahsiswa , padx=10,pady=5 )
		self.SubmitTambahMahasiswa.grid(row=3,column = 0, sticky='W', pady=10,padx=40)



	def hapusMahasiswa(self):
		self.clearAllForm()

		# function untuk mengambil semua nama mahasiswa pada table yang ada
		def semuaNama():
			semuanama = []
			path = str(self.posSemester.get())
			with open(path) as f:
				reader = csv.DictReader(f, delimiter=',')
				# z untuk memberuikan no pada no table
				z = 1
				for row in reader:
					no = z
					nama = row['Nama']
					semuanama.append(nama)
					z += 1

			return semuanama

		# function untuk eksekusi hapus mahasiswa
		def hapusSiswa():
			nama = self.hapusMahasiswaEntry.get()
			MsgBox = tk.messagebox.askquestion ('hapus mahasiswa','hapus {} ?'.format(nama),icon = 'warning')
			
			if MsgBox == 'yes':
				PandasFile.PandasCsv(str(self.posSemester.get())).deleteSiswa(nama)
				semuanama = semuaNama()
				self.hapusMahasiswaEntry['value'] = tuple(semuanama)
				pesan = 'Info (!)  : Mahasiswa ' + str(nama) + ' berhasil dihapus'
				Label(self.CanvasFormLeft, text=pesan).place(y=170,x=300, anchor=CENTER)

			else:
				tk.messagebox.showinfo('info','mahasiswa tidak dihapus')
				
			self.UpdateTable()
		
		# mengambil semua nama yang tersedia pada funtion diatas
		semuanama 	= semuaNama()

		# form kiri
		self.CanvasFormLeft 	= Canvas(self.root, height=250, width=600)
		self.CanvasFormLeft.place(y=600, x=290, anchor=CENTER)

		self.LabelHapusSiswa 	= Label(self.CanvasFormLeft, text='Pilih Mahasiswa : ')
		self.LabelHapusSiswa.place(y=90, x=100, anchor=CENTER)
		
		self.hapusMahasiswaValue = tk.StringVar()
		self.hapusMahasiswaEntry = ttk.Combobox(self.CanvasFormLeft, width= 30, textvariable=self.hapusMahasiswaValue)
		self.hapusMahasiswaEntry['value'] = tuple(semuanama)
		self.hapusMahasiswaEntry.place(y=90, x=350, anchor=CENTER)

		self.hapusMahasiswaBtn 	= Button(self.CanvasFormLeft, text='Submit' , command=hapusSiswa)
		self.hapusMahasiswaBtn.place(y=90, x=550, anchor=CENTER)




	def rekupSiswabtn(self):
		self.clearAllForm()

		def RekapMahasiswa(semuanama):
			pertemuan = self.ChooseSesi.get()
			pertemuan = pertemuan[::1]
			
			# jika pertemuan lebih dari pertemuan 9 dan ambil nilai pertemuan
			if (len(pertemuan) > 11) :
				pertemuan = pertemuan[-2:]
			else :
				pertemuan = pertemuan[-1:]
			
			# ambil form rekap
			kehadiran = []
			for i in range(len(semuanama)):
				kehadiran.append(self.entryKehadiranRekap[i].get())
			# save file
			PandasFile.PandasCsv(str(self.posSemester.get())).RekapMahasiswa(pertemuan,kehadiran,mahasiswa=semuanama)

			# tambah info ketika berhasil rekap
			pesan = 'Info (!)  : Berhasil rekup mahasiswa pada sesi ke ' + str(pertemuan)
			self.labelInfoRekapMahasiswa = Label(self.CanvasFormRight, text=pesan).place(rely=0.7, x =60)
			self.UpdateTable()


		# right form
		self.CanvasFormLeft 	= Canvas(self.root, height=250, width=600)
		self.CanvasFormLeft.place(y=600, x=290, anchor=CENTER)
		

		semuanama 	= []
		path 		= str(self.posSemester.get())
		
		with open(path) as file:
			reader 	= csv.DictReader(file, delimiter=',')
			z 		= 1
			
			for row in reader:
				no 		= z
				nama 	= row['Nama']
				semuanama.append(nama)

				z += 1		

		self.container 	= tk.Frame(self.CanvasFormLeft)
		self.container.pack()

		# tambahkan canvas pada frame
		self.canvasrekap = tk.Canvas(self.container , height=180, width=500)
		self.canvasrekap.pack(side="left", fill="both", expand=True)

		# Link untuk memberikan scroolbar effect pada list canvas daftar mahasiswa 
		self.scrollbar 	= tk.Scrollbar(self.container, orient="vertical",command=self.canvasrekap.yview)
		self.scrollable_frame = ttk.Frame(self.canvasrekap)
		self.scrollbar.pack(side=RIGHT,fill=Y, expand=1)

		self.scrollable_frame.bind(
		    "<Configure>",
		    lambda e: self.canvasrekap.configure(
		        scrollregion=self.canvasrekap.bbox("all")
		    )
		)

		# persediaan untuk looping dan menginisiasi agar lebih efektif
		self.labelNamaRekap 		= []
		self.entryKehadiranRekap 	= []
		
		for i in range(len(semuanama)):
			self.labelNamaRekap.append(Label(self.scrollable_frame, text=semuanama[i]))
			self.labelNamaRekap[i].grid(column=0, row=i, sticky='W', pady=5, padx=30 )

			self.entryKehadiranRekap.append(ttk.Combobox(self.scrollable_frame, width=27,))
			self.entryKehadiranRekap[i]['values'] = ('hadir','alpha','sakit',' - ')
			self.entryKehadiranRekap[i].current(0)
			self.entryKehadiranRekap[i].grid(column = 2, row = i, pady=5 , padx=30)

		self.canvasrekap.create_window((0, 0), window=self.scrollable_frame, anchor='nw')
		self.canvasrekap.configure(yscrollcommand=self.scrollbar.set)

		# Camvas Right
		self.CanvasFormRight = Canvas(self.root, height=180, width=500)
		self.CanvasFormRight.place(y=580, x=1000, anchor=CENTER)
		
		self.LabelSesi 		= Label(self.CanvasFormRight, text='Pilih Sesi ke : ')
		self.LabelSesi.place(rely=0.5, x=100, anchor=CENTER)
		
		self.rekapSesi 		= tk.StringVar()
		self.ChooseSesi 	= ttk.Combobox(self.CanvasFormRight, width= 17, textvariable=self.rekapSesi)
		self.ChooseSesi['values'] = ('pertemuan 1','pertemuan 2','pertemuan 3','pertemuan 4','pertemuan 5','pertemuan 6','pertemuan 7','pertemuan 8','pertemuan 9','pertemuan 10','pertemuan 11','pertemuan 12','pertemuan 13','pertemuan 14','pertemuan 15','pertemuan 16',)
		self.ChooseSesi.current(0)
		self.ChooseSesi.place(rely=0.5,x=235,anchor=CENTER)
		
		self.buttonrekap 	= Button(self.CanvasFormRight, text='Submit Form', command=lambda:RekapMahasiswa(semuanama))
		self.buttonrekap.place(rely=0.5, x =405, anchor=CENTER)


	def ubahKehadiranMahasiswaBtn(self):
		self.clearAllForm()

		# function untuk button
		def UbahKehadiranMahasiswa():
			namaMahasiswa 	= self.NamaMahasiswaUbahKehadiranEntry.get()
			pertemuan 		= self.SesiKehadiranMahasiswa_UbahEntry.get()
			
			if (len(pertemuan) > 11) :
				pertemuan = pertemuan[-2:]
			
			else :
				pertemuan = pertemuan[-1:]

			kehadiran = self.SesiKehadiranMahasiswa_UbahKehadiran.get()

			# goto class
			PandasFile.PandasCsv(str(self.posSemester.get())).UbahKehadiranMahasiswa(namaMahasiswa,pertemuan,kehadiran)

			# Info text
			pesan = 'Info (!)  : Sesi ke ' + pertemuan + ' ' + namaMahasiswa + ' berhasil dirubah' 
			self.labelInfoRekapMahasiswa = Label(self.CanvasFormLeft, text=pesan).place(y=180, x=370, anchor=CENTER)
			
			self.UpdateTable()

		semuanama 	= []
		path 		= str(self.posSemester.get())
		
		with open(path) as file:
			reader = csv.DictReader(file, delimiter=',')
			
			for row in reader:
				nama = row['Nama']
				semuanama.append(nama)
			
		# left form
		self.CanvasFormLeft = Canvas(self.root, height=250, width=600)
		self.CanvasFormLeft.place(y=600, x=290, anchor=CENTER)

		self.LabelUbahKehadiranMahasiswaNama 	= Label(self.CanvasFormLeft, text='Pilih Mahasiswa ')
		self.LabelUbahKehadiranMahasiswaNama.place(y=60, x=100, anchor=CENTER)
		
		self.NamaMahasiswaUbahKehadiran 		= tk.StringVar()
		self.NamaMahasiswaUbahKehadiranEntry 	= ttk.Combobox(self.CanvasFormLeft, width=30, textvariable=self.NamaMahasiswaUbahKehadiran)
		self.NamaMahasiswaUbahKehadiranEntry['value'] = tuple(semuanama)
		self.NamaMahasiswaUbahKehadiranEntry.place(y=60, x=350, anchor=CENTER)


		self.LabelUbahKehadiranMahasiswaEntry 	= Label(self.CanvasFormLeft, text='Pilih Sesi ')
		self.LabelUbahKehadiranMahasiswaEntry.place(y=100, x=100, anchor=CENTER)
		
		self.SesiKehadiranMahasiswa_Ubah 		= tk.StringVar()
		self.SesiKehadiranMahasiswa_UbahEntry 	= ttk.Combobox(self.CanvasFormLeft, width=30, textvariable=self.SesiKehadiranMahasiswa_Ubah)
		self.SesiKehadiranMahasiswa_UbahEntry['value'] = ('pertemuan 1','pertemuan 2','pertemuan 3','pertemuan 4','pertemuan 5','pertemuan 6','pertemuan 7','pertemuan 8','pertemuan 9','pertemuan 10','pertemuan 11','pertemuan 12','pertemuan 13','pertemuan 14','pertemuan 15','pertemuan 16',)
		self.SesiKehadiranMahasiswa_UbahEntry.current(0)
		self.SesiKehadiranMahasiswa_UbahEntry.place(y=100, x=350, anchor=CENTER)


		self.LabelUbahKehadiranMahasiswaKehadiran = Label(self.CanvasFormLeft, text='Kehadiran ')
		self.LabelUbahKehadiranMahasiswaKehadiran.place(y=140,x=100, anchor=CENTER)
		
		self.SesiKehadiranMahasiswa_Kehadiran 	= tk.StringVar()
		self.SesiKehadiranMahasiswa_UbahKehadiran = ttk.Combobox(self.CanvasFormLeft, width=30, textvariable=self.SesiKehadiranMahasiswa_Kehadiran)
		self.SesiKehadiranMahasiswa_UbahKehadiran['value'] = ('hadir','alpha','sakit','-')
		self.SesiKehadiranMahasiswa_UbahKehadiran.current(0)
		self.SesiKehadiranMahasiswa_UbahKehadiran.place(y=140, x=350, anchor=CENTER)

		self.ubahKehadiranSubmit 	= Button(self.CanvasFormLeft, text='Submit', command=UbahKehadiranMahasiswa)
		self.ubahKehadiranSubmit.place(y=180, x=100, anchor=CENTER)




	def ubahpasswordbtn(self):
		# clear form if exist
		self.clearAllForm()

		def UbahPassword():
			username = self.InputUsername.get()
			password = self.InputPassword.get()
			
			MsgBox = tk.messagebox.askquestion ('Reset Password','Apa anda yakin ingin mengubah username & password',icon = 'warning')
			
			if MsgBox == 'yes':
				con = sqlite3.connect('log.db')
				con.execute("UPDATE LOGIN SET username = ? WHERE username = ? ",(str(username),str(self.username)))
				con.commit()
				super(absensi, self).__init__()
				con.execute("UPDATE LOGIN SET password = ? WHERE password = ? ",(str(password),str(self.password)))
				con.commit()
				con.close()
				tk.messagebox.showinfo('info','akun berhasil diubah')
				
				self.username = username
				self.password = password

			else:
				tk.messagebox.showinfo('info','akun tidak dirubah')
		

		# Canvas Left
		# canvas form adding siswa
		self.CanvasForm 	= Canvas(self.root, height=180, width=500)
		self.CanvasForm.place(y=600, x=270, anchor=CENTER)

		self.LabelUsername 	= Label(self.CanvasForm, text='Username')
		self.LabelUsername.place(y=20, x=100, anchor=CENTER)
		self.InputUsername 	= Entry(self.CanvasForm, width="30")
		self.InputUsername.place(y=20, x=330, anchor=CENTER)
		self.Labelpassword 	= Label(self.CanvasForm, text= 'Password')
		self.Labelpassword.place(y=55, x=100, anchor=CENTER)
		self.InputPassword 	= Entry(self.CanvasForm, width="30",show="*")
		self.InputPassword.place(y=55, x=330, anchor=CENTER)
		
		self.SubmitUbahPassword = Button(self.CanvasForm, text='Submit', command=UbahPassword)
		self.SubmitUbahPassword.place(y=100 , x=80, anchor=CENTER)

	# clear all form 
	def clearAllForm(self):
		self.CanvasForm.place_forget()
		self.CanvasFormLeft.place_forget()
		self.CanvasFormRight.place_forget()	
		self.JumlahLabelTambah = 0

	# if wont to closing
	def on_closing(self):
		if messagebox.askokcancel("Quit", "Do you want to quit?"):
			self.myFilesCsv.Quit()
			self.root.destroy()
	    	



class myCsv():

	def __init__(self):

		try:
			os.chdir('file/')
			# try to open files if exist
			with zipfile.ZipFile('DataFile.zip','r') as zf:
				path = zf.extractall()
			
			# # get all filename in ZipFile and extract it
			with zipfile.ZipFile('DataFile.zip') as zf:
				path = zf.extractall()
				self.allfiles = zf.namelist()
		except:
			self.CreateFiles()

	def OpenFile(self):
		return self.allfiles


	def CreateFiles(self):
		mydataCsv = {
				'Nama':[],
				'Nim':[],
				'Jurusan':[],
				'1':[],'2':[],'3':[],'4':[],'5':[],'6':[],'7':[],'8':[],'9':[],'10':[],'11':[],'12':[],'13':[],'14':[],'15':[],'16':[],
				}
		
		data = pd.DataFrame(mydataCsv)
		
		data.to_csv('Semester 1 Ganjil.csv', index=False)
		data.to_csv('Semester 1 Genap.csv', index=False)
		data.to_csv('Semester 2 Ganjil.csv', index=False)
		data.to_csv('Semester 2 Genap.csv', index=False)
		data.to_csv('Semester 3 Ganjil.csv', index=False)
		data.to_csv('Semester 3 Genap.csv', index=False)
		data.to_csv('Semester 4 Ganjil.csv', index=False)
		data.to_csv('Semester 4 Genap.csv', index=False)
		
		self.allfiles = ['Semester 1 Ganjil.csv','Semester 1 Genap.csv','Semester 2 Ganjil.csv','Semester 2 Genap.csv', 'Semester 3 Ganjil.csv','Semester 3 Genap.csv','Semester 4 Ganjil.csv','Semester 4 Genap.csv']
		
		messagebox.showinfo("Info","Lembar kerja tidak ditemukan\nberhasil membuat lembar kerja baru")
		return list(self.allfiles)


	def Quit(self):

		#create zip
		zf = zipfile.ZipFile("DataFile.zip",'w')
		# add to file
		for i in self.allfiles :
			path = str(i)
			# write file
			zf.write(path)

			# del files
			os.remove(path)

		zf.close()

		
