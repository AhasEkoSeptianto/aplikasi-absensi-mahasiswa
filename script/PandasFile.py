import pandas as pd 


class PandasCsv():
	
	def __init__(self,path):
		self.path = path
		self.mydatafile = pd.read_csv(path)

	def saveCsv(self,nama,nim,jurusan):		
		tambahsiswa = {
			'Nim':nim,
			'Nama':nama,
			'Jurusan':jurusan,
			'1':'-',
			'2':'-',
			'3':'-',
			'4':'-',
			'5':'-',
			'6':'-',
			'7':'-',
			'8':'-',
			'9':'-',
			'10':'-',
			'11':'-',
			'12':'-',
			'13':'-',
			'14':'-',
			'15':'-',
			'16':'-'
		}

		self.mydatafile = self.mydatafile.append(tambahsiswa,ignore_index=True)
		self.mydatafile.to_csv(str(self.path),index=False)

	def deleteSiswa(self,nama):
		data = pd.read_csv(str(self.path))
		siswa = data[((data.Nama == nama))].index
		data.drop(siswa,inplace=True)
		data.to_csv(str(self.path),index=False)

	def RekapMahasiswa(self,pertemuan,kehadiran,mahasiswa):
		data = pd.read_csv(str(self.path))
		for i in range(len(mahasiswa)):
			data.loc[i, str(pertemuan)] = kehadiran[i]
		data.to_csv(str(self.path), index=False)
	
	def UbahKehadiranMahasiswa(self,nama,pertemuan,kehadiran):
		data = pd.read_csv(str(self.path))
		data.loc[(data['Nama'] == nama),str(pertemuan)]= kehadiran
		data.to_csv(str(self.path), index=False)

	def SoringCsv(self,sortby):
		data = pd.read_csv(str(self.path))
		data = data.sort_values(sortby)
		data.to_csv(str(self.path), index=False)