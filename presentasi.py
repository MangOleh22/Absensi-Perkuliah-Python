from tkinter import *
from openpyxl import Workbook
from openpyxl.styles import Font,Alignment,Border,Side
from tkinter import font as tkfont
 
root = Tk()
root.title(" Absensi Perkuliahan ")
root.resizable(width=False,height=False)
workbook = Workbook()
sheet = workbook.active

styling = tkfont.Font(family='Helvetica',weight='bold', size=15)
styling2 = tkfont.Font(family='Helvetica', size=9)

font = Font(bold=True)
border = Border(left=Side(border_style='thin',color='00000000'),
                right=Side(border_style='thin',color='00000000'),
                top=Side(border_style='thin',color='00000000'),
                bottom=Side(border_style='thin',color='00000000'))

alignment = Alignment(horizontal='center', vertical='center')

HEIGHT = 500
WIDTH= 600
canvas= Canvas(root,  height= HEIGHT , width= WIDTH, bg= 'lightblue')
canvas.pack ()

sheet['A1'] = "Mata Kuliah\t:"
A1 = sheet['A1']
A1.font = font
sheet['A2'] = "Tanggal Perkuliahan\t:"
A2 = sheet['A2']
A2.font = font

sheet['A3'] = "No"
A3 = sheet['A3']
A3.font = font
A3.border = border
A3.alignment = alignment

sheet['B3'] = "Nama"
B3 = sheet['B3']
B3.font = font
B3.border = border
B3.alignment = alignment

sheet['C3'] = "NIM"
C3 = sheet['C3']
C3.font = font
C3.border = border
C3.alignment = alignment

sheet['D3'] = "Jurusan"
D3 = sheet['D3']
D3.font = font
D3.border = border
D3.alignment = alignment

num = 0


def InsertData():
    global num
    num = num + 1
    sheetnum = num + 3

    sheet['A'+str(sheetnum)] = num
    DataNo = sheet['A'+str(sheetnum)]
    DataNo.border = border
    DataNo.alignment = alignment

    sheet['B'+str(sheetnum)] = MasukanNama.get()
    DataNama = sheet['B'+str(sheetnum)]
    DataNama.border = border
    DataNama.alignment = alignment

    sheet['C' + str(sheetnum)] = MasukkanNim.get()
    DataNIM = sheet['C' + str(sheetnum)]
    DataNIM.border = border
    DataNIM.alignment = alignment

    sheet['D' + str(sheetnum)] = jurusanEntry.get()
    DataJurusan = sheet['D' + str(sheetnum)]
    DataJurusan.border = border
    DataJurusan.alignment = alignment

    sheet['B1'] = masukanMatkul.get()
    sheet['B2'] = Masukantanggal.get()

    MasukanNama.delete(0, END)
    MasukkanNim.delete(0, END)
    jurusanEntry.delete(0, END)

def SaveData():
    global informasi
    workbook.save(filename=str(masukanMatkul.get())+"_"+str(Masukantanggal.get())+".xlsx")
    informasi['text'] = "Data absen telah di save!\nNama file: "+str(masukanMatkul.get())+"_"+str(Masukantanggal.get())+".xlsx"

def CreateNewData():
    global informasi, num
    informasi['text'] = 'Klik Insert untuk semua mahasiswa, kemudian klik Save jika semua telah diabsen.'
    MasukanNama.delete(0, END)
    MasukkanNim.delete(0, END)
    jurusanEntry.delete(0, END)
    masukanMatkul.delete(0, END)
    Masukantanggal.delete(0, END)
    num = 0

frame_judul= Frame(root, bg= 'white' )
frame_judul.place(rely=0.025,relx=0.5,relheight=0.1,relwidth=0.8,anchor='n')
judul= Label(frame_judul, bg= 'white',text= ' Absensi Perkuliahan')
judul.place( relheight= 1 , relwidth= 1,)

framematkul= Frame(root, bg= 'white')
framematkul.place(rely=0.2,relx=0.5,relheight=0.06,relwidth=0.8,anchor='n')
namamatkul= Label(framematkul, bg= 'white',text= ' Mata Kuliah', font= styling2)
namamatkul.place ( relheight= 1, relwidth= 0.4)
masukanMatkul = Entry(framematkul)
masukanMatkul.place(relx=0.4,relheight=1,relwidth=0.6)

frametanggal= Frame(root, bg= 'blue')
frametanggal.place(rely=0.27,relx=0.5,relheight=0.06,relwidth=0.8,anchor='n')
infotanggal= Label(frametanggal, bg='white', text='Tanggal perkuliahan', font=styling2 )
infotanggal.place (relwidth=0.4,relheight=1)
Masukantanggal = Entry(frametanggal)
Masukantanggal.place(relx=0.4,relheight=1,relwidth=0.6)
Masukantanggal.get()

frameNama= Frame(root, bg= 'blue')
frameNama.place(rely=0.34,relx=0.5,relheight=0.06,relwidth=0.8,anchor='n')
infonama = Label(frameNama, bg='white', text='Nama', font=styling2)
infonama.place(relwidth=0.4,relheight=1)
MasukanNama = Entry(frameNama)
MasukanNama.place(relx=0.4,relheight=1,relwidth=0.6)
MasukanNama.get()

frameNIM= Frame(root, bg= 'blue')
frameNIM.place(rely=0.41,relx=0.5,relheight=0.06,relwidth=0.8,anchor='n')
infonim = Label(frameNIM, bg='white', text='NIM', font=styling2)
infonim.place(relwidth=0.4,relheight=1)
MasukkanNim= Entry(frameNIM)
MasukkanNim.place(relx=0.4,relheight=1,relwidth=0.6)
MasukkanNim.get()

frameJurusan= Frame(root, bg= 'blue')
frameJurusan.place(rely=0.48,relx=0.5,relheight=0.06,relwidth=0.8,anchor='n')
jurusaninfo = Label(frameJurusan, bg='white', text='Jurusan', font=styling2)
jurusaninfo.place(relwidth=0.4,relheight=1)
jurusanEntry = Entry(frameJurusan)
jurusanEntry.place(relx=0.4,relheight=1,relwidth=0.6)
jurusanEntry.get()

informasi = Label(root, bg='white', font=styling2, text='Klik Insert untuk semua mahasiswa, kemudian klik Save jika semua telah diabsen.')
informasi.place(rely=0.56,relx=0.5,relheight=0.1,relwidth=0.8,anchor='n')


framebutton= Frame(root , bg= 'black')
framebutton.place (rely=0.675,relx=0.5,relheight=0.3,relwidth=0.3,anchor='n')
insert = Button(framebutton, text='Insert', command=InsertData)
insert.place(rely=0,relx=0.5,relheight=0.25,relwidth=1,anchor='n')
save = Button(framebutton, text='Save', command=SaveData)
save.place(rely=0.25,relx=0.5,relheight=0.25,relwidth=1,anchor='n')
createNewData = Button(framebutton, text='Create New', command=CreateNewData)
createNewData.place(rely=0.5,relx=0.5,relheight=0.25,relwidth=1,anchor='n')
Exit = Button(framebutton, text='Exit', command=root.quit)
Exit.place(rely=0.75,relx=0.5,relheight=0.25,relwidth=1,anchor='n')

root.mainloop()