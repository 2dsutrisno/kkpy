import datetime

def parse(input):
    dic_bulan = {
        'Januari':1,
        'Februari':2,
        'Maret':3,
        'April':4,
        'Mei':5,
        'Juni':6,
        'Juli':7,
        'Agustus':8,
        'September':9,
        'Oktober':10,
        'Nopember':11,
        'Desember':12
    }
    list_tanggal = input.split(', ')[1].split(' ')
    thn = int(list_tanggal[2])
    bln = dic_bulan[list_tanggal[1]]
    tgl = int(list_tanggal[0])
    return datetime.datetime(thn, bln, tgl)
