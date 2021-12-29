import pandas as pd
import numpy as np
import os
import glob
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import parse_date





class KKPy:
    def __init__(self, npwp, masa_awal, masa_akhir, thn_pajak):
        self.npwp = npwp
        self.masa_awal =  masa_awal
        self.masa_akhir = masa_akhir
        self.thn_pajak = thn_pajak
        self.wb = load_workbook('template.xlsx')
    
    def a1_to_n21(self, masa):
        print("copy ekspor masa " + masa)

        path = os.path.join(os.getcwd(), "Data", self.npwp,  "SPT PPN MASA " + masa +" TAHUN PAJAK " + self.thn_pajak, "SPT MASA PPN LAMPIRAN A1")
        
        if os.path.exists(path):
            filenames = glob.glob(path + os.sep + "*.csv")
            
            tipe_data = {
                'ID_SPT':np.str,
                'NO_FAKTUR':np.str
                }
            df = pd.read_csv(filenames[0], sep=";", dtype=tipe_data)
            df['MASA'] = int(masa)
            df['TANGGAL_FAKTUR'] = df['TANGGAL_FAKTUR'].apply(lambda x: parse_date.parse(x))

            siap_ekspor = df[['NAMA_PARTNER', 'NO_FAKTUR', 'TANGGAL_FAKTUR', 'JUMLAH_DPP', 'KET', 'MASA']].reset_index(drop=True)
            siap_ekspor.index = siap_ekspor.index + 1

            ws = self.wb['N.2.1']
            for r in dataframe_to_rows(siap_ekspor, index=True, header=False):
                ws.append(r)
            print("- copy selesai")
        else:
            print("- masa " + masa + " tidak ada")

    def a2_to_n31(self, masa)    :
        print("copy dipungut sendiri masa " + masa)
        path = os.path.join(os.getcwd(), "Data", self.npwp ,"SPT PPN MASA " + masa +" TAHUN PAJAK " + self.thn_pajak, "SPT MASA PPN LAMPIRAN A2")
        if os.path.exists(path):
            filenames = glob.glob(path + os.sep + "*.csv")
            tipe_data = {
                'ID_SPT':np.str,
                'NPWP_PARTNER':np.str,
                'NPWP_TETAP_PARTNER':np.str,
                'KPP_ADMINISTRASI_PARTNER':np.str,
                'KD_TRX':np.str,
                'NO_FAKTUR':np.str,
                'NO_FAKTUR_PENGGANTI':np.str
                }

            df = pd.read_csv(filenames[0], sep=";", dtype=tipe_data)
            df['MASA'] = int(masa)
            df['TANGGAL_FAKTUR'] = df['TANGGAL_FAKTUR'].apply(lambda x: parse_date.parse(x))

            df01 = df[df['KD_TRX'] == '01']
            siap = df01[['NAMA_PARTNER', 'NPWP_PARTNER', 'NO_FAKTUR', 'TANGGAL_FAKTUR', 'JUMLAH_DPP', 'JUMLAH_PPN', 'JUMLAH_PPNBM', 'NO_FAKTUR_PENGGANTI', 'MASA']].reset_index(drop=True)
            siap.index = siap.index + 1
            
            ws = self.wb['N.3.1']
            for r in dataframe_to_rows(siap, index=True, header=False):
                ws.append(r)
            print("- copy selesai")
        else:
            print("- masa " + masa + " tidak ada")


    def a2_to_n51(self, masa)    :
        print("copy tidak dipungut masa " + masa)
        path = os.path.join(os.getcwd(),  "Data", self.npwp, "SPT PPN MASA " + masa +" TAHUN PAJAK " + self.thn_pajak, "SPT MASA PPN LAMPIRAN A2")
        if os.path.exists(path):
            filenames = glob.glob(path + os.sep + "*.csv")

            tipe_data = {
                'ID_SPT':np.str,
                'NPWP_PARTNER':np.str,
                'NPWP_TETAP_PARTNER':np.str,
                'KPP_ADMINISTRASI_PARTNER':np.str,
                'KD_TRX':np.str,
                'NO_FAKTUR':np.str,
                'NO_FAKTUR_PENGGANTI':np.str
                }
            df = pd.read_csv(filenames[0], sep=";", dtype=tipe_data)
            df['MASA'] = int(masa)
            df['TANGGAL_FAKTUR'] = df['TANGGAL_FAKTUR'].apply(lambda x: parse_date.parse(x))

            df01 = df[df['KD_TRX'] == '07']
        
            siap = df01[['NAMA_PARTNER', 'NPWP_PARTNER', 'NO_FAKTUR', 'TANGGAL_FAKTUR', 'JUMLAH_DPP', 'JUMLAH_PPN', 'JUMLAH_PPNBM', 'NO_FAKTUR_PENGGANTI', 'MASA']].reset_index(drop=True)
            siap.index = siap.index + 1
            ws = self.wb['N.5.1']

            
            for r in dataframe_to_rows(siap, index=True, header=False):
                ws.append(r)
            print("- copy selesai")
        else:
            print("- masa " + masa + " tidak ada")

    def save(self, nama_file):
        self.wb.save(nama_file)
   
        





