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
    
    def copy_ekspor(self, masa): #ke N.2.1
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

    def copy_pungut_sendiri(self, masa): #ke N.3.1 kode 01, 04 & 09
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

            df01 = df[(df.KD_TRX == '01') | (df.KD_TRX == '04') | (df.KD_TRX == '09')]
            siap = df01[['NAMA_PARTNER', 'NPWP_PARTNER', 'NO_FAKTUR', 'TANGGAL_FAKTUR', 'JUMLAH_DPP', 'JUMLAH_PPN', 'JUMLAH_PPNBM', 'NO_FAKTUR_PENGGANTI', 'MASA']].reset_index(drop=True)
            siap.index = siap.index + 1
            
            ws = self.wb['N.3.1']
            for r in dataframe_to_rows(siap, index=True, header=False):
                ws.append(r)
            print("- copy selesai")
        else:
            print("- masa " + masa + " tidak ada")

    def copy_dipungut_pemungut(self, masa): #ke N.4.1 kode 02 & 03
        print("copy dipungut pemungut masa " + masa)
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

            df01 = df[(df.KD_TRX == '02') | (df.KD_TRX == '03') ]
            siap = df01[['NAMA_PARTNER', 'NPWP_PARTNER', 'NO_FAKTUR', 'TANGGAL_FAKTUR', 'JUMLAH_DPP', 'JUMLAH_PPN', 'JUMLAH_PPNBM', 'NO_FAKTUR_PENGGANTI', 'MASA']].reset_index(drop=True)
            siap.index = siap.index + 1
            
            ws = self.wb['N.4.1']
            for r in dataframe_to_rows(siap, index=True, header=False):
                ws.append(r)
            print("- copy selesai")
        else:
            print("- masa " + masa + " tidak ada")
        
    def copy_tdk_dipungut(self, masa): #ke N.5.1 kode 07
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

    def copy_dibebaskan(self, masa): #ke N.6.1 kode 08
        print("copy dibebaskan masa " + masa)
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

            df01 = df[df['KD_TRX'] == '08']
        
            siap = df01[['NAMA_PARTNER', 'NPWP_PARTNER', 'NO_FAKTUR', 'TANGGAL_FAKTUR', 'JUMLAH_DPP', 'JUMLAH_PPN', 'JUMLAH_PPNBM', 'NO_FAKTUR_PENGGANTI', 'MASA']].reset_index(drop=True)
            siap.index = siap.index + 1
            ws = self.wb['N.6.1']

            
            for r in dataframe_to_rows(siap, index=True, header=False):
                ws.append(r)
            print("- copy selesai")
        else:
            print("- masa " + masa + " tidak ada")

    def copy_impor(self, masa): #ke N.8.1
        print("copy impor masa " + masa)
        path = os.path.join(os.getcwd(),  "Data", self.npwp, "SPT PPN MASA " + masa +" TAHUN PAJAK " + self.thn_pajak, "SPT MASA PPN LAMPIRAN B1")

        if os.path.exists(path):
            filenames = glob.glob(path + os.sep + "*.csv")

            tipe_data = {
                'ID_SPT':np.str,
                'NO_FAKTUR':np.str,
                }
            df = pd.read_csv(filenames[0], sep=";", dtype=tipe_data)
            df['MASA'] = int(masa)
            df['TANGGAL_FAKTUR'] = df['TANGGAL_FAKTUR'].apply(lambda x: parse_date.parse(x))

            siap = df[['NM_PARTNER', 'NO_FAKTUR', 'TANGGAL_FAKTUR', 'JUMLAH_DPP', 'JUMLAH_PPN', 'JUMLAH_PPNBM', 'MASA']].reset_index(drop=True)
            siap.index = siap.index + 1
            ws = self.wb['N.8.1']
            
            for r in dataframe_to_rows(siap, index=True, header=False):
                ws.append(r)
            print("- copy selesai")
        else:
            print("- masa " + masa + " tidak ada")

    def save(self, nama_file): #save workbook
        self.wb.save(nama_file)
   
        





