import kkpy


def main():
    # npwp = input('NPWP : ')
    # masa_awal = input('Masa Awal :')
    # masa_akhir = input('Masa Akhir :')
    
    Test = kkpy.KKPy('010619096057000', '01', '12', '2020')

    for i in range(int(Test.masa_awal), int(Test.masa_akhir) + 1):
       Test.copy_ekspor(str(i))
       Test.copy_pungut_sendiri(str(i))
       Test.copy_dipungut_pemungut(str(i))
       Test.copy_tdk_dipungut(str(i))
       Test.copy_dibebaskan(str(i))
       Test.copy_impor(str(i))
    Test.save("KKP Baru.xlsx")


if __name__ == '__main__':
    main()