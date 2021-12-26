import kkpy


def main():
    # npwp = input('NPWP : ')
    # masa_awal = input('Masa Awal :')
    # masa_akhir = input('Masa Akhir :')
    
    Test = kkpy.KKPy('010619096057000', '01', '12', '2020')
    for i in range(int(Test.masa_awal), int(Test.masa_akhir) + 1):
        Test.a1_to_n21(str(i))

    for i in range(int(Test.masa_awal),int(Test.masa_akhir) + 1 ):
        Test.a2_to_n31(str(i))

    Test.save("KKP Baru.xlsx")


if __name__ == '__main__':
    main()