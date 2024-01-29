from automate_report import gen_reporte
import pandas as pd

# def test():
#     # Use a breakpoint in the code line below to debug your script.
#     read_report('Kam y Asociados - Mayo 2023 Siniestralidades.xlsx')


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    gen_reporte('May2023', [7, 2022], [10, 2022], num_pol=[1234, 5678, 91011])

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
