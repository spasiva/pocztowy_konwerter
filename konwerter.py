#!/usr/bin/python3
#
# Copyright (C) 2017 Łukasz Kopacz
#
# This file is part of Pocztowy konwerter.
#
# Pocztowy konwerter is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# Pocztowy konwerter is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with Pocztowy konwerter. If not, see <http://www.gnu.org/licenses/>.

import xlwt
import xml.etree.ElementTree as ET
import os


def convert_to_template(filenamefull):
    try:
        tree = ET.parse(filenamefull)
        root = tree.getroot()

        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet('Arkusz1')

        nazwy = ['NumerNadania', 'AdresatNazwa', 'AdresatNazwaCd', 'AdresatUlica', 'AdresatNumerDomu', 'AdresatNumerLokalu',
                 'AdresatKodPocztowy', 'AdresatMiejscowosc', 'AdresatKraj', 'AdresatEmail', 'AdresatMobile', 'AdresatTelefon',
                 'AdresatPosteRestante', 'AdresatOsobaKontaktowa', 'AdresatNIP', 'Gabaryt', 'KategoriaLubGwarancjaTerminu',
                 'Masa', 'KwotaPobrania', 'NRB', 'TytulPobrania', 'SprawdzenieZawartosci', 'Wartosc', 'Ubezpieczenie',
                 'GodzinaDoreczenia', 'Ostroznie', 'Uwagi', 'Zawartosc', 'PotwierdzenieOdbioru',
                 'ElektronicznePotwierdzenieOdbioru', 'OdbiorWPunkcie', 'OdbiorNazwa', 'OdbiorNazwaCd', 'OdbiorUlica',
                 'OdbiorNumerDomu', 'OdbiorNumerLokalu', 'OdbiorKodPocztowy', 'OdbiorMiejscowosc', 'OdbiorOsobaKontaktowa',
                 'DoreczenieNazwa', 'DoreczenieNazwaCd', 'DoreczenieUlica', 'DoreczenieNumerDomu', 'DoreczenieNumerLokalu',
                 'DoreczenieKodPocztowy', 'DoreczenieMiejscowosc', 'DoreczenieOsobaKontaktowa', 'Platnik', 'PlatnikNazwa',
                 'PlatnikNazwaCd', 'PlatnikUlica', 'PlatnikNumerDomu', 'PlatnikNumerLokalu', 'PlatnikKodPocztowy',
                 'PlatnikMiejscowosc', 'PlatnikNIP', 'RodzajPalety', 'SzerokoscPalety', 'DlugoscPalety', 'WysokoscPalety',
                 'IloscZwracanychPaletEuro', 'ZalaczonaFaktura', 'ZalaczonyWZ', 'ZalaczoneDokumenty', 'ZwracanaFaktura',
                 'ZwracanyWZ', 'ZwracaneDokumenty', 'DataZaladunku', 'DataDostawy', 'IdMultiPack', 'PowiadomienieNadawcy',
                 'PowiadomienieOdbiorcy', 'TypOpakowania', 'UiszczaOplate', 'Ponadgabaryt', 'DoreczenieWDzien',
                 'DoreczenieWSobote', 'OdbiorWSobote', 'DoreczenieDoRakWlasnych', 'DokumentyZwrotne',
                 'DoreczenieWNiedziele/Swieto', 'Doreczenie20:00-7:00', 'OdbiorWNiedziele/Swieto', 'Odbior20:00-7:00',
                 'Miejscowa', 'ObszarMiasto', 'DeklaracjaWaluta', 'DeklaracjaNrRefImportera', 'DeklaracjaNrTelImportera',
                 'DeklaracjaOplatyPocztowe', 'DeklaracjaUwagi', 'DeklaracjaLicencja', 'DeklaracjaSwiadectwo',
                 'DeklaracjaFaktura', 'DeklaracjaPodarunek', 'DeklaracjaDokument', 'DeklaracjaProbkaHandlowa',
                 'DeklaracjaZwrotTowaru', 'DeklaracjaTowary', 'DeklaracjaInny', 'DeklaracjaWyjasnienie',
                 'DeklaracjaOkreslenieZawartosci1', 'DeklaracjaIlosc1', 'DeklaracjaMasa1', 'DeklaracjaWartosc1',
                 'DeklaracjaNrTaryfowy1', 'DeklaracjaKrajPochodzenia1', 'DeklaracjaOkreslenieZawartosci2', 'DeklaracjaIlosc2',
                 'DeklaracjaMasa2', 'DeklaracjaWartosc2', 'DeklaracjaNrTaryfowy2', 'DeklaracjaKrajPochodzenia2',
                 'DeklaracjaOkreslenieZawartosci3', 'DeklaracjaIlosc3', 'DeklaracjaMasa3', 'DeklaracjaWartosc3',
                 'DeklaracjaNrTaryfowy3', 'DeklaracjaKrajPochodzenia3', 'DeklaracjaOkreslenieZawartosci4', 'DeklaracjaIlosc4',
                 'DeklaracjaMasa4', 'DeklaracjaWartosc4', 'DeklaracjaNrTaryfowy4', 'DeklaracjaKrajPochodzenia4',
                 'DeklaracjaOkreslenieZawartosci5', 'DeklaracjaIlosc5', 'DeklaracjaMasa5', 'DeklaracjaWartosc5',
                 'DeklaracjaNrTaryfowy5', 'DeklaracjaKrajPochodzenia5', 'MPK', 'ZasadySpecjalne',
                 'DokumentyZwrotneProfilAdresowyNazwa', 'Serwis', 'PROCEDURA_SERWIS', 'PROCEDURA_ZawartoscNazwa',
                 'PROCEDURA_ListaCzynnosciNazwa', 'numerPrzesylkiKlienta'
        ]

        for i in range(0, len(nazwy)):
            sheet.write(0, i, nazwy[i])

        k = 0
        for child in root:
            for child2 in child:
                k += 1
                # print('tag: {}    attrib: {}'.format(child2.tag, child2.attrib))
                for child3 in child2:
                    # print('tag: {}    attrib: {}'.format(child3.tag, child3.attrib))
                    if child3.attrib['Nazwa'] == 'Nazwa':
                        sheet.write(k, 1, child3.text)
                    if child3.attrib['Nazwa'] == 'NazwaII':
                        sheet.write(k, 2, child3.text)
                    if child3.attrib['Nazwa'] == 'Ulica':
                        sheet.write(k, 3, child3.text)
                    if child3.attrib['Nazwa'] == 'Dom':
                        sheet.write(k, 4, child3.text)
                    if child3.attrib['Nazwa'] == 'Lokal':
                        sheet.write(k, 5, child3.text)
                    if child3.attrib['Nazwa'] == 'Kod':
                        sheet.write(k, 6, child3.text[:2] + '-' + child3.text[2:])
                    if child3.attrib['Nazwa'] == 'Miejscowosc':
                        sheet.write(k, 7, child3.text)
                    if child3.attrib['Nazwa'] == 'Kraj':
                        sheet.write(k, 8, child3.text)
                    if child3.attrib['Nazwa'] == 'CzyEZwrot':
                        sheet.write(k, 12, child3.text)
                    if child3.attrib['Nazwa'] == 'Strefa':
                        sheet.write(k, 15, child3.text)
                    if child3.attrib['Nazwa'] == 'Kategoria':
                        sheet.write(k, 16, child3.text)
                    if child3.attrib['Nazwa'] == 'Masa':
                        sheet.write(k, 17, int(child3.text) * 0.001 )
                    if child3.attrib['Nazwa'] == 'IloscPotwOdb':
                        if int(child3.text) > 0:
                            sheet.write(k, 28, 'T')
                        else:
                            sheet.write(k, 28, 'N')

        filename, file_extension = os.path.splitext(filenamefull)
        output_full = filename + '_wynik.xls'

        # Checking for duplicates & changing name of duplicates
        if os.path.isfile(output_full):
            base, extension = os.path.splitext(output_full)
            file_num = 2
            while os.path.isfile('{}_{}{}'.format(base, file_num, extension)):
                file_num += 1
            output_full = '{}_{}{}'.format(base, file_num, extension)

        workbook.save(output_full)
        if os.path.isfile(output_full):
            return output_full
        else:
            return False
    except ET.ParseError:
        return -1
