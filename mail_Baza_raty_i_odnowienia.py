from openpyxl import load_workbook
from datetime import date, timedelta
import datetime
import time
import re
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from l_h_mailing.py import passw

start_time = time.time()

print('MAILING - Przypomnienia o ratach - szuka...')


class MailingRaty:

    def __init__(self):
        self.wb = load_workbook(filename="M:/Agent baza/2014 BAZA MAGRO.xlsx", read_only=True)
        self.ws = self.wb['BAZA 2014']
        self.cells = self.ws['T8000':'BA20000']
        today = date.today()
        self.week_period = today - timedelta(-5)

    def read_excel(self):
        for email, j1, j2, marka, model, nr_rej, rok_prod, SU, j3, j4, j5, pocz, j7, j8, j9, j10, j11, j12, tu, \
            rodz_ub, nr_polisy, j16, j17, j18, j19, j20, j21, j22, j23, data_raty, kwota, j26, j27, nr_raty \
                in self.cells:
            self.data_raty = data_raty.value
            self.kwota = kwota.value
            self.nr_raty = nr_raty
            self.marka = marka.value
            if self.marka is None:
                self.marka = ''
            self.model = model.value
            if self.model is None:
                self.model = ''
            self.nr_rej = nr_rej.value
            if self.nr_rej is None:
                self.nr_rej = ''
            self.rok_prod = rok_prod.value
            if self.rok_prod is None:
                self.rok_prod = 'b/d'
            self.tu = tu.value
            self.nr_polisy = nr_polisy.value
            self.email = email.value

            yield self.data_raty

    def select_cells(self):
        for self.data_raty in self.read_excel():
            if self.data_raty is not None and re.search('[0-9]', str(self.data_raty)) and not \
                    re.search('[AWV()=.]', str(self.data_raty)):
                data_r = str(self.data_raty)
                self.termin_płatności = data_r[:10]
                if datetime.datetime.strptime(str(self.termin_płatności), '%Y-%m-%d').date() == self.week_period and \
                        int(self.nr_raty.value) > 1:
                    if self.email is not None:
                        di = {'ALL': 'Allianz', 'AXA': 'AXA', 'COM': 'Compensa', 'EPZU': 'PZU', 'GEN': 'Generali',
                              'GOT': 'Gothaer', 'HDI': 'HDI', 'HES': 'Ergo Hestia', 'IGS': 'IGS', 'INT': 'INTER',
                              'LIN': 'LINK 4', 'PRO': 'Proama', 'PZU': 'PZU', 'RIS': 'InterRisk', 'TUW': 'TUW',
                              'TUZ': 'TUZ', 'UNI': 'Uniqa', 'WAR': 'Warta', 'WIE': 'Wiener'}
                        self.tu = di.get(self.tu)

                        yield self.termin_płatności

    def iterate_funct(self):
        for i in self.select_cells():
            print()
            print('Data raty: ' + str(self.termin_płatności))
            print('Marka/kod poczt: ' + str(self.marka))
            print('Model/miasto: ' + str(self.model))
            print('Nr.rej./ulica: ' + str(self.nr_rej))
            print('Rok prod.: ' + str(self.rok_prod))
            print('TU: ' + str(self.tu))
            print('Nr. polisy: ' + str(self.nr_polisy))
            print('Kwota: ' + str(self.kwota))
            print(self.email)

            dash = " - "
            comma_year = ", rok "
            if self.marka == '' and self.model == '':
                dash = ""
                comma_year = ""

            SUBJECT = "Przypomnienie o ... "
            TEXT = "</i></b><br>" \
                   "<br><div style='padding:16px;background:#efefef;'>"\
                   + "Przypomnienie o płatności raty. <br><br><br> " + " Dnia " + str(self.termin_płatności) \
                   + " upływa termin wpłaty raty za polisę nr.: " + str(self.nr_polisy) + " - T.U. " + self.tu \
                   + ", na kwotę " + str(self.kwota) + " zł. <br>" + str(self.marka) + " " + str(self.model) + dash \
                   + str(self.nr_rej) + comma_year + str(self.rok_prod) \
                   + "<br><br> Prosimy o terminową wpłatę." \
                   + "<br><br><br>Z wyrazami szacunku,"\
                   + "<br>MAGRO Ubezpieczenia" \
                   + "<br><br>===================================================================" \
                   + "<br>KONTAKT: https://ubezpieczenia-magro.pl/?pl_biura-ubezpieczen-kontakt,2"

            html = TEXT
            text = 'MAGRO Ubezpieczenia Sp. z o.o.'

            mail = MIMEMultipart('alternative')
            mail['From'] = 'przypomnienia@ubezpieczenia-magro.pl'
            mail['To'] = self.email
            mail['Cc'] = 'ubezpieczenia.magro@gmail.com'
            mail['Subject'] = 'MAGRO Ubezpieczenia Sp. z o.o.'

            part1 = MIMEText(text, 'plain')
            part2 = MIMEText(html, 'html')
            mail.attach(part1)
            mail.attach(part2)
            msg_full = mail.as_string().encode('utf-8')
            server = smtplib.SMTP('ubezpieczenia-magro.home.pl:25')
            server.starttls()
            server.login('przypomnienia@ubezpieczenia-magro.pl', 'dsrhsR3P')
            server.sendmail('przypomnienia@ubezpieczenia-magro.pl', [self.email], msg_full)
            server.quit()
            print()
            if self.email is not None:
                print('Wysłane przypomnienia o nadchodzących ratach!')
            else:
                print('Brak adresu email')


raty = MailingRaty()
raty.read_excel()
raty.select_cells()
raty.iterate_funct()

end_time = time.time() - start_time
print()
print()
print('Czas wykonania: {:.0f} sekund'.format(end_time))

########################################################################################################################

start_time = time.time()

print()
print()
print()
print('MAILING - Przypomnienia o wznowieniach - szuka...')


class MailingOdn:

    def __init__(self):
        self.wb = load_workbook(filename="M:/Agent baza/2014 BAZA MAGRO.xlsx", read_only=True)
        self.ws = self.wb['BAZA 2014']
        self.cells = self.ws['G8000':'BA20000']
        today = date.today()
        self.week_period_odn = today - timedelta(-12)


    def read_excel(self):
        for rozlicz, H, I, J, K, L, M, N, O, P, Q, R, nr_tel, email, U, V, marka, model, przedmiot_ub, rok_prod, SU, \
                AB, AC, AD, pocz, koniec, AG, AH, AI, AJ, AK, tu, rodz_ub, nr_polisy, AO, AP, AQ, AR, AS, AT, AU, \
                    przypis, data_raty, kwota, AY, AZ, nr_raty in self.cells:
            self.rozlicz = rozlicz.value
            self.email = email.value
            self.marka = marka.value
            if self.marka is None:
                self.marka = ''
            self.model = model.value
            if self.model is None:
                self.model = ''
            self.przedmiot_ub = przedmiot_ub.value
            if self.przedmiot_ub is None:
                self.przedmiot_ub = ''
            self.rok_prod = rok_prod.value
            if self.rok_prod is None:
                self.rok_prod = 'b/d'
            self.koniec = koniec.value
            self.tu = tu.value
            self.rodz_ub = rodz_ub.value
            self.nr_polisy = nr_polisy.value
            self.przypis = przypis.value
            self.data_raty = data_raty.value
            self.kwota = kwota.value
            self.nr_raty = nr_raty

            # yield self.data_raty
            yield self.koniec

    def select_cells_odn(self):
        for self.koniec in self.read_excel():
            if self.koniec is not None and re.search('[0-9]', str(self.koniec)):
                # and not \
                #     re.search('[AWV()=.]', str(self.koniec)):

                koniec_okresu = str(self.koniec)
                self.koniec_okresu_bez_sec = koniec_okresu[:10]
                if datetime.datetime.strptime(str(self.koniec_okresu_bez_sec), '%Y-%m-%d').date() == self.week_period_odn:
                    if self.email is not None and self.rodz_ub != 'życ' and self.przypis is not None:
                        d = {'Filipiak': 'Ultimatum, tel. 694888197', 'Pankiewicz': 'R. Pankiewiczem, tel. 577839889',
                             'Wawrzyniak': 'A. Wawrzyniak, tel. 691602675', 'Wołowski': 'M. Wołowskim, tel. 692830084',
                             'Robert': 'naszym biurem, tel. 572 810 576'}
                        if self.rozlicz in d:
                            self.rozlicz = d.get(self.rozlicz)
                        else:
                            self.rozlicz = 'naszym biurem, tel. 602 752 893 lub 42 637 18 42'

                        di = {'ALL': 'Allianz', 'AXA': 'AXA', 'COM': 'Compensa', 'EPZU': 'PZU', 'GEN': 'Generali',
                              'GOT': 'Gothaer', 'HDI': 'HDI', 'HES': 'Ergo Hestia', 'IGS': 'IGS', 'INT': 'INTER',
                              'LIN': 'LINK 4', 'PRO': 'Proama', 'PZU': 'PZU', 'RIS': 'InterRisk', 'TUW': 'TUW',
                              'TUZ': 'TUZ', 'UNI': 'Uniqa', 'WAR': 'Warta', 'WIE': 'Wiener'}
                        self.tu = di.get(self.tu)

                        yield self.koniec_okresu_bez_sec



    def iterate_funct_odn(self):
        for i in self.select_cells_odn():
            print()
            print('Data końca polisy: ' + self.koniec_okresu_bez_sec)
            print('Marka/kod poczt: ' + str(self.marka))
            print('Model/miasto: ' + str(self.model))
            print('Nr.rej./ulica: ' + str(self.przedmiot_ub))
            print('Rok prod.: ' + str(self.rok_prod))
            print('TU: ' + str(self.tu))
            print('Nr. polisy: ' + str(self.nr_polisy))
            print('Przypis z polisy: ' + str(self.przypis))
            print(self.email)


            dash = " - "
            comma_year = ", rok "
            if self.marka == '' and self.model == '':
                dash = ""
                comma_year = ""

            TEXT = "</i></b><br>" \
                   "<br><div style='padding:16px;background:#efefef;'>"\
                   + "Przypomnienie o końcu ochrony ubezpieczeniowej. <br><br><br> " + " Dnia " + str(self.koniec_okresu_bez_sec) \
                   + " dobiega końca Twoja polisa ubezpieczeniowa, nr. " + str(self.nr_polisy) + " - T.U. " + self.tu \
                   + "<br>" + str(self.marka) + " " + str(self.model) + dash \
                   + str(self.przedmiot_ub) + comma_year + str(self.rok_prod) \
                   + "<br><br> W sprawie odnowienia prosimy o kontakt z " + str(self.rozlicz) \
                   + "<br><br><br>Z wyrazami szacunku,"\
                   + "<br>MAGRO Ubezpieczenia" \
                   + "<br><br>===================================================================" \
                   + "<br>https://ubezpieczenia-magro.pl"

            html = TEXT
            text = 'MAGRO Ubezpieczenia Sp. z o.o.'

            mail = MIMEMultipart('alternative')
            mail['From'] = 'przypomnienia@ubezpieczenia-magro.pl'
            mail['To'] = self.email
            mail['Cc'] = 'ubezpieczenia.magro@gmail.com'
            mail['Subject'] = 'MAGRO Ubezpieczenia Sp. z o.o.'

            part1 = MIMEText(text, 'plain')
            part2 = MIMEText(html, 'html')
            mail.attach(part1)
            mail.attach(part2)
            msg_full = mail.as_string().encode('utf-8')
            server = smtplib.SMTP('ubezpieczenia-magro.home.pl:25')
            server.starttls()
            server.login('przypomnienia@ubezpieczenia-magro.pl', passw)
            server.sendmail('przypomnienia@ubezpieczenia-magro.pl', [self.email], msg_full) ##
            server.quit()
            print()
            if self.email is not None:
                print('Wysłane przypomnienia o końcu polis!')
            if self.email == '':
                print('Brak adresu email')


raty = MailingOdn()
raty.read_excel()
raty.select_cells_odn()
raty.iterate_funct_odn()


end_time = time.time() - start_time
print()
print()
print('Czas wykonania: {:.0f} sekund'.format(end_time))
time.sleep(120)
