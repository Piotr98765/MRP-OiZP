import pandas as pd
writer = pd.ExcelWriter('strukturazlozona.xlsx', engine='xlsxwriter')


class Produkt:
    def __init__(self, nazwa, zamowienie, tydzien_dostawy, stan_zapasu, cykl_dostawy_produkcji, zapas_zabezpieczajacy
                 , podprodukty, zapotrzebowanie_zalezne, zapotrzebowanie_netto, dostawy_w_drodze, dodatkowe_zamowienie, tydzien_d_z):
        self.nazwa = nazwa
        self.zamowienie = zamowienie
        self.tydzien_dostawy = tydzien_dostawy
        self.stan_zapasu = stan_zapasu
        self.cykl_dostawy_produkcji = cykl_dostawy_produkcji
        self.zapas_zabezpieczajacy = zapas_zabezpieczajacy
        self.podprodukty = podprodukty
        self.zapotrzebowanie_zalezne = zapotrzebowanie_zalezne
        self.zapotrzebowanie_netto = zapotrzebowanie_netto
        self.dostawy_w_drodze = dostawy_w_drodze
        self.dodatkowe_zamowienie = dodatkowe_zamowienie
        self.tydzien_d_z = tydzien_d_z
        

p1 = Produkt('A1', 350, 9, 25, 1, 20, [], 0, 0, 0, 0, 0)
p2 = Produkt('B1', 200, 8, 30, 1, 20, [], 0, 0, 0, 0, 0)
p3 = Produkt('A2', 0, 0, 25, 2, 0, [], 1, 0, 0, 0, 0)
p4 = Produkt('A3', 0, 0, 45, 1, 50, [], 4, 0, 0, 100, 7)
p5 = Produkt('B2', 0, 0, 45, 2, 0, [], 2, 0, 0, 0, 0)
p6 = Produkt('A4', 0, 0, 10, 1, 0, [], 2, 0, 0, 0, 0)
p7 = Produkt('A5', 0, 0, 18, 2, 0, [], 1, 0, 0, 0, 0)
p8 = Produkt('A6', 0, 0, 20, 1, 0, [], 1, 0, 0, 0, 0)
p9 = Produkt('A7', 0, 0, 85, 1, 0, [], 1, 0, 0, 0, 0)
p10 = Produkt('A8', 0, 0, 65, 1, 50, [], 1, 0, 0, 0, 0)
p11 = Produkt('A9', 0, 0, 50, 1, 100, [], 16, 0, 0, 0, 0)
p12 = Produkt('B3', 0, 0, 50, 3, 50, [], 1, 0, 500, 0, 0)
p13 = Produkt('A10', 0, 0, 300, 2, 200, [], 1, 0, 0, 0, 0)
p14 = Produkt('A11', 0, 0, 100, 2, 100, [], 1.6, 0, 0, 0, 0)

parents_p3 = [p1, p2]
p1.zapotrzebowanie_netto = p1.zamowienie - p1.stan_zapasu + p1.zapas_zabezpieczajacy
p2.zapotrzebowanie_netto = p2.zamowienie - p2.stan_zapasu + p2.zapas_zabezpieczajacy


produkty = {
    "p1": p1,
    "p2": p2,
    "p3": p3,
    "p4": p4,
    "p5": p5,
    "p6": p6,
    "p7": p7,
    "p8": p8,
    "p9": p9,
    "p10": p10,
    "p11": p11,
    "p12": p12,
    "p13": p13,
    "p14": p14
}




produktyR = {
    "A1": p1,
    "B1": p2,
    "A2": p3,
    "A11": p14,
    "A3": p4,
    "B2": p2,
    "A4": p6,
    "A5": p7,
    "A6":p8,
    "A7":p9
    
}



rodzice_p4 = [p1, p2]
rodzice_p3 = [p1, p2]
rodzice_p5 = [p2]
rodzice_p6 = [p3]
rodzice_p7 = [p3]
rodzice_p8 = [p3]
rodzice_p9 = [p3]
rodzice_p10 = [p3]
rodzice_p11 = [p1, p2]
rodzice_p12 = [p2]
rodzice_p14 = [p10]
rodzice_p13 = [p4, p5, p6, p7, p8, p9]




lista_rodziców = {"p4": rodzice_p4, "p3": rodzice_p3, "p5": rodzice_p5, 
                  "p6": rodzice_p6, "p7": rodzice_p7, "p8": rodzice_p8,
                 "p9": rodzice_p9, "p10": rodzice_p10, "p11": rodzice_p11, "p12": rodzice_p12, 
                 "p14": rodzice_p14, "p13": rodzice_p13}

#Tabela
tygodnie = 9
kategorie = ['ZZ', 'ZA', 'DO', 'ZB', 'ZN', 'PZ']
dane = {'Tydzień {}'.format(i+1): [0]*len(kategorie) for i in range(tygodnie)}

#Stół
ZZ_p1 = p1.zapas_zabezpieczajacy
ZA_p1 = p1.stan_zapasu
ZB_p1 = p1.zamowienie
ZN_p1 = p1.zapotrzebowanie_netto
PZ_p1 = ZN_p1

df_p1 = pd.DataFrame(dane, index=kategorie)
df_p1.at['ZB', 'Tydzień {}'.format(p1.tydzien_dostawy)] = ZB_p1
df_p1.at['ZN', 'Tydzień {}'.format(p1.tydzien_dostawy)] = ZN_p1
df_p1.at['PZ', 'Tydzień {}'.format(p1.tydzien_dostawy-p1.cykl_dostawy_produkcji)] = PZ_p1
for i in range(p1.tydzien_dostawy-2):
    df_p1.at['ZA', 'Tydzień {}'.format(i+1)] = ZA_p1
# Wyświetlenie tabeli
print('Tabela dla produktu: {}'.format(p1.nazwa))
print(df_p1)
df_p1.to_excel(writer, sheet_name=p1.nazwa)

ZZ_p2 = p2.zapas_zabezpieczajacy
ZA_p2 = p2.stan_zapasu
ZB_p2 = p2.zamowienie
ZN_p2 = p2.zapotrzebowanie_netto
PZ_p2 = ZN_p2

df_p2 = pd.DataFrame(dane, index=kategorie)
df_p2.at['ZB', 'Tydzień {}'.format(p2.tydzien_dostawy)] = ZB_p2
df_p2.at['ZN', 'Tydzień {}'.format(p2.tydzien_dostawy)] = ZN_p2
df_p2.at['PZ', 'Tydzień {}'.format(p2.tydzien_dostawy-p2.cykl_dostawy_produkcji)] = PZ_p2
for i in range(p2.tydzien_dostawy-2):
    df_p2.at['ZA', 'Tydzień {}'.format(i+1)] = ZA_p2
# Wyświetlenie tabeli
print('Tabela dla produktu: {}'.format(p2.nazwa))
print(df_p2)

df_p2.to_excel(writer, sheet_name=p2.nazwa)
lokalizacja_A2 = {}
wartosciR={}
tabela_dict = {}
slownik_tabel={}
lista_wykorzystanych_rodziców =[]
for produkt, rodzice in lista_rodziców.items():
    produkt_liczony = produkty[produkt]
    posortowane_rodzice = sorted(rodzice, key=lambda x: x.tydzien_dostawy)
    df_p = pd.DataFrame(dane, index=kategorie)
    #print(produkt_liczony.dodatkowe_zamowienie)
    #print(produkt_liczony.tydzien_d_z)
    lista_wykorzystanych_rodziców.append(produkt_liczony.nazwa)
    first_iteration = True
    for p in posortowane_rodzice:
        if p.nazwa not in lista_wykorzystanych_rodziców:
            produkt_liczony.zamowienie = p.zapotrzebowanie_netto*produkt_liczony.zapotrzebowanie_zalezne
            produkt_liczony.tydzien_dostawy = p.tydzien_dostawy - 1
            ZZ_produkt_liczony = produkt_liczony.zapas_zabezpieczajacy
            if first_iteration:
                ZZ_produkt_liczony = produkt_liczony.zapas_zabezpieczajacy
                ZA_produkt_liczony = produkt_liczony.stan_zapasu
                first_iteration = False
            else:
                ZZ_produkt_liczony = produkt_liczony.zapas_zabezpieczajacy
                ZA_produkt_liczony = produkt_liczony.zapas_zabezpieczajacy
            if produkt_liczony.tydzien_dostawy == produkt_liczony.tydzien_d_z:
                ZB_produkt_liczony = produkt_liczony.zamowienie  + produkt_liczony.dodatkowe_zamowienie
            else:
                ZB_produkt_liczony = produkt_liczony.zamowienie
            DD_produkt_liczony = produkt_liczony.dostawy_w_drodze
            ZN_produkt_liczony = ZB_produkt_liczony - ZA_produkt_liczony - DD_produkt_liczony + ZZ_produkt_liczony
            if ZN_produkt_liczony < 0:
                ZN_produkt_liczony = 0
            PZ_produkt_liczony = ZN_produkt_liczony
            print(ZN_produkt_liczony)
            #print(p.nazwa, ZN_produkt_liczony)
            wartosciR[produkt_liczony.nazwa, p.nazwa, ZN_produkt_liczony ] = (produkt_liczony.tydzien_dostawy-produkt_liczony.cykl_dostawy_produkcji)
            df_p.at['ZB', 'Tydzień {}'.format(produkt_liczony.tydzien_dostawy)] = ZB_produkt_liczony
            df_p.at['ZN', 'Tydzień {}'.format(produkt_liczony.tydzien_dostawy)] = ZN_produkt_liczony
            df_p.at['PZ', 'Tydzień {}'.format(produkt_liczony.tydzien_dostawy-produkt_liczony.cykl_dostawy_produkcji)] = PZ_produkt_liczony
            for i in range(produkt_liczony.tydzien_dostawy-1):
                df_p.at['ZA', 'Tydzień {}'.format(i+1)] = produkt_liczony.stan_zapasu
            for i in range(produkt_liczony.tydzien_dostawy, 9):
                df_p.at['ZA', 'Tydzień {}'.format(i+1)] = ZZ_produkt_liczony
            for i in range(produkt_liczony.tydzien_dostawy+1):
                df_p.at['ZZ', 'Tydzień {}'.format(i+1)] = ZZ_produkt_liczony

        else:
            
            #print(slownik_tabel[p.nazwa])
            #dla wyrobu B1
            first_iteration_R = True
            for key, value in wartosciR.items():
                if key[0] == 'A2' or key[0] == 'A8':
                    print(key, value)
                    produkt_liczonyR = produktyR[key[1]]
                    produkt_liczony.zamowienie  = key[2]*produkt_liczony.zapotrzebowanie_zalezne
                    ZB_produkt_liczony = produkt_liczony.zamowienie
                    produkt_liczony.tydzien_dostawy = value
                    
                    if first_iteration_R:
                        ZZ_produkt_liczony = produkt_liczony.zapas_zabezpieczajacy
                        ZA_produkt_liczony = produkt_liczony.stan_zapasu
                        ZN_produkt_liczony = ZB_produkt_liczony - ZA_produkt_liczony + ZZ_produkt_liczony
                        PZ_produkt_liczony = ZN_produkt_liczony
                        first_iteration_R = False

                    else:
                        ZZ_produkt_liczony = produkt_liczony.zapas_zabezpieczajacy
                        ZA_produkt_liczony = produkt_liczony.zapas_zabezpieczajacy
                        ZN_produkt_liczony = ZB_produkt_liczony
                        PZ_produkt_liczony = ZN_produkt_liczony

                    df_p.at['ZB', 'Tydzień {}'.format(produkt_liczony.tydzien_dostawy)] = ZB_produkt_liczony
                    df_p.at['ZN', 'Tydzień {}'.format(produkt_liczony.tydzien_dostawy)] = ZN_produkt_liczony
                    df_p.at['PZ', 'Tydzień {}'.format(produkt_liczony.tydzien_dostawy-produkt_liczony.cykl_dostawy_produkcji)] = PZ_produkt_liczony
                    for i in range(produkt_liczony.tydzien_dostawy-1):
                        df_p.at['ZA', 'Tydzień {}'.format(i+1)] = produkt_liczony.stan_zapasu
                    for i in range(produkt_liczony.tydzien_dostawy, 9):
                        df_p.at['ZA', 'Tydzień {}'.format(i+1)] = ZZ_produkt_liczony
                    for i in range(produkt_liczony.tydzien_dostawy+1):
                        df_p.at['ZZ', 'Tydzień {}'.format(i+1)] = ZZ_produkt_liczony
    print('Tabela dla produktu: {}'.format(produkt_liczony.nazwa))
    print(df_p)
    slownik_tabel[produkt_liczony.nazwa] = df_p.copy()


#zn_wartosc = lista_tabel[0].at['ZA', 'Tydzień 1']
#print(zn_wartosc)