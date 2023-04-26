import pandas as pd
writer = pd.ExcelWriter('strukturaprosta.xlsx', engine='xlsxwriter')


class Produkt:
    def __init__(self, nazwa, zamowienie, tydzien_dostawy, stan_zapasu, cykl_dostawy_produkcji, zapas_zabezpieczajacy
                 , podprodukty, zapotrzebowanie_zalezne, zapotrzebowanie_netto):
        self.nazwa = nazwa
        self.zamowienie = zamowienie
        self.tydzien_dostawy = tydzien_dostawy
        self.stan_zapasu = stan_zapasu
        self.cykl_dostawy_produkcji = cykl_dostawy_produkcji
        self.zapas_zabezpieczajacy = zapas_zabezpieczajacy
        self.podprodukty = podprodukty
        self.zapotrzebowanie_zalezne = zapotrzebowanie_zalezne
        self.zapotrzebowanie_netto = zapotrzebowanie_netto

p1 = Produkt('Stół', 100, 7, 0, 1, 0, [], 0, 0,)
p2 = Produkt('Blat', 0, 0, 20, 1, 0, [], 1, 0)
p3 = Produkt('Nogi', 0, 0, 0, 1, 0, [], 4, 0)
p4 = Produkt('Płyta', 0, 0, 0, 2, 0, [], 1.2, 0)
p5 = Produkt('Kantówka', 0, 0, 0, 3, 0, [], 0.8, 0)
p6 = Produkt('Okucia', 0, 0, 50, 5, 0, [], 1, 0)

p1.zapotrzebowanie_netto = p1.zamowienie - p1.stan_zapasu + p1.zapas_zabezpieczajacy


produkty = {"p1": p1, "p2": p2, "p3": p3, "p4": p4, "p5": p5, "p6": p6}

podprodukty_p1 = [p2, p3, p6]
podprodukty_p2 = [p4]
podprodukty_p3 = [p5]

lista_podproduktów = {"p1": podprodukty_p1, "p2": podprodukty_p2, "p3": podprodukty_p3}

#Tabela
tygodnie = 11
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
#print('Tabela dla produktu: {}'.format(p1.nazwa))
#print(df_p1)
df_p1.to_excel(writer, sheet_name=p1.nazwa)


for parent, podprodukty in lista_podproduktów.items():
    parent_product = produkty[parent]
    for p in podprodukty:
        p.zamowienie = parent_product.zapotrzebowanie_netto*p.zapotrzebowanie_zalezne
        p.tydzien_dostawy = parent_product.tydzien_dostawy - 1
        ZZ_p = p.zapas_zabezpieczajacy
        ZA_p = p.stan_zapasu
        ZB_p = p.zamowienie 
        ZN_p = p.zamowienie - p.stan_zapasu + p.zapas_zabezpieczajacy
        PZ_p = ZN_p
        p.zapotrzebowanie_netto = ZN_p
        df_p = pd.DataFrame(dane, index=kategorie)
        df_p.at['ZB', 'Tydzień {}'.format(p.tydzien_dostawy)] = ZB_p
        df_p.at['ZN', 'Tydzień {}'.format(p.tydzien_dostawy)] = ZN_p
        df_p.at['PZ', 'Tydzień {}'.format(p.tydzien_dostawy-p.cykl_dostawy_produkcji)] = PZ_p
        for i in range(p.tydzien_dostawy-1):
            df_p.at['ZA', 'Tydzień {}'.format(i+1)] = ZA_p
        # Wyświetlenie tabeli
        print('Tabela dla produktu: {}'.format(p.nazwa))
        print(df_p)
        df_p.to_excel(writer, sheet_name=p.nazwa)
#writer.save()