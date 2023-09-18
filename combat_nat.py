import matplotlib.pyplot as plt
import pandas as pd
import yfinance as yf
import numpy as np
pd.set_option("mode.chained_assignment",None)

print(" +-+-+-+-+ ")
print(" |C|O|D|E|")
print(" +-+-+-+-+ \n")
print(" +-+-+ ")
print(" |B|Y| ")
print(" +-+-+ \n")
print(" +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+ ")
print(" |O|U|G|I|_|L|E|S|G|_|A|F|A|R|D| ")
print(" +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+ \n")

def tranform(matrice): #supprime colonnes inutiles et ajoute specialdate
    matrice = pd.DataFrame(matrice)
    A = np.random.random()
    B = np.random.random()
    C = np.random.random()
    D = np.random.random()
    nombre = int((A*B+C+D*A)*100000)
    nom = "save" +str(nombre) + ".csv"
    matrice.to_csv(nom)
    matrice = pd.read_csv(nom)

    matrice["Valeur"] = matrice["Close"]
    matrice["specialdate"] = 0
    for i in range(0, len(matrice)):
        matrice["specialdate"][i] = int(str(matrice["Date"][i][:4]) + str(matrice["Date"][i][5:7]) + str(matrice["Date"][i][8:10]))
    for i in range(0, len(matrice)):
        matrice["Date"][i] = str(matrice["specialdate"][i])[0:4] + "-" + str(matrice["specialdate"][i])[4:6]

    matrice = matrice.drop(["Volume", "Low", "Open", "High", "Adj Close", "Close"], axis=1)
    matrice.to_csv(nom)
    return matrice

def Firstmonth(matrice): #garde le premier jour du mois
    temp = matrice
    rejection = []
    mois = str(matrice["specialdate"][0])[4:6]
    for i in range(1,len(matrice)):
        if int(mois) == int(str(matrice["specialdate"][i])[4:6]):
            rejection.append(i)

        else:
            if (int(mois) == 12) and (int(str(matrice["specialdate"][i])[4:6]) == 1):
                mois = str("01")
            else:
                mois = str(matrice["specialdate"][i])[4:6]

    temp.drop(rejection,inplace=True)
    temp = temp.reset_index(drop=True)
    for i in range(0, len(temp)):
        temp["specialdate"][i] = str(temp["specialdate"][i])[:6]
    return temp

def finmatrice(matrice):  #calculation de l'évolution de 100€
    temp1 = matrice
    temp1["valorisation"] = 0
    temp1["valorisation"][0] = 100
    temp1["unit"] = 100/temp1["Valeur"][0]
    for i in range(1,len(temp1)):
        temp1["valorisation"][i] = temp1["unit"][i] * temp1["Valeur"][i]
    return temp1

def creer_liste(): #liste des tickers
    matrice = []
    T = True
    while T == True:
        in1 = input("quel est le ticker, si fin alors rentrer 007 : ")
        if in1 != "007":
            matrice.append(in1)
        else:
            T = False
    print("\n")
    return matrice

def nomfinal(): #nom final pour le ficher .csv 
    T = True
    while T:
        nomfichersave = input("quel est le nom que à donner au ficher csv sous la forme XXX.csv ? ")
        format = nomfichersave[len(nomfichersave) - 4:]
        if format == '.csv' and not len(nomfichersave) < 5:
            T = False
        else:
            print("Le nom n'est pas correct, il manque le '.csv', ou il est trop court")
    print("\n\n")
    return nomfichersave

def demandeDate(): #la date pour commencer la bataille
    T = True
    while T:
        int1 = input("entrer la date de depart sous format YYYY-MM-DD : ")
        temp1 = (int(int1[0:4]) >= 1850) and (int(int1[0:4]) <= 2023)
        temp2 = (int(int1[5:7]) >= 1) and (int(int1[5:7]) <= 12)
        temp3 = (int(int1[8:10]) >= 1) and (int(int1[8:10]) <= 31)
        if ((temp1 and temp2) and temp3) == True:
            T = False
            print("La date {0} est correcte".format(int1))
            return int1
        else:
            print("La date insérée est incorrecte")

def telechargement(ticker, date):
    matrice = []
    if len(ticker) > 1:
        for i in range(0,len(ticker)):
            temp = yf.download(ticker[i], start=date, interval="1d")
            matrice.append(temp)
            

    else:
        matrice = yf.download(ticker, start=date, interval="1d")
    return matrice

def MyBattle():
    liste = creer_liste()

    date = demandeDate()
    print("\n")

    matrice = telechargement(liste, date)

    NbTicker = len(matrice)
    for i in range(0,NbTicker):
        matrice[i] = tranform(matrice[i])
        matrice[i] = Firstmonth(matrice[i])
        matrice[i] = finmatrice(matrice[i])

    print("\n")
    nameFile = nomfinal()
    Final = pd.DataFrame(columns=liste)

    for i in range(0, NbTicker):
        Final[liste[i]] = matrice[i]["valorisation"]
    Final["Date"]=0
    for i in range(0, len(matrice[0])):
        Final["Date"][i] = matrice[0]["Date"][i]

    print("\n")
    column = Final.pop("Date")
    Final.insert(0,"Date", column)

    PER = []
    for i in range(0, len(Final)):
        Final["Date"][i]=(float(Final["Date"][i][:4])+float(float(Final["Date"][i][5:])/13))
    color = ["b","g","r", "c", "m","k","y"]
    for i in range(0, len(liste)):
        plot = plt.plot(Final["Date"], Final[liste[i]], c = color[i])
    plt.figlegend(liste)
    plt.show()

    #Final.to_csv(nameFile)
    

MyBattle()




