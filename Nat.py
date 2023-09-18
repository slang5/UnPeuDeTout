import pandas as pd
import yfinance as yf
import time as TTTT
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


def definir_nom():
    true = True
    while true:
        nomfichersave = input("quel est le nom que à donner au ficher csv sous la forme XXX.csv ? ")
        format = nomfichersave[len(nomfichersave) - 4:]
        if format == '.csv' and not len(nomfichersave) < 5:
            true = False
        else:
            print("Le nom n'est pas correct, il manque le '.csv' ou il est trop court")
    print("\n\n")
    return nomfichersave

def telechargement_action(TYPE, nomfichersave, value):
    if str(value) == "0":
        temp2 = True
        while temp2 == True:
            nomAction = input("quel est le ticker de l'action sur yahoo ? ")
            if len(yf.download(nomAction)) > 1:
                temp2 = False
        print("\n\n")
    else:
        nomAction = value
    
    if int(TYPE) == int(1):
            temp3 = True
            while temp3 == True:
                debut = input("debut sous format YYYY-MM-DD : ")
                temp4 = (int(debut[0:4]) > 1895 and int(debut[0:4]) < 2023)
                temp5 = (int(debut[5:7]) > 0 and int(debut[5:7]) < 13)
                temp6 = (int(debut[8:10]) > 0 and int(debut[8:10]) < 32)
                temp7 = ((temp4 and temp5) and temp6)
                if  temp7 == True:
                    data = yf.download(nomAction, period="max", start=debut)
                    temp3 = False
                else:
                    print("La date est mauvaise")
        
    else:
        data = yf.download(nomAction, period="max")
    print("historique est telecharge\n\n")

    donnees = pd.DataFrame(data)
    return donnees.to_csv(nomfichersave)

def checkcsvfile(binaire):
    if binaire == "0":
        temp8 = True
        while temp8 == True:
            Nom = input("nom du ficher CSV à traiter : ")
            if (str(Nom[len(Nom)-4:]) == ".csv"):
                temp8 = False
            else:
                print("Il manque le '.csv' ")
    else:
        Nom = binaire
    print("\n")
    return Nom

def injectionMensuelle():
    temp9 = True
    while temp9 == True:
        mensuel = float(input("quel montant mensuel doit être injecté ? "))
        if ((mensuel > 0) and (mensuel < 9999999999999)) == True:
            temp9 = False
        else:
            print("Le montant inséré est incorrect")
    print("\n")
    return mensuel

def yieldSaving():
    temp10 = True
    while temp10 == True:
        taux_annuel = float(input("quel taux annuel de rémunération ? sous la forme X% "))
        if ((taux_annuel > 0) and (taux_annuel < 9999999999999)) == True:
            temp10 = False
        else:
            print("Le taux inséré est incorrect")
    print("\n")
    return taux_annuel

def defineFinalName():
    temp11 = True
    while temp11 == True:
        Nom_final = input("quel est le nom du fichier excel final ? ")
        count = 0
        for i in Nom_final:
            for z in "&é"'(-èçà)='"^¨$£¤*µ;.,?/:§!€":
                if i == z:
                    count +=1
                    print(z, i)
        if count == 0:
            temp11 = False
        else:
            print("le nom donné est incorrect")
    print("\n")
    return Nom_final

def preprocessData(input1):
    UN = checkcsvfile(input1)
    DEUX = injectionMensuelle()
    TROIS = yieldSaving()
    QUATRE = defineFinalName()
    return UN, DEUX, TROIS, QUATRE

def tranferer(finalName, matrice):
    nameCSV = str(finalName + ".csv")
    nameEXCEL = str(finalName + ".xlsx")
    matrice.to_csv(nameCSV)
    matrice.to_excel(nameEXCEL)
    return nameCSV, nameEXCEL

def processData(matrix):
    csvname = matrix[0]
    monthly = matrix[1]
    rendement = matrix[2]
    finalName = matrix[3]

    DoNotChange = pd.read_csv(csvname,delimiter=",")
    Final = pd.DataFrame(DoNotChange)
    taille = len(Final)
    Nom = pd.DataFrame(Final)

    Excel = pd.DataFrame()
    Excel["Valeur"] = Final["Close"]
    Excel["Special_date"] = 0

    for i in range(0,taille):
        Excel["Special_date"][i] = int(str(Final["Date"][i][:4]) + str(Final["Date"][i][5:7]) + str(Final["Date"][i][8:10]))

    Excel["firstOfTheMonth"] = 0
    mois_actuel = 0
    for i in range(0,taille):
        mois = int(str(Excel["Special_date"][i])[4:6])
        if mois_actuel + 1 == mois:
            Excel["firstOfTheMonth"][i] = 1
            mois_actuel = (mois_actuel + 1) % 12

    for i in range(0,taille):
        value = Excel["firstOfTheMonth"][i]
        if value == 0:
            Excel.drop([i], axis=0, inplace=True)

    taille = len(Excel)
    Excel = Excel.reset_index(drop=True)

    for i in range(0, taille):
        date = str(Excel["Special_date"][i])
        date = int(date[:6])
        Excel["Special_date"][i] = date

    Excel.drop("firstOfTheMonth",axis=1, inplace=True)

    Excel["Nb_action"], Excel["Valo"] = float(0), float(0)
    Excel["Nb_action"][0] = monthly / Excel["Valeur"][0]
    Excel["Valo"][0] = Excel["Nb_action"][0] * Excel["Valeur"][0]

    for i in range(1,taille):
        Stock_volume = float(monthly / Excel["Valeur"][i])
        Excel["Nb_action"][i] = Excel["Nb_action"][i-1] + Stock_volume
        Excel["Valo"][i] = Excel["Nb_action"][i] * Excel["Valeur"][i]

    Excel["Epargne"] = float(0)
    Excel["Epargne"][0] = monthly
    count = 0
    for i in range(1, taille):
        Excel["Epargne"][i] = Excel["Epargne"][i - 1] + monthly
        date = int(str(Excel["Special_date"][i])[4:])
        if  date == 12:
            Excel["Epargne"][i] = Excel["Epargne"][i] * (1 + rendement / 100)

    Excel["Date"] = str(0)
    for i in range(0,taille):
        Excel["Date"][i] = str(Excel["Special_date"][i])[:4] + "-" + str(Excel["Special_date"][i])[4:]

    Excel.drop("Special_date", axis=1, inplace=True)
    verif = tranferer(finalName, Excel)
    print('Les fichers {0} et {1} sont téléchargés'.format(verif[0], verif[1]))

def onlyDownload1():
    temp1 = definir_nom()
    print("le nom est {0}\n\n".format(temp1))

    inside = input("si vous voulez une date de début particulière alors tapez 1 sinon tout sauf 1 : ")
    print("\n")
    telechargement_action(inside, temp1, "0")
    return temp1

def faire_choix():
    def question1():
        print("1 - travailler sur un seul actif \n2 - travailler sur plusieurs actifs (uniquement telecharger)")
        in1 = input("1 ou 2 ==> ")
        return in1
    def question2():
        print("1 - télécharger des données\n2 - traiter des données\n3- faire les deux")
        in2 = input("1, 2 ou 3 ==> ")
        return in2
    in1 = 12
    in2 = 12

    while in1 <= 0 or in1 >= 3:
        in1 = int(question1())
    print("\n")
    while in2 <= 0 or in2 >=4:
        in2 = int(question2())

    return in1, in2




application = True
TTTT.sleep(2)
while application:
    CHOIX = faire_choix()

    if CHOIX[0] == 1 and CHOIX[1] == 1: #uniquement telecharger
        onlyDownload1() #OK

    elif (CHOIX[0] == 1 and CHOIX[1] == 2): #uniquement traiter
        temp11 = preprocessData("0")
        processData(temp11)

    elif CHOIX[0] == 1 and CHOIX[1] == 3: # telecharger et traiter un documnet
        value = onlyDownload1()
        temp11 = preprocessData(value)
        processData(temp11)

    elif CHOIX[0] == 2 and CHOIX[1] == 1: #telecharger plusieurs actions
        listTickers = []
        varStr = input("rentrer un ticker ou 0 pour terminer la liste : ")
        while varStr != "0":
            listTickers.append(varStr)
            #print("{0} est ajouté".format(varStr))
            varStr = input("rentrer un ticker ou 0 pour terminer la liste : ")
        for i in listTickers:
            #print(i, str(i)+".csv")
            telechargement_action("0", str(i)+".csv", str(i))

    if input("mettre fin insérer 1 sinon autre ==> ") == "1":
        application = False
        break