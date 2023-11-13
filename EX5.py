import numpy as np
import pandas as ps
import matplotlib.pyplot as plt
import scipy.stats as sc
import beepy as bp

def emettreSon():
    bp.beep(5)
    return 0

def tirage(number):
    X = np.random.rand(1,number)[0]
    count = 0
    for i in X:
        if i < 0.5:
            temp = 0
        else:
            temp = 1
        X[count] = int(temp)
        count += 1
    return X

def BuyAndHold(RowMatrix):
    Matrice = RowMatrix
    countUp = 0
    countDown = 0
    for i in Matrice:
        if i == 1:
            countUp += 1
        else:
            countDown += 1
    Profit = countUp - countDown
    return Profit    

def Momentum(RowMatrix):
    Matrice = RowMatrix
    countUp = 0
    countDown = 0
    
    for i in range(1,len(Matrice)):
        if Matrice[i-1] == Matrice[i]:
            countUp += 1
        else:
            countDown += 1
    Profit = countUp - countDown
    return Profit 

def Contrarian(RowMatrix):
    Matrice = RowMatrix
    countUp = 0
    countDown = 0
    
    for i in range(1,len(Matrice)):
        if Matrice[i-1] != Matrice[i]:
            countUp += 1
        else:
            countDown += 1
    Profit = countUp - countDown
    return Profit

def result():
    X1 = tirage(1024)
    X2 = tirage(1024)
    X3 = tirage(1024)
    resultat = [BuyAndHold(X1), Momentum(X2), Contrarian(X3)]
    return resultat

def ranking(Mat):
    Matrice = ps.DataFrame(Mat)
    rank = Matrice.rank(axis=0,method="average",ascending=False)
    Rang = [int(rank[0][i]) for i in range(3)]
    return Rang


def number_of_procedure(number):
    array = [ranking(result()) for i in range(0,number)]
    return array

def DisplayResult(matrice_of_matrice):
    timeFirst=[0,0,0]
    for i in matrice_of_matrice:
        count = 0
        for j in i:
            if j == 1:
                timeFirst[count] += 1
                break
            count += 1
            
    return [timeFirst[i]/len(matrice_of_matrice) for i in range(0,3)]

def afficher_graphique():
    plt.bar(["Buy and Hold", "Momentum", "Contrarian"],(DisplayResult(number_of_procedure(60))))
    plt.show()

def independancetest():
    index = [0,0,0]
    while True:
        mat = DisplayResult(number_of_procedure(60))
        if (mat[0] < mat[1]) or (mat[0] < mat[2]):
            index[0] +=1
        else:
            break

    while True:
        mat = DisplayResult(number_of_procedure(60))
        if (mat[1] < mat[0]) or (mat[1] < mat[2]):
            index[1] +=1
        else:
            break

    while True:
        mat = DisplayResult(number_of_procedure(60))
        if (mat[2] < mat[1]) or (mat[2] < mat[0]):
            index[2] +=1
        else:
            break
    return index

def test_stat():
    matrice = [independancetest() for i in range(0,50)]
    value = [sc.chi2_contingency(matrice)[:3]]
    return (value)

def run_test():
    for i in range(0,10):
        print(test_stat())
    return 0


def moyenne_plrs_tirages(N_procedures, N_tests):
    BHo = []
    Mom = []
    Con = []
    for i in range(0,N_tests):
        result = DisplayResult(number_of_procedure(N_procedures))
        BHo.append(result[0])
        Mom.append(result[1])
        Con.append(result[2])

    print(np.mean(BHo))
    print(np.mean(Mom))
    print(np.mean(Con))

moyenne_plrs_tirages(60, 1000)

emettreSon()