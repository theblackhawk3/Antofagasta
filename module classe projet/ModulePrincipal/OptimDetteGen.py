from random import *
import random
import numpy as np
#Données initiales de départ 
RNets = [10000,25000,30000,10000,50000,60000,30000,80000,40000]
Dette = 80000
TauxActu = 0.05
FP0 =  90000
AlphaTilde = 0
#Heuristiques de départ
def mutation(population):
    for i in range(0,len(population[0]),2):
        Indexmin = random.randint(0,len(population[0]))
        Indexmax = random.randint(Indexmin,len(population[0]))
        Rest1 = sum(population[i][0:Indexmin]+population[i][Indexmax:])
        Rest2 = sum(population[i+1][0:Indexmin]+population[i+1][Indexmax:])
        population[i][Indexmin:Indexmax],population[i+1][Indexmin:Indexmax] = population[i+1][Indexmin:Indexmax],population[i][Indexmin:Indexmax]
        for i in range(Indexmin):
            pass
        for i in range(Indexmax):
            pass
def solutionValide(solution,AlphaT):
    Cf = [0]*len(RNets)
    for i in range(len(Cf)):
        Cf[i] = RNets[i]-solution[i]
    if min(Cf) >= AlphaT:
        return True
    else:
        return False
        
def fitness(solution):
    Cfs = [0]*len(RNets)
    for i in range(len(Cfs)):
        Cfs[i] = RNets[i]-solution[i]
    return sum(Cfs)/((1+TauxActu)**(lni(solution)+1))
    
def lni(l):
    index = 0
    for i in range(len(l)):
        if l[i] != 0:
            index = i
    return index

def generatePopulation(Taille):
    Population = []
    n=0
    while n < Taille:
        a = 100
        solution = []
        for i in range(len(RNets)-1):
            rdm = random.randint(0,a)
            solution.append(rdm/100)
            a = a -rdm
        solution.append(a)
        solution = list(np.multiply(Dette,solution))
        if solutionValide(solution,AlphaTilde):
            Population.append(solution)
            n+=1
    return Population