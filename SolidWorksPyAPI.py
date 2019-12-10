# -*- coding: utf-8 -*-
"""
Created on Mon Oct  7 17:22:55 2019

@author: henrique.ferreira
"""

import win32com.client
#import pythoncom

sw = win32com.client.Dispatch("SldWorks.Application.26")
model = sw.ActiveDoc

#Const
angle = 34/260*2.0*3.1416
distance = 0.2

junta1 = model.Parameter("D1@Junta1")
junta2 = model.Parameter("D1@Junta2")

def moveAngle(angle): #Method to change angle of a mate
    junta1.SystemValue = angle
    bol = model.EditRebuild3

def moveLin(distance): #Method to change dimension of a mate
    junta2.SystemValue = distance
    bol = model.EditRebuild3

def getmates(debug = 0): #It Returns a MateTable Array of Solidworks Mates and their properties
    MateTable = []
    c = model.FeatureManager.GetFeatures(True)[-1]
    if c.GetTypeName == 'MateGroup':
        print('Working.',end='')
        b = c.GetFirstSubFeature
        while b != None:
            a = b.GetSpecificFeature2 #IMate2
            ent1 = a.MateEntity(0)
            ent2 = a.MateEntity(1)
            ref1 = list(ent1.EntityParams)
            ref2 = list(ent2.EntityParams)
            if a.Type == 0:
                TypeS = 'Coincidente'
            elif a.Type == 1:
                TypeS = 'Concentrico'
            else:
                TypeS = 'Outro'
            if debug == 1:        
                print('Analisando feature: ',a.Name)
                print('Nome do mate: ',a.Name)
                print('Tipo do mate: ',TypeS)
                print('   A Entity 1 esta na peça: ',ent1.ReferenceComponent.Name)    
                print('     Com parametros: ', ref1,'do tipo',ent1.ReferenceType2)        
                print('   A Entity 2 esta na peça: ',ent2.ReferenceComponent.Name)        
                print('     Com parametros: ', ref2,'do tipo',ent2.ReferenceType2)
                print('\n')
            entry = [a.Name,TypeS,[ent1.ReferenceComponent.Name,ref1,ent1.ReferenceType2],[ent2.ReferenceComponent.Name,ref2,ent2.ReferenceType2]]
            MateTable.append(entry)
            print('.',end='')
            b = b.GetNextSubFeature
    print('Mates extraction complete.')
    return MateTable

def UpdateBodyPosition(): #In development
 swMUtil = sw.GetmathUtility
 return None


def createjointtable(MateTable):
    JointTable = []    
    for i in range(len(MateTable)):#Run over mates to see links pairs.
        JointFraction = [MateTable[i][2][0],MateTable[i][3][0],MateTable[i]]
        AddFraction = True #as default, this JointFracton will create a new 2-links entry.
        for j in range(len(JointTable)): #Check if theese specific 2 links already have a mate-group to join this JointFraction Mate
            if JointFraction[0]+JointFraction[1] == JointTable[j][0][0]+JointTable[j][0][1] or JointFraction[1]+JointFraction[0] == JointTable[j][0][0]+JointTable[j][0][1]:
                JointTable[j].append([JointFraction[2]])
                AddFraction = False
        if AddFraction == True: #Add this mate to mate-group for theese specific 2 links
            JointTable.append([[JointFraction[0],JointFraction[1]],JointFraction[2]])
    return JointTable

            

        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        