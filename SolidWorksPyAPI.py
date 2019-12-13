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

def GetMates(debug = 0): #It Returns a MateTable Array of Solidworks Mates and their properties
    MateTable = []
    ComObj = model.FeatureManager.GetFeatures(True)[-1]
    if ComObj.GetTypeName == 'MateGroup':
        print('Working.',end='')
        ComObj = ComObj.GetFirstSubFeature
        while ComObj != None:
            MateFeature = ComObj.GetSpecificFeature2 #IMate2
            ent1 = MateFeature.MateEntity(0)
            ent2 = MateFeature.MateEntity(1)
            ref1 = RoundVec(ent1.EntityParams)
            ref2 = RoundVec(ent2.EntityParams)
            
            if MateFeature.Type == 0:
                TypeS = 'Coincidente'
            elif MateFeature.Type == 1:
                TypeS = 'Concentrico'
            else:
                TypeS = 'Outro'
            if debug == 1:        
                print('Analisando feature: ',MateFeature.Name)
                print('Nome do mate: ',MateFeature.Name)
                print('Tipo do mate: ',TypeS)
                print('   A Entity 1 esta na peça: ',ent1.ReferenceComponent.Name)    
                print('     Com parametros: ', ref1,'do tipo',ent1.ReferenceType2)        
                print('   A Entity 2 esta na peça: ',ent2.ReferenceComponent.Name)        
                print('     Com parametros: ', ref2,'do tipo',ent2.ReferenceType2)
                print('\n')
            entry = [MateFeature.Name,TypeS,[ent1.ReferenceComponent.Name,ref1,ent1.ReferenceType2],[ent2.ReferenceComponent.Name,ref2,ent2.ReferenceType2]]
            MateTable.append(entry)
            print('.',end='')
            ComObj = ComObj.GetNextSubFeature
    print('Mates extraction complete.')
    return MateTable

def GetPartPos():
    PartPos = []
    Features = model.FeatureManager.GetFeatures(True)
    for i in Features:
        FeatType = i.GetTypeName2
        if FeatType == 'Reference':
            vector = i.GetSpecificFeature2
            vector = vector.Transform2.ArrayData
            vector = RoundVec(vector)
            #The first 9 elements define the 3x3 rotation matrix. The next 
            #3 elements define the translation component. The next element 
            #defines the scaling component. The last 3 elements are unused.
            PartPos.append([i.Name,vector])
    return PartPos

def RoundVec(vector):#Round a Vector
    if type(vector) == list or type(vector) == tuple:
        ToupleToList = vector
        vector = []
        for j in ToupleToList:
            j = round(j,6)
            vector.append(j)
    else:
        return None
    return vector

def UpdateBodyPosition(): #In development
 #swMUtil = sw.GetmathUtility
 return None

def CreateJointTable(MateTable):
    JointTable = []    
    JointNumber = 1
    for i in range(len(MateTable)):#Run over mates to see links pairs.
        JointFraction = [MateTable[i][2][0],MateTable[i][3][0],MateTable[i]]
        AddFraction = True #as default, this JointFracton will create a new 2-links entry.
        for j in range(len(JointTable)): #Check if theese specific 2 links already have a mate-group to join this JointFraction Mate
            if JointFraction[0]+JointFraction[1] == JointTable[j][1][0]+JointTable[j][1][1] or JointFraction[1]+JointFraction[0] == JointTable[j][1][0]+JointTable[j][1][1]:
                JointTable[j].append([JointFraction[2]])
                AddFraction = False
        if AddFraction == True: #Add this mate to mate-group for theese specific 2 links
            JointTable.append([['Junta_'+str(JointNumber)],[JointFraction[0],JointFraction[1]],JointFraction[2]])
            JointNumber += 1
    return JointTable

def Run():  
    global MateTable
    global JointTable
    global PartPos
    MateTable = GetMates()
    JointTable = CreateJointTable(MateTable)
    PartPos = GetPartPos()
    return None



        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        