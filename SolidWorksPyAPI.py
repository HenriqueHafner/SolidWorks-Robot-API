# -*- coding: utf-8 -*-
"""
Created on Mon Oct  7 17:22:55 2019

@author: henrique.ferreira
"""

import win32com.client
import numpy as np
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
        mate_index = 1
        partposA,partposB = 'partposA','partposB'
        parent_index = 'parent_index'
        ComObj = ComObj.GetFirstSubFeature

        while ComObj != None:
            MateFeature = ComObj.GetSpecificFeature2 #IMate2
            ent1 = MateFeature.MateEntity(0) 
            ent2 = MateFeature.MateEntity(1)
            ref1 = RoundVec(ent1.EntityParams) #[pointX, pointY, pointZ, vectorI, vectorJ, vectorK, radius1, radius2]
            ref2 = RoundVec(ent2.EntityParams)           
            PosLimts = ['None','None']
            if MateFeature.Type == 0:
                TypeS = 'CoincidentPlane'
            elif MateFeature.Type == 1:
                TypeS = 'ConcentricCylinder'
            elif MateFeature.Type == 5:
                TypeS = 'LimitedSliding'
                PosLimts = [MateFeature.MinimumVariation,MateFeature.MaximumVariation]
            else:
                TypeS = 'Outro '+str(MateFeature.Type)
            if debug == 1:        
                print('Analisando feature: ',MateFeature.Name)
                print('Nome do mate: ',MateFeature.Name)
                print('Tipo do mate: ',TypeS)
                print('   A Entity 1 esta na peça: ',ent1.ReferenceComponent.Name)    
                print('     Com parametros: ', ref1,'do tipo',ent1.ReferenceType2)        
                print('   A Entity 2 esta na peça: ',ent2.ReferenceComponent.Name)        
                print('     Com parametros: ', ref2,'do tipo',ent2.ReferenceType2)
                print('\n')
            entry = [MateFeature.Name,TypeS,[ent1.ReferenceComponent.Name,ref1,ent1.ReferenceType2],[ent2.ReferenceComponent.Name,ref2,ent2.ReferenceType2],PosLimts]
            entry_b = [mate_index,ent2.ReferenceComponent.Name,ent2.ReferenceComponent.Name,partposA,partposB,TypeS,[ref1,ref2,PosLimts],parent_index]
            MateTable.append(entry_b)
            print('.',end='')
            mate_index += 1
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
            #The first 9 elements define 3 axis vectors. The next 
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
            JointTable.append([['Junta_'+str(JointNumber)],[JointFraction[0],JointFraction[1],'ChildPos'],JointFraction[2]])
            JointNumber += 1
    
    return JointTable

def SetParentTree(JointTable):
    ParentTable = []
    for i in range(len(JointTable)):
        ParentTable.append(JointTable[i][1])      
    
    return ParentTable
    

def VectorTransf(ChildBase,Vector):
    a = [] #tensor from ChildBase #[[x1,x2,x3],[y1,y2,y3],[z1,z2,z3]]
    b = [] #Vector in referential base
    VectorBaseChild = np.linlag.solve(a,b)
    VectorBaseChild = VectorBaseChild.round(6)
    return VectorBaseChild

def Run():  
    global MateTable
    global JointTable
    global PartPos

    MateTable = GetMates()
    PartPos = GetPartPos()
    BaseLinkName = PartPos[0][0]
    JointTable = CreateJointTable(MateTable)
    return None

#JointTable[n] -> Joint number 'n'
#JointTable[n][0] -> JointName

        
        
        
        















        
        
        
        
        
        
        


























        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        