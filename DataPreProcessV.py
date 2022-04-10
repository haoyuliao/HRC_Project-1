#!/usr/bin/env python
# coding: utf-8

# In[40]:


import pandas as pd
import numpy as np

def dataPreProcess(filesName, ColumnsName, sheetName, leadtime=True , t = 3):
    for fn in range(len(filesName)):
        #print(filesName[fn])
        #Load the dataset.
        Marks = pd.read_excel(filesName[fn], sheet_name='Markers')
        sN = Marks['Frame'].iloc[len(Marks['Frame'])-10]+1 ####
        lsN = Marks['Frame'].iloc[len(Marks['Frame'])-2]+1 ####
        #print(len(Marks['Frame']))
        for i in range(len(sheetName)):
            Data = pd.read_excel(filesName[fn], sheet_name=sheetName[i])
            if i == 0:
                df = Data[ColumnsName[i]][sN:lsN]
            else:
                df = np.hstack((df, Data[ColumnsName[i]][sN:lsN]))
        #print(Marks['Frame'])
        #print(lsN)
        lbN = 1 #Name of label
        for i in range(len(Marks['Frame'])-9, len(Marks['Frame'])-1): ###
            if i == (len(Marks['Frame'])-9): ###
                n = Marks['Frame'][i]- Marks['Frame'][i-1]
                labels = np.zeros((n,1), dtype=int)
                labels[-45:] = lbN ## For 9 labels
                #print(n)
                continue
                
            if i == (len(Marks['Frame'])-5) or i == (len(Marks['Frame'])-3): ### 5 oo 7
                continue
            if i == (len(Marks['Frame'])-6) or i == (len(Marks['Frame'])-4): ### 4 or 6
                n = Marks['Frame'][i+1] - Marks['Frame'][i-1]
            else:
                n = Marks['Frame'][i] - Marks['Frame'][i-1]
            #print(n)
            addNewLabels = np.zeros((n,1), dtype=int)+(lbN+1)
            addNewLabels[:45] = lbN 
            lbN += 2
            if i != (len(Marks['Frame'])-2):
                addNewLabels[-45:] = lbN ## For 9 labels
            labels = np.vstack((labels, addNewLabels))  
            
        if fn == 0:
            if leadtime:
                a = np.array(df)
                #t = 3
                for i in range(t):
                    if i == 0:
                        b = np.hstack((a[t:], a[(t-1):-1]))# np.hstack((a[3:], a[2:-1])) t = 3 #t-1
                    else:
                        b = np.hstack((b, a[(t-1-i):(-1-i)])) #np.hstack((b, a[1:-2])) #t-2 
                allLabels = labels[t:]
                allData = b 
            else:
                allLabels = labels
                allData = np.array(df)                       
        else:
            if leadtime:
                # lead time = 3
                a = np.array(df)
                #t = 3
                for i in range(t):
                    if i == 0:
                        b = np.hstack((a[t:], a[(t-1):-1]))# np.hstack((a[3:], a[2:-1])) t = 3 #t-1
                    else:
                        b = np.hstack((b, a[(t-1-i):(-1-i)])) #np.hstack((b, a[1:-2])) #t-2           
                allLabels = np.vstack((allLabels, labels[t:]))
                allData = np.vstack((allData, b))
            else:
                allLabels = np.vstack((allLabels, labels))
                allData = np.vstack((allData, np.array(df)))


    return allData, allLabels

'''
testFilesName = ['P001-005.xlsx']
ColumnsName = [['Right Elbow Ulnar Deviation/Radial Deviation', 'Right Elbow Pronation/Supination', 'Right Elbow Flexion/Extension',
                'Left Elbow Ulnar Deviation/Radial Deviation', 'Left Elbow Pronation/Supination', 'Left Elbow Flexion/Extension'],
             ['Right Hand x', 'Right Hand y','Right Hand z',
              'Left Hand x', 'Left Hand y', 'Left Hand z']]
sheetName = ['Joint Angles ZXY', 'Segment Position']
testAllData, testAllLabels = dataPreProcess(testFilesName, labels, sheetName, leadtime=False)
trainData = np.hstack((testAllData, testAllLabels))
## convert your array into a dataframe
df = pd.DataFrame (trainData)

## save to xlsx file

filepath = 'Data.xlsx'

df.to_excel(filepath, index=False)

trainFilesName, validFilesName, testFilesName = [], [], []

for i in range(22):
    for j in range(3):
        fn = '../P00%s/P00%s-00%s.xlsx' %(i+1, i+1, j+5)
        trainFilesName.append(fn)
    validFilesName.append('../P00%s/P00%s-00%s.xlsx' %(i+1, i+1, j+6))
    testFilesName.append('../P00%s/P00%s-00%s.xlsx' %(i+1, i+1, j+7))
print(trainFilesName)
print(validFilesName)
print(testFilesName)
ColumnsName = [['Right Elbow Ulnar Deviation/Radial Deviation', 'Right Elbow Pronation/Supination', 'Right Elbow Flexion/Extension',
                'Left Elbow Ulnar Deviation/Radial Deviation', 'Left Elbow Pronation/Supination', 'Left Elbow Flexion/Extension'],
             ['Right Hand x', 'Right Hand y','Right Hand z',
              'Left Hand x', 'Left Hand y', 'Left Hand z']]

trainAllData, trainAllLabels = dataPreProcess(trainFilesName, ColumnsName, sheetName, leadtime=False, t=3)


testAllData, testAllLabels = dataPreProcessT(testFilesName, ColumnsName, sheetName, leadtime=True, t=3)
'''

