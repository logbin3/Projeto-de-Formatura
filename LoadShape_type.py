# -*- coding: utf-8 -*-
"""
Created on Sun Apr 18 15:54:35 2021

@author: Fabio
"""
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
from sklearn.neighbors import KNeighborsClassifier
from sklearn.model_selection import train_test_split
from sklearn import metrics
import joblib

def train_model(loadshapes_array, classification, loadshapes_names, folder_analysis, folder_program):
    f = open(folder_analysis + "\KNeighbors_loadshape_classifier_report.txt", "w")
    f.write( "Codefication: 1: residential ; 2: comercial/industrial; 3: Street lighting\n")
    print("loadshapes_array.shape = %s\n"%str(loadshapes_array.shape)) 
    headers=['h=0','h=1', 'h=2', 'h=3', 'h=4', 'h=5', 'h=6', 'h=7', 'h=8', 'h=9', 'h=10', 'h=11', 'h=12', 'h=13', 'h=14', 'h=15', 'h=16', 'h=17', 'h=18', 'h=19', 'h=20', 'h=21', 'h=22', 'h=23', 'Type']
    dataset=np.zeros((loadshapes_array.shape[0], loadshapes_array.shape[1] +1))
    dataset[:,0:loadshapes_array.shape[1]]=loadshapes_array
    #print("loadshapes_array.shape\n = %s"%str(dataset.shape))   
    dataset[:,loadshapes_array.shape[1]]=classification  
    f.write("Dataset:")
    f.write("\nNumber of rows: %d ; Number of columns (hours or minuts): %d"%(loadshapes_array.shape[1], loadshapes_array.shape[0]))      
    df=pd.DataFrame(dataset, index=loadshapes_names, columns=headers)  
    #convert the Pandas data frame to a Numpy array
    x = df[headers[0:(len(headers)-1)]].astype(float) 
    y = df['Type'].astype(int)
    
    # ## Normalize Data
    #x = preprocessing.StandardScaler().fit(x).transform(x.astype(float))
    
    ### Train Test Split
    x_train, x_test, y_train, y_test = train_test_split(x, y, test_size=0.2, random_state=2)
    f.write("\nTraining set size: %d; Testing set size: %d"%(x_train.shape[0], x_test.shape[0]))
        
    #Train Model   
    k = 2
    neigh = KNeighborsClassifier(n_neighbors = k).fit(x_train,y_train)

    #### Finding best K
    f.write("\ntesting the accuracy of the model with different values of k")
    Ks = 22
    mean_acc = np.zeros((Ks-1))
    std_acc = np.zeros((Ks-1))
    for n in range(1,Ks):    
        #Train Model and Predict  
        neigh = KNeighborsClassifier(n_neighbors = n).fit(x_train,y_train)
        yhat=neigh.predict(x_test)
        mean_acc[n-1] = metrics.accuracy_score(y_test, yhat)    
        std_acc[n-1]=np.std(yhat==y_test)/np.sqrt(yhat.shape[0])    
    f.write("\nThe best accuracy was %.1f%% with k= %d" %(mean_acc.max()*100, mean_acc.argmax()+1))
    f.close()
    #### Plot  model accuracy  for Different number of Neighbors
    plt.figure()
    plt.plot(range(1,Ks),mean_acc,'g')
    plt.fill_between(range(1,Ks),mean_acc - 1 * std_acc,mean_acc + 1 * std_acc, alpha=0.10)
    plt.fill_between(range(1,Ks),mean_acc - 3 * std_acc,mean_acc + 3 * std_acc, alpha=0.10,color="green")
    plt.legend(('Accuracy ', '+/- 1xstd','+/- 3xstd'))
    plt.ylabel('Accuracy ')
    plt.xlabel('Number of Neighbors (K)')
    plt.tight_layout()
    plt.savefig(folder_analysis + '/KNeighbors_loadshape_classifier.png')
    plt.show()
  
    #training the model with the best k
    k=mean_acc.argmax()+1
    neigh = KNeighborsClassifier(n_neighbors = k).fit(x_train,y_train)
    
    #saving the model
    with open (folder_program + "\\KNeighbors_loadshape_classifier", 'wb') as f:
        joblib.dump(neigh, f)

def predict_loadshape_Type(loadShape_array_list, folder_program): 
    with open (folder_program + "\\KNeighbors_loadshape_classifier", 'rb') as f:
        neigh=joblib.load(f)
    classification= neigh.predict(loadShape_array_list)
    return classification
    
    
if __name__ == "__main__":
    
    loadshapes_array=np.array([[1,1,1,1,1, 0,0,0,0,0, 0,0,0,0,0], [0,0,0,0,0, 0,0,0,0,0, 1,1,1,1,1],[1,1,1,1,1, 0,0,0,0,0, 0,0,0,0,0], [0,0,0,0,0, 1,1,1,1,1, 0,0,0,0,0], [1,1,1,1,1, 0,0,0,0,0, 0,0,0,0,0], [0,0,0,0,0, 1,1,1,1,1, 0,0,0,0,0], [1,1,1,1,1 ,0,0,0,0,0, 0,0,0,0,0], [0,0,0,0,0, 0,0,0,0,0, 1,1,1,1,1], [0,0,0,0,0, 1,1,1,1,1, 0,0,0,0,0], [0,0,0,0,0, 0,0,0,0,0, 1,1,1,1,1]])
    print("loadshapes_array.shape",loadshapes_array.shape)
    loadshape_names=["loadshape_1", "loadshape_2","loadshape_3","loadshape_4","loadshape_5","loadshape_6","loadshape_7","loadshape_8","loadshape_9","loadshape_10"]
    headers=['h=0','h=1', 'h=2', 'h=3', 'h=4', 'h=5', 'h=6', 'h=7', 'h=8', 'h=9', 'h=10', 'h=11', 'h=12', 'h=13', 'h=14', 'Type']
    classification=np.array([0,2,0,1,0,1,0,2,1,2])
    dataset=np.zeros((loadshapes_array.shape[0], loadshapes_array.shape[1] +1))
    dataset[:,0:loadshapes_array.shape[1]]=loadshapes_array
    print("dataset.shape",dataset.shape)
    dataset[:,loadshapes_array.shape[1]]=classification
    print("dataset.shape",dataset.shape)    
    
    train_model(loadshapes_array, classification, loadshape_names)

    
    