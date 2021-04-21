# -*- coding: utf-8 -*-
import os
import glob
import pandas as pd
import math

result=[]

def open_files(way):

    all_data= pd.DataFrame()
    for x in os.walk(way):
        for y in glob.glob(os.path.join(x[0], '*.xlsx')):
            betreuer=get_Betreuer(y)
            einordnung=get_Einordnung(y)
            
            df=pd.read_excel(y)
            
            studg=get_Studiengaenge(df)
            sem=get_Semester(df)
            
            all_data=all_data.append(df, ignore_index=True)
            result.append([betreuer, einordnung, studg, sem])
    liste=[]
    liste =sorted(result,key=lambda x: (x[1],x[0]))
    return liste


def get_Betreuer(name):
    """ 'Name' """
    start=name.find("- ")+len("- ")
    end=name.find(".xlsx")
    betreuer=name[start:end]
    return betreuer
    
def get_Einordnung(name):
    """ 'Alte Geschichte' oder 'Mittelalterliche Geschichte' oder 'Neuere Geschichte'"""
    if "Alte Geschichte" in name or "AG" in name:
        einordnung="AG"
    elif "Mittelalterliche Geschichte" in name or "MA" in name:
        einordnung="MA"
    elif "Neuere Geschichte" in name or "NZ" in name:
        einordnung="NZ"
    return einordnung

def get_Studiengaenge(df):
    counter_L2=df.Studiengang.str.contains("L2", case=False).sum()
    counter_L3=df["Studiengang"].str.contains("L3", case=False).sum()
    counter_L5=df["Studiengang"].str.contains("L5", case=False).sum()
    counter_BA_HF=df["Studiengang"].str.contains("BA HF", case=False).sum()
    counter_BA_NF=df["Studiengang"].str.contains("BA NF", case=False).sum()
    counter_sonst=len(df.index) - counter_L2 - counter_L3 - counter_L5 - counter_BA_HF - counter_BA_NF
    counter_alle= counter_L2 + counter_L3 + counter_L5 + counter_BA_HF + counter_BA_NF + counter_sonst
    studiengaenge=[counter_alle, counter_L2, counter_L3, counter_L5, counter_BA_HF, counter_BA_NF, counter_sonst]
    return studiengaenge   

def get_Semester(df):
    liste=list(zip(df['Studiengang'],df['Fachsemester']))
    counter_L2=0
    counter_L3=0
    counter_L5=0
    counter_BA_HF=0
    counter_BA_NF=0
    counter_sonst=0
    for (i,j) in liste:
        if type(i)!=str:
            continue
        if "L2" in i and j <=7:
            counter_L2+=1
        elif "L3" in i and j<=7:
            counter_L3+=1
        elif "L5" in i and j<=7:
            counter_L5+=1
        elif "BA HF" in i and j<=7:
            counter_BA_HF+=1
        elif "BA NF" in i and j<=7:
            counter_BA_NF+=1
        elif "L2" not in i and "L3" not in i and "L5" not in i and "BA HF" not in i and "BA NF" not in i and j<=7:
            counter_sonst+=1
        else:
            continue
    counter_alle= counter_L2 + counter_L3 + counter_L5 + counter_BA_HF + counter_BA_NF + counter_sonst
    semester=[counter_alle, counter_L2, counter_L3, counter_L5, counter_BA_HF, counter_BA_NF, counter_sonst]
    return semester

def calc_einordnung(result):
    ag=["Gesamt:", "AG", [0,0,0,0,0,0,0], [0,0,0,0,0,0,0]]
    c_1=0
    ma=["Gesamt:", "MA" , [0,0,0,0,0,0,0], [0,0,0,0,0,0,0]]
    c_2=0
    nz=["Gesamt:", "NZ" , [0,0,0,0,0,0,0], [0,0,0,0,0,0,0]]
    c_3=0
    for i in range(len(result)):
        if result[i][1]=="AG":
            c_1+=1
            for j in range(7):
                ag[2][j]+=result[i][2][j]
                ag[3][j]+=result[i][3][j]
        if result[i][1]=="MA":
            c_2+=1
            for j in range(7):
                ma[2][j]+=result[i][2][j]
                ma[3][j]+=result[i][3][j]
        if result[i][1]=="NZ":
            c_3+=1
            for j in range(7):
                nz[2][j]+=result[i][2][j]
                nz[3][j]+=result[i][3][j]
    result.insert(c_1, ag)
    result.insert(c_1 + c_2 + 1, ma)
    result.insert(c_1+c_2+c_3 + 2, nz)
    return 
            
            
def make_Output(result):
    out=input("Bitte geben Sie das Speicherverzeichnis für die Ausgabe an:")
    df=pd.DataFrame(result, columns=["Kurs", "Einordnung", "# Studenten", "L2", "L3", "L5", "BA HF", "BA NF", "Sonstige", "# Semester unter 7", "L2", "L3", "L5", "BA HF", "BA NF", "Sonstige"])

    df.to_excel(out + "dataframe.xlsx", index=0)
                       
result=open_files(input("Bitte geben Sie den Pfad des Verzeichnisses der Exceldatein an: "))
calc_einordnung(result)
make_Output(result)


            
"""Liste=[Kursleiter, Einordnung (AG, MA, NG), [#Studenten, #L2, #L3, #L5,#BA HF, #BA NF, #sonst] #bis 7 Semester, #L2, #L3, #L5, #BA HF, #BA NF, #sonst]"""

