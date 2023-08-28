from genericpath import exists
import serial
import pandas as pd
import datetime
import numpy as np
import sqlite3
import smtplib, ssl
import copy
import time



def connexion_BD():
    try:
        path = r'C:\Projet SGA\DB\BD.db'
        return sqlite3.connect(path)
    except:
        print("connexion impossible!")
        return None


def recupererID():
    try:
        lecture=ser.readline()
        if len(lecture) != 0 :
            ID=str(lecture).split(",")[1][0]
            Salle = str(lecture).split(",")[0][2:]
            #print("Hatim ", lecture)
            return(ID,Salle)
        else:
            return None
    except:
        print ("ARDUINO CONNECTION ERROR")

    


def Rechercher_Existe_Etudiant(ID):
    global connexion
    global Niveau
    requete = "select Niveau FROM Etudiants Where ID_Etu='" + ID + "'"
    #print(requete)
    Niveau = connexion.execute(requete);
    Niveau= Niveau.fetchall()
    #print("Niveau =",Niveau)
    if (len(Niveau) == 0):
        return None
    else:
        Niveau=Niveau[0][0]
        return Niveau

def Rechercher_Seance(date,Niveau):
    global connexion,considerer_minute_present
    path = r'C:\Projet SGA\Planning\P'
    
    
        
    if (Niveau == "JMA1"):
        path = path + 'JMA1.xlsx'
    elif (Niveau == "JMA2"):
        path = path + 'JMA2.xlsx'
    elif (Niveau == "JMA3"):
        path = path + 'JMA3.xlsx'
    else:

        print("NiveAau")
        return None
    Planning = pd.read_excel (path)
    annee = date.year
    mois = date.month
    jour = date.day
    heure = date.hour
    minute = date.minute
    jour =str(annee)+"/"+"0"+str(mois)+"/"+"0"+ str(jour)
    try:
        seances = Planning[jour]
    except:
        print(Planning[jour])

        return None
    heurs1 = heure - np.array([10 , 20 , 18 , 14])
    heurs2 = heure - np.array([9 , 11 , 17 , 15])
    indice1 = np.where(heurs1 == 0)[0]
    indice2 = np.where(heurs2 == 0)[0]
    mi=30
    if considerer_minute_present==False:
        mi=60
    if (len(indice1) != 0 and minute >= 45 ):
        seance = seances[indice1[0]]
        seance = seance.split('\n')
        
        if ( seance[0].split(':')[1] == ""):
            return None
        return [heure + 1,seance[0].split(":")[1], seance[1].split(":")[1], seance[2].split(":")[1] , 0, "Present", jour]
    elif ( len(indice2) != 0 and minute <= mi):
        seance = seances[indice2[0]]
        seance = seance.split('\n')
        considerer_minute_present=True
        if ( seance[0].split(':')[1] == ""):
            return None
        return [heure, seance[0].split(":")[1], seance[1].split(":")[1], seance[2].split(":")[1] , minute, "Present", jour]
    else:
        print("CELUI LA")
        return None
   



def Recuperer_Cours_Retard(date,salle,niveau):
    path = r'C:\Projet SGA\Planning\P'
    if (niveau == "JMA1"):
        path = path + 'JMA1.xlsx'
    elif (niveau == "JMA2"):
        path = path + 'JMA2.xlsx'
    elif (niveau == "JMA3"):
        path = path + 'JMA3.xlsx'
    else:
        
        return None
    Planning = pd.read_excel (path)
    annee = date.year
    mois = date.month
    jour = date.day
    heure = date.hour
    minute = date.minute
    jour =str(annee)+"/"+"0"+str(mois)+"/"+"0"+ str(jour)
    try:
        seances = Planning[jour]
    except:
        print("ejr")
        return None
    heurs1 = heure - np.array([10 , 18 , 16 , 14])
    heurs2 = heure - np.array([9 , 11 , 17 ,15])
    indice1 = np.where(heurs1 == 0)[0]
    indice2 = np.where(heurs2 == 0)[0]
    if (len(indice1) != 0 and minute >= 45 ):
        seance = seances[indice1[0]]
        seance = seance.split('\n')
        if ( seance[0].split(':')[1] == ""):
            print(seance)
            return None
        if (seance[2].split(":")[1] != salle):

            return None
        return [heure + 1,seance[0].split(":")[1], seance[1].split(":")[1], seance[2].split(":")[1] , 0, "Present", jour]
    elif ( len(indice2) != 0 and minute <= 55):
        seance = seances[indice2[0]]
        seance = seance.split('\n')
        if ( seance[0].split(':')[1] == ""):
            return None
        if (seance[2].split(":")[1] != salle):
            return None
        return [heure, seance[0].split(":")[1], seance[1].split(":")[1], seance[2].split(":")[1] , minute, "Present", jour]
    else:
        print("eazer")
        return None





def Recherche_Nom_Prenom_Email(ID):
    global connexion
    if (Rechercher_Existe_Etudiant(ID) == None):
        print("pp")
        return None
    
    requete = "select Nom, Prénom, Mail FROM Etudiants Where ID_Etu='" + ID + "'"
    print("p")
    Info = connexion.execute(requete);
    Info= Info.fetchall()
    if (len(Info) == 0):
        return None
    else:
        return Info[0]
#Partie envoi de mail


def Envoyer_Mail(ID,message):
    (Nom,Prenom,receiver_email)=Recherche_Nom_Prenom_Email(ID)
    port = 587 # # For SSL
    smtp_server = "smtp.office365.com"
    sender_email = "simpleattempt001234@hotmail.com" #landrysanou11@gmail.com" # Enter your address
    requete = "select Mail FROM Etudiants Where ID_Etu='" + ID + "'"
    mail = connexion.execute(requete);
    mail=mail.fetchall()
    receiver_email = mail[0][0] # Enter receiver address
    password = '123456789123456789MH' #GENIE=MC^2PAI'

    context = ssl.create_default_context()
    print("RB")
    with smtplib.SMTP(smtp_server, port) as server:
        print("BR")
        server.set_debuglevel(1)
        server.starttls()
        server.login(sender_email, password)
        message = chr(10) + message
        server.sendmail(sender_email,receiver_email, message.encode('utf8'))
        print("BRRR")
        print(message)
        server.quit
    return "ok"


def Generer_Mail_Absence(Informations):
    global Niveau
    Heure = Informations[0]
    Matiere = Informations[1]
    Prof = Informations[2]
    Salle = Informations[3]
    Retard = Informations[4]
    Etat_Presence= "Absent_Non_Justifie"
    Date = Informations[6]
    print("AV",Niveau)
    Niveau = Informations[8]
    ID_Etu = Informations[7].split("/")[0][:7]
    NPE = Recherche_Nom_Prenom_Email(ID_Etu)
    print("ql",Niveau)
    if (NPE == None):
        print("ppp")
        return None
    (Nom , Prenom, Mail) = NPE


    return ("Bonjour " + Nom + " " + Prenom + "\n Merci de justifier l'absencce de la séance de " + Matiere + " du " + Date + " à " + str(Heure) + "h 00min.\n Bien cordialement")



def Enregistrement_Present_Retard(ID_Etu,salle,maintenant):
   
    Niveau = Rechercher_Existe_Etudiant(ID_Etu)
    if ( Niveau == None ):

        return None

    #maintenant = datetime.datetime.now()
    print (maintenant)
    info = Recuperer_Cours_Retard(maintenant,salle , Niveau)

    if info == None:
         return "informations vide"
    if connexion == None:
        return "connexion impossible"
    info.append(ID_Etu)
    info.append(Niveau)
    Heure = info[0]
    Matiere = info[1]
    Prof = info[2]
    Salle = info[3]
    Retard = info[4]
    Etat_Presence= info[5]
    Date = info[6]
    ID_Etu = info[7]
    Niveau = info[8]
    ID = info[7] + Date + str(Heure)
    print(ID)
    requete = "INSERT INTO Tableaura ( ID_RA,Matière, Prof, Retard, Date, Etat_Abscence, ID_Etrangère, Heure, Niveau )  VALUES  (' "
    requete += ID + "','"
    requete += Matiere + "','"
    requete += Prof + "','"
    requete += str(Retard) + "','"
    requete += Date + "','"
    requete += Etat_Presence + "','"
    requete += ID_Etu + "','"
    requete += str(Heure) + "','"
    requete += Niveau + "' )"
    print(requete)
    try:
        connexion.execute(requete);
    except:
        return "Enregistrement déja fait!"
    connexion.commit()
    return "Enregistrement de" + ID + "est effectué avec succes!"




def Enregistrement_Absences(date,Niveau):
    global connexion
    #print(Rechercher_Seance(date , Niveau, classe))
    
    Informations1 = Rechercher_Seance(date ,Niveau)
    print(Niveau)
    if (Informations1 == None ):
        print("3ZER3")
        return None
    Nombre_Etudiants = 0
    print("je")
    if( Niveau == "JMA1" ):
        liste=pd.read_excel(r'C:\Projet SGA\Listes\JMA1.xlsx')
        Nombre_Etudiants = len(liste)
    elif ( Niveau == "JMA2" ):
          liste=pd.read_excel(r'C:\Projet SGA\Listes\JMA2.xlsx')
          Nombre_Etudiants = len(liste)
    elif( Niveau == "JMA3" ):  
        liste=pd.read_excel(r'C:\Projet SGA\Listes\JMA3.xlsx')
        Nombre_Etudiants = len(liste)
   
    print("yue")
    annee = date.year
    mois = date.month
    jour = date.day
    heure = date.hour
    jour =str(annee)+"/"+"0"+str(mois)+"/"+"0"+ str(jour)
    jour = jour + str(heure)
    print("haz")
    print(Nombre_Etudiants)
    for i in range(0, Nombre_Etudiants ):
        print("hzz")
        Id_etu =str(liste["ID"][i])
        ID = Id_etu + jour
        Informations = copy.deepcopy(Informations1)
        Informations.append(ID)
        print("il,",Niveau)
        Informations.append(Niveau)
        Heure = Informations[0]
        Matiere = Informations[1]
        Prof = Informations[2]
        
        Salle = Informations[3]
        Retard = Informations[4]
        Etat_Presence= "Absent_Non_Justifie"
        print("hzje")
        Date = Informations[6]
        Niveau = Informations[8]
        requete = "INSERT INTO Tableaura ( ID_RA,Matière, Prof, Retard, Date, Etat_Abscence, ID_Etrangère, Heure, Niveau )  VALUES  (' "
        requete += ID + "','"
        requete += Matiere + "','"
        requete += Prof + "','"
        requete += str(Retard) + "','"
        requete += Date + "','"
        requete += Etat_Presence + "','"
        requete += Id_etu + "','"
        requete += str(Heure) + "','"
        print(Informations)
        requete += Niveau + "' )"
        print(requete)

        try:
            
            print(connexion.execute(requete))
            print("Enregistrement de " + ID + " comme absent est effectué avec succes!")
            #print("Generer_Mail_Absence(Informations)")
            message=Generer_Mail_Absence(Informations) #Envoi de mail
            print(message)
            #print(type(message)) #Envoi de mail
            Envoyer_Mail(Id_etu,message) #Envoi de mail
            #print(Informations)
        except:
            print("exi")
        connexion.commit()
    #time.sleep(60)


connexion = connexion_BD()

ser = serial.Serial("COM3", 9600,timeout=2)




while 1:
    Niveau = None
    
    
    a=recupererID()
    maintenant = datetime.datetime.now()
    if a != None :
        ID , Salle = a
        considerer_minute_present=True
        Niveau = Rechercher_Existe_Etudiant(ID)
        print(ID,Salle,Niveau)
        print(Enregistrement_Present_Retard(ID,Salle,maintenant))
    #print("a")
    H=[9,11,18,15]
    for i in H:

        if Niveau==None and maintenant.hour==i and maintenant.minute==7:
            Liste_Niveau=["JMA1", "JMA2", "JMA3"]
            for i in Liste_Niveau:
                considerer_minute_present=False
                Enregistrement_Absences( maintenant,i)

    
