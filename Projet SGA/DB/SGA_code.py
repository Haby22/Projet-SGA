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
    global connexion
    path = r'C:\Projet SGA\Planning\P'
    
    if Niveau != None :
        
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
        jour = str(jour)+"/"+str(mois)+"/"+str(annee)
        try:
            seances = Planning[jour]
        except:

            return None
        heurs1 = heure - np.array([8 , 10 , 13 , 23])
        heurs2 = heure - np.array([9 , 11 , 14 , 0])
        indice1 = np.where(heurs1 == 0)[0]
        indice2 = np.where(heurs2 == 0)[0]
        if (len(indice1) != 0 and minute >= 45 ):
            seance = seances[indice1[0]]
            seance = seance.split('\n')
            if ( seance[0].split(':')[1] == ""):
                return None
            return [heure + 1,seance[0].split(":")[1], seance[1].split(":")[1], seance[2].split(":")[1] , 0, "Present", jour]
        elif ( len(indice2) != 0 and minute <= 30):
            seance = seances[indice2[0]]
            seance = seance.split('\n')
            if ( seance[0].split(':')[1] == ""):
                return None
            return [heure, seance[0].split(":")[1], seance[1].split(":")[1], seance[2].split(":")[1] , minute, "Present", jour]
        else:
            print("CELUI LA")
            return None

    else:
        
        #n1_str : niveau 1 en string pour avoir la main sur le niveau 
        n1_str='JMA1'
        n2_str='JMA2'
        n3_str='JMA3'
        Planning1 = pd.read_excel ( path + n1_str+'.xlsx')
        Planning2 = pd.read_excel ( path + n2_str+'.xlsx')
        Planning3 = pd.read_excel ( path + n3_str+'.xlsx')
        
        annee = date.year
        mois = date.month
        jour = date.day
        heure = date.hour
        minute = date.minute
        jour = str(jour)+"/"+str(mois)+"/"+str(annee)
        heurs1 = heure - np.array([9 , 11 , 14 , 0])
        indice1 = np.where(heurs1 == 0)[0]
        try:
            
            seances1 = Planning1[jour]
            seances2 = Planning2[jour]
            seances3 = Planning3[jour]
            print("a")
        except:
            #print(Planning2)
            return 
        seance1 = seances1[indice1[0]]
        seance1 = seance1.split('\n')
        
        seance2 = seances2[indice1[0]]
        seance2 = seance2.split('\n')
        #
        seance3 = seances3[indice1[0]]
        seance3 = seance3.split('\n')
        
        if seance1[0].split(":")[1] != "" :
            Matiere=seance1[0].split(":")[1]
            Prof=seance1[1].split(":")[1]
            Retard=0
            Heure=date.hour
            Etat_Presence="Absent_Non_Justifie"
            annee = date.year
            mois = date.month
            jour = date.day
    
            
            jour = str(jour)+"/"+str(mois)+"/"+str(annee)
            requete = f"select ID_Etu FROM Etudiants where Niveau = '{str(n1_str)}'"
            print(requete)
            id=connexion.execute(requete)
            
            id=id.fetchall()
            print(id)
            
            for i in id:
                try :
                    ID_Etrangère=i[0]
                    
                    ID_RA = ID_Etrangère + jour + str(Heure)
                    requete = "INSERT INTO Tableaura ( ID_RA,Matière, Prof, Retard, Date, Etat_Abscence, ID_Etrangère, Heure, Niveau )  VALUES  (' "
                    requete += ID_RA + "','"
                    requete += Matiere + "','"
                    requete += Prof + "','"
                    requete += str(Retard) + "','"
                    requete += jour + "','"
                    requete += Etat_Presence + "','"
                    requete += ID_Etrangère + "','"
                    requete += str(Heure) + "','"
                    print("azh")
                    requete += n1_str + "' )"
                    
                    print(requete)
                    connexion.execute(requete)
                    print("d")
                except:
                    return("Etudiant "+str(ID_RA)+"n'est pas abscent ")
                connexion.commit()
                
        if seance2[0].split(":")[1] != "" :
            Matiere=seance2[0].split(":")[1]
            Prof=seance2[1].split(":")[1]
            Retard=0
            Heure=date.hour
            Etat_Presence="Absent_Non_Justifie"
            annee = date.year
            mois = date.month
            jour = date.day
    
            
            jour = str(jour)+"/"+str(mois)+"/"+str(annee)
            requete = f"select ID_Etu FROM Etudiants where Niveau = '{str(n2_str)}'"
            id=connexion.execute(requete) 
            id=id.fetchall()
            
            for i in id:
                try :
                    ID_Etrangère=i[0]
                    ID_RA = ID_Etrangère + jour + str(Heure)
                    requete = "INSERT INTO Tableaura ( ID_RA,Matière, Prof, Retard, Date, Etat_Abscence, ID_Etrangère, Heure, Niveau )  VALUES  (' "
                    requete += ID_RA + "','"
                    requete += Matiere + "','"
                    requete += Prof + "','"
                    requete += str(Retard) + "','"
                    requete += jour + "','"
                    requete += Etat_Presence + "','"
                    requete += ID_Etrangère + "','"
                    requete += str(Heure) + "','"
                    requete += n2_str + "' )"
                    connexion.execute(requete)
                except:
                    return("Etudiant "+str(ID_RA)+"n'est pas abscent ")
                connexion.commit()
                
        if seance3[0].split(":")[1] !="" :
            Matiere=seance3[0].split(":")[1]
            Prof=seance3[1].split(":")[1]
            Retard=0
            Heure=date.hour
            Etat_Presence="Absent_Non_Justifie"
            annee = date.year
            mois = date.month
            jour = date.day
    
            
            jour = str(jour)+"/"+str(mois)+"/"+str(annee)
            requete = f"select ID_Etu FROM Etudiants where Niveau = '{str(n3_str)}'"
            id=connexion.execute(requete) 
            id=id.fetchall()
            
            for i in id:
                try :
                    ID_Etrangère=i[0]
                    ID_RA = ID_Etrangère + jour + str(Heure)
                    requete = "INSERT INTO Tableaura ( ID_RA,Matière, Prof, Retard, Date, Etat_Abscence, ID_Etrangère, Heure, Niveau )  VALUES  (' "
                    requete += ID_RA + "','"
                    requete += Matiere + "','"
                    requete += Prof + "','"
                    requete += str(Retard) + "','"
                    requete += jour + "','"
                    requete += Etat_Presence + "','"
                    requete += ID_Etrangère + "','"
                    requete += str(Heure) + "','"
                    requete += n3_str + "' )"
                    connexion.execute(requete)
                except:
                    return("Etudiant "+str(ID_RA)+"n'est pas abscent ")
                connexion.commit()
        

        
     



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
    jour = str(jour)+"/"+str(mois)+"/"+str(annee)
    try:
        seances = Planning[jour]
    except:

        return None
    heurs1 = heure - np.array([8 , 10 , 13 , 23])
    heurs2 = heure - np.array([9 , 11 , 14 ,0])
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
    elif ( len(indice2) != 0 and minute <= 30):
        seance = seances[indice2[0]]
        seance = seance.split('\n')
        if ( seance[0].split(':')[1] == ""):
            return None
        if (seance[2].split(":")[1] != salle):
            return None
        return [heure, seance[0].split(":")[1], seance[1].split(":")[1], seance[2].split(":")[1] , minute, "Present", jour]
    else:

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
    receiver_email = "yayehaby2000@gmail.com" # Enter receiver address
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



def Enregistrement_Present_Retard(ID_Etu,salle):
    Niveau = Rechercher_Existe_Etudiant(ID_Etu)
    if ( Niveau == None ):

        return None

    #maintenant = datetime.datetime.now()

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
    elif( Niveau == "JMA2" ):  
        liste=pd.read_excel(r'C:\Projet SGA\Listes\JMA3.xlsx')
        Nombre_Etudiants = len(liste)
    print("yue")
    annee = date.year
    mois = date.month
    jour = date.day
    heure = date.hour
    jour = str(jour)+"/"+str(mois)+"/"+str(annee)
    jour = jour + str(heure)
    print("haz")
    print(Nombre_Etudiants)
    for i in range(0, Nombre_Etudiants ):
        print("hzz")
        Id_etu =liste["ID"][i]
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
        H=[9,11,14,0]
        for i in H:

            if date.hour==i and date.minute==6:
                print("quf")
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
    time.sleep(60)


connexion = connexion_BD()

ser = serial.Serial("COM11", 9600,timeout=2)




while 1:
    Niveau = None
    
    
    a=recupererID()
    maintenant = datetime.datetime.now()
    if a != None :
        ID , Salle = a

        Niveau = Rechercher_Existe_Etudiant(ID)
        print(ID,Salle,Niveau)
        Enregistrement_Present_Retard(ID,Salle)
    #print("a")
    Enregistrement_Absences( maintenant,Niveau)

    
