import sqlite3
def connexion_BD():
    try:
        path = r'C:\Projet SGA\DB\BD.db'
        return sqlite3.connect(path)
    except:
        print("connexion impossible!")
        return None

connexion = connexion_BD()
requete="select ID_Etu FROM Etudiants where Niveau = 'JMA1'"
Niveau = connexion.execute(requete);
Niveau= Niveau.fetchall()
print(Niveau)