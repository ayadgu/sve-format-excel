# Importation de librairies
from logging import Formatter
import sys
sys.tracebacklimit = 0
import xlrd
import pandas as pd
from datetime import datetime
import numpy as np
from openpyxl import load_workbook
pd.set_option('display.max_colwidth', None)
import PySimpleGUI as sg
import logging

DEBUGMODE=True

logging.basicConfig(level=logging.DEBUG)
log=logging.getLogger('=>')    

# pd.set_option('display.width', 1000)
FILEOPENOPTIONS = dict(defaultextension=".xls", 
                       filetypes=[('xls file', '*.xls')])

class Formatter():
    def __init__(self):
        self.init_variable()
        pass    

    def init_variable(self):
        now = datetime.now()
        date_string = now.strftime("%d-%m-%Y-%H-%M-%S")
        self.date_string=date_string
        pass

    def open_xls(self):

        pass

    def format_compte_mandat(self,df,filename):
        col_to_drop=self.find_droppable_columns(df)
        try:
            col_to_drop=[]
            df_height,_=df.shape

            intitule=df[14].iloc[3]
            # print(intitule)

            for col in df.columns:
                # On élimine les colonnes qui ne servent à rien
                count = df[col].isna().sum()
                if (count/df_height)>0.95:
                    col_to_drop.append(col)
            
            # On drop les premières lignes qui ne servent à rien
            df_drop=df.drop(col_to_drop, axis=1)
            df_drop = df_drop.iloc[4: , :]

            # On créer un alias pour éviter de perdre la donnée déjà obtenue et travailler sur une copie
            df_alias=df_drop

            # On remplace le mot "Pièce" par "" (rien) pour éviter que le tableau soit trop chargé inutillement
            df_alias.replace(np.nan,"",regex=True,inplace=True)

            # On créer un alias pour éviter de perdre la donnée déjà obtenue et travailler sur une copie
            df_rename=df_alias

            # On renomme les colonnes pour avoir des libellés clairs
            df_rename = self.rename_columns(df_rename,{0: "Pièce",3: "Date",6: "C.A",7: "Compte",9: "Jal",11: "Intitulé",13: "",17: "N° Chèque",21: "Débit",25: "Crédit",27: "Solde"})

            # On exporte en un fichier excel propre
            self.generate_excel_file(df_rename,filename)

            # On affiche dans la fenêtre un message de succès
            self.display_success_message()

        except Exception as e:
            self.display_error_message()
            return

    def rename_columns(self,df,array):
        for key in array:
            df.rename(columns={key: array[key]},inplace=True)
        return df

    def format_journaux_provisoire(self,df,filename):
        col_to_drop=self.find_droppable_columns(df,custom_drop_array=[10,12,13])
        try:
            # Custom drop

            df_drop=df.drop(col_to_drop, axis=1)
            df_drop = df_drop.iloc[3: , :]

            # On créer un alias pour sauvegarder la donnée existante
            df_alias=df_drop

            # Dans la colonne 1 (Pièce), s'il y a une cellule vide, je la renseine à partir de la prochaine valeur non-vide situé en dessous d'elle.
            df_alias[1].ffill(inplace=True)

            # On supprime toutes les lignes dont la colonne 0 (Date) est nulle
            df_alias = df_alias.drop(df_alias[df_alias[0].isnull() & df_alias[14].isnull()].index)

            # On remplace le mot "Pièce" et NaN par "" (rien) pour éviter que le tableau soit trop chargé inutillement
            # Et autre cleaning...
            df_alias.replace("Pièce : ","",regex=True,inplace=True)
            df_alias.replace(np.nan,"",regex=True,inplace=True)
            df_alias[14].replace('Sous total pièce [0-9]+ :',"Sous-Total",regex=True,inplace=True)

            # Si la cellule situé à "Date" (0) est null, alors on annule sa colonne voisine "Pièce" (1)
            df_alias.loc[df[0].isnull(),[1]] = ""

            # On créer un autre alias pour sauvegarder la donnée existante
            df_rename=df_alias

            # On renomme les colonnes pour avoir des libellés clairs
            df_rename = self.rename_columns(df_rename,{0: "Date", 1: "Pièce", 4: "Compte", 6: "Ctr-Par", 9: "Jal", 11: "Libellé", 12: "Solde", 14: "", 19: "Débit", 22: "Crédit"})

            # On exporte en un fichier excel propre
            self.generate_excel_file(df_rename,filename)

            # On affiche dans la fenêtre un message de succès
            self.display_success_message()

        except Exception as e:
            self.display_erroor_message()
            return
            
    def display_success_message(self):
        self.window["OUTPUT"].Update(value="Fichier généré ! (Vérifiez tout de même si vous avez sélectionné le bon type de formatage)")

    def display_error_message(self):
        self.window["OUTPUT"].Update(value="Erreur, avez-vous choisi le bon type de fichier ?")
    
    def generate_excel_file(self,df,filename):
        df.to_excel(filename.split('.')[0]+'_clean_'+self.date_string+'.xlsx', engine='xlsxwriter',index=False)     
    
    def find_droppable_columns(self,df,custom_drop_array=[]):
        col_to_drop=[]
        col_to_drop.extend(custom_drop_array)
        df_height,_=df.shape
        for col in df.columns:
            count = df[col].isna().sum()
            if (count/df_height)>0.95:
                col_to_drop.append(col)
        return col_to_drop

    def init_dataframe(self,filename):
        wb = xlrd.open_workbook(filename, encoding_override='iso-8859-1')

        df = pd.read_excel(wb, engine="xlrd", header = None)
        return df

    def init_gui(self):
        sg.theme('DarkAmber')   # Add a touch of color
        # All the stuff inside your window.
        layout = [  
                    [sg.Text("Choisissez votre fichier: "), sg.FileBrowse("Ouvrir un fichier",file_types=(("Fichier Excel 97-2003", "*.xls"),))],
                    [sg.Text('Quel type de fichier ?')],
                    # [sg.Text("Choisissez la nature du fichier: ")],
                    [sg.Radio('Journaux Provisoire', "RADIO1", default=True)],
                    [sg.Radio('Compte Mandat', "RADIO1", default=False)],
                    [sg.Button('Formatter'), sg.Button('Quitter')],
                    [sg.Text('', size=(0, 1),key='OUTPUT')],
                    ]
        # Create the Window
        self.window = sg.Window('Formatteur Excel', layout)

    def apply(self):
        self.init_gui()
        
        # Event Loop to process "events" and get the "values" of the inputs
        while True:
            event, values = self.window.read()
            # print(values)

            if event == sg.WIN_CLOSED or event == 'Quitter': # if user closes self.window or clicks Quitter
                break

            if event=="Formatter" and (values[0] or values[1]) and values['Ouvrir un fichier']:
                self.init_variable()
                df=""
                filename = values["Ouvrir un fichier"]
                df=self.init_dataframe(filename)
                self.window["OUTPUT"].Update(value="Formattage en cours...")
                
                # Journaux provisoire
                if values[0]:
                    # print("Journaux provisoire")
                    self.format_journaux_provisoire(df,filename)

                # Compte Mandat
                if values[1]:
                    # print("Compte Mandat")
                    self.format_compte_mandat(df,filename)


            else:
                self.window["OUTPUT"].Update(value="Saisie incomplète.")

            

        self.window.close()

def main():
    # print("Bonjour... :)")
    df_formatter=Formatter()
    df_formatter.apply()

    


    # df_formatter.open_xls()
    pass

if __name__ == "__main__":
    main()
