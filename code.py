import tkinter as tk
import pandas as pd
import os
from tkinter import filedialog

# création de la fenêtre principale
root = tk.Tk()
root.title("Copier Coller")

# définir les dimensions de la fenêtre
root.geometry("500x400")

# ajouter un label pour le fichier source
label_source = tk.Label(root, text="Sélectionnez le fichier source Excel")
label_source.pack()

# ajouter un label pour afficher le nom du fichier source sélectionné
label_file_source = tk.Label(root, text="")
label_file_source.pack()

# ajouter un bouton pour sélectionner le fichier source
def select_file_source():
    file_path = filedialog.askopenfilename()
    if file_path:
        label_file_source.config(text=f"Fichier source sélectionné : {file_path}")
        global df_source
        df_source = pd.read_excel(file_path)
        df_source['Nom du fichier'] = os.path.basename(file_path)

button_source = tk.Button(root, text="Sélectionner un fichier source", command=select_file_source)
button_source.pack()

# ajouter un label pour le fichier de destination
label_destination = tk.Label(root, text="Sélectionnez le fichier de destination Excel")
label_destination.pack()

# ajouter un label pour afficher le nom du fichier de destination sélectionné
label_file_destination = tk.Label(root, text="")
label_file_destination.pack()

# ajouter un bouton pour sélectionner le fichier de destination
def select_file_destination():
    file_path = filedialog.askopenfilename()
    if file_path:
        label_file_destination.config(text=f"Fichier de destination sélectionné : {file_path}")
        global df_destination
        df_destination = pd.read_excel(file_path)
        df_destination['Nom du fichier'] = os.path.basename(file_path)

button_destination = tk.Button(root, text="Sélectionner un fichier de destination", command=select_file_destination)
button_destination.pack()

# ajouter un bouton pour copier les données
def copy_data():
    global df_source, df_destination
    df_destination = pd.concat([df_destination, df_source], axis=0)
    label_result.config(text="Copie des données effectuée avec succès !")

button_copy = tk.Button(root, text="Copier les données", command=copy_data)
button_copy.pack()

# ajouter un label pour afficher le résultat de la copie
label_result = tk.Label(root, text="")
label_result.pack()

# ajouter un bouton pour enregistrer le fichier final
def save_file():
    global df_destination
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
    if file_path:
        df_destination.to_excel(file_path, index=False)
        label_result.config(text="Fichier enregistré avec succès !")

button_save = tk.Button(root, text="Enregistrer le fichier final", command=save_file)
button_save.pack()
# démarrer la boucle principale de la fenêtre

root.mainloop()
