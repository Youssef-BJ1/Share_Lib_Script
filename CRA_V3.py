import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time
import os

def initialize_driver():
    dr = webdriver.Chrome()
    return dr

def execute_invitation_process(email, identifiant, password, library_links):
    try:
        driver = initialize_driver()
        wait = WebDriverWait(driver, 5)
        for idx, library_link in enumerate(library_links):
            if idx == 0:
                connect_to_library(driver, email, identifiant, password, library_link)
            else:
                navigate_to_library(driver, library_link)
            perform_invitation_process(driver, wait)
        driver.quit()
    except Exception as e:
        messagebox.showerror("Erreur", f"Une erreur s'est produite : {str(e)}")

def connect_to_library(driver, email, identifiant, password, library_link):
    driver.get(library_link)
    wait = WebDriverWait(driver, 5)
    email_input = wait.until(EC.presence_of_element_located((By.NAME, "username")))
    email_input.send_keys(email)
    continue_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".spectrum-Button.spectrum-Button--cta.SpinnerButton.SpinnerButton--right")))
    continue_button.click()
    wait.until(EC.presence_of_element_located((By.NAME, "pf.username")))
    identifiant_input = driver.find_element(By.NAME, "pf.username")
    identifiant_input.send_keys(identifiant)
    wait.until(EC.presence_of_element_located((By.NAME, "pf.pass")))
    password_input = driver.find_element(By.NAME, "pf.pass")
    password_input.send_keys(password)
    soumettre_button = wait.until(EC.element_to_be_clickable((By.ID, "signOnButton")))
    soumettre_button.click()
    time.sleep(8)

def navigate_to_library(driver, library_link):
    driver.get(library_link)

def perform_invitation_process(driver, wait):
    try:
        share_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[aria-label="Partager"]')))
        share_button.click()
    except Exception as e:
        print(f"Erreur lors de la recherche du bouton : {e}")
    try:
        Invite_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="react-spectrum-26"]/li[1]')))
        Invite_button.click()
    except Exception as e:
        print(f"Erreur lors de la recherche du bouton : {e}")
    email_list = get_emails_from_excel()
    Invite_input = wait.until(EC.presence_of_element_located((By.ID, 'ccx-ss-flex-input-textarea')))
    for idx, email in enumerate(email_list):
        if idx != 0:
            Invite_input.send_keys(" ")
        Invite_input.send_keys(email)
    try:
        Invite_button2 = wait.until(EC.element_to_be_clickable((By.ID, 'ccx-ss-invite-send-btn')))
        Invite_button2.click()
    except Exception as e:
        print(f"Erreur lors de l'invitation de l'e-mail : {e}")
    try:
        Close_button = wait.until(EC.element_to_be_clickable((By.ID, 'ccx-ss-invite-close-btn')))
        Close_button.click()
    except Exception as e:
        print(f"Erreur lors du clic sur l'autre bouton : {e}")

def on_submit():
    email = email_entry.get()
    identifiant = identifiant_entry.get()
    password = password_entry.get()
    library_file_path = library_file_path_entry.get()
    library_links = get_library_links_from_excel(library_file_path)
    
    if library_links:
        execute_invitation_process(email, identifiant, password, library_links)

def get_emails_from_excel():
    try:
        file_path = file_path_entry.get()
        if file_path:
            df = pd.read_excel(file_path)
            if "Email" in df.columns:
                return df["Email"].tolist()
            else:
                messagebox.showerror("Erreur", "La colonne 'Email' n'a pas été trouvée dans le fichier Excel.")
    except Exception as e:
        messagebox.showerror("Erreur", f"Impossible de charger le fichier Excel : {str(e)}")
    return []

def get_library_links_from_excel(library_file_path):
    try:
        if library_file_path:
            df = pd.read_excel(library_file_path)
            if "LibraryLink" in df.columns:
                return df["LibraryLink"].tolist()
            else:
                messagebox.showerror("Erreur", "La colonne 'LibraryLink' n'a pas été trouvée dans le fichier Excel des liens de bibliothèque.")
    except Exception as e:
        messagebox.showerror("Erreur", f"Impossible de charger le fichier Excel des liens de bibliothèque : {str(e)}")
    return []

def browse_excel_file():
    file_path = filedialog.askopenfilename(filetypes=[("Fichiers Excel", "*.xlsx;*.xls")])
    file_path_entry.delete(0, tk.END)
    file_path_entry.insert(0, file_path)

def browse_library_file():
    library_file_path = filedialog.askopenfilename(filetypes=[("Fichiers Excel", "*.xlsx;*.xls")])
    library_file_path_entry.delete(0, tk.END)
    library_file_path_entry.insert(0, library_file_path)

root = tk.Tk()
root.title("Saisie d'informations")

style = ttk.Style()
style.configure('TButton', font=('calibri', 10, 'bold'), borderwidth='4')
label_style = ttk.Style()
label_style.configure('TLabel', font=('calibri', 12, 'bold'))
entry_style = ttk.Style()
entry_style.configure('TEntry', font=('calibri', 12, 'normal'))
frame_style = ttk.Style()
frame_style.configure('TFrame', background='#E5E8E8')

main_frame = ttk.Frame(root, style='TFrame')
main_frame.pack(padx=20, pady=20)

email_label = ttk.Label(main_frame, text="Email:", style='TLabel')
email_label.grid(row=0, column=0, padx=10, pady=10)
default_email = "@stellantis.com"
email_entry = ttk.Entry(main_frame, font=('calibri', 12, 'normal'), width=20)
email_entry.insert(0, default_email)
email_entry.grid(row=0, column=1, padx=10, pady=10)

identifiant_label = ttk.Label(main_frame, text="Identifiant:", style='TLabel')
identifiant_label.grid(row=1, column=0, padx=10, pady=10)
default_identifiant = ""
identifiant_entry = ttk.Entry(main_frame, font=('calibri', 12, 'normal'), width=20)
identifiant_entry.insert(0, default_identifiant)
identifiant_entry.grid(row=1, column=1, padx=10, pady=10)

password_label = ttk.Label(main_frame, text="Mot de passe:", style='TLabel')
password_label.grid(row=2, column=0, padx=10, pady=10)
password_entry = ttk.Entry(main_frame, font=('calibri', 12, 'normal'), width=20, show='*')
password_entry.grid(row=2, column=1, padx=10, pady=10)

file_path_label = ttk.Label(main_frame, text="Chemin du fichier Excel des emails:", style='TLabel')
file_path_label.grid(row=3, column=0, padx=10, pady=10)
file_path_entry = ttk.Entry(main_frame, font=('calibri', 12, 'normal'), width=20)
file_path_entry.grid(row=3, column=1, padx=10, pady=10)
browse_button = ttk.Button(main_frame, text="Parcourir", style='TButton', command=browse_excel_file)
browse_button.grid(row=3, column=2, padx=10, pady=10)

library_file_path_label = ttk.Label(main_frame, text="Chemin du fichier Excel des liens de bibliothèque:", style='TLabel')
library_file_path_label.grid(row=4, column=0, padx=10, pady=10)
library_file_path_entry = ttk.Entry(main_frame, font=('calibri', 12, 'normal'), width=20)
library_file_path_entry.grid(row=4, column=1, padx=10, pady=10)
browse_library_button = ttk.Button(main_frame, text="Parcourir", style='TButton', command=browse_library_file)
browse_library_button.grid(row=4, column=2, padx=10, pady=10)

submit_button = ttk.Button(main_frame, text="Soumettre", style='TButton', command=on_submit)
submit_button.grid(row=5, columnspan=2, pady=20)

root.mainloop()
