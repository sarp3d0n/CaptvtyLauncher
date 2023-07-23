import requests
import os
import zipfile
import subprocess
import logging
import win32api
import win32con
import psutil

def is_captvty_running():
    # Vérifier si Captvty est déjà lancé
    for process in psutil.process_iter(['pid', 'name']):
        if process.info['name'].lower() == 'captvty.exe':
            return True
    return False

def terminate_captvty():
    # Fermer le processus Captvty s'il est en cours d'exécution
    for process in psutil.process_iter(['pid', 'name']):
        if process.info['name'].lower() == 'captvty.exe':
            process.terminate()

def get_latest_version_url():
    url = "http://captvty.fr/"
    response = requests.get(url)
    html = response.text

    # Trouver le lien de téléchargement à partir du code source de la page
    start_index = html.find('href="//releases.captvty.fr/') + len('href="')
    end_index = html.find('.zip"', start_index)
    if start_index != -1 and end_index != -1:
        latest_version_url = 'http:' + html[start_index:end_index] + '.zip'
        return latest_version_url
    else:
        return None

def download_and_extract(url):
    file_name = url.split("/")[-1]
    download_path = os.path.join(".", file_name)

    print("Téléchargement du fichier...")
    response = requests.get(url)
    with open(download_path, 'wb') as file:
        file.write(response.content)

    print("Extraction des fichiers...")
    with zipfile.ZipFile(download_path, "r") as zip_ref:
        # Extraire les fichiers directement dans le répertoire courant avec remplacement forcé
        zip_ref.extractall(".", pwd=None, members=None)

    print("Fichier téléchargé et extrait avec succès!")
    os.remove(download_path)

def launch_captvty():
    captvty_path = "Captvty.exe"  # Utilisation du chemin relatif (chemin courant)

    if os.path.exists(captvty_path):
        subprocess.Popen(captvty_path)
    else:
        print(f"Le programme Captvty.exe n'a pas été trouvé dans le répertoire : {captvty_path}")

def get_current_version():
    captvty_path = "Captvty.exe"  # Utilisation du chemin relatif (chemin courant)
    if os.path.exists(captvty_path):
        # Lire les propriétés du fichier Captvty.exe
        info = win32api.GetFileVersionInfo(captvty_path, "\\")
        ms = info['FileVersionMS']
        ls = info['FileVersionLS']

        # Extraire les niveaux de version
        major_version = win32api.HIWORD(ms)
        minor_version = win32api.LOWORD(ms)
        revision = win32api.HIWORD(ls)
        build = win32api.LOWORD(ls)

        # Formater la version dans un format lisible
        current_version = f"{major_version}.{minor_version}.{revision}.{build}"

        return current_version

    return None

def create_log(version_site, version_actuelle):
    log_file = "captvty_log.txt"
    logging.basicConfig(filename=log_file, level=logging.INFO,
                        format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')

    logging.info("---- Log Captvty ----")
    logging.info(f"Version du site : {version_site}")
    logging.info(f"Version actuelle : {version_actuelle}")
    logging.info("--------------------")

if __name__ == "__main__":
    if is_captvty_running():
        print("Captvty est déjà lancé. Fermeture du processus...")
        terminate_captvty()

    latest_version_url = get_latest_version_url()
    version_url = latest_version_url.split('/')[-1][:-4]
    latest_version = version_url.split("-")[1]
    
    if latest_version_url:
        # Vérifier ici si la version actuelle correspond à la dernière version
        current_version = get_current_version()

        # Exemple de vérification : si la version actuelle est différente de la dernière version
        if current_version != latest_version:
            download_and_extract(latest_version_url)

    create_log(latest_version, current_version)
    launch_captvty()
