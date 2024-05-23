import requests
import zipfile
import os
import shutil


def get_latest_release_info():
    url = "https://api.github.com/repos/hydraangel/erfpedro/releases/latest"
    response = requests.get(url)
    response.raise_for_status()  # Check if the request was successful
    return response.json()


def check_for_updates():
    latest_release = get_latest_release_info()
    latest_version = latest_release["tag_name"]
    update_url = latest_release["assets"][0]["browser_download_url"]
    with open("versaoatual.txt", 'r') as ver:
        client_version = ver.read()
    print(client_version, latest_version)
    if latest_version != client_version:
        return latest_version, update_url
    return None, None


def download_update(update_url, download_path="ERF.zip"):
    response = requests.get(update_url, stream=True)
    with open(download_path, "wb") as file:
        for chunk in response.iter_content(chunk_size=8192):
            file.write(chunk)
    return download_path


