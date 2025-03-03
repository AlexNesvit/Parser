import os
import requests
import json
import csv
from datetime import datetime
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment

# URL de la page
url = ""
headers = {"User-Agent": "Mozilla/5.0"}

# Chargement de la page
response = requests.get(url, headers=headers)
soup = BeautifulSoup(response.text, "html.parser")

# Fonction pour obtenir le texte d'un sélecteur
def get_text(selector):
    element = soup.select_one(selector)
    return element.text.strip() if element else "Non disponible"

def get_href(selector):
    element = soup.select_one(selector)
    return element['href'].strip() if element and 'href' in element.attrs else "Non disponible"

# ✅ **Function to get text from multiple selectors**
def get_text_multiple(selectors):
    """Пробует несколько селекторов и возвращает первый найденный текст."""
    for selector in selectors:
        element = soup.select_one(selector)
        if element:
            return element.get_text(strip=True)
    return "Non disponible" 

# 1️⃣ **Nom du promoteur**
nom_promoteur = get_text("h1.elementor-heading-title").replace(" ", "_")  # Change " " to "_"

# 2️⃣ **Site web**
site_web_element = soup.select_one("a.elementor-button.elementor-button-link.elementor-size-xs")
site_web = site_web_element["href"] if site_web_element else "Non disponible"

# 3️⃣ **Année de création**
annee_creation = get_text(".elementor-element-a230edc > div:nth-child(1)")

# 4️⃣ **Programmes sur Otaree**
programmes = get_text("div.elementor-element:nth-child(9) > div:nth-child(1)")

# 5️⃣ **Nombre de lots** 
nombre_lots = get_text(".badge_ul > li:nth-child(3)")


# 6️⃣ **À propos**
a_propos = "\n".join([
    p.text.strip() for p in soup.select("div.elementor-widget-container p")
    if "Site internet" not in p.text
    and "Réseaux Sociaux" not in p.text
    and "Portails Annonces" not in p.text
])

# 7️⃣ **Nos services**
services = "\n".join([li.text.strip() for li in soup.select("div.elementor-widget-container ul li")])

# 8️⃣ **Nos valeurs**
# valeurs = "\n".join([li.text.strip() for li in soup.select("div.elementor-widget-container ul li")])

# 9️⃣ **Interlocuteurs**
interlocuteurs = []

for team in soup.select("div.card-team"):
    full_name = team.select_one(".team_name").text.strip() if team.select_one(".team_name") else None
    fonction = team.select_one(".team_function").text.strip() if team.select_one(".team_function") else None
    email_attr = team.select_one(".team_button_mail")["data-tf-hidden"] if team.select_one(".team_button_mail") else ""
    email = email_attr.split("email_promoteur=")[1].split(",")[0] if "email_promoteur=" in email_attr else None

    if full_name and fonction and email:
        # Diviser le nom complet en nom et prénom
        name_parts = full_name.split(" ")
        prenom = name_parts[0]  
        nom = " ".join(name_parts[1:]) if len(name_parts) > 1 else ""  

        interlocuteurs.append({
            "Prénom": prenom,
            "Nom": nom,
            "Fonction": fonction,
            "Email": email
        })

#  **Parsing "Lots sur Otaree"**
lots = {}

for row in soup.select(".elementor-widget-container tr"):
    cols = row.find_all("td")  # Toutes les colonnes `td`
    if len(cols) >= 2:
        lot_name = cols[0].text.strip()  # Nom
        lot_count = cols[1].text.strip()  # Quantité
        if lot_count.isdigit():
            lots[lot_name] = int(lot_count)

# Store the number of lots
nombre_lots = len(lots)

# ✅ **Ajout des paramètres supplémentaires**
extra_params = {
    "Durée estimative d’une dénonciation": get_text_multiple([  
        "div.elementor-element:nth-child(11) > div:nth-child(1)", # Selecteur principal
        "div.elementor-element:nth-child(12) > div:nth-child(1)",  # Selecteur alternatif
    ]),
    # "Durée estimative d’une dénonciation": get_text("div.elementor-element:nth-child(12) > div:nth-child(1)"),
"Durée estimative d’une option": get_text_multiple([
    "div.elementor-element:nth-child(13) > div:nth-child(1)", # Selecteur principal
    "div.elementor-element:nth-child(14) > div:nth-child(1)",  # Selecteur alternatif
]),
    "Possibilité de prolonger une option": get_text_multiple([
    "div.elementor-element:nth-child(15) > div:nth-child(1)",  # Selecteur principal 
    "div.elementor-element:nth-child(16) > div:nth-child(1)",  # Selecteur alternatif
]),
    "Montant du dépôt de garantie": get_text(
        # "div.elementor-element:nth-child(18) > div:nth-child(1)",
        "div.elementor-element:nth-child(17) > div:nth-child(1)"
),
    # "Délai estimatif paiement honoraires": get_text("div.elementor-element:nth-child(24) > div:nth-child(1)")
    "Délai estimatif paiement honoraires": get_text_multiple([  
        "div.elementor-element:nth-child(24) > div:nth-child(1)",  # Selecteur principal
        "div.elementor-element:nth-child(23) > div:nth-child(1)",  # Selecteur alternatif
        ".elementor-element-xxxxx"  # Classe CSS supplémentaire
    ])
}



# ✅ **Parsing "Communication autorisée par le promoteur"**
communication_element = soup.select_one(".elementor-element-d1bed84 img")
communication_autorisee = "Site internet"  # Atribut par défaut

if communication_element:
    alt_text = communication_element.get("alt", "").strip().lower()
    if "check" in alt_text:
        communication_autorisee = "Autorisé"
    elif "cross" in alt_text:
        communication_autorisee = "Non autorisé"

# **Formation du dictionnaire**
data = {
    "Nom du promoteur": nom_promoteur,
    "Site web": site_web,
    "Année de création": annee_creation,
    "Programmes sur Otaree": programmes,
    "À propos": a_propos,
    "Nos services": services,
    # "Nos valeurs": valeurs,
    "Interlocuteurs": interlocuteurs,
    "Communication autorisée par le promoteur": communication_autorisee
}


# ✅ Ajout de "Lots sur Otaree" si ils existent
if lots:
    data["Lots sur Otaree"] = lots

# ✅ Ajout des paramètres supplémentaires
data.update({key: value for key, value in extra_params.items() if value})

# **Genérer le nom du fichier**
date_str = datetime.now().strftime("%d-%m-%Y")  # Date actuelle
version = 1

# Verifier si le fichier existe déjà
while os.path.exists(f"{nom_promoteur}_{date_str}_v{version}.json"):
    version += 1

file_name = f"{nom_promoteur}_{date_str}_v{version}"

# **JSON**
with open(f"{file_name}.json", "w", encoding="utf-8") as json_file:
    json.dump(data, json_file, ensure_ascii=False, indent=4)

# **CSV**
csv_columns = list(data.keys())

with open(f"{file_name}.csv", "w", encoding="utf-8-sig", newline="") as csv_file:
    writer = csv.DictWriter(csv_file, fieldnames=csv_columns, delimiter=";")  # <-- Здесь меняем разделитель
    writer.writeheader()
    writer.writerow(data)

#Excel
wb = Workbook()
ws = wb.active

# Titre
ws.append(csv_columns)

# Functions
def format_cell_value(value):
    """ Преобразуем вложенные структуры данных (списки, словари) в многострочный текст """
    if isinstance(value, list):
        return "\n".join(
            [f"{item['Prénom']} {item['Nom']}: {item['Fonction']} ({item['Email']})"
             if isinstance(item, dict) else str(item)
             for item in value]
        )
    elif isinstance(value, dict):
        return "\n".join([f"{key}: {val}" for key, val in value.items()])
    return str(value)

# **Fussion des `p` et `li**
def clean_text(elements):
    """ Объединяет все `p` и `li` в один текст, без пропуска начальных строк """
    return "\n".join(
        [el.get_text(strip=True, separator=" ") for el in elements]
    )

# **À propos, Nos services, Nos valeurs**
data["À propos"] = clean_text(soup.select("div.elementor-widget-container p"))
data["Nos services"] = clean_text(soup.select("div.elementor-widget-container ul li"))
data["Nos valeurs"] = clean_text(soup.select("div.elementor-widget-container ul li"))

# **Garanties**
def ensure_multiline(value):
    """ Linguee le texte si il est trop long """
    if isinstance(value, str) and len(value) > 100:  # Si le texte est trop long
        return value.replace(". ", ".\n")  # Ajoutez un saut de ligne apres chaque point
    return value

# **Structure des données**
row_values = [ensure_multiline(format_cell_value(data.get(col, ""))) for col in csv_columns]

# **Structure des données**
ws.append(row_values)

# **Largeur des colonnes**
for col in ws.columns:
    max_length = 0
    col_letter = col[0].column_letter  # Lettre de la colonne

    for cell in col:
        if cell.value:
            cell.alignment = Alignment(wrap_text=True)  
            max_length = max(max_length, len(str(cell.value)))

    # **Largeur de la colonne 60**
    ws.column_dimensions[col_letter].width = min(max_length // 2, 60)

# **Fixer la deuxième ligne**
ws.freeze_panes = "A2"

# **Automatic row height**
ws.row_dimensions[2].height = 150  # Hauteur de la deuxième ligne

# **Enregistrer le fichier Excel**
wb.save(f"{file_name}.xlsx")

print(f"✅ Données enregistrées sous {file_name}.json et {file_name}.csv et {file_name}.xlsx 🚀")