"""
Convertisseur de logs Fortigate vers format CSV créé par Hyperion.
"""


# Import Interface Graphique
import os.path
import re
import sys
from tkinter import Tk
from tkinter.filedialog import askopenfilename, askdirectory


def __load_file(filename: str) -> list:
    """
    Récupération du contenu du fichier de configuration
    :param filename : Chemin d'accès vers le fichier de configuration à convertir
    :return : Liste contenant toutes les lignes du fichier de configuration
    """
    file = open(filename, 'rb')
    brut = file.readlines()
    file.close()
    content = []
    for line in brut:
        content.append(line.decode('utf8'))
    return content


def check_file_already_exist(input_pathfile: str, output_path: str, prefix: str, overwrite: bool, extension: str) -> str:
    """
    Vérification si le fichier de sortie existe déjà et ouvre une fenêtre de demande d'écrasement.
    Renomme le fichier sinon.
    :param input_pathfile : Chemin d'accès du fichier d'entrée
    :param output_path : Chemin d'accès du dossier de sortie
    :param prefix : Permet d'ajouter un préfixe au fichier pour distinguer les différents types de log
    :param overwrite : Doit-on écraser automatiquement les fichiers déjà existants
    :param extension : Précise l'extension du fichier de sortie
    :return : Chemin d'accès du fichier de sortie à écrire
    """
    input_name = os.path.splitext(os.path.basename(input_pathfile))[0]
    output_filename = "{}/output_{}_{}.{}".format(
        output_path,
        prefix,
        input_name,
        extension
    )
    if os.path.isfile(output_filename):
        if not overwrite:
            answer = input("Voulez-vous écraser le fichier {} ? (y/n)".format(os.path.basename(output_filename)))
            if answer in ("y", "yes", "o", "oui"):
                print("Ecrasement du fichier {}".format(os.path.basename(output_filename)))
                return output_filename
        else:
            print("Ecrasement du fichier {}".format(os.path.basename(output_filename)))
            return output_filename
        j = 1
        output_filename = "{}/output_{}_{} ({}).{}".format(
            output_path,
            prefix,
            input_name,
            j,
            extension
        )

        while os.path.isfile(output_filename):
            print("Fichier {} déjà existant".format(output_filename))
            j += 1
            output_filename = "{}/output_{}_{} ({}).{}".format(
                output_path,
                prefix,
                input_name,
                j,
                extension
            )

    return output_filename


def __find_type(matches: list) -> str:
    for match in matches:
        if match[0] == "type":
            return match[1]
    return ""


def __find_sub_type(matches: list) -> str:
    for match in matches:
        if match[0] == "subtype":
            return match[1]
    return ""


def __init_sub_type(length: int) -> dict:
    data_init = {}
    for n in range(length - 2):
        data_init[str(n)] = []
    data_init['keys'] = ["{}".format(x) for x in range(length - 2)]
    return data_init


def __init_list_values(data_values: dict) -> list:
    max_length = 0
    for data_key in data_values.keys():
        if len(data_values[data_key]) > max_length:
            max_length = len(data_values[data_key])

    line_list = []
    for z in range(max_length):
        line_list.append("")

    return line_list


if __name__ == "__main__":
    # Récupération du fichier de config
    Tk().withdraw()
    input_file = askopenfilename()

    # Test du fichier avant ouverture
    if not os.path.isfile(input_file):
        print("Le fichier sélectionné est introuvable")
        sys.exit(2)

    # Emplacement de sortie
    output_dir = askdirectory(title="Emplacement de sortie")


    regex = r"\"([^=\"]+)=[\"]*([^=\"]+)[\"]*\""

    # Début de la conversion du fichier
    content = __load_file(input_file)
    data = {}

    print("Analyse du fichier en cours...")

    for line in content:
        items = line.split(",")
        data_tuples = []
        log_type = ""
        log_sub_type = ""

        for item in items:
            if item == "\"\"":
                data_tuples.append(("", ""))
            else:
                match = re.findall(regex, item)
                if len(match):
                    if match[0][0] == "type":
                        log_type = match[0][1]
                    elif match[0][0] == "subtype":
                        log_sub_type = match[0][1]
                    else:
                        data_tuples.append(match[0])


        if log_type not in data.keys():
            data[log_type] = {}
        if log_sub_type not in data[log_type].keys():
            data[log_type][log_sub_type] = __init_sub_type(len(items))

        for i in range(len(data_tuples)):
            if data_tuples[i][0]:
                if data_tuples[i][0] not in data[log_type][log_sub_type].keys():
                    data[log_type][log_sub_type][data_tuples[i][0]] = data[log_type][log_sub_type][str(i)]
                    del data[log_type][log_sub_type][str(i)]
                    data[log_type][log_sub_type]["keys"][i] = data_tuples[i][0]
                data[log_type][log_sub_type][data_tuples[i][0]].append(data_tuples[i][1])
            else:
                data[log_type][log_sub_type][data[log_type][log_sub_type]["keys"][i]].append("")

    print("Création des fichiers CSV en cours...")
    # Ecriture des fichiers Excel
    for prefix in data.keys():
        for sheet_name in data[prefix].keys():
            output_file = check_file_already_exist(input_file, output_dir, "{}_{}".format(prefix, sheet_name), False, "csv")
            print("Traitement de {} - {}".format(prefix, sheet_name))

            cols = []
            for key in data[prefix][sheet_name]["keys"]:
                try:
                    int(key)
                except ValueError:
                    cols.append(key)

            content = cols[0]
            list_lines = __init_list_values(data[prefix][sheet_name])

            # On écrit les entêtes
            for col in cols[1:]:
                content += ';' + col
            content = content[:-1]
            content += '\n'

            # On écrit les valeurs
            for col in cols:
                i = 0
                for value in data[prefix][sheet_name][col]:
                    list_lines[i] += ';' + value
                    i += 1

            for i in range(len(list_lines)):
                list_lines[i] = list_lines[i][1:]

            content += '\n'.join(list_lines)

            print("Ecriture de {} - {}".format(prefix, sheet_name))
            csv_file = open(output_file, 'w+')
            csv_file.write(content)
