import os
import re
import sys
import argparse

def sanitize_windows_filename(name):
    """Remplace tous les caractères interdits Windows par ' '."""
    # Nettoie uniquement le nom de base, garde l’extension
    if '.' in name:
        basename, ext = os.path.splitext(name)
        clean = re.sub(r'[ <>:"/\\|?*\x00-\x1F]', '', basename)
        return clean.strip().strip('.') + ext
    else:
        return re.sub(r'[ <>:"/\\|?*\x00-\x1F]', '', name).strip().strip('.')

def rename_images_in_folder(folder):
    files = os.listdir(folder)
    renamed_count = 0
    for filename in files:
        old_path = os.path.join(folder, filename)
        if not os.path.isfile(old_path):
            continue

        new_name = sanitize_windows_filename(filename)
        new_path = os.path.join(folder, new_name)

        if filename == new_name:
            continue

        # Gestion de collision de nom
        count = 1
        final_path = new_path
        while os.path.exists(final_path):
            basename, ext = os.path.splitext(new_name)
            final_path = os.path.join(folder, f"{basename}_{count}{ext}")
            count += 1

        print(f"Renommage : {filename} -> {os.path.basename(final_path)}")
        os.rename(old_path, final_path)
        renamed_count += 1
    print(f"\nTotal fichiers renommés : {renamed_count}")

def main():
    parser = argparse.ArgumentParser(
        description="Nettoie les noms de fichiers images pour compatibilité Windows."
    )
    parser.add_argument(
        "folder",
        nargs="?",
        default=r"D:\commun\PICS\10026",
        help="Dossier à traiter (défaut : D:\\commun\\PICS\\10026)"
    )
    args = parser.parse_args()

    folder = args.folder
    if not os.path.isdir(folder):
        print(f"Erreur : le dossier '{folder}' n'existe pas !")
        sys.exit(1)

    rename_images_in_folder(folder)

if __name__ == "__main__":
    main()
