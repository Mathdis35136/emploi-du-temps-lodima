from flask import Flask, render_template, request, send_file
import os
from werkzeug.utils import secure_filename
from Programme import generer_excel

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        fichiers = {}
        for nom in ['file1', 'file2', 'file3']:
            fichier = request.files[nom]
            filename = secure_filename(fichier.filename)
            path = os.path.join(UPLOAD_FOLDER, filename)
            fichier.save(path)

            nom_fichier = filename.lower()
            if "mois" in nom_fichier:
                fichiers["mois"] = path
            elif "prof" in nom_fichier:
                fichiers["prof"] = path
            elif "heure" in nom_fichier:
                fichiers["heures"] = path

        if not all(k in fichiers for k in ["mois", "prof", "heures"]):
            return "❌ Erreur : un des fichiers est manquant ou mal nommé (doit contenir 'mois', 'prof', 'heure')."

        output_path = os.path.abspath(os.path.join(OUTPUT_FOLDER, 'resultat.xlsx'))
        generer_excel(fichiers["heures"], fichiers["prof"], fichiers["mois"], output_path)

        if not os.path.exists(output_path):
            print("❌ Fichier non trouvé :", output_path)
            return "❌ Erreur lors de la génération du fichier."

        return send_file(output_path, as_attachment=True)

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
