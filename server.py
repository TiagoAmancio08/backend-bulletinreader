from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from werkzeug.utils import secure_filename
import os
from processamento import gerar_pdf

app = Flask(__name__)
CORS(app)

BASE_DIR = os.path.dirname(__file__)
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route("/processar", methods=["POST"])
def processar():
    try:
        if "pdf" not in request.files:
            return jsonify({"erro": "Nenhum arquivo enviado"}), 400

        pdf_file = request.files["pdf"]
        bimestres = request.form.get("bimestres", None)

        safe_name = secure_filename(pdf_file.filename or "arquivo.pdf")
        caminho_entrada = os.path.join(UPLOAD_FOLDER, safe_name)
        pdf_file.save(caminho_entrada)
        print("Arquivo salvo em:", caminho_entrada)

        caminho_saida = gerar_pdf(caminho_entrada, bimestres)
        print("Arquivo processado em:", caminho_saida)

        # Retorna URL completa para facilitar download direto
        nome_saida = os.path.basename(caminho_saida)
        url_download = f"http://192.168.2.7:5000/uploads/{nome_saida}"


        return jsonify({"url": url_download}), 200

    except Exception as e:
        print("Erro no servidor:", e)
        return jsonify({"erro": str(e)}), 500

@app.route("/uploads/<path:filename>")
def serve_upload(filename):
    return send_from_directory(UPLOAD_FOLDER, filename, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
