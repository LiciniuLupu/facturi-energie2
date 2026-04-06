"""
Analizator Facturi Energie + Estimare Fotovoltaica
Instalare si rulare: vezi README.txt
"""

import os
import base64
import json
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import anthropic

app = Flask(__name__, static_folder=".")
CORS(app)

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")

@app.route("/")
def index():
    return send_from_directory(".", "index.html")

@app.route("/analyze", methods=["POST"])
def analyze():
    if not ANTHROPIC_API_KEY:
        return jsonify({"error": "ANTHROPIC_API_KEY nu este setat pe server."}), 500

    year = request.form.get("year", "2024")
    files = request.files.getlist("files")

    if not files:
        return jsonify({"error": "Nu au fost trimise fisiere."}), 400

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    content_parts = [
        {
            "type": "text",
            "text": f"""Esti expert in analiza facturi energie electrica din Romania.
Analizeaza TOATE documentele atasate (facturi PDF, poze, scanari).
Pentru fiecare factura extrage:
1. Luna si anul PERIOADEI DE CONSUM (nu data emiterii facturii)
2. Consumul in kWh (cauta: consum activ, energie activa, kWh consumati, diferenta indecsi)

Daca o factura acopera mai multe luni, listeaza fiecare luna separat.
Concentreaza-te pe facturile din {year}.

Raspunde EXCLUSIV cu JSON valid, fara markdown, fara text suplimentar, fara backticks:
{{"invoices":[{{"luna":1,"an":{year},"kwh":150}},{{"luna":2,"an":{year},"kwh":220}}]}}"""
        }
    ]

    for f in files:
        raw = f.read()
        b64 = base64.standard_b64encode(raw).decode("utf-8")
        mime = f.content_type or "application/octet-stream"

        if mime.startswith("image/"):
            content_parts.append({
                "type": "image",
                "source": {"type": "base64", "media_type": mime, "data": b64}
            })
        elif mime == "application/pdf" or f.filename.lower().endswith(".pdf"):
            content_parts.append({
                "type": "document",
                "source": {"type": "base64", "media_type": "application/pdf", "data": b64}
            })
        else:
            try:
                text = raw.decode("utf-8", errors="replace")
                content_parts.append({
                    "type": "text",
                    "text": f"\n--- {f.filename} ---\n{text}"
                })
            except Exception:
                pass

    try:
        message = client.messages.create(
            model="claude-opus-4-5",
            max_tokens=1024,
            messages=[{"role": "user", "content": content_parts}]
        )
        raw_text = "".join(b.text for b in message.content if hasattr(b, "text"))
        raw_text = raw_text.strip().replace("```json", "").replace("```", "").strip()
        parsed = json.loads(raw_text)
        return jsonify(parsed)
    except json.JSONDecodeError as e:
        return jsonify({"error": f"Raspuns invalid de la AI: {str(e)}", "raw": raw_text}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"Server pornit pe http://localhost:{port}")
    app.run(host="0.0.0.0", port=port, debug=False)
