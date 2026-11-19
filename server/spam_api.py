import os
import re

from flask import Flask, jsonify, request
import joblib

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

clf = joblib.load(os.path.join(BASE_DIR, "spam_classifier.joblib"))
vectorizer = joblib.load(os.path.join(BASE_DIR, "tfidf_vectorizer.joblib"))


def clean_text(text: str) -> str:
    text = text.lower()
    text = re.sub(r"http\S+|www\.\S+", " url ", text)
    text = re.sub(r"\S+@\S+", " email ", text)
    text = re.sub(r"\d+", " num ", text)
    text = re.sub(r"[^\w\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


app = Flask(__name__)

# En dessous de ce nombre de mots reconnus par le vocabulaire TF-IDF, le vecteur
# est trop vide pour que le modele ait un vrai signal : il se rabat alors sur son
# biais interne et penche vers "spam" par defaut. On considere ces cas comme
# "signal insuffisant" et on classe prudemment en non-spam plutot que de se fier
# a une prediction non fondee (ex: textes tres courts, ou dans une langue non vue
# a l'entrainement comme le francais).
MIN_SIGNAL_TOKENS = 4


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


@app.route("/predict", methods=["POST"])
def predict():
    data = request.get_json(force=True, silent=True) or {}
    subject = data.get("subject", "") or ""
    body = data.get("body", "") or ""
    text = f"{subject} {body}"

    cleaned = clean_text(text)
    X = vectorizer.transform([cleaned])

    if X.nnz < MIN_SIGNAL_TOKENS:
        return jsonify({"label": "ham", "proba_spam": 0.0, "low_signal": True})

    proba_spam = float(clf.predict_proba(X)[0, 1])
    label = "spam" if proba_spam >= 0.5 else "ham"

    return jsonify({"label": label, "proba_spam": proba_spam, "low_signal": False})


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000)
