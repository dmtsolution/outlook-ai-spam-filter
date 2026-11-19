# Outlook AI Spam Filter

Classifieur de spam par apprentissage automatique, entraîné sur un corpus d'emails réels, exporté puis intégré directement dans le client Outlook de bureau pour une classification en temps réel, à l'arrivée de chaque message.

## Contexte et objectif

Les filtres anti-spam natifs des messageries (Outlook, Gmail) reposent sur des règles et des signaux propriétaires qui ne détectent pas systématiquement tous les messages indésirables. Lors de tests informels menés pendant ce projet, plusieurs emails construits pour ressembler à des tentatives de fraude (gains fictifs, urgence artificielle, appels à l'action suspects) n'ont pas été filtrés par Gmail, alors qu'ils ont été correctement identifiés comme spam par le modèle développé ici.

Ce constat illustre un point simple : un filtre supplémentaire, entraîné et contrôlé localement, peut combler des angles morts des solutions standards. C'est un argument pertinent pour des organisations qui souhaitent renforcer leur premier niveau de défense contre le phishing et les emails frauduleux, en complément (et non en remplacement) des protections déjà en place.

Ce dépôt documente l'ensemble du pipeline : entraînement du modèle, évaluation, export, puis intégration dans Outlook, afin que le travail soit reproductible et auditable.

## Architecture

```
Kaggle (entraînement)              Poste local (Windows)
┌─────────────────────┐            ┌──────────────────────────────────┐
│ Dataset Enron-Spam   │            │  spam_api.py (Flask, 127.0.0.1)   │
│ Nettoyage + TF-IDF   │  export    │  charge le modèle + vectoriseur   │
│ Entraînement +       │ ─────────► │  expose POST /predict             │
│ comparaison modèles  │            └───────────────┬────────────────────┘
│ Export .joblib       │                            │ HTTP (local)
└─────────────────────┘                             ▼
                                    ┌──────────────────────────────────┐
                                    │  Outlook (VBA, ThisOutlookSession)│
                                    │  Application_NewMailEx            │
                                    │  → appelle /predict à chaque mail │
                                    │  → tague "IA-Spam"                │
                                    │  → déplace vers Courrier indés.   │
                                    └──────────────────────────────────┘
```

Le modèle est entraîné et évalué une fois dans un notebook Kaggle, puis exporté. En production locale, aucune étape de ré-entraînement n'a lieu : le serveur se contente de charger les artefacts exportés et de servir des prédictions.

## Dataset et modèle

- **Dataset** : [SetFit/enron_spam](https://huggingface.co/datasets/SetFit/enron_spam), version nettoyée du corpus Enron-Spam (Metsis et al.), l'un des jeux de données les plus utilisés en classification de spam sur emails réels. Environ 30 000 emails après dédoublonnage, répartition quasi équilibrée (52 % ham / 48 % spam).
- **Prétraitement** : normalisation du texte (minuscules, URLs, adresses email et nombres remplacés par des tokens génériques, ponctuation retirée).
- **Représentation** : TF-IDF, unigrammes et bigrammes, vocabulaire plafonné à 30 000 termes.
- **Modèles comparés** (validation croisée à 5 plis, score F1) :

  | Modèle | F1 (validation croisée) |
  |---|---|
  | Multinomial Naive Bayes | 0.9847 |
  | Logistic Regression | 0.9898 |
  | **LinearSVC (retenu)** | **0.9902** |

- **Performance du modèle retenu sur le jeu de test** (6 093 emails jamais vus à l'entraînement) :

  | Métrique | Valeur |
  |---|---|
  | Accuracy | 0.9921 |
  | Precision | 0.9907 |
  | Recall | 0.9928 |
  | F1 | 0.9918 |
  | ROC-AUC | 0.9994 |

Le notebook complet (installation, import, statistiques descriptives, visualisations, prétraitement, entraînement, évaluation détaillée, analyse des erreurs, export, interface de test Gradio) se trouve dans [`notebook/spam_classifier_kaggle.ipynb`](notebook/spam_classifier_kaggle.ipynb).

## Structure du dépôt

```
notebook/
  spam_classifier_kaggle.ipynb   Notebook d'entraînement complet (à exécuter sur Kaggle)
model/
  spam_classifier.joblib         Modèle entraîné (LinearSVC calibré)
  tfidf_vectorizer.joblib        Vectoriseur TF-IDF associé
  metadata.json                  Métriques et métadonnées du modèle exporté
server/
  spam_api.py                    API locale Flask qui sert les prédictions
outlook/
  outlook_macro.vba              Macro VBA à intégrer dans Outlook
requirements.txt                 Dépendances Python du serveur local
```

## Installation et utilisation

### 1. Entraîner le modèle (optionnel — un modèle déjà entraîné est fourni dans `model/`)

1. Importer `notebook/spam_classifier_kaggle.ipynb` sur Kaggle.
2. Activer l'accès Internet dans les paramètres du notebook (le GPU n'est pas nécessaire : le modèle retenu est un modèle linéaire sur représentation TF-IDF).
3. Exécuter les cellules dans l'ordre. La dernière section lance une interface Gradio pour tester le modèle manuellement avant export.
4. Récupérer `spam_classifier.joblib`, `tfidf_vectorizer.joblib` et `metadata.json` depuis `/kaggle/working/outputs/model/`.

### 2. Lancer le serveur local

```bash
pip install -r requirements.txt
python server/spam_api.py
```

Le serveur écoute sur `http://127.0.0.1:5000` et expose :
- `GET /health` : vérification de disponibilité
- `POST /predict` : classification, attend un JSON `{"subject": "...", "body": "..."}`, retourne `{"label": "spam"|"ham", "proba_spam": float, "low_signal": bool}`

Le serveur doit rester actif pour que l'intégration Outlook fonctionne. Un lancement automatique à l'ouverture de session Windows peut être configuré en plaçant un raccourci vers `pythonw.exe server/spam_api.py` dans le dossier Démarrage (`shell:startup`).

### 3. Intégrer la macro dans Outlook

Prérequis : Outlook de bureau classique (l'intégration VBA n'est pas disponible dans le "nouveau Outlook").

1. Ouvrir Outlook, puis l'éditeur VBA (`Alt+F11`).
2. Dans l'arborescence du projet, ouvrir `ThisOutlookSession`.
3. Coller le contenu de [`outlook/outlook_macro.vba`](outlook/outlook_macro.vba).
4. Enregistrer (`Ctrl+S`).
5. Dans Outlook : Fichier → Options → Centre de gestion de la confidentialité → Paramètres du Centre de gestion de la confidentialité → Paramètres des macros → sélectionner "Notifications pour toutes les macros".
6. Redémarrer Outlook complètement (l'événement `Application_NewMailEx` ne s'active qu'au démarrage).

À chaque mail reçu, la macro interroge le serveur local, tague les messages classés spam avec la catégorie "IA-Spam", puis les déplace vers le dossier Courrier indésirable du compte concerné.

## Limites connues

- **Modèle entraîné exclusivement sur des emails en anglais.** Le dataset Enron-Spam ne contient aucun contenu en français ou dans d'autres langues. Sur du texte non anglais, le vecteur TF-IDF obtenu est quasi vide, et sans garde-fou le modèle se rabattrait sur son biais interne au lieu de produire une prédiction fondée. Un garde-fou a été ajouté côté serveur (`MIN_SIGNAL_TOKENS` dans `spam_api.py`) : en dessous d'un seuil de termes reconnus, le message est classé non-spam par défaut plutôt que de risquer un faux positif non fondé. Ce n'est qu'un filet de sécurité, pas une solution multilingue.
- **Pas de couverture multilingue.** Une version pour le français (et d'autres langues) nécessiterait un pipeline équivalent entraîné sur un corpus d'emails dans la langue cible, avec son propre vectoriseur et éventuellement ses propres modèles candidats.
- **Intégration locale, mono-poste.** Le serveur tourne en local sur la machine de l'utilisateur ; il ne s'agit pas d'une solution déployée au niveau d'un serveur de messagerie ou d'une passerelle d'entreprise.
- **Évaluation du cas Gmail non systématique.** L'observation selon laquelle certains emails test n'étaient pas filtrés par Gmail provient de tests manuels ponctuels, pas d'un benchmark comparatif rigoureux et reproductible.

## Pistes d'amélioration

- Pipeline équivalent pour le français et d'autres langues (corpus dédié, éventuellement un vectoriseur multilingue).
- Exploration d'une approche par transformer (par exemple un modèle de type BERT finement ajusté) pour comparer la robustesse face à des formulations plus subtiles ou adverses, en complément du modèle TF-IDF actuel.
- Déploiement du serveur local sous forme de service Windows plutôt qu'un simple raccourci de démarrage, pour une meilleure résilience (redémarrage automatique en cas d'échec).
- Tests automatisés et jeu d'évaluation continu pour suivre la dérive du modèle dans le temps.
- Intégration via un add-in Office.js pour prendre en charge le nouveau Outlook et Outlook sur le Web, qui ne supportent pas les macros VBA.

## Avertissement

Ce projet est un prototype à vocation éducative et de démonstration en sécurité. Il ne remplace pas les protections anti-spam et anti-phishing officielles d'un client de messagerie ou d'un fournisseur de messagerie d'entreprise, et ne doit pas être considéré comme une solution de sécurité de production sans audit, tests et durcissement complémentaires.

## Licence

Distribué sous licence MIT. Voir [LICENSE](LICENSE).
