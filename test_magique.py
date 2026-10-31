"""
Test de qualite pour la "Recherche magique" (a ne pas garder).
Calcule l'empreinte de quelques phrases (en francais brut ET traduites en anglais)
avec le modele texte CLIP compatible avec les empreintes d'images deja en base.
Resultat : resultats_test_magique.json
"""
import json
import numpy as np
from fastembed import TextEmbedding

REQUETES = [
    ("vergetures",  "crème contre les vergetures",            "cream for stretch marks"),
    ("barbe",       "produit pour faire pousser la barbe",    "beard growth oil"),
    ("dos",         "ceinture contre le mal de dos",          "back pain relief belt"),
    ("dentifrice",  "dentifrice blanchissant",                "whitening toothpaste"),
    ("minceur",     "patch minceur pour maigrir",             "slimming patch weight loss"),
    ("ecouteurs",   "écouteurs sans fil",                     "wireless earbuds"),
    ("chaussures",  "chaussures pour femme",                  "women's shoes"),
    ("virilite",    "café pour la virilité des hommes",       "coffee for men vitality"),
    ("hemorroides", "produit contre les hémorroïdes",         "hemorrhoid treatment"),
    ("ventilateur", "ventilateur rechargeable",               "rechargeable fan"),
]

modele = TextEmbedding(model_name="Qdrant/clip-ViT-B-32-text", cache_dir="./modele_cache")


def vecteur(texte):
    v = np.asarray(list(modele.embed([texte]))[0], dtype="float64")
    n = np.linalg.norm(v)
    if n > 0:
        v = v / n
    return "[" + ",".join(f"{x:.4f}" for x in v) + "]"


sortie = []
for cle, fr, en in REQUETES:
    sortie.append({"cle": cle, "fr": fr, "en": en, "vec_en": vecteur(en), "vec_fr": vecteur(fr)})
    print("ok", cle, flush=True)

with open("resultats_test_magique.json", "w", encoding="utf-8") as f:
    json.dump(sortie, f, ensure_ascii=False)
print("Termine.")
