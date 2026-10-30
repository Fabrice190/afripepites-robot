"""
Calcule l'« empreinte visuelle » de chaque pub (image ou vignette de video)
et la range dans la base Supabase, pour la recherche par image.

A lancer APRES la collecte (les liens d'images Facebook expirent en ~5 jours).

Variables d'environnement : SUPABASE_URL, SUPABASE_KEY, CODE_ROBOT
Options : --lot 100 (images traitees par paquet)   --max 1500 (total par lancement)
"""
import argparse
import base64
import io
import json
import os
import sys
import time
import urllib.request
import urllib.error

import numpy as np
from PIL import Image
from fastembed import ImageEmbedding

MODELE = "Qdrant/clip-ViT-B-32-vision"   # 512 dimensions (CLIP ViT-B/32)
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"


def rpc(nom, corps):
    url = os.environ["SUPABASE_URL"].rstrip("/") + "/rest/v1/rpc/" + nom
    cle = os.environ["SUPABASE_KEY"]
    req = urllib.request.Request(
        url, data=json.dumps(corps).encode("utf-8"), method="POST",
        headers={"apikey": cle, "Authorization": "Bearer " + cle,
                 "Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=120) as r:
        return json.loads(r.read().decode("utf-8") or "null")


def telecharger(url):
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=20) as r:
        donnees = r.read()
    img = Image.open(io.BytesIO(donnees))
    img.load()
    if img.mode != "RGB":
        img = img.convert("RGB")
    return img


def miniature(img):
    m = img.copy()
    m.thumbnail((200, 200))
    tampon = io.BytesIO()
    m.save(tampon, format="JPEG", quality=55, optimize=True)
    return "data:image/jpeg;base64," + base64.b64encode(tampon.getvalue()).decode("ascii")


def vecteur_texte(v):
    v = np.asarray(v, dtype="float64")
    n = np.linalg.norm(v)
    if n > 0:
        v = v / n
    return "[" + ",".join(f"{x:.6f}" for x in v) + "]"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--lot", type=int, default=100)
    ap.add_argument("--max", type=int, default=1500)
    args = ap.parse_args()

    for nom in ("SUPABASE_URL", "SUPABASE_KEY", "CODE_ROBOT"):
        if not os.environ.get(nom):
            sys.exit(f"Variable manquante : {nom}")
    code = os.environ["CODE_ROBOT"]

    print("Chargement du modele d'images...", flush=True)
    modele = ImageEmbedding(model_name=MODELE, cache_dir="./modele_cache")

    traites = reussis = echecs = 0
    debut = time.time()
    while traites < args.max:
        a_faire = rpc("pubs_a_indexer", {"code": code, "n": min(args.lot, args.max - traites)})
        if not a_faire:
            print("Plus rien a indexer.")
            break
        images, cles, ko = [], [], []
        for ligne in a_faire:
            try:
                images.append(telecharger(ligne["media_url"]))
                cles.append(ligne["cle"])
            except Exception as e:            # lien expire, image absente...
                ko.append({"cle": ligne["cle"], "ko": True})
        envoi = list(ko)
        if images:
            try:
                vecs = list(modele.embed(images))
                for cle, img, v in zip(cles, images, vecs):
                    envoi.append({"cle": cle, "empreinte": vecteur_texte(v),
                                  "miniature": miniature(img), "ko": False})
                reussis += len(images)
            except Exception as e:
                print("Erreur de calcul sur un paquet :", e, flush=True)
                envoi += [{"cle": c, "ko": True} for c in cles]
        echecs += len(ko)
        traites += len(a_faire)
        for i in range(0, len(envoi), 40):
            rpc("enregistrer_empreintes", {"code": code, "lignes": envoi[i:i + 40]})
        print(f"{traites} traitees ({reussis} indexees, {echecs} liens inutilisables) "
              f"- {int(time.time() - debut)} s", flush=True)

    print(f"Termine : {reussis} pubs indexees, {echecs} images introuvables.")


if __name__ == "__main__":
    main()
