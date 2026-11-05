"""
Robot collecteur 4 - Facebook Ad Library par NICHES (mots uniques)

Ce que fait ce script :
  1. Pour chaque PAYS et chaque MOT-CLE, ouvre la Facebook Ad Library
  2. Releve les pubs ACTIVES (annonceur, date de debut, texte, lien)
  3. Garde tout dans une base (base_pubs.csv) : les pubs deja vues ne sont pas dupliquees
  4. Fabrique un fichier Excel propre : pubs_veille.xlsx
       - feuille "Pubs"       : toutes les pubs, les plus anciennes en haut
       - feuille "Annonceurs" : les annonceurs qui insistent le plus

Lancement simple (pays et mots-cles par defaut) :
  python collecteur3.py

Lancement personnalise :
  python collecteur3.py --pays CI SN --mots "livraison gratuite" "promo" --defilements 4
"""

import argparse
import csv
import json
import os
import random
import re
import time
from datetime import date
from pathlib import Path
import urllib.request
import urllib.error
from urllib.parse import quote_plus

PAYS_DEFAUT = ["BJ"]          # Benin d'abord ; ajoute d'autres pays avec --pays
ZONES = {
    "afrique": ["BJ", "TG", "CI", "SN", "CM"],
    "occident": ["FR", "US"],
}
LANGUE_PAYS = {"FR": "fr", "BE": "fr", "CA": "fr",
               "ML": "fr", "BF": "fr", "NE": "fr", "GN": "fr", "MR": "fr",
               "GA": "fr", "CG": "fr", "CD": "fr", "TD": "fr", "CF": "fr",
               "US": "en", "GB": "en", "NG": "en", "GH": "en", "SL": "en", "LR": "en", "GM": "en"}

# Des MOTS UNIQUES (recherches larges), classes par niche.
NICHES_FR = {
    "Santé": ["douleur", "diabète", "tension", "prostate", "érection", "articulations",
              "arthrose", "hémorroïdes", "ulcère", "fatigue", "insomnie",
              "paludisme", "cholestérol", "digestion", "constipation", "minceur", "ventre",
              "gynécomastie",
              # expressions exactes (le signe = devant veut dire : phrase exacte)
              "=douleurs musculaires et articulaires", "=soulagez naturellement", "=plus de confort",
              "=formule 100% naturelle sans effet secondaire", "=solution naturelle"],
    "Beauté / soin": ["savon", "crème", "huile", "sérum", "vergetures", "boutons",
                      "taches", "rides", "éclaircissant", "dentifrice", "ongles",
                      "cils", "parfum",
                      "=sourire éclatant", "=vous cachez votre sourire", "=tu caches ton sourire",
                      "=dents jaunes", "=les mauvaises odeurs"],
    "Cheveux / barbe": ["cheveux", "chute", "repousse", "barbe", "perruque", "calvitie"],
    "Gadgets / maison": ["gadget", "montre", "écouteurs", "caméra", "lampe", "mixeur",
                         "ventilateur", "massage", "chargeur",
                         # cuisine
                         "ustensile", "casserole", "poêle", "couteau", "friteuse", "bouilloire",
                         "rangement", "organisateur",
                         # voiture
                         "accessoire auto", "autoradio", "dashcam", "démarreur", "compresseur",
                         "tapis de voiture", "support téléphone", "support voiture", "chargeur sans fil",
                         "chargeur voiture", "support magnétique",
                         # maison
                         "nettoyage", "aspirateur", "humidificateur", "lampe solaire", "antimoustique",
                         # électronique
                         "enceinte", "casque", "powerbank", "smartwatch",
                         # autres
                         "outil", "perceuse", "jardin", "rasoir", "tondeuse"],
    "Vitalité / bien-être": ["vitalité", "énergie", "endurance", "virilité", "tonus",
                             "libido", "testostérone", "puissance", "cordyceps",
                             "café", "tisane",
                             "=retrouvez l'énergie", "=manque d'énergie", "=plus d'énergie ou d'endurance",
                             "=booste ton énergie", "=boostez votre quotidien", "=prends soin de ta virilité",
                             "=le secret des hommes", "=retrouve ta puissance", "=redeviens l'homme",
                             "=tu veux assurer", "=fatigué", "=vous manquez de",
                             "=retrouvez plus de confiance", "=finis les performances"],
    "Autres": ["bébé", "chaussures"],
    "Cadeaux": ["=cadeau de noël", "=offrez-lui ce collier", "=le cadeau parfait", "=un cadeau offert"],
    "Accroches": ["=ne laissez plus", "=ne laisse plus", "=ne laissez pas", "=il est temps de changer",
                  "=il est temps de réagir", "=il est temps de reprendre le contrôle", "=marre de",
                  "=assez de", "=dites adieu à", "=offre spéciale", "=offre exceptionnelle",
                  "=stock limité", "=profitez maintenant", "=ne manquez pas", "=résultat visible",
                  "=devenez le boss", "=découvrez le secret", "=simple à utiliser",
                  "=facile à appliquer", "=offrez à votre corps"],
}
NICHES_EN = {
    "Santé": ["pain", "diabetes", "blood pressure", "prostate", "joint", "arthritis",
              "hemorrhoids", "fatigue", "sleep", "insomnia", "cholesterol", "digestion",
              "constipation", "weight", "belly", "back"],
    "Beauté / soin": ["soap", "cream", "oil", "serum", "stretch marks", "acne",
                      "dark spots", "wrinkles", "whitening", "toothpaste", "nails",
                      "lashes", "perfume", "skin"],
    "Cheveux / barbe": ["hair", "hair loss", "regrowth", "beard", "wig"],
    "Gadgets / maison": ["gadget", "watch", "earbuds", "camera", "lamp", "blender",
                         "fan", "massage", "charger", "alarm", "knife",
                         "kitchen utensil", "cookware", "frying pan", "air fryer", "kettle",
                         "organizer", "car accessory", "car stereo", "dashcam", "jump starter",
                         "tire inflator", "phone holder", "car phone mount", "wireless charger", "vacuum cleaner", "humidifier",
                         "solar lamp", "mosquito", "speaker", "headphones", "power bank",
                         "smartwatch", "tools", "drill", "garden", "shaver", "trimmer"],
    "Autres": ["baby", "shoes", "bag", "car", "motorcycle", "kitchen", "cleaning"],
}
NOMS_PAYS = {"CI": "Côte d'Ivoire", "SN": "Sénégal", "CM": "Cameroun",
             "BJ": "Bénin", "TG": "Togo", "ML": "Mali", "BF": "Burkina Faso",
             "CD": "RD Congo", "GA": "Gabon", "NG": "Nigeria", "GH": "Ghana",
             "FR": "France", "BE": "Belgique", "CA": "Canada", "US": "États-Unis",
             "GB": "Royaume-Uni", "NE": "Niger", "GN": "Guinée", "MR": "Mauritanie",
             "CG": "Congo", "TD": "Tchad", "CF": "Centrafrique",
             "SL": "Sierra Leone", "LR": "Liberia", "GM": "Gambie"}

# Mots qui font courir un risque de refus sur Facebook (promesses de sante)
RISQUE_ELEVE = re.compile(
    r"gu[eé]ri|soigne|cancer|diab[eè]te|prostat|[eé]rection|impuissan|[eé]jacul|"
    r"\bsida\b|\bvih\b|h[eé]morro|ulc[eè]re|[eé]rectile|st[eé]rilit|paludisme|"
    r"cholest[eé]rol|\bcure\b|heal|cure for|diabetes|erectile", re.I)
RISQUE_MOYEN = re.compile(
    r"[eé]claircissan|blanchi|maigrir|minceur|perdre du poids|anti-?[aâ]ge|vergetur|"
    r"\brides?\b|douleur|articulation|arthrose|chute des cheveux|grossir|"
    r"whitening|weight loss|wrinkle|stretch mark|hair loss|\bpain\b|arthritis", re.I)

FICHIER_BASE = "base_pubs.csv"
FICHIER_EXCEL = "pubs_veille.xlsx"

# ---------------------------------------------------------------
# Lecture des dates (anglais et francais)
# ---------------------------------------------------------------
MOIS = [
    ("jan", 1), ("feb", 2), ("fev", 2), ("fév", 2), ("mar", 3),
    ("apr", 4), ("avr", 4), ("may", 5), ("mai", 5), ("juin", 6),
    ("jun", 6), ("juil", 7), ("jul", 7), ("aug", 8), ("aou", 8),
    ("aoû", 8), ("sep", 9), ("oct", 10), ("nov", 11), ("dec", 12),
    ("déc", 12),
]


def numero_mois(mot):
    mot = mot.strip(".").lower()
    for debut, numero in MOIS:
        if mot.startswith(debut):
            return numero
    return None


def lire_date(texte):
    """Transforme 'Sep 30, 2026' ou '30 sept. 2026' en date. None si illisible."""
    t = texte.lower()
    m = re.search(r"(\d{1,2})\s+([a-zéûè\.]+)\s+(\d{4})", t)
    if m:
        mois = numero_mois(m.group(2))
        if mois:
            try:
                return date(int(m.group(3)), mois, int(m.group(1)))
            except ValueError:
                return None
    m = re.search(r"([a-zéûè\.]+)\s+(\d{1,2}),?\s+(\d{4})", t)
    if m:
        mois = numero_mois(m.group(1))
        if mois:
            try:
                return date(int(m.group(3)), mois, int(m.group(2)))
            except ValueError:
                return None
    return None


# ---------------------------------------------------------------
# Lecture du texte de la page : une pub = un bloc "Library ID"
# ---------------------------------------------------------------
def lire_pubs(texte_page, pays, mot, aujourdhui=None):
    aujourdhui = aujourdhui or date.today()
    pubs = []
    morceaux = texte_page.split("Library ID:")
    for morceau in morceaux[1:]:
        m = re.match(r"\s*(\d{6,})", morceau)
        if not m:
            continue
        pub_id = m.group(1)
        lignes = [l.strip() for l in morceau.split("\n") if l.strip()]

        debut = None
        for l in lignes:
            if "started running on" in l.lower() or "en diffusion depuis" in l.lower():
                debut = lire_date(l)
                break

        annonceur = ""
        texte_pub = ""
        for i, l in enumerate(lignes):
            if l.lower() in ("sponsored", "sponsorisé", "sponsorisée"):
                if i > 0:
                    annonceur = lignes[i - 1]
                texte_pub = " ".join(lignes[i + 1:i + 16])[:700]
                break

        identiques = ""
        m2 = re.search(r"(\d+)\s+ads? use this creative", morceau, re.IGNORECASE)
        if m2:
            identiques = m2.group(1)

        jours = (aujourdhui - debut).days if debut else ""
        pubs.append({
            "id": pub_id,
            "pays": pays,
            "mot_cle": mot,
            "annonceur": annonceur,
            "date_debut": debut.isoformat() if debut else "",
            "jours_actif": jours,
            "pubs_identiques": identiques,
            "texte": texte_pub,
            "lien": f"https://www.facebook.com/ads/library/?id={pub_id}",
            "vue_le": aujourdhui.isoformat(),
        })
    return pubs


# ---------------------------------------------------------------
# Navigation dans la page
# ---------------------------------------------------------------
BOUTONS_COOKIES = [
    "Decline optional cookies", "Allow all cookies",
    "Refuser les cookies optionnels", "Autoriser tous les cookies",
]


JS_MEDIAS = """
() => {
  const sortie = {};
  const snap = document.evaluate("//*[contains(text(),'Library ID:')]", document, null,
                                 XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
  for (let i = 0; i < snap.snapshotLength; i++) {
    const el = snap.snapshotItem(i);
    const m = (el.textContent || '').match(/Library ID:\\s*(\\d{6,})/);
    if (!m) continue;
    const id = m[1];
    let node = el, carte = null;
    for (let k = 0; k < 16 && node.parentElement; k++) {
      node = node.parentElement;
      const n = ((node.innerText || '').match(/Library ID:/g) || []).length;
      if (n > 1) break;
      carte = node;
    }
    if (!carte) continue;
    let video = null, poster = '';
    const vids = carte.querySelectorAll('video');
    if (vids.length) {
      const v = vids[0];
      poster = v.poster || '';
      const src = v.currentSrc || v.src || (v.querySelector('source') || {}).src || '';
      if (src.startsWith('http')) video = src;
    }
    let meilleure = '', aire = 0;
    carte.querySelectorAll('img').forEach(im => {
      const w = im.naturalWidth || im.width || 0, h = im.naturalHeight || im.height || 0;
      if (w >= 120 && h >= 120 && w * h > aire && (im.src || '').startsWith('http')) {
        aire = w * h; meilleure = im.src;
      }
    });
    const type = vids.length ? 'video' : (meilleure ? 'image' : '');
    // Adresse de destination du bouton de la pub (site, boutique ou WhatsApp du vendeur)
    let dest = '';
    const hotesFb = /(^|\\.)(facebook\\.com|fb\\.com|fb\\.me|instagram\\.com|fbcdn\\.net|meta\\.com)$/;
    for (const a of carte.querySelectorAll('a[href]')) {
      let h = a.href || '';
      try {
        const u = new URL(h);
        if (/(^|\\.)l\\.facebook\\.com$/.test(u.hostname) || /(^|\\.)lm\\.facebook\\.com$/.test(u.hostname)) {
          h = u.searchParams.get('u') || '';
        }
        if (!h.startsWith('http')) continue;
        const v = new URL(h);
        if (hotesFb.test(v.hostname)) continue;
        v.searchParams.delete('fbclid');
        dest = v.toString();
        break;
      } catch (e) {}
    }
    // Numero de la page Facebook de l'annonceur (cache dans les liens de la creative)
    let pageId = '';
    for (const a of carte.querySelectorAll('a[href]')) {
      const h = a.href || '';
      let m2 = h.match(/view_all_page_id=(\\d{6,})/);
      if (!m2) m2 = h.match(/^https?:\\/\\/(?:[a-z0-9-]+\\.)?facebook\\.com\\/(?:profile\\.php\\?id=)?(\\d{8,})(?:[\\/?&#]|$)/);
      if (m2) { pageId = m2[1]; break; }
    }
    sortie[id] = {type: type, vignette: poster || meilleure, video: video || '', dest: dest, page_id: pageId};
  }
  return sortie;
}
"""


def recuperer_texte(page, pays, mot, defilements, pause, sortie_rapide=False, page_id=None):
    if page_id:
        # Toutes les pubs d'UNE page Facebook precise (par son numero)
        url = (
            "https://www.facebook.com/ads/library/?active_status=active&ad_type=all"
            f"&country={pays}&view_all_page_id={page_id}"
            "&search_type=page&media_type=all"
        )
    elif mot.startswith("="):
        # Expression exacte : les mots dans cet ordre
        url = (
            "https://www.facebook.com/ads/library/?active_status=active&ad_type=all"
            f"&country={pays}&q={quote_plus(chr(34) + mot[1:] + chr(34))}"
            "&search_type=keyword_exact_phrase&media_type=all"
        )
    else:
        url = (
            "https://www.facebook.com/ads/library/?active_status=active&ad_type=all"
            f"&country={pays}&q={quote_plus(mot)}"
            "&search_type=keyword_unordered&media_type=all"
        )
    page.goto(url, wait_until="domcontentloaded", timeout=60000)
    page.wait_for_timeout(6000)
    for nom in BOUTONS_COOKIES:
        try:
            page.get_by_role("button", name=nom).first.click(timeout=2000)
            page.wait_for_timeout(1500)
            break
        except Exception:
            pass
    if sortie_rapide:
        # Aucune pub dans ce pays : inutile de descendre dans la page
        try:
            if "Library ID:" not in page.inner_text("body"):
                return "", {}
        except Exception:
            pass
    for _ in range(defilements):
        page.mouse.wheel(0, 7000)
        page.wait_for_timeout(int(pause * 1000))
    try:
        medias = page.evaluate(JS_MEDIAS)
    except Exception:
        medias = {}
    return page.inner_text("body"), medias


# ---------------------------------------------------------------
# Base de donnees (fichier CSV, avec point-virgule pour Excel)
# ---------------------------------------------------------------
COLONNES = [
    "id", "pays", "niche", "mot_cle", "annonceur", "date_debut", "jours_actif",
    "pubs_identiques", "texte", "lien", "vue_le", "premiere_vue",
    "media_type", "media_url", "video_url", "lien_dest",
]


def cle(ligne):
    return f"{ligne['pays']}-{ligne['id']}"


def charger_base(chemin):
    if not Path(chemin).exists():
        return {}
    with open(chemin, newline="", encoding="utf-8-sig") as f:
        premiere = f.readline()
        f.seek(0)
        sep = ";" if premiere.count(";") > premiere.count(",") else ","
        return {cle(l): l for l in csv.DictReader(f, delimiter=sep)}


def enregistrer_base(chemin, base):
    with open(chemin, "w", newline="", encoding="utf-8-sig") as f:
        ecrivain = csv.DictWriter(f, fieldnames=COLONNES, delimiter=";")
        ecrivain.writeheader()
        for ligne in base.values():
            ecrivain.writerow({c: ligne.get(c, "") for c in COLONNES})


def _fusion(ancien, nouveau):
    liste = [m for m in (ancien or "").split(" | ") if m]
    if nouveau and nouveau not in liste:
        liste.append(nouveau)
    return " | ".join(liste)


def ajouter(base, pub):
    """Ajoute une pub, ou met a jour si on la connait deja."""
    k = cle(pub)
    if k in base:
        ancien = base[k]
        pub["mot_cle"] = _fusion(ancien.get("mot_cle"), pub["mot_cle"])
        pub["niche"] = _fusion(ancien.get("niche"), pub.get("niche", ""))
        pub["premiere_vue"] = ancien.get("premiere_vue") or ancien.get("vue_le") or pub["vue_le"]
        for champ in ("media_type", "media_url", "video_url", "lien_dest"):
            if not pub.get(champ):
                pub[champ] = ancien.get(champ, "")
        nouvelle = False
    else:
        pub["premiere_vue"] = pub["vue_le"]
        nouvelle = True
    base[k] = pub
    return nouvelle


def risque(texte):
    if RISQUE_ELEVE.search(texte or ""):
        return "Élevé"
    if RISQUE_MOYEN.search(texte or ""):
        return "Moyen"
    return "Faible"


# ---------------------------------------------------------------
# Envoi vers Supabase (la base en ligne qui alimente la page de filtres)
# ---------------------------------------------------------------
FICHIER_CONFIG = "config_supabase.json"


def charger_config(chemin=FICHIER_CONFIG):
    """Reglages Supabase. Sur GitHub : variables secretes SUPABASE_URL, SUPABASE_KEY, CODE_ROBOT.
    Sur ton PC : le fichier config_supabase.json. Renvoie None si rien n'est trouve."""
    url = os.environ.get("SUPABASE_URL", "").strip()
    cle = os.environ.get("SUPABASE_KEY", "").strip()
    code = os.environ.get("CODE_ROBOT", "").strip()
    if url and cle and code:
        return {"url": url, "cle_publique": cle, "code_robot": code}
    if Path(chemin).exists():
        return json.loads(Path(chemin).read_text(encoding="utf-8"))
    return None


# --- Filtre "signal de vente" pour la niche Accroches (phrases tres larges => beaucoup de bruit) ---
_SIGNAL_TEXTE = re.compile(
    r"(\bprix\b|\btarif|livraison|\bcommand|whatsapp|wa\.me|\bpromo|fcfa|f\s?cfa|\bcfa\b|"
    r"\d\s?(\$|€|usd|eur|fc\b|frs?\b|f\b)|[$€]\s?\d|\bachet|paiement|boutique|\binbox\b|\bcommande)", re.I)
_SIGNAL_LIEN = re.compile(
    r"(myshopify|youcan|storeep|/products?/|/product/|\.shop\b|shop|store|boutique|wa\.me|api\.whatsapp|t\.me)", re.I)


def a_signal_de_vente(l):
    """Vrai si la crea ressemble a une vraie vente (prix, livraison, WhatsApp, lien boutique)."""
    return bool(_SIGNAL_TEXTE.search(str(l.get("texte", ""))) or _SIGNAL_LIEN.search(str(l.get("lien_dest", ""))))


# Livres (romans Kindle) et annonceurs Amazon : pas des produits a revendre en cash a la livraison.
_ANNONCEUR_AMAZON = re.compile(r"(^amazon\b|with amazon associates$)", re.I)
_TEXTE_LIVRE = re.compile(r"kindle unlimited", re.I)
LIVRES_EXCLUS = {
    "christian romantic suspense", "lara adrian books", "siobhan davis", "olivia darcy autora",
    "skye webb - herbal medicine & wellness author",
}


def est_livre_ou_amazon(l):
    nom = str(l.get("annonceur", "")).strip()
    return bool(nom.lower() in LIVRES_EXCLUS or _ANNONCEUR_AMAZON.search(nom)
                or _TEXTE_LIVRE.search(str(l.get("texte", ""))))


def envoyer_supabase(base, chemin_config=FICHIER_CONFIG, paquet=50):
    """Envoie toute la base dans Supabase. Ne fait rien si la config manque."""
    conf = charger_config(chemin_config)
    if not conf:
        print(f"(Pas de {chemin_config} ni de variables secretes : envoi vers la base en ligne ignore.)")
        return False
    url = conf["url"].rstrip("/") + "/rest/v1/rpc/envoyer_pubs"
    entetes = {"apikey": conf["cle_publique"],
               "Authorization": "Bearer " + conf["cle_publique"],
               "Content-Type": "application/json"}
    lignes = []
    ignorees = 0
    exclus = 0
    sans_vente = 0
    livres_amazon = 0
    aujourdhui_envoi = date.today()
    for k, l in base.items():
        age = jours_depuis(l, aujourdhui_envoi)
        if age is not None and age > AGE_MAX_JOURS:
            ignorees += 1
            continue
        if str(l.get("annonceur", "")).strip().lower() in ANNONCEURS_EXCLUS:
            exclus += 1
            continue
        if est_livre_ou_amazon(l):
            livres_amazon += 1
            continue
        if str(l.get("niche", "")).strip() == "Accroches" and not a_signal_de_vente(l):
            sans_vente += 1
            continue
        lignes.append({
            "cle": k, "id": l.get("id", ""), "pays": l.get("pays", ""),
            "niche": nettoyer(l.get("niche")), "mot_cle": nettoyer(l.get("mot_cle")),
            "annonceur": nettoyer(l.get("annonceur")),
            "date_debut": l.get("date_debut", ""),
            "pubs_identiques": str(l.get("pubs_identiques", "")),
            "texte": nettoyer(l.get("texte")), "lien": l.get("lien", ""),
            "vue_le": l.get("vue_le", ""),
            "premiere_vue": l.get("premiere_vue") or l.get("vue_le", ""),
            "risque": risque(l.get("texte")),
            "media_type": l.get("media_type", ""),
            "media_url": l.get("media_url", ""),
            "video_url": l.get("video_url", ""),
            "lien_dest": l.get("lien_dest", ""),
        })
    if exclus:
        print(f"({exclus} creas d'annonceurs exclus (logiciels, applis, formations) non envoyees a la base.)")
    if livres_amazon:
        print(f"({livres_amazon} creas de livres / annonceurs Amazon non envoyees.)")
    if sans_vente:
        print(f"({sans_vente} creas 'Accroches' sans signal de vente (prix, livraison, WhatsApp, boutique) non envoyees.)")
    if ignorees:
        print(f"({ignorees} creas de plus de {AGE_MAX_JOURS} jours ignorees, non envoyees a la base.)")

    def envoyer_lot(lot):
        """Envoie un lot. Si la base repond 'trop long' (timeout), on coupe le lot en deux et on recommence."""
        corps = json.dumps({"code": conf["code_robot"], "lignes": lot}).encode("utf-8")
        req = urllib.request.Request(url, data=corps, headers=entetes, method="POST")
        try:
            with urllib.request.urlopen(req, timeout=60) as r:
                r.read()
            return len(lot)
        except urllib.error.HTTPError as e:
            detail = e.read().decode("utf-8", "ignore")[:200]
            if e.code >= 500 and len(lot) > 4:
                milieu = len(lot) // 2
                time.sleep(1)
                a = envoyer_lot(lot[:milieu])
                b = envoyer_lot(lot[milieu:]) if a is not None else None
                return None if (a is None or b is None) else a + b
            print(f"   Envoi refuse ({e.code}) : {detail}")
            return None
        except Exception as e:
            print(f"   Envoi impossible : {e}")
            return None

    envoyees = 0
    for i in range(0, len(lignes), paquet):
        n = envoyer_lot(lignes[i:i + paquet])
        if n is None:
            return False
        envoyees += n
    print(f"Base en ligne mise a jour : {envoyees} pubs envoyees.")
    return True


# ---------------------------------------------------------------
# Excel soigne
# ---------------------------------------------------------------
# Annonceurs qui ne vendent pas de produit a revendre (logiciels, applis, formations, places de marche) :
# ils ne sont pas envoyes a la base. Noms en minuscules.
ANNONCEURS_EXCLUS = {
    "alibaba.com", "turboscribe", "hearty app", "freakaudio.store", "scale certified",
    "cinematic composing", "qeravio", "translate now ios", "promuse.ai", "facelessreels",
    "iman gadzhi", "ewa: learn languages", "sleep monitor - sleep recorder, talk & cycle",
    "videoexpress.ai", "woocommerce", "elevenlabs", "vid.ai", "moonly", "polsia",
    "clonevoiceai", "family rosary", "trending crowdfunding projects",
    # grandes marques (pas revendables en dropshipping)
    "colgate", "colgate canada", "nivea", "pedigree canada", "silk canada", "silk hair",
}
AGE_MAX_JOURS = 400  # les creas actives depuis plus longtemps ne sont plus gardees dans la base en ligne
ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def nettoyer(texte):
    return ILLEGAL.sub("", str(texte or ""))


def jours_depuis(ligne, aujourdhui):
    try:
        debut = date.fromisoformat(ligne.get("date_debut", ""))
    except ValueError:
        return None
    return (aujourdhui - debut).days


def exporter_excel(base, chemin, aujourdhui=None):
    from openpyxl import Workbook
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    aujourdhui = aujourdhui or date.today()
    lignes = []
    for l in base.values():
        l = dict(l)
        l["jours"] = jours_depuis(l, aujourdhui)
        lignes.append(l)
    lignes.sort(key=lambda l: (l["jours"] is None, -(l["jours"] or 0)))

    bleu = "1F3A5F"
    gris = "D9DEE7"
    bord = Border(*(Side(style="thin", color=gris),) * 4)
    entete_fill = PatternFill("solid", fgColor=bleu)
    entete_font = Font(bold=True, color="FFFFFF", size=11)
    zebra = PatternFill("solid", fgColor="F5F7FB")

    wb = Workbook()

    # ---------- Feuille 1 : Pubs ----------
    ws = wb.active
    ws.title = "Pubs"
    titres = ["Pays", "Niche", "Annonceur", "Jours actif", "Date début",
              "Pubs identiques", "Texte de la pub", "Lien", "Mot-clé",
              "Risque Facebook", "Nouveau", "Vue le"]
    largeurs = [14, 16, 28, 11, 12, 10, 70, 14, 24, 12, 11, 12]
    ws.append(titres)
    for i, l in enumerate(largeurs, start=1):
        ws.column_dimensions[chr(64 + i)].width = l
    for c in ws[1]:
        c.fill, c.font, c.border = entete_fill, entete_font, bord
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 32
    couleur_risque = {"Élevé": "F8CBAD", "Moyen": "FFE699", "Faible": "C6E0B4"}

    for n, l in enumerate(lignes, start=2):
        try:
            d_debut = date.fromisoformat(l.get("date_debut", ""))
        except ValueError:
            d_debut = None
        try:
            d_vue = date.fromisoformat(l.get("vue_le", ""))
        except ValueError:
            d_vue = None
        ident = int(l["pubs_identiques"]) if str(l.get("pubs_identiques", "")).isdigit() else None
        premiere = l.get("premiere_vue") or l.get("vue_le") or ""
        est_nouvelle = (premiere == aujourdhui.isoformat())
        niv = risque(l.get("texte"))
        ws.append([
            NOMS_PAYS.get(l["pays"], l["pays"]),
            nettoyer(l.get("niche")),
            nettoyer(l.get("annonceur")),
            l["jours"],
            d_debut,
            ident,
            nettoyer(l.get("texte")),
            "Voir la pub",
            nettoyer(l.get("mot_cle")),
            niv,
            "NOUVEAU" if est_nouvelle else "",
            d_vue,
        ])
        ws.cell(row=n, column=8).hyperlink = l.get("lien", "")
        ws.cell(row=n, column=8).font = Font(color="0563C1", underline="single")
        ws.cell(row=n, column=5).number_format = "DD/MM/YYYY"
        ws.cell(row=n, column=12).number_format = "DD/MM/YYYY"
        ws.row_dimensions[n].height = 60
        for col in range(1, 13):
            c = ws.cell(row=n, column=col)
            c.border = bord
            if n % 2 == 1:
                c.fill = zebra
            horiz = "center" if col in (1, 4, 5, 6, 8, 10, 11, 12) else "left"
            c.alignment = Alignment(horizontal=horiz, vertical="center",
                                    wrap_text=col in (2, 3, 7, 9))
        ws.cell(row=n, column=4).font = Font(bold=True, size=12)
        ws.cell(row=n, column=10).fill = PatternFill("solid", fgColor=couleur_risque[niv])
        if est_nouvelle:
            ws.cell(row=n, column=11).font = Font(bold=True, color="C00000")

    fin = len(lignes) + 1
    if fin >= 2:
        plage = f"D2:D{fin}"
        ws.conditional_formatting.add(plage, CellIsRule(
            operator="greaterThanOrEqual", formula=["60"],
            fill=PatternFill("solid", bgColor="C6EFCE", fgColor="C6EFCE"),
            font=Font(bold=True, color="006100")))
        ws.conditional_formatting.add(plage, CellIsRule(
            operator="between", formula=["30", "59"],
            fill=PatternFill("solid", bgColor="FFF2CC", fgColor="FFF2CC"),
            font=Font(bold=True, color="7F6000")))
    ws.freeze_panes = "D2"
    ws.auto_filter.ref = f"A1:L{max(fin, 2)}"
    ws.sheet_view.zoomScale = 100

    # ---------- Feuille 2 : Annonceurs ----------
    groupes = {}
    for l in lignes:
        nom = nettoyer(l.get("annonceur")).strip()
        if not nom:
            continue
        g = groupes.setdefault(nom, {"total": 0, "longues": 0, "max": -1,
                                     "pays": set(), "lien": ""})
        g["total"] += 1
        g["pays"].add(NOMS_PAYS.get(l["pays"], l["pays"]))
        j = l["jours"]
        if j is not None:
            if j >= 30:
                g["longues"] += 1
            if j > g["max"]:
                g["max"], g["lien"] = j, l.get("lien", "")

    wa = wb.create_sheet("Annonceurs")
    wa.append(["Annonceur", "Pays", "Pubs actives", "Dont 30 jours et +",
               "Pub la plus ancienne (jours)", "Lien"])
    for i, w in enumerate([34, 28, 13, 16, 20, 14], start=1):
        wa.column_dimensions[chr(64 + i)].width = w
    for c in wa[1]:
        c.fill, c.font, c.border = entete_fill, entete_font, bord
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    wa.row_dimensions[1].height = 34
    classes = sorted(groupes.items(),
                     key=lambda kv: (-kv[1]["longues"], -kv[1]["max"], kv[0]))
    for n, (nom, g) in enumerate(classes, start=2):
        wa.append([nom, ", ".join(sorted(g["pays"])), g["total"], g["longues"],
                   g["max"] if g["max"] >= 0 else None, "Voir la pub"])
        wa.cell(row=n, column=6).hyperlink = g["lien"]
        wa.cell(row=n, column=6).font = Font(color="0563C1", underline="single")
        wa.row_dimensions[n].height = 26
        for col in range(1, 7):
            c = wa.cell(row=n, column=col)
            c.border = bord
            if n % 2 == 1:
                c.fill = zebra
            c.alignment = Alignment(horizontal="left" if col in (1, 2) else "center",
                                    vertical="center", wrap_text=col == 2)
    wa.freeze_panes = "B2"
    wa.auto_filter.ref = f"A1:F{max(len(classes) + 1, 2)}"

    wb.save(chemin)


# ---------------------------------------------------------------
# Programme principal
# ---------------------------------------------------------------
FICHIER_SUIVI = "concurrents.txt"


def lire_concurrents(chemin=FICHIER_SUIVI, chemin_config=FICHIER_CONFIG):
    """Noms des vendeurs a suivre : ceux suivis dans AfriPepites (base en ligne)
    + ceux ecrits dans concurrents.txt (un nom par ligne)."""
    noms = []
    # 1) La liste en ligne (alimentee par le bouton "Suivre" de l'application)
    conf = charger_config(chemin_config)
    if conf:
        try:
            # Acces securise avec le code du robot : liste de tous les vendeurs suivis
            corps = json.dumps({"code": conf["code_robot"]}).encode("utf-8")
            req = urllib.request.Request(
                conf["url"].rstrip("/") + "/rest/v1/rpc/lister_suivis",
                data=corps, method="POST",
                headers={"apikey": conf["cle_publique"],
                         "Authorization": "Bearer " + conf["cle_publique"],
                         "Content-Type": "application/json"})
            with urllib.request.urlopen(req, timeout=30) as r:
                en_ligne = [str(x) for x in json.loads(r.read().decode("utf-8"))]
            print(f"{len(en_ligne)} vendeur(s) suivi(s) dans AfriPepites.")
            noms += en_ligne
        except Exception as e:
            print(f"(Impossible de lire la liste en ligne : {e})")
    # 2) Le fichier texte (en secours / en plus)
    if Path(chemin).exists():
        for ligne in Path(chemin).read_text(encoding="utf-8-sig").splitlines():
            ligne = ligne.strip()
            if ligne and not ligne.startswith("#"):
                noms.append(ligne)
    # Sans doublons, dans l'ordre
    vus, uniques = set(), []
    for n in noms:
        k = n.lower().strip()
        if k not in vus:
            vus.add(k)
            uniques.append(n.strip())
    return uniques


TOUS_PAYS = ["BJ", "TG", "CI", "SN", "CM", "ML", "BF", "NE", "GN", "MR", "GA", "CG",
             "CD", "TD", "CF", "NG", "GH", "SL", "LR", "GM", "FR", "CA"]


def _appel_robot(fonction, corps):
    """Appelle une fonction de la base avec le code du robot. Renvoie la reponse JSON."""
    conf = charger_config(FICHIER_CONFIG)
    if not conf:
        raise RuntimeError("reglages Supabase absents")
    corps = dict(corps, code=conf["code_robot"])
    req = urllib.request.Request(
        conf["url"].rstrip("/") + "/rest/v1/rpc/" + fonction,
        data=json.dumps(corps).encode("utf-8"), method="POST",
        headers={"apikey": conf["cle_publique"],
                 "Authorization": "Bearer " + conf["cle_publique"],
                 "Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=30) as r:
        texte = r.read().decode("utf-8")
    return json.loads(texte) if texte else None


def lire_file_attente():
    """Annonceurs que quelqu'un vient de suivre et dont on n'a pas encore ramene les pubs."""
    try:
        return [str(x) for x in (_appel_robot("suivis_a_traiter", {}) or [])]
    except Exception as e:
        print(f"(Impossible de lire la file d'attente : {e})")
        return []


def marquer_traite(nom):
    try:
        _appel_robot("marquer_suivi_traite", {"p_nom": nom})
    except Exception as e:
        print(f"(Impossible de marquer {nom!r} comme traite : {e})")


def pays_connus(nom):
    try:
        return [str(x) for x in (_appel_robot("pays_de_suivi", {"p_nom": nom}) or [])]
    except Exception:
        return []


def pages_de_suivi(nom):
    """Numeros de pages Facebook deja connus pour cet annonceur suivi (liste vide si inconnus)."""
    try:
        return [str(x) for x in (_appel_robot("pages_de_suivi", {"p_nom": nom}) or [])]
    except Exception:
        return []


def creas_exemple(nom):
    """Quelques creas deja en base pour cet annonceur (numeros de la bibliotheque Facebook)."""
    try:
        return [str(x) for x in (_appel_robot("creas_exemple_suivi", {"p_nom": nom}) or [])]
    except Exception:
        return []


def pages_depuis_creas(page, nom, pause):
    """Comme un humain : on part d'une pub DEJA connue de cet annonceur, on ouvre cette pub
    sur Facebook et on lit le numero de SA page. Evite de se tromper quand plusieurs pages
    portent le meme nom."""
    compte = {}
    for ad_id in creas_exemple(nom):
        try:
            page.goto(f"https://www.facebook.com/ads/library/?id={ad_id}", wait_until="domcontentloaded", timeout=60000)
            page.wait_for_timeout(6000)
            hrefs = page.evaluate("() => Array.from(document.querySelectorAll('a[href]')).map(a => a.href)")
        except Exception as erreur:
            print(f"   Pub {ad_id} illisible : {erreur}")
            continue
        for h in hrefs:
            m = re.search(r"view_all_page_id=(\d{6,})", h)
            if m:
                compte[m.group(1)] = compte.get(m.group(1), 0) + 1
                break
        time.sleep(random.uniform(1, 2))
    ids = [pid for pid, _ in sorted(compte.items(), key=lambda x: -x[1])[:3]]
    print(f"   Numeros de page lus depuis des pubs connues de {nom!r} : {ids if ids else 'aucun'}")
    return ids


def enregistrer_pages_suivi(nom, ids):
    try:
        _appel_robot("enregistrer_pages_suivi", {"p_nom": nom, "p_ids": list(ids)})
    except Exception as e:
        print(f"(Impossible d'enregistrer les numeros de page de {nom!r} : {e})")


def choisir_pages(pubs, medias, nom, maxi=3):
    """Parmi les pubs trouvees pour ce nom, retient les pages dont le nom est EXACTEMENT
    celui de l'annonceur suivi (les plus actives d'abord, 3 au maximum)."""
    voulu = _norm(nom)
    if not voulu:
        return []
    compte = {}
    for pub in pubs:
        pid = (medias.get(pub["id"], {}) or {}).get("page_id", "")
        if not pid:
            continue
        if _norm(pub.get("annonceur", "")) == voulu:
            compte[pid] = compte.get(pid, 0) + 1
    return [pid for pid, _ in sorted(compte.items(), key=lambda x: -x[1])[:maxi]]


def trouver_pages(page, nom, pause):
    """Cherche l'annonceur dans TOUS les pays et lit le numero de sa page Facebook."""
    try:
        texte, medias = recuperer_texte(page, "ALL", nom, 4, pause)
    except Exception as erreur:
        print(f"   Recherche du numero de page impossible : {erreur}")
        return []
    pubs = lire_pubs(texte, "ALL", nom)
    ids = choisir_pages(pubs, medias, nom)
    print(f"   Numeros de page trouves pour {nom!r} : {ids if ids else 'aucun'} ({len(pubs)} pubs lues)")
    return ids


def ma_part(noms, part):
    """Plusieurs robots en parallele : 'i/N' = le robot numero i (sur N) prend 1 annonceur sur N."""
    try:
        i, n = [int(x) for x in str(part or "1/1").split("/")]
    except Exception:
        i, n = 1, 1
    if n < 1 or not 1 <= i <= n:
        i, n = 1, 1
    return noms[i - 1::n]


def _norm(t):
    """Minuscules, sans accents ni signes : pour comparer des noms."""
    import unicodedata
    t = unicodedata.normalize("NFD", t or "")
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9]+", " ", t.lower()).strip()


def construire_recherches(args):
    """Liste de (pays, niche, mot) a faire."""
    if args.file:
        noms = ma_part(lire_file_attente(), args.part)
        args.noms_file = noms
        # Premier passage : tous les pays couverts par le robot
        return [(pays, "", nom) for nom in noms for pays in TOUS_PAYS]
    if args.suivi:
        noms = ma_part(lire_concurrents(), args.part)
        if not noms:
            print(f"Aucun vendeur a traiter (fichier {FICHIER_SUIVI} ou base en ligne).")
        if args.pays_connus:
            # Passages de routine : seulement les pays ou l'annonceur a deja des pubs chez nous
            recherches = []
            for nom in noms:
                liste = sorted(set(pays_connus(nom)) | {"BJ"})
                recherches += [(pays, "", nom) for pays in liste]
            return recherches
        return [(pays, "", nom) for pays in args.pays for nom in noms]
    pays_liste = args.pays
    if args.zone:
        pays_liste = ZONES[args.zone]
    voulues = {n.lower() for n in args.niches} if args.niches else None
    recherches = []
    for pays in pays_liste:
        if args.mots:
            recherches += [(pays, "Perso", m) for m in args.mots]
            continue
        langue = LANGUE_PAYS.get(pays, "fr")
        dico = NICHES_EN if langue == "en" else NICHES_FR
        for niche, mots in dico.items():
            if voulues and niche.lower() not in voulues:
                continue
            recherches += [(pays, niche, m) for m in mots]
    return recherches


def main():
    p = argparse.ArgumentParser(description="Collecteur Ad Library par niches")
    p.add_argument("--pays", nargs="+", default=PAYS_DEFAUT,
                   help="Codes pays : BJ TG CI SN CM FR US ...")
    p.add_argument("--zone", choices=list(ZONES), default=None,
                   help="afrique (BJ TG CI SN CM) ou occident (FR US)")
    p.add_argument("--niches", nargs="+", default=None,
                   help='Ex : --niches "Santé" "Beauté / soin"')
    p.add_argument("--mots", nargs="+", default=None,
                   help="Tes propres mots (remplace les niches)")
    p.add_argument("--defilements", type=int, default=6,
                   help="Combien de fois on descend dans la page")
    p.add_argument("--pause", type=float, default=2.5,
                   help="Secondes d'attente entre chaque defilement")
    p.add_argument("--min-jours", type=int, default=30)
    p.add_argument("--sans-envoi", action="store_true",
                   help="Ne pas envoyer les pubs vers la base en ligne")
    p.add_argument("--envoyer-seulement", action="store_true",
                   help="Envoie la base actuelle vers Supabase sans rien collecter")
    p.add_argument("--suivi", action="store_true",
                   help=f"Mode suivi : cherche chaque vendeur ecrit dans {FICHIER_SUIVI}")
    p.add_argument("--file", action="store_true",
                   help="File d'attente : annonceurs tout juste suivis, dans tous les pays")
    p.add_argument("--part", default="1/1",
                   help="Robots en parallele : 2/3 = je suis le robot 2 sur 3")
    p.add_argument("--pays-connus", action="store_true", dest="pays_connus",
                   help="Avec --suivi : seulement les pays ou l'annonceur a deja des pubs chez nous")
    p.add_argument("--debug", action="store_true",
                   help="Sauvegarde le texte brut de chaque page")
    args = p.parse_args()
    if args.suivi or args.file:
        args.defilements = max(args.defilements, 12)
    args.noms_file = []

    if args.envoyer_seulement:
        base = charger_base(FICHIER_BASE)
        try:
            exporter_excel(base, FICHIER_EXCEL)
            print(f"Excel pret : {FICHIER_EXCEL}")
        except PermissionError:
            print(f"IMPOSSIBLE d'ecrire {FICHIER_EXCEL} : ferme-le dans Excel et relance.")
        envoyer_supabase(base)
        return

    from playwright.sync_api import sync_playwright

    recherches = construire_recherches(args)
    if args.file and not recherches:
        print("File d'attente vide : rien a faire.")
        return
    base = charger_base(FICHIER_BASE)
    avant = len(base)
    total = len(recherches)
    print(f"{total} recherches a faire.")
    erreurs = {}

    try:
        with sync_playwright() as pw:
            navigateur = pw.chromium.launch(headless=False)
            contexte = navigateur.new_context(
                locale="en-US", viewport={"width": 1300, "height": 900}
            )
            page = contexte.new_page()
            captures = 0
            # Annonceurs suivis : on retrouve d'abord le numero de leur page Facebook
            pages_par_nom = {}
            if args.suivi or args.file:
                for nom_suivi in dict.fromkeys(m for (_, _, m) in recherches):
                    ids = pages_de_suivi(nom_suivi)
                    if not ids:
                        ids = pages_depuis_creas(page, nom_suivi, args.pause)
                    if not ids:
                        ids = trouver_pages(page, nom_suivi, args.pause)
                        if ids:
                            enregistrer_pages_suivi(nom_suivi, ids)
                    pages_par_nom[nom_suivi] = ids
            for fait, (pays, niche, mot) in enumerate(recherches, start=1):
                print(f"\n[{fait}/{total}] {pays} - {niche} - {mot!r}")
                ids_page = pages_par_nom.get(mot) or [None]
                pubs = []
                echec = False
                for pid in ids_page:
                    try:
                        texte, medias = recuperer_texte(page, pays, mot,
                                                        args.defilements, args.pause,
                                                        sortie_rapide=bool(args.file or args.suivi),
                                                        page_id=pid)
                    except Exception as erreur:
                        print(f"   Probleme sur cette recherche : {erreur}")
                        erreurs[mot] = erreurs.get(mot, 0) + 1
                        echec = True
                        continue
                    if args.debug:
                        nom = f"debug_{pays}_" + re.sub(r"\W+", "_", mot) + ".txt"
                        Path(nom).write_text(texte, encoding="utf-8")
                    morceau = lire_pubs(texte, pays, mot[1:] if mot.startswith("=") else mot)
                    if (args.suivi or args.file) and not pid:
                        # Sans numero de page : on ne garde que les pubs de CE vendeur (par son nom)
                        voulu = _norm(mot)
                        morceau = [x for x in morceau if voulu and voulu in _norm(x.get("annonceur", ""))]
                    for x in morceau:
                        x["_info"] = medias.get(x["id"], {})
                    pubs += morceau
                if not pubs and not echec and (args.suivi or args.file) and ids_page and ids_page[0]:
                    # Le numero de page n'a rien donne ici : on essaie aussi par le NOM exact
                    try:
                        texte, medias = recuperer_texte(page, pays, mot, 4, args.pause,
                                                        sortie_rapide=True, page_id=None)
                        voulu = _norm(mot)
                        morceau = [x for x in lire_pubs(texte, pays, mot) if voulu and voulu == _norm(x.get("annonceur", ""))]
                        for x in morceau:
                            x["_info"] = medias.get(x["id"], {})
                        pubs += morceau
                    except Exception as erreur:
                        print(f"   Recherche par nom impossible : {erreur}")
                if echec and not pubs:
                    continue
                medias = {x["id"]: x.get("_info", {}) for x in pubs}
                avec_media = 0
                avec_dest = 0
                for pub in pubs:
                    pub["niche"] = niche
                    pub.pop("_info", None)
                    info = medias.get(pub["id"], {})
                    pub["media_type"] = info.get("type", "")
                    pub["media_url"] = info.get("vignette", "")
                    pub["video_url"] = info.get("video", "")
                    pub["lien_dest"] = info.get("dest", "")
                    if pub["media_url"]:
                        avec_media += 1
                    if pub["lien_dest"]:
                        avec_dest += 1
                nouvelles = sum(1 for pub in pubs if ajouter(base, pub))
                print(f"   {len(pubs)} pubs lues, {nouvelles} nouvelles, {avec_media} avec image/video, {avec_dest} avec lien du vendeur")
                if not pubs:
                    print("   (0 pub : Facebook a peut-etre bloque ou change sa page)")
                    # Sur GitHub : on garde 3 photos d'ecran pour comprendre ce qui s'est passe
                    if os.environ.get("ROBOT_CAPTURE") and captures < 3:
                        captures += 1
                        try:
                            page.screenshot(path=f"capture_{pays}_{captures}.png")
                        except Exception:
                            pass
                if fait % 5 == 0:
                    enregistrer_base(FICHIER_BASE, base)
                time.sleep(random.uniform(1, 2) if (args.file or args.suivi) and not pubs else random.uniform(3, 6))
            navigateur.close()
    except KeyboardInterrupt:
        print("\nArret demande : on sauvegarde ce qui a ete collecte.")
        args.noms_file = []
    finally:
        enregistrer_base(FICHIER_BASE, base)

    print(f"\nBase : {len(base)} pubs ({len(base) - avant} nouvelles)")
    try:
        exporter_excel(base, FICHIER_EXCEL)
        print(f"Excel pret : {FICHIER_EXCEL}")
    except PermissionError:
        print(f"\nIMPOSSIBLE d'ecrire {FICHIER_EXCEL} : il est ouvert dans Excel.")
        print("Ferme Excel puis relance la commande (les donnees sont sauvegardees).")

    if not args.sans_envoi:
        envoye = envoyer_supabase(base)
        # File d'attente : on coche l'annonceur seulement si l'envoi a reussi
        # et si Facebook n'a pas bloque plus de la moitie des recherches
        if args.file and envoye:
            for nom in args.noms_file:
                if erreurs.get(nom, 0) <= len(TOUS_PAYS) // 2:
                    marquer_traite(nom)
                    print(f"Annonceur traite : {nom}")
                else:
                    print(f"Annonceur a refaire (trop d'erreurs) : {nom}")

    aujourdhui = date.today()
    longues = [(jours_depuis(l, aujourdhui), l) for l in base.values()]
    longues = [(j, l) for j, l in longues if j is not None and j >= args.min_jours]
    print(f"Pubs actives depuis {args.min_jours} jours ou plus : {len(longues)}")


if __name__ == "__main__":
    main()
