"""Essai : recherche par image sur AliExpress avec un vrai navigateur (sans écran)."""
import json, os, re, sys, urllib.request
from playwright.sync_api import sync_playwright

url_image = os.environ["URL_IMAGE"]
os.makedirs("sortie", exist_ok=True)
req = urllib.request.Request(url_image, headers={"User-Agent": "Mozilla/5.0"})
open("sortie/produit.jpg", "wb").write(urllib.request.urlopen(req, timeout=30).read())
print("Image téléchargée :", os.path.getsize("sortie/produit.jpg"), "octets")

with sync_playwright() as p:
    nav = p.chromium.launch(headless=True, args=["--disable-blink-features=AutomationControlled"])
    ctx = nav.new_context(locale="fr-FR", viewport={"width": 1366, "height": 900},
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36")
    page = ctx.new_page()
    reponses = []
    page.on("response", lambda r: reponses.append((r.status, r.url)) if re.search(r"image|search|imgsearch|similar", r.url, re.I) else None)
    try:
        page.goto("https://www.aliexpress.com/", wait_until="domcontentloaded", timeout=60000)
        page.wait_for_timeout(5000)
        print("Page d'accueil :", page.title(), "|", page.url)
        page.screenshot(path="sortie/1_accueil.png")
        # Fermer d'éventuelles fenêtres par-dessus (cookies, pubs, coupons)
        for sel in ['button:has-text("Accept")', 'button:has-text("Accepter")', '[class*="close"]', '[aria-label="close"]', '[aria-label="Close"]']:
            try:
                if page.locator(sel).count():
                    page.locator(sel).first.click(timeout=2000, force=True)
                    print("Fenêtre fermée :", sel)
                    page.wait_for_timeout(800)
            except Exception:
                pass
        page.keyboard.press("Escape")
        envoye = False
        try:
            with page.expect_file_chooser(timeout=8000) as fc:
                page.locator('[class*="picture-search-container"]').first.click(timeout=5000, force=True)
            fc.value.set_files("sortie/produit.jpg")
            envoye = True
            print("Image déposée via le sélecteur de fichier")
        except Exception as e:
            print("Sélecteur de fichier :", repr(e)[:200])
            champs = page.locator('input[type="file"]')
            print("Champs de fichier :", champs.count())
            if champs.count():
                champs.first.set_input_files("sortie/produit.jpg")
                envoye = True
        if envoye:
            page.wait_for_timeout(14000)
            print("Après envoi :", page.title(), "|", page.url)
            page.screenshot(path="sortie/2_resultats.png", full_page=False)
            liens = page.eval_on_selector_all('a[href*="/item/"]',
                "els => els.slice(0,8).map(e => ({href: e.href.split('?')[0], texte: (e.innerText||'').replace(/\\s+/g,' ').slice(0,160)}))")
            print("RESULTATS :", json.dumps(liens, ensure_ascii=False, indent=1))
        else:
            print("Impossible de trouver l'endroit où déposer l'image.")
            print("Texte de la page :", page.inner_text("body")[:600].replace("\n", " "))
    except Exception as e:
        print("ERREUR :", repr(e)[:400])
    print("Réponses réseau liées à la recherche :", reponses[:15])
    nav.close()
