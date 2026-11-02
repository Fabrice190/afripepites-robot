"""Essai : recherche par image sur une plateforme (alibaba ou dhgate)."""
import json, os, re, urllib.request
from playwright.sync_api import sync_playwright

SITE = os.environ["SITE"]
ACCUEIL = {"alibaba": "https://www.alibaba.com/", "dhgate": "https://www.dhgate.com/"}[SITE]
os.makedirs("sortie", exist_ok=True)
req = urllib.request.Request(os.environ["URL_IMAGE"], headers={"User-Agent": "Mozilla/5.0"})
open("sortie/produit.jpg", "wb").write(urllib.request.urlopen(req, timeout=30).read())

with sync_playwright() as p:
    nav = p.chromium.launch(headless=True, args=["--disable-blink-features=AutomationControlled"])
    ctx = nav.new_context(locale="en-US", viewport={"width": 1366, "height": 900},
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36")
    page = ctx.new_page()
    try:
        page.goto(ACCUEIL, wait_until="domcontentloaded", timeout=60000)
        page.wait_for_timeout(6000)
        print(SITE, "accueil :", page.title()[:80], "|", page.url[:100])
        page.screenshot(path=f"sortie/{SITE}_1.png")
        for sel in ['button:has-text("Accept")', '[class*="close"]', '[aria-label*="lose"]']:
            try:
                if page.locator(sel).count():
                    page.locator(sel).first.click(timeout=2000, force=True); page.wait_for_timeout(700)
            except Exception:
                pass
        page.keyboard.press("Escape")
        envoye = False
        print("Champs fichier au départ :", page.locator('input[type="file"]').count())
        candidats = ['[class*="image-search"]', '[class*="imageSearch"]', '[class*="camera"]', '[class*="picture-search"]',
                     '[class*="photo"]', '[class*="visual"]', '[class*="upload"]', '[class*="search-img"]', '[class*="searchImg"]']
        if page.locator('input[type="file"]').count() == 0:
            for sel in candidats:
                if page.locator(sel).count():
                    try:
                        page.locator(sel).first.click(timeout=3000, force=True); page.wait_for_timeout(1500)
                        print("Clic :", sel, "-> champs fichier :", page.locator('input[type="file"]').count())
                        if page.locator('input[type="file"]').count(): break
                    except Exception as e:
                        print("Clic raté :", sel, repr(e)[:80])
        if page.locator('input[type="file"]').count():
            page.locator('input[type="file"]').first.set_input_files("sortie/produit.jpg")
            envoye = True
        if envoye:
            page.wait_for_timeout(15000)
            print("Après envoi :", page.title()[:80], "|", page.url[:160])
            page.screenshot(path=f"sortie/{SITE}_2.png")
            res = page.eval_on_selector_all("a",
                "els => els.filter(e => /(US\\\\s?\\\\$|\\\\$|€)\\\\s?\\\\d/.test(e.innerText||'')).slice(0,8).map(e => ({href: e.href.split('?')[0].slice(0,120), texte: (e.innerText||'').replace(/\\\\s+/g,' ').slice(0,170)}))")
            print("RESULTATS :", json.dumps(res, ensure_ascii=False, indent=1))
            liens = page.eval_on_selector_all('a[href*="product-detail"], a[href*="/product/"], a[href*="/item/"]',
                "els => els.slice(0,8).map(e => ({href: e.href.split('?')[0].slice(0,120), texte: (e.innerText||'').replace(/\\s+/g,' ').slice(0,200)}))")
            print("LIENS PRODUITS :", json.dumps(liens, ensure_ascii=False, indent=1))
            print("TEXTE PAGE :", page.inner_text("body")[:1500].replace("\n", " | "))
        else:
            print("Aucun endroit pour déposer l'image trouvé.")
            html = page.eval_on_selector_all('[class*="search"], [id*="search"]', "els => els.slice(0,12).map(e => e.tagName+'.'+e.className+'#'+e.id+' :: '+e.outerHTML.slice(0,160))")
            print("ZONE RECHERCHE :", json.dumps(html, ensure_ascii=False, indent=1))
            imgs = page.eval_on_selector_all('img[src*="camera"], img[src*="photo"], img[src*="upload"], [class*="camera"], [class*="pic"]', "els => els.slice(0,8).map(e => e.outerHTML.slice(0,200))")
            print("ICONES :", json.dumps(imgs, ensure_ascii=False))
            print("Texte :", page.inner_text("body")[:400].replace("\n", " "))
    except Exception as e:
        print("ERREUR :", repr(e)[:400])
    nav.close()
