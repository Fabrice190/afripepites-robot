"""Robot Pépites Agent IA : pour chaque agent relié à Telegram, cherche le prix d'achat du produit
(AliExpress + Alibaba, recherche par image), lit le prix de vente de la boutique, puis demande à la base d'envoyer le message."""
import json, os, re, statistics, sys, urllib.error, urllib.request
from playwright.sync_api import sync_playwright

URL = os.environ["SUPABASE_URL"].rstrip("/")
KEY = os.environ["SUPABASE_KEY"]
CODE = os.environ["CODE_ROBOT"]
TEST_USER = os.environ.get("TEST_USER") or None
TEST_ONLY = os.environ.get("TEST_ONLY", "") == "oui"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"


def rpc(nom, corps):
    req = urllib.request.Request(f"{URL}/rest/v1/rpc/{nom}", data=json.dumps(corps).encode(), method="POST",
                                 headers={"apikey": KEY, "Authorization": "Bearer " + KEY, "Content-Type": "application/json"})
    try:
        return json.loads(urllib.request.urlopen(req, timeout=60).read() or "null")
    except urllib.error.HTTPError as e:
        print("ERREUR de la base sur", nom, ":", e.code, e.read()[:500])
        raise


def taux_usd_fcfa():
    try:
        d = json.loads(urllib.request.urlopen("https://open.er-api.com/v6/latest/USD", timeout=20).read())
        return float(d["rates"]["XOF"])
    except Exception as e:
        print("Taux du jour indisponible, taux de secours utilisé :", e)
        return 590.0


def nombre(t):
    return float(t.replace(",", ""))


def resume(prix, taux):
    """Prix en dollars -> fourchette en FCFA (on écarte le plus bas et le plus haut quand il y en a assez)."""
    p = sorted(prix)
    if len(p) >= 5:
        p = p[1:-1]
    if not p:
        return None
    return {"fcfa_min": round(p[0] * taux), "fcfa_max": round(p[-1] * taux)}


def prix_vente(lien):
    try:
        req = urllib.request.Request(lien, headers={"User-Agent": UA, "Accept-Language": "fr,en;q=0.8"})
        h = urllib.request.urlopen(req, timeout=25).read().decode("utf-8", "ignore")
    except Exception as e:
        print("  Boutique illisible :", repr(e)[:100]); return ""
    for a, c in ((r'property="(?:og|product):price:amount"\s+content="([\d.,]+)"', r'property="(?:og|product):price:currency"\s+content="(\w+)"'),
                 (r'content="([\d.,]+)"\s+property="(?:og|product):price:amount"', r'content="(\w+)"\s+property="(?:og|product):price:currency"')):
        m, d = re.search(a, h), re.search(c, h)
        if m: return f"{m.group(1)} {d.group(1) if d else ''}".strip()
    m, d = re.search(r'"price"\s*:\s*"?([\d.,]+)"?', h), re.search(r'"priceCurrency"\s*:\s*"(\w+)"', h)
    if m and d: return f"{m.group(1)} {d.group(1)}"
    return ""


def cherche_aliexpress(ctx, image, taux):
    page = ctx.new_page()
    try:
        page.goto("https://www.aliexpress.com/", wait_until="domcontentloaded", timeout=60000)
        page.wait_for_timeout(5000)
        try: page.locator('[class*="close"]').first.click(timeout=2000, force=True)
        except Exception: pass
        page.keyboard.press("Escape")
        if page.locator('input[type="file"]').count() == 0:
            return None
        page.locator('input[type="file"]').first.set_input_files(image)
        page.wait_for_timeout(14000)
        res = page.eval_on_selector_all('a[href*="/item/"]',
            "els => els.slice(0,10).map(e => ({href: e.href.split('?')[0], texte: (e.innerText||'').replace(/\\s+/g,' ')}))")
        prix, premier, note = [], None, None
        for r in res:
            m = re.search(r"(\d\.\d)\s+(?:US\s?)?\$\s?([\d.,]+)", r["texte"]) or re.search(r"\$\s?([\d.,]+)", r["texte"])
            if not m: continue
            valeur = nombre(m.group(m.lastindex))
            if valeur <= 0: continue
            prix.append(valeur)
            if premier is None:
                premier = r["href"]; note = m.group(1) if m.lastindex == 2 else None
            if len(prix) >= 6: break
        out = resume(prix, taux)
        if out:
            out.update({"url": premier, "note": note}); print("  AliExpress :", out)
        return out
    except Exception as e:
        print("  AliExpress en échec :", repr(e)[:150]); return None
    finally:
        page.close()


def cherche_alibaba(ctx, image, taux):
    page = ctx.new_page()
    try:
        page.goto("https://www.alibaba.com/", wait_until="domcontentloaded", timeout=60000)
        page.wait_for_timeout(6000)
        try: page.locator('[class*="close"]').first.click(timeout=2000, force=True)
        except Exception: pass
        page.keyboard.press("Escape")
        if page.locator('input[type="file"]').count() == 0:
            try:
                page.locator('[class*="image-search"]').first.click(timeout=3000, force=True); page.wait_for_timeout(1500)
            except Exception: pass
        if page.locator('input[type="file"]').count() == 0:
            return None
        page.locator('input[type="file"]').first.set_input_files(image)
        page.wait_for_timeout(15000)
        res = page.eval_on_selector_all('a[href*="product-detail"]',
            "els => els.map(e => ({href: e.href.split('?')[0], texte: (e.innerText||'').replace(/\\s+/g,' ')}))")
        ordre, textes = [], {}
        for r in res:
            if r["href"] not in textes:
                ordre.append(r["href"]); textes[r["href"]] = ""
            textes[r["href"]] += " " + r["texte"]
        prix, premier, moq = [], None, None
        for h in ordre:
            m = re.search(r"\$\s?([\d,]+\.?\d*)(?:\s?-\s?\$?([\d,]+\.?\d*))?\s*Min\.? order:?\s*([\d,]+)", textes[h])
            if not m: continue
            prix.append(nombre(m.group(1)))
            if premier is None: premier, moq = h, m.group(3)
            if len(prix) >= 6: break
        out = resume(prix, taux)
        if out:
            out.update({"url": premier, "moq": moq}); print("  Alibaba :", out)
        return out
    except Exception as e:
        print("  Alibaba en échec :", repr(e)[:150]); return None
    finally:
        page.close()


def main():
    taux = taux_usd_fcfa()
    print("Taux du jour : 1 $ =", taux, "FCFA")
    agents = rpc("agents_du_jour", {"code": CODE, "p_test_user": TEST_USER})
    print(len(agents), "pépite(s) à envoyer")
    if not agents:
        return
    os.makedirs("sortie", exist_ok=True)
    with sync_playwright() as p:
        nav = p.chromium.launch(headless=True, args=["--disable-blink-features=AutomationControlled"])
        ctx = nav.new_context(locale="en-US", viewport={"width": 1366, "height": 900}, user_agent=UA)
        for i, a in enumerate(agents):
            print(f"[{i+1}/{len(agents)}] {a['cle']}")
            extra = {}
            try:
                image = f"sortie/produit_{i}.jpg"
                req = urllib.request.Request(a["media_url"], headers={"User-Agent": "Mozilla/5.0"})
                open(image, "wb").write(urllib.request.urlopen(req, timeout=30).read())
                ali = cherche_aliexpress(ctx, image, taux)
                abb = cherche_alibaba(ctx, image, taux)
                if ali: extra["ali"] = ali
                if abb: extra["alibaba"] = abb
            except Exception as e:
                print("  Recherche en échec :", repr(e)[:150])
            pv = prix_vente(a["lien_dest"])
            if pv: extra["prix_vente"] = pv; print("  Prix de vente :", pv)
            print("  Envoi :", rpc("agent_publier", {"code": CODE, "p_user": a["user_id"], "k": a["cle"], "p_extra": extra, "p_test": TEST_ONLY}))
        nav.close()


main()
