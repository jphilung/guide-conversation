#!/usr/bin/env python3
"""Lit le Google Sheets (ou un .xlsx), complète les phonétiques vides, génère l'audio
et produit data.json + audio/ dans le dossier du site.

Installation (une fois) :
    pip install openpyxl pypinyin pykakasi korean-romanizer pycantonese edge-tts
    + espeak-ng installé sur le système (pour l'API automatique, optionnel)

Utilisation :
    python build_data.py ID_OU_URL_DU_SHEETS       # écrit dans site/
    python build_data.py guide_conversation.xlsx   # même chose depuis un fichier
    options : --no-audio   --force-audio (régénère tous les MP3)   --out dossier
              --write-back : écrit les colonnes api et romanisation générées dans le Sheets (compte de
              service en Éditeur requis) ; seulement des cellules vides, jamais ce que tu as saisi

Accès au Sheets, au choix :
  - compte de service Google (recommandé, Sheets privé) : variable GOOGLE_SERVICE_ACCOUNT_JSON
    ou option --credentials fichier.json ; le Sheets est partagé avec l'adresse du compte (lecteur) ;
  - sans identifiants : le Sheets doit être partagé « Toute personne disposant du lien : lecteur ».
Une valeur saisie à la main n'est JAMAIS écrasée. Les valeurs générées sont listées dans
a_verifier.csv : relis-les et colle-les dans le Sheets, elles deviennent alors les tiennes.
"""

import argparse, asyncio, csv, hashlib, io, json, os, re, shutil, subprocess, sys, urllib.parse, urllib.request
from pathlib import Path
from openpyxl import load_workbook

# ---------- Générateurs de phonétique (tous optionnels) ----------
def romanize(lang, text):
    lang = lang.split("-")[0]  # "zh-tw" -> "zh"
    try:
        if lang == "zh":
            from pypinyin import pinyin, Style
            return " ".join(s[0] for s in pinyin(text, style=Style.TONE))
        if lang == "yue":  # cantonais : code "yue" (et non zh-hk, qui donnerait du pinyin)
            import pycantonese
            return " ".join(j or w for w, j in pycantonese.characters_to_jyutping(text))
        if lang == "ja":
            import pykakasi
            return " ".join(w["hepburn"] for w in pykakasi.kakasi().convert(text))
        if lang == "ko":
            from korean_romanizer.romanizer import Romanizer
            return Romanizer(text).romanize()
    except ImportError as e:
        print(f"  (module manquant pour {lang}: {e.name})")
    return ""

ESPEAK = {"es-mx": "es-419", "pt-br": "pt-br",  # variantes régionales d'abord
          "yue": "yue", "zh": "cmn", "ja": "ja", "ko": "ko", "es": "es", "fr": "fr-fr", "en": "en-us",
          "it": "it", "de": "de", "pt": "pt", "th": "th", "ar": "ar", "vi": "vi"}

def ipa(lang, text):
    """IPA approximatif via espeak-ng : à relire absolument."""
    voice = ESPEAK.get(lang) or ESPEAK.get(lang.split("-")[0])
    if not shutil.which("espeak-ng") or not voice:
        return ""
    out = subprocess.run(["espeak-ng", "-q", "--ipa", "-v", voice, text],
                         capture_output=True, text=True)
    return out.stdout.strip().replace("\n", " ")

# ---------- Audio (régénéré seulement si le texte ou la voix change) ----------
async def make_audio(items, outdir, manifest_path, force=False):
    import edge_tts
    manifest = json.loads(manifest_path.read_text()) if manifest_path.exists() else {}
    for id_, text, voice, rate in items:
        h = hashlib.sha1(f"{text}|{voice}|{rate}".encode()).hexdigest()
        f = outdir / f"{id_}.mp3"
        if not force and manifest.get(id_) == h and f.exists():
            continue
        print(f"  audio : {id_}")
        try:
            await edge_tts.Communicate(text, voice, rate=rate).save(str(f))
        except Exception as e:  # on continue : l'app se rabat sur la synthèse du navigateur
            print(f"  ! échec audio {id_} : {e}"); continue
        manifest[id_] = h
    manifest_path.write_text(json.dumps(manifest, indent=1))

# ---------- Sources de données ----------
def _csv_rows(sid, tab):
    url = ("https://docs.google.com/spreadsheets/d/%s/gviz/tq?tqx=out:csv&headers=1&sheet=%s"
           % (sid, urllib.parse.quote(tab)))
    try:
        with urllib.request.urlopen(url) as r:
            txt = r.read().decode("utf-8")
    except Exception as e:
        sys.exit(f"Impossible de lire l'onglet « {tab} » ({e}). Lecture publique utilisée : aucun compte "
                 "de service détecté. Si tu en as configuré un, le secret GOOGLE_SERVICE_ACCOUNT_JSON est vide "
                 "ou introuvable. Sinon, le Sheets doit être partagé en lecture avec toute personne ayant le lien.")
    if txt.lstrip().startswith("<"):
        sys.exit("Google a renvoyé une page web : partage le Sheets en lecture avec toute personne ayant le lien.")
    return list(csv.reader(io.StringIO(txt)))

def _session(info, write=False):
    from google.oauth2 import service_account
    from google.auth.transport.requests import AuthorizedSession
    scope = "https://www.googleapis.com/auth/spreadsheets" + ("" if write else ".readonly")
    return AuthorizedSession(service_account.Credentials.from_service_account_info(info, scopes=[scope]))

def _api_rows(sid, tab, info):
    """Lecture via l'API Google Sheets avec un compte de service (Sheets privé, toutes les lignes)."""
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{sid}/values/{urllib.parse.quote(tab)}"
    r = _session(info).get(url, params={"valueRenderOption": "FORMATTED_VALUE"})
    if r.status_code != 200:
        sys.exit(f"Erreur de l'API Google Sheets ({r.status_code}) sur l'onglet « {tab} » : {r.text[:400]}")
    return r.json().get("values", [])

def _col(i):  # 0 -> A, 25 -> Z, 26 -> AA
    out, i = "", i + 1
    while i:
        i, r = divmod(i - 1, 26); out = chr(65 + r) + out
    return out

def write_back(sid, info, writes):
    """writes = [(id, colonne_index, valeur)]. Écrit UNIQUEMENT dans des cellules encore vides, et seulement
    si la ligne porte toujours le même id (le Sheets a pu être modifié ou trié depuis la lecture)."""
    cur = _api_rows(sid, "Expressions", info)  # relecture fraîche juste avant d'écrire
    idc = [h.strip() for h in cur[0]].index("id")
    pos = {str(r[idc]).strip(): n for n, r in enumerate(cur) if n and len(r) > idc}
    data = []
    for id_, ci, val in writes:
        n = pos.get(id_)
        if n is None or (ci < len(cur[n]) and str(cur[n][ci]).strip()):
            continue
        data.append({"range": f"Expressions!{_col(ci)}{n + 1}", "values": [[val]]})
    if not data:
        return 0
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{sid}/values:batchUpdate"
    r = _session(info, write=True).post(url, json={"valueInputOption": "RAW", "data": data})
    if r.status_code != 200:
        print(f"! écriture dans le Sheets impossible ({r.status_code}) : {r.text[:300]}")
        return 0
    return len(data)

def load(source, info=None):
    """Retourne (lignes de Langues sans en-tête, lignes de Expressions avec en-tête)."""
    if source.lower().endswith(".xlsx"):
        from openpyxl import load_workbook
        wb = load_workbook(source, data_only=True)
        grab = lambda n: [["" if c is None else str(c) for c in r] for r in wb[n].iter_rows(values_only=True)]
        return grab("Langues")[1:], grab("Expressions")
    m = re.search(r"/d/([\w-]+)", source)
    sid = m.group(1) if m else source
    if info:
        return _api_rows(sid, "Langues", info)[1:], _api_rows(sid, "Expressions", info)
    return _csv_rows(sid, "Langues")[1:], _csv_rows(sid, "Expressions")

def rate(s):
    """« -10 % », « 5% », « −10% » (Sheets en français) → format attendu par edge-tts."""
    s = re.sub(r"\s", "", str(s)).replace("\u2212", "-").replace(",", ".")
    if not s:
        return "+0%"
    if not s.endswith("%"):
        s += "%"
    return s if s[0] in "+-" else "+" + s

# ---------- Principal ----------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("source", help="ID ou URL du Google Sheets, ou chemin d'un .xlsx")
    ap.add_argument("--credentials", help="fichier JSON du compte de service (ou variable GOOGLE_SERVICE_ACCOUNT_JSON)")
    ap.add_argument("--write-back", action="store_true", help="écrit api/romanisation générés dans le Sheets")
    ap.add_argument("--no-audio", action="store_true")
    ap.add_argument("--force-audio", action="store_true", help="régénère tous les MP3"); ap.add_argument("--out", default="site")
    a = ap.parse_args()
    out = Path(a.out); (out / "audio").mkdir(parents=True, exist_ok=True)

    raw = Path(a.credentials).read_text(encoding="utf-8") if a.credentials else os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "")
    lrows, erows = load(a.source, json.loads(raw) if raw.strip() else None)
    if not erows:
        sys.exit("L'onglet Expressions est vide ou illisible.")
    langs = {}
    for r in lrows:
        r = list(r) + [""] * 7
        if r[0].strip():
            langs[r[0].strip()] = dict(code=r[0].strip(), nom=r[1], local=r[2], systeme=r[3],
                                       voix_nav=r[4], voix_audio=r[5], vitesse=r[6])
    col = {n.strip(): i for i, n in enumerate(erows[0])}
    rows, audio_items, review, seen, writes = [], [], [], set(), []
    for r in erows[1:]:
        def v(n):
            i = col.get(n)
            return str(r[i]).strip() if i is not None and i < len(r) else ""
        id_, lang, text = v("id"), v("langue"), v("texte_local")
        if not (id_ and lang and text):
            continue
        if id_ in seen:
            print(f"ID en double « {id_} » (ligne ignorée) : corrige-le dans le Sheets"); continue
        seen.add(id_)
        if lang not in langs:
            print(f"Langue inconnue « {lang} » ({id_})"); continue
        row = dict(id=id_, langue=lang, categorie=v("catégorie"), texte=text, fr=v("trad_fr"),
                   en=v("trad_en"), api=v("api"), roman=v("romanisation"),
                   commentaire=v("commentaire"), verifie=v("vérifié").lower() == "oui")
        for key, fn, name in (("roman", romanize, "romanisation"), ("api", ipa, "api")):
            if not row[key]:
                g = fn(lang, text)
                if g:
                    row[key] = g
                    review.append([id_, text, key, g])
                    if a.write_back and col.get(name) is not None:
                        writes.append((id_, col[name], g))
        # « à vérifier » tant que la ligne n'est pas marquée vérifiée dans le Sheets
        row["auto"] = [] if row["verifie"] else [k for k in ("api", "roman") if row[k]]
        row["audio"] = f"audio/{id_}.mp3"
        rows.append(row)
        if langs[lang]["voix_audio"]:
            audio_items.append((id_, text, langs[lang]["voix_audio"], rate(langs[lang]["vitesse"])))

    if review:
        with open("a_verifier.csv", "w", newline="", encoding="utf-8") as f:
            csv.writer(f).writerows([["id", "texte", "champ", "valeur_generee"]] + review)
        print(f"{len(review)} valeurs générées à relire → a_verifier.csv")

    if a.write_back and writes:
        if not raw.strip() or a.source.lower().endswith(".xlsx"):
            print("--write-back nécessite un compte de service et un Google Sheets : ignoré")
        else:
            m = re.search(r"/d/([\w-]+)", a.source)
            n = write_back(m.group(1) if m else a.source, json.loads(raw), writes)
            print(f"{n} cellule(s) remplie(s) dans le Sheets")

    if not a.no_audio and audio_items:
        asyncio.run(make_audio(audio_items, out / "audio", out / "audio" / "manifest.json", a.force_audio))

    for row in rows:  # "" si pas d'enregistrement (l'app passe alors directement à la voix du navigateur)
        f = out / "audio" / f"{row['id']}.mp3"
        row["audio"] = f"audio/{f.name}?v={hashlib.sha1(f.read_bytes()).hexdigest()[:8]}" if f.exists() else ""
    (out / "data.json").write_text(json.dumps(
        dict(langues=list(langs.values()), expressions=rows), ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"{len(rows)} expressions → {out / 'data.json'}")

    mf = out / "audio" / "manifest.json"  # supprime les MP3 générés dont l'id n'existe plus (jamais tes enregistrements)
    if mf.exists() and rows:
        m = json.loads(mf.read_text()); ids = {r["id"] for r in rows}
        for old in [k for k in m if k not in ids]:
            (out / "audio" / f"{old}.mp3").unlink(missing_ok=True); del m[old]
            print(f"  audio orphelin supprimé : {old}")
        mf.write_text(json.dumps(m, indent=1))

    sw = out / "sw.js"  # nouvelle version du cache hors ligne dès que données ou audio changent
    if sw.exists():
        mf = out / "audio" / "manifest.json"
        blob = (out / "data.json").read_bytes() + (mf.read_bytes() if mf.exists() else b"")
        h = hashlib.sha1(blob).hexdigest()[:10]
        sw.write_text(re.sub(r'const V="[^"]*"', f'const V="v-{h}"', sw.read_text(encoding="utf-8")), encoding="utf-8")

if __name__ == "__main__":
    main()
