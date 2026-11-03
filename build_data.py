#!/usr/bin/env python3
"""Lit le Google Sheets (ou un .xlsx), complète les phonétiques vides, génère l'audio
et produit data.json + audio/ dans le dossier du site.

Installation (une fois) :
    pip install openpyxl pypinyin pykakasi korean-romanizer pycantonese edge-tts
    + espeak-ng installé sur le système (pour l'API automatique, optionnel)

Utilisation :
    python build_data.py ID_OU_URL_DU_SHEETS       # écrit dans site/
    python build_data.py guide_conversation.xlsx   # même chose depuis un fichier
    options : --no-audio   --out dossier

Le Sheets doit être partagé « Toute personne disposant du lien : lecteur ».
Une valeur saisie à la main n'est JAMAIS écrasée. Les valeurs générées sont listées dans
a_verifier.csv : relis-les et colle-les dans le Sheets, elles deviennent alors les tiennes.
"""

import argparse, asyncio, csv, hashlib, io, json, re, shutil, subprocess, sys, urllib.parse, urllib.request
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
async def make_audio(items, outdir, manifest_path):
    import edge_tts
    manifest = json.loads(manifest_path.read_text()) if manifest_path.exists() else {}
    for id_, text, voice, rate in items:
        h = hashlib.sha1(f"{text}|{voice}|{rate}".encode()).hexdigest()
        f = outdir / f"{id_}.mp3"
        if manifest.get(id_) == h and f.exists():
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
        sys.exit(f"Impossible de lire l'onglet « {tab} » ({e}). Le Sheets est-il partagé "
                 "en lecture avec toute personne ayant le lien ?")
    if txt.lstrip().startswith("<"):
        sys.exit("Google a renvoyé une page web : partage le Sheets en lecture avec toute personne ayant le lien.")
    return list(csv.reader(io.StringIO(txt)))

def load(source):
    """Retourne (lignes de Langues sans en-tête, lignes de Expressions avec en-tête)."""
    if source.lower().endswith(".xlsx"):
        from openpyxl import load_workbook
        wb = load_workbook(source, data_only=True)
        grab = lambda n: [["" if c is None else str(c) for c in r] for r in wb[n].iter_rows(values_only=True)]
        return grab("Langues")[1:], grab("Expressions")
    m = re.search(r"/d/([\w-]+)", source)
    sid = m.group(1) if m else source
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
    ap.add_argument("--no-audio", action="store_true"); ap.add_argument("--out", default="site")
    a = ap.parse_args()
    out = Path(a.out); (out / "audio").mkdir(parents=True, exist_ok=True)

    lrows, erows = load(a.source)
    langs = {}
    for r in lrows:
        r = list(r) + [""] * 7
        if r[0].strip():
            langs[r[0].strip()] = dict(code=r[0].strip(), nom=r[1], local=r[2], systeme=r[3],
                                       voix_nav=r[4], voix_audio=r[5], vitesse=r[6])
    col = {n.strip(): i for i, n in enumerate(erows[0])}
    rows, audio_items, review = [], [], []
    for r in erows[1:]:
        def v(n):
            i = col.get(n)
            return str(r[i]).strip() if i is not None and i < len(r) else ""
        id_, lang, text = v("id"), v("langue"), v("texte_local")
        if not (id_ and lang and text):
            continue
        if lang not in langs:
            print(f"Langue inconnue « {lang} » ({id_})"); continue
        row = dict(id=id_, langue=lang, categorie=v("catégorie"), texte=text, fr=v("trad_fr"),
                   en=v("trad_en"), api=v("api"), roman=v("romanisation"),
                   commentaire=v("commentaire"), verifie=v("vérifié").lower() == "oui")
        auto = []
        for key, fn in (("roman", romanize), ("api", ipa)):
            if not row[key]:
                g = fn(lang, text)
                if g:
                    row[key] = g; auto.append(key); review.append([id_, text, key, g])
        row["auto"] = auto
        row["audio"] = f"audio/{id_}.mp3"
        rows.append(row)
        if langs[lang]["voix_audio"]:
            audio_items.append((id_, text, langs[lang]["voix_audio"], rate(langs[lang]["vitesse"])))

    (out / "data.json").write_text(json.dumps(
        dict(langues=list(langs.values()), expressions=rows), ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"{len(rows)} expressions → {out / 'data.json'}")

    if review:
        with open("a_verifier.csv", "w", newline="", encoding="utf-8") as f:
            csv.writer(f).writerows([["id", "texte", "champ", "valeur_generee"]] + review)
        print(f"{len(review)} valeurs générées à relire → a_verifier.csv")

    if not a.no_audio and audio_items:
        asyncio.run(make_audio(audio_items, out / "audio", out / "audio" / "manifest.json"))

    sw = out / "sw.js"  # nouvelle version du cache hors ligne dès que données ou audio changent
    if sw.exists():
        mf = out / "audio" / "manifest.json"
        blob = (out / "data.json").read_bytes() + (mf.read_bytes() if mf.exists() else b"")
        h = hashlib.sha1(blob).hexdigest()[:10]
        sw.write_text(re.sub(r'const V="[^"]*"', f'const V="v-{h}"', sw.read_text(encoding="utf-8")), encoding="utf-8")

if __name__ == "__main__":
    main()
