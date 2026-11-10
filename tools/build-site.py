"""Construit le portail web d'OpenWrite (accueil/telechargement, fonctionnalites, aide en ligne) dans un dossier.

Utilise localement (tools/publish-update.sh, sous WSL) et par les pipelines GitHub Actions (Windows).

  python tools/build-site.py --out work/site --exe dist/OpenWrite/OpenWrite.exe --msi dist/OpenWrite-0.2.0.msi \\
         (--manifest latest.json | --sign-key cle.pem --version 0.2.0 --notes "...")

- --manifest : reutilise un latest.json deja signe (pipeline du portail) ; version et notes en sont lues.
- --sign-key : signe un nouveau latest.json avec la cle privee (pipeline de l'application, publication locale).
L'executable sert aux captures d'ecran (document de demonstration neutre, fabrique ici : jamais un document du
corpus personnel) et a la verification des liens d'aide.

Deux versions telechargeables, quota Azure respecte (offre Free : 250 Mo pour tout le site) :
- la version courante : son MSI est sur le site, a cote de latest.json (mises a jour automatiques) ;
- la version precedente : lien vers l'installeur de sa release GitHub (rien sur Azure). Elle est trouvee
  automatiquement : la plus haute release publiee (ni brouillon ni pre-version) de numero inferieur a la version
  courante, via `gh api` (GH_TOKEN en CI) ou, a defaut, l'API publique de GitHub. Sans reponse : pas de lien.
- le dossier produit ne doit pas depasser --max-mb (230 Mo par defaut), sinon echec.

La page des fonctionnalites est decrite dans tools/site-template/fonctionnalites.json.
"""
import argparse
import datetime
import glob
import html
import json
import os
import platform
import re
import shutil
import subprocess
import sys
import urllib.request

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE = os.path.join(ROOT, "tools", "site-template")
WSL = "microsoft" in platform.release().lower()
DEFAULT_REPO = "prolice/open-write"
MB = 1048576

# captures de l'aide : nom du fichier -> scene de OpenWrite.exe --screenshot
DOC_SCENES = {
    "accueil": "TabHome", "onglet-accueil": "TabHome", "onglet-insertion": "TabInsert", "onglet-affichage": "TabView",
    "menu-fichier": "file-menu", "a-propos": "about",
}

SWA_CONFIG = {
    # fonction geree /api/crash (tools/site-api) : relais des rapports d'erreur vers les issues GitHub
    "platform": {"apiRuntime": "node:22"},
    "mimeTypes": {".msi": "application/x-msi", ".json": "application/json", ".woff2": "font/woff2"},
    "routes": [
        {"route": "/latest.json", "headers": {"Cache-Control": "no-cache, no-store, must-revalidate"}},
        {"route": "/*.msi", "headers": {"Cache-Control": "public, max-age=31536000, immutable"}},
        {"route": "/fonts/*", "headers": {"Cache-Control": "public, max-age=31536000, immutable"}},
        # pages qui citent la version courante : toujours revalidees (sinon lien vers un MSI retire du site)
        {"route": "/", "headers": {"Cache-Control": "no-cache"}},
        {"route": "/index.html", "headers": {"Cache-Control": "no-cache"}},
        {"route": "/fonctionnalites.html", "headers": {"Cache-Control": "no-cache"}},
    ],
}

MOIS = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre",
        "novembre", "décembre"]

# icones des domaines (trait 24 x 24, degrade de la charte #ow-grad)
ICONS = {
    "docx": "M6 2h9l5 5v15H6zM14 2v6h6M9 13l1.5 5 1.5-4 1.5 4 1.5-5",
    "texte": "M4 7V5h16v2M12 5v14M9 19h6",
    "styles": "M4 6h16M4 12h10M4 18h13",
    "tableaux": "M3 5h18v14H3zM3 10h18M3 15h18M9 5v14M15 5v14",
    "images": "M3 5h18v14H3zM3 16l5-5 4 4 3-3 6 6M15 9a1.5 1.5 0 1 0 3 0 1.5 1.5 0 1 0-3 0",
    "revision": "M4 20h4L19 9l-4-4L4 16zM13 7l4 4",
    "surete": "M12 3l8 3v6c0 4.5-3.4 8-8 9-4.6-1-8-4.5-8-9V6zM8.5 12l2.5 2.5 4.5-5",
    "impression": "M7 9V3h10v6M7 17H4v-7h16v7h-3M7 14h10v7H7z",
    "mises-a-jour": "M20 11A8 8 0 0 0 5.5 6.5L4 8M4 4v4h4M4 13a8 8 0 0 0 14.5 4.5L20 16M20 20v-4h-4",
}
ICON_DEFS = ('<svg width="0" height="0" style="position:absolute" aria-hidden="true" focusable="false"><defs>'
             '<linearGradient id="ow-grad" gradientUnits="userSpaceOnUse" x1="2" y1="2" x2="22" y2="22">'
             '<stop offset="0" stop-color="#087cff"/><stop offset=".55" stop-color="#00c9d5"/>'
             '<stop offset="1" stop-color="#ff8712"/></linearGradient></defs></svg>')
DOWNLOAD_SVG = ('<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" '
                'stroke-linejoin="round" aria-hidden="true"><path d="M12 3v12m-5-5 5 5 5-5M4 16v4h16v-4"/></svg>')


def native(path):
    """Chemin passe a l'executable Windows (sous WSL : C:\\...)."""
    path = os.path.abspath(path)
    if WSL:
        return subprocess.run(["wslpath", "-w", path], capture_output=True, text=True, check=True).stdout.strip()
    return path


def run_exe(exe, *args):
    out = subprocess.run([exe, *args], capture_output=True, text=True, encoding="utf-8", errors="replace")
    if out.returncode != 0:
        sys.exit(f"echec de {os.path.basename(exe)} {args[0]} :\n{out.stdout}\n{out.stderr}")
    return out.stdout.replace("\r", "")


def warn(msg):
    # annotation visible dans le resume de la pipeline GitHub
    print(("::warning::" if os.environ.get("GITHUB_ACTIONS") else "attention : ") + msg)


def vtuple(v):
    m = re.fullmatch(r"v?(\d+)\.(\d+)\.(\d+)", (v or "").strip())
    return tuple(int(x) for x in m.groups()) if m else None


def size_mo(n):
    return f"{n / MB:.1f}".replace(".", ",") + " Mo"


def date_fr(iso):
    try:
        d = datetime.date.fromisoformat((iso or "")[:10])
    except ValueError:
        return ""
    return f"{d.day}{'er' if d.day == 1 else ''} {MOIS[d.month - 1]} {d.year}"


# ------------------------------------------------------------------ version precedente (releases GitHub)

def gh_candidates():
    if os.environ.get("GH_EXE"):
        yield os.environ["GH_EXE"]
    for name in ("gh", "gh.exe"):
        if shutil.which(name):
            yield shutil.which(name)
    if WSL:  # gh pour Windows, hors du PATH de WSL
        yield from sorted(glob.glob("/mnt/c/Users/*/AppData/Local/Programs/gh/bin/gh.exe"))
        yield from glob.glob("/mnt/c/Program Files/GitHub CLI/gh.exe")


def github_releases(repo):
    """Liste des releases du depot, ou None (avec un avertissement) si GitHub ne repond pas."""
    path = f"repos/{repo}/releases?per_page=100"
    for gh in gh_candidates():
        try:
            out = subprocess.run([gh, "api", path], capture_output=True, text=True, encoding="utf-8", timeout=60)
            if out.returncode == 0:
                return json.loads(out.stdout)
        except (OSError, subprocess.TimeoutExpired, json.JSONDecodeError):
            pass
    # repli : API publique (depot public ; jeton facultatif)
    req = urllib.request.Request("https://api.github.com/" + path, headers={"Accept": "application/vnd.github+json",
                                                                            "User-Agent": "openwrite-build-site"})
    token = os.environ.get("GH_TOKEN") or os.environ.get("GITHUB_TOKEN")
    if token:
        req.add_header("Authorization", "Bearer " + token)
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return json.load(r)
    except Exception as ex:  # reseau absent, limite de requetes... : la page sera produite sans version precedente
        warn(f"releases GitHub de {repo} illisibles ({ex}) : pas de lien vers la version precedente")
        return None


def previous_release(releases, current, forced=None):
    """Release publiee la plus recente de version strictement inferieure a `current` et ayant son MSI."""
    cur = vtuple(current)
    best = None
    for r in releases or []:
        v = vtuple(r.get("tag_name"))
        if not v or r.get("draft") or r.get("prerelease") or v >= cur:
            continue
        if forced and v != vtuple(forced):
            continue
        name = "OpenWrite-{}.{}.{}.msi".format(*v)
        asset = next((a for a in r.get("assets", []) if a.get("name") == name), None)
        if not asset or (best and v <= best["v"]):
            continue
        best = {"v": v, "version": "{}.{}.{}".format(*v), "url": asset["browser_download_url"],
                "size": asset.get("size") or 0, "date": r.get("published_at") or "", "page": r.get("html_url") or ""}
    return best


# ------------------------------------------------------------------ pages

def icon(key, large=False):
    if key not in ICONS:
        sys.exit(f"icone inconnue dans fonctionnalites.json : {key}")
    return (f'<span class="ico{" lg" if large else ""}" aria-hidden="true"><svg viewBox="0 0 24 24">'
            f'<path d="{ICONS[key]}"/></svg></span>')


def load_features(version, preview):
    """Sections de fonctionnalites.json, sans les elements livres apres `version` (sauf apercu)."""
    data = json.load(open(os.path.join(TEMPLATE, "fonctionnalites.json"), encoding="utf-8"))
    cur = vtuple(version)
    sections = []
    for s in data["sections"]:
        items = []
        for it in s["elements"]:
            since = vtuple(it.get("depuis"))
            if it.get("depuis") and not since:
                sys.exit(f"fonctionnalites.json : version 'depuis' invalide pour « {it['titre']} »")
            state = "a-venir" if since and since > cur else ("nouveau" if since and since[:2] == cur[:2] else "")
            if state == "a-venir" and not preview:
                continue
            items.append(dict(it, etat=state))
        if items:
            sections.append(dict(s, elements=items))
    return sections


def features_html(sections, captures):
    out, flip = [], False
    for s in sections:
        lis = []
        for it in s["elements"]:
            badge = ""
            if it["etat"] == "nouveau":
                badge = ' <span class="badge">Nouveau</span>'
            elif it["etat"] == "a-venir":
                badge = f' <span class="badge soon">À venir · {html.escape(it["depuis"])}</span>'
            lis.append(f'<li><h3>{html.escape(it["titre"])}{badge}</h3><p>{html.escape(it["texte"])}</p></li>')
        help_link = (f'<a class="help" href="{html.escape(s["aide"])}">Lire l\'aide : {html.escape(s["titre"])} '
                     f'<span aria-hidden="true">→</span></a>') if s.get("aide") else ""
        text = (f'<div class="fsec-text"><div class="fsec-head">{icon(s["icone"], True)}'
                f'<h2 id="t-{s["id"]}">{html.escape(s["titre"])}</h2></div>'
                f'<p class="lead">{html.escape(s["accroche"])}</p><ul class="flist">{"".join(lis)}</ul>{help_link}</div>')
        cap = s.get("capture")
        if cap:
            src, w, h = captures[cap["scene"]]
            # petite capture (menu, fenetre) : jamais agrandie au-dela de sa taille reelle
            narrow = f' style="max-width:{w + 22}px;margin-inline:auto"' if w < 900 else ""
            fig = (f'<figure class="shot-frame"{narrow}><img src="{src}" width="{w}" height="{h}" loading="lazy" '
                   f'alt="{html.escape(cap["alt"])}"><figcaption>{html.escape(cap["alt"])}</figcaption></figure>')
            cls = "fsec flip" if flip else "fsec"
            flip = not flip
            out.append(f'<section class="{cls}" id="{s["id"]}" aria-labelledby="t-{s["id"]}"><div class="container fsec-grid">'
                       f'{text}<div class="fsec-visual">{fig}</div></div></section>')
        else:
            out.append(f'<section class="fsec solo" id="{s["id"]}" aria-labelledby="t-{s["id"]}">'
                       f'<div class="container fsec-grid">{text}</div></section>')
    return "\n".join(out)


def chips_html(sections, page=""):
    return "".join(f'<li><a href="{page}#{s["id"]}">{icon(s["icone"])}{html.escape(s["titre"])}</a></li>'
                   for s in sections)


def notes_html(version, notes):
    # « OpenWrite 0.3.0 : pagination ; ... » -> liste, sans repeter le numero de version
    text = re.sub(r"^\s*OpenWrite\s+" + re.escape(version) + r"\s*[:\-–—]\s*", "", notes or "").strip()
    parts = [p.strip() for p in re.split(r"\s*;\s*|\n+", text) if p.strip()]
    if not parts:
        return "<p>Améliorations et corrections.</p>"
    parts = [p[0].upper() + p[1:] for p in parts]
    if len(parts) == 1:
        return f"<p>{html.escape(parts[0])}.</p>" if not parts[0].endswith((".", "!", "»")) else f"<p>{html.escape(parts[0])}</p>"
    return '<ul class="news-list">' + "".join(f"<li>{html.escape(p)}</li>" for p in parts) + "</ul>"


def previous_html(prev, repo):
    if not prev:
        return "", ""
    v, when = html.escape(prev["version"]), date_fr(prev["date"])
    link = f'<a class="prev-link" href="#version-precedente">Version précédente : {v}</a>'
    all_versions = f"https://github.com/{repo}/releases"
    card = (f'<article class="info-card prev" id="version-precedente"><h2><span class="dot" aria-hidden="true"></span>'
            f'Version précédente</h2><p>OpenWrite {v}{", publiée le " + when if when else ""}, reste disponible si vous '
            f'devez revenir en arrière.</p><a class="prev-dl" href="{html.escape(prev["url"])}" '
            f'aria-label="Télécharger OpenWrite {v}, version précédente, {size_mo(prev["size"])}, depuis GitHub">'
            f'{DOWNLOAD_SVG}OpenWrite {v} · {size_mo(prev["size"])}</a>'
            f'<p class="small" style="margin-top:12px">Téléchargé depuis GitHub. Désinstallez d\'abord la version actuelle '
            f'(Paramètres ▸ Applications) : l\'installeur ne remplace pas une version plus récente. Vos documents ne sont '
            f'pas touchés ; choisissez ensuite « Ignorer cette version » dans le bandeau de mise à jour. '
            f'<a class="more" href="{all_versions}">Toutes les versions</a></p></article>')
    return link, card


def fill(name, values):
    page = open(os.path.join(TEMPLATE, name), encoding="utf-8").read()
    for k, v in values.items():
        page = page.replace("{{" + k + "}}", v)
    if "{{" in page:
        sys.exit(f"champ non rempli dans {name} : " + ", ".join(sorted(set(re.findall(r"\{\{\w+\}\}", page)))))
    return page


def png_size(path):
    with open(path, "rb") as f:
        head = f.read(24)
    return int.from_bytes(head[16:20], "big"), int.from_bytes(head[20:24], "big")


def dir_size(path):
    return sum(os.path.getsize(os.path.join(d, f)) for d, _, files in os.walk(path) for f in files)


# ------------------------------------------------------------------ document de demonstration (neutre)

DEMO_PARAS = [
    ("Title", "Compte rendu de réunion"),
    ("Subtitle", "Association des jardins partagés — exemple de document"),
    ("Heading1", "Ordre du jour"),
    ("ListParagraph", "Bilan de la saison et récolte d'automne"),
    ("ListParagraph", "Calendrier des ateliers de l'hiver"),
    ("ListParagraph", "Questions diverses"),
    ("Heading1", "Bilan de la saison"),
    (None, "La saison s'est bien déroulée : les parcelles ont été entretenues régulièrement et la récolte de "
           "courges a dépassé les attentes. Les bénévoles remercient chaleureusement les familles qui ont participé "
           "aux journées d'arrosage pendant l'été."),
    (None, "Le composteur collectif sera vidé au printemps ; le terreau obtenu sera partagé entre les parcelles."),
    ("Heading1", "Ateliers de l'hiver"),
]
DEMO_TABLE = [("Date", "Atelier", "Animation"), ("Samedi 6 décembre", "Taille des arbres fruitiers", "Équipe verger"),
              ("Samedi 10 janvier", "Semis sous abri", "Équipe potager"), ("Samedi 7 février", "Fabrication de nichoirs", "Équipe bricolage")]


def write_demo_docx(path):
    """.docx de demonstration ecrit ici (contenu fictif) : les captures n'utilisent jamais de vrai document."""
    import zipfile
    W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

    def run(text, bold=False):
        rpr = "<w:rPr><w:b/></w:rPr>" if bold else ""
        return f'<w:r>{rpr}<w:t xml:space="preserve">{html.escape(text, quote=False)}</w:t></w:r>'

    def para(style, text, bold=False):
        ppr = f'<w:pPr><w:pStyle w:val="{style}"/>' + ('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' if style == "ListParagraph" else "") + "</w:pPr>" if style else ""
        return f"<w:p>{ppr}{run(text, bold)}</w:p>"

    rows = "".join("<w:tr>" + "".join(f'<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>{para(None, c, i == 0)}</w:tc>' for c in r) + "</w:tr>"
                   for i, r in enumerate(DEMO_TABLE))
    border = "".join(f'<w:{b} w:val="single" w:sz="4" w:color="9DB7E8"/>' for b in ("top", "left", "bottom", "right", "insideH", "insideV"))
    body = "".join(para(st, t) for st, t in DEMO_PARAS)
    body += (f'<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblBorders>{border}</w:tblBorders></w:tblPr>'
             f'<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>{rows}</w:tbl>')
    body += para(None, "Prochaine réunion : premier samedi du mois, à 10 heures, au local des jardins.")
    body += ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" '
             'w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>')
    document = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}</w:body></w:document>'

    def style(sid, name, size, color=None, bold=False, before=0, after=160, based="Normal"):
        rpr = (f'<w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/>' if sid != "Normal" else "") + \
              ("<w:b/>" if bold else "") + (f'<w:color w:val="{color}"/>' if color else "") + f'<w:sz w:val="{size}"/>'
        basedon = f'<w:basedOn w:val="{based}"/>' if sid != "Normal" else ""
        return (f'<w:style w:type="paragraph" w:styleId="{sid}"><w:name w:val="{name}"/>{basedon}'
                f'<w:pPr><w:spacing w:before="{before}" w:after="{after}"/></w:pPr><w:rPr>{rpr}</w:rPr></w:style>')

    styles = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr>'
              '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:lang w:val="fr-FR"/>'
              '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr>'
              '</w:pPrDefault></w:docDefaults>'
              + style("Normal", "Normal", 22)
              + style("Title", "Title", 52, "111A47", before=0, after=60)
              + style("Subtitle", "Subtitle", 26, "087CFF", after=320)
              + style("Heading1", "heading 1", 30, "087CFF", bold=True, before=320, after=120)
              + '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/>'
                '<w:pPr><w:spacing w:after="60"/><w:ind w:left="720"/></w:pPr></w:style></w:styles>')
    numbering = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering {W}><w:abstractNum w:abstractNumId="0">'
                 '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/>'
                 '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>'
                 '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>')
    R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    parts = {
        "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
            '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
            '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
            '</Types>',
        "_rels/.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{R}/officeDocument" Target="word/document.xml"/></Relationships>',
        "word/_rels/document.xml.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{R}/styles" Target="styles.xml"/>'
            f'<Relationship Id="rId2" Type="{R}/numbering" Target="numbering.xml"/>'
            f'<Relationship Id="rId3" Type="{R}/settings" Target="settings.xml"/></Relationships>',
        "word/document.xml": document, "word/styles.xml": styles, "word/numbering.xml": numbering,
        # mode Word 2013 et suivants : pas de « [Mode de compatibilité] » dans la capture
        "word/settings.xml": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:compat>'
            '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
            '</w:compat></w:settings>',
    }
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)


# ------------------------------------------------------------------ construction

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True)
    ap.add_argument("--exe", required=True, help="OpenWrite.exe autonome (captures, liens d'aide, signature)")
    ap.add_argument("--msi", required=True)
    ap.add_argument("--manifest", help="latest.json deja signe a reutiliser")
    ap.add_argument("--sign-key", help="cle privee PEM pour signer un nouveau latest.json")
    ap.add_argument("--version")
    ap.add_argument("--notes", default="")
    ap.add_argument("--repo", default=os.environ.get("GITHUB_REPOSITORY") or DEFAULT_REPO,
                    help="depot GitHub des releases (version precedente)")
    ap.add_argument("--previous", help="X.Y.Z : version precedente imposee ; 'none' : pas de version precedente")
    ap.add_argument("--max-mb", type=float, default=230, help="taille maximale du site (quota Azure Free : 250 Mo)")
    ap.add_argument("--apercu-a-venir", action="store_true",
                    help="relecture : affiche aussi les fonctionnalites pas encore livrees (jamais pour publier)")
    a = ap.parse_args()
    if not (a.manifest or a.sign_key):
        sys.exit("--manifest ou --sign-key est requis")

    out = os.path.abspath(a.out)
    shutil.rmtree(out, ignore_errors=True)
    os.makedirs(out)
    for item in ("assets", "fonts"):
        shutil.copytree(os.path.join(TEMPLATE, item), os.path.join(out, item))
    shutil.copy(os.path.join(TEMPLATE, "styles.css"), out)
    msi_name = os.path.basename(a.msi)
    shutil.copy(a.msi, os.path.join(out, msi_name))
    exe = os.path.abspath(a.exe)

    # --- manifeste de mise a jour (url relative au flux : le MSI courant, a la racine)
    manifest_path = os.path.join(out, "latest.json")
    if a.manifest:
        shutil.copy(a.manifest, manifest_path)
    else:
        if not a.version:
            sys.exit("--version est requis avec --sign-key")
        print(run_exe(exe, "--sign-update", native(os.path.join(out, msi_name)), a.version, a.notes,
                      native(a.sign_key), native(manifest_path)).strip())
    manifest = json.load(open(manifest_path, encoding="utf-8-sig"))
    version, notes = manifest["version"], manifest.get("notes") or ""
    if manifest.get("url") != msi_name:
        sys.exit(f"latest.json pointe vers {manifest.get('url')}, l'installeur fourni est {msi_name}")
    if not vtuple(version):
        sys.exit(f"version invalide dans latest.json : {version}")
    check = run_exe(exe, "--verify-update", native(manifest_path))
    if "INVALIDE" in check or "NON CONFORME" in check:
        sys.exit("latest.json invalide : " + check)
    print(check.strip())

    # --- version precedente : installeur de la release GitHub (aucun octet sur Azure)
    prev = None
    if (a.previous or "").lower() != "none":
        prev = previous_release(github_releases(a.repo), version, a.previous)
        if a.previous and not prev:
            sys.exit(f"release v{a.previous} introuvable (ou sans OpenWrite-{a.previous}.msi) dans {a.repo}")
    print(f"version precedente : {prev['version'] + ' -> ' + prev['url'] if prev else 'aucune'}")

    # --- aide en ligne (avant les pages : ses captures servent aussi aux fonctionnalites)
    subprocess.run([sys.executable, os.path.join(ROOT, "tools", "build-docs.py"), out, version], check=True)
    demo_path = os.path.join(out, "_demo", "demonstration.docx")
    write_demo_docx(demo_path)
    demo = native(demo_path)
    for name, scene in DOC_SCENES.items():
        png = os.path.join(out, "aide", "img", f"{name}.png")
        run_exe(exe, "--screenshot", demo, native(png), "1402", "900", scene)
        if not os.path.exists(png):
            sys.exit(f"capture manquante : {name}")
    anchors_file = os.path.join(out, "aide", "ancres.txt")
    anchors = set(open(anchors_file, encoding="utf-8").read().split())
    topics = set(run_exe(exe, "--help-topics").split())
    missing = sorted(topics - anchors)
    if missing:
        sys.exit("liens d'aide de l'application sans section correspondante :\n  " + "\n  ".join(missing))
    os.remove(anchors_file)
    print(f"aide : {len([f for f in os.listdir(os.path.join(out, 'aide')) if f.endswith('.html')])} pages, "
          f"{len(DOC_SCENES)} captures, {len(topics)} liens contextuels verifies")

    # --- page des fonctionnalites (captures de l'aide reutilisees, autres scenes dans assets/captures)
    sections = load_features(version, a.apercu_a_venir)
    captures = {}
    for s in sections:
        scene = (s.get("capture") or {}).get("scene")
        if not scene or scene in captures:
            continue
        doc_name = next((n for n, sc in DOC_SCENES.items() if sc == scene), None)
        if doc_name:
            rel = f"aide/img/{doc_name}.png"
        else:
            rel = f"assets/captures/{scene}.png"
            os.makedirs(os.path.join(out, "assets", "captures"), exist_ok=True)
            run_exe(exe, "--screenshot", demo, native(os.path.join(out, rel)), "1402", "900", scene)
        if not os.path.exists(os.path.join(out, rel)):
            sys.exit(f"capture manquante : {scene}")
        captures[scene] = (rel, *png_size(os.path.join(out, rel)))
    for s in sections:
        if s.get("aide") and s["aide"].split("#")[0].replace("aide/", "") not in os.listdir(os.path.join(out, "aide")):
            sys.exit(f"fonctionnalites.json : page d'aide introuvable {s['aide']}")
    n_items = sum(len(s["elements"]) for s in sections)
    open(os.path.join(out, "fonctionnalites.html"), "w", encoding="utf-8").write(fill("fonctionnalites.html", {
        "VERSION": html.escape(version), "CHIPS": chips_html(sections), "SECTIONS": features_html(sections, captures),
        "ICON_DEFS": ICON_DEFS}))
    print(f"fonctionnalites : {len(sections)} domaines, {n_items} fonctions"
          + (" (APERCU : fonctions a venir incluses, ne pas publier)" if a.apercu_a_venir else ""))

    # --- accueil (telechargement)
    run_exe(exe, "--screenshot", demo, native(os.path.join(out, "assets", "apercu-openwrite.png")), "1402", "948", "TabHome")
    shutil.rmtree(os.path.join(out, "_demo"))
    size = size_mo(os.path.getsize(a.msi))
    when = date_fr(manifest.get("published"))
    prev_link, prev_card = previous_html(prev, a.repo)
    open(os.path.join(out, "index.html"), "w", encoding="utf-8").write(fill("index.html", {
        "VERSION": html.escape(version), "MSI": html.escape(msi_name), "SIZE": size,
        "DATE_BADGE": f" · {when}" if when else "", "NOTES_BLOCK": notes_html(version, notes),
        "PREVIOUS_LINK": prev_link, "PREVIOUS_CARD": prev_card,
        "CHIPS": chips_html(sections, "fonctionnalites.html"), "ICON_DEFS": ICON_DEFS}))
    print(f"accueil : version {version}, {size}")

    json.dump(SWA_CONFIG, open(os.path.join(out, "staticwebapp.config.json"), "w"), indent=2)

    # --- budget de taille (Azure Static Web Apps Free : 250 Mo pour tout le site)
    heavy = [f for f in os.listdir(out) if f.lower().endswith((".msi", ".exe"))]
    if heavy != [msi_name]:
        sys.exit(f"seul l'installeur courant doit etre sur le site, trouve : {heavy}")
    total = dir_size(out)
    parts = {"installeur": os.path.getsize(os.path.join(out, msi_name)), "aide": dir_size(os.path.join(out, "aide")),
             "images et polices": dir_size(os.path.join(out, "assets")) + dir_size(os.path.join(out, "fonts"))}
    parts["pages et divers"] = total - sum(parts.values())
    print("taille : " + ", ".join(f"{k} {size_mo(v)}" for k, v in parts.items())
          + f" -> total {size_mo(total)} / limite {size_mo(a.max_mb * MB)} (quota Azure 250 Mo)")
    if total > a.max_mb * MB:
        sys.exit(f"site trop volumineux : {size_mo(total)} > {size_mo(a.max_mb * MB)} (quota Azure Free 250 Mo)")
    print(f"portail pret : {out}")


if __name__ == "__main__":
    main()
