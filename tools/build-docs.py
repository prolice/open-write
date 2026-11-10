"""Genere la documentation en ligne d'OpenWrite (dossier aide/ du site).

Usage : python3 -I tools/build-docs.py DOSSIER_SITE [VERSION]
Produit DOSSIER_SITE/aide/*.html, aide.css, recherche.js. Les captures sont attendues dans
DOSSIER_SITE/aide/img/<scene>.png (produites par OpenWrite.exe --screenshot, voir build-site.py).

Les ancres (#...) sont utilisees par les liens contextuels de l'application (Help/HelpService.cs) : ne pas les renommer.
"""
import html
import json
import os
import re
import sys

OUT = sys.argv[1] if len(sys.argv) > 1 else "work/site"
VERSION = sys.argv[2] if len(sys.argv) > 2 else ""


def k(*keys):
    """Touches clavier : k("Ctrl", "S") -> <kbd>Ctrl</kbd>+<kbd>S</kbd>"""
    return "+".join(f"<kbd>{html.escape(x)}</kbd>" for x in keys)


def tip(text):
    return f'<div class="callout tip"><strong>Astuce</strong><p>{text}</p></div>'


def warn(text):
    return f'<div class="callout warn"><strong>Attention</strong><p>{text}</p></div>'


def shot(scene, alt, caption=None):
    cap = f"<figcaption>{caption}</figcaption>" if caption else ""
    return f'<figure class="shot"><img src="img/{scene}.png" alt="{html.escape(alt)}" loading="lazy">{cap}</figure>'


def steps(*items):
    return "<ol class=\"steps\">" + "".join(f"<li>{i}</li>" for i in items) + "</ol>"


def table(rows, head=("Action", "Raccourci")):
    h = "".join(f"<th>{c}</th>" for c in head)
    b = "".join("<tr>" + "".join(f"<td>{c}</td>" for c in r) + "</tr>" for r in rows)
    return f'<div class="table-wrap"><table><thead><tr>{h}</tr></thead><tbody>{b}</tbody></table></div>'


# ------------------------------------------------------------------ contenu
# (fichier, titre, description, icone, sections[(ancre, titre, html)])
PAGES = [
    ("index.html", "Bienvenue dans OpenWrite", "Prendre en main OpenWrite en quelques minutes.", "🏠", [
        ("premiers-pas", "Premiers pas", f"""
<p>OpenWrite est un traitement de texte pour Windows qui <strong>ouvre et enregistre directement les fichiers Word</strong>
(.docx, .docm, .dotx). Il n'a pas de format à lui : vos documents restent des fichiers Word, que vos correspondants
ouvrent comme d'habitude.</p>
{steps("Lancez OpenWrite depuis le menu Démarrer.",
       f"Ouvrez un document avec <strong>Fichier ▸ Ouvrir…</strong> ({k('Ctrl', 'O')}), ou glissez un fichier .docx sur la fenêtre.",
       "Le document s'affiche page par page, mis en page comme dans Word : polices, styles, listes, tableaux, en-têtes et pieds, notes.",
       f"Enregistrez avec {k('Ctrl', 'S')}. Tout ce qu'OpenWrite ne sait pas encore modifier est conservé tel quel.")}
{shot("accueil", "Fenêtre principale d'OpenWrite", "La fenêtre d'OpenWrite : ruban en haut, plan à gauche, mise en forme à droite.")}
"""),
        ("version-actuelle", "Ce que fait cette version", """
<p>OpenWrite avance par versions, chacune ajoutant un pan de Word. Aujourd'hui, il <strong>affiche un document page
par page</strong>, mis en page comme Word le fait (mêmes coupures de page sur nos documents de test dans la grande
majorité des cas), permet de <strong>cocher les cases des formulaires</strong>, d'annuler, de rétablir et
d'<strong>enregistrer sans perte</strong>. La saisie du texte et la mise en forme arrivent avec les versions suivantes :
chaque groupe du ruban indique la version qui l'activera.</p>
<p>Voir la <a href="feuille-de-route.html">feuille de route</a>.</p>
"""),
        ("aide-contextuelle", "Obtenir de l'aide", f"""
<p>Appuyez sur {k('F1')} à tout moment : OpenWrite ouvre la page d'aide de ce que vous êtes en train de faire.</p>
"""),
        ("sommaire", "Sommaire", """
<div class="cards">
<a class="card" href="interface.html"><b>Interface et ruban</b><span>Fichier, onglets, volets, barre d'état</span></a>
<a class="card" href="documents.html"><b>Documents</b><span>Ouvrir, enregistrer, annuler, copies de secours</span></a>
<a class="card" href="feuille-de-route.html"><b>Feuille de route</b><span>Ce que chaque version apporte</span></a>
<a class="card" href="mises-a-jour.html"><b>Mises à jour</b><span>Automatiques et sécurisées</span></a>
<a class="card" href="raccourcis.html"><b>Raccourcis clavier</b><span>Tous les raccourcis</span></a>
<a class="card" href="faq.html"><b>Questions fréquentes</b><span>Compatibilité Word, sauvegardes…</span></a>
</div>
"""),
    ]),
    ("interface.html", "Interface et ruban", "Le ruban, le menu Fichier, les volets et la barre d'état.", "🧭", [
        ("ruban", "Le ruban", f"""
<p>Le ruban reprend les onglets de Word, dans le même ordre : Accueil, Insertion, Conception, Mise en page,
Références, Publipostage, Révision, Affichage. Chaque groupe porte un petit badge avec la <strong>version
d'OpenWrite</strong> qui l'activera ; les groupes déjà actifs n'en ont pas.</p>
{shot("onglet-insertion", "Onglet Insertion du ruban", "Les groupes à venir indiquent leur version.")}
"""),
        ("fichier", "Le menu Fichier", f"""
<p>Le bouton <strong>Fichier</strong> ouvre les commandes du document (ouvrir, enregistrer, informations), les
documents récents et les réglages : mises à jour automatiques, copies de secours, version précédente (.bak) et
rapports d'erreur anonymes.</p>
{shot("menu-fichier", "Menu Fichier d'OpenWrite")}
"""),
        ("accueil", "L'onglet Accueil", f"""
<p>Presse-papiers, police, paragraphe, styles et recherche : comme dans Word. Le groupe <strong>Styles</strong>
montre une galerie des styles du document ; le groupe <strong>Paragraphe</strong> les puces, les numéros et la liste de
titres. L'onglet <strong>Insertion</strong> propose déjà Lien, Symbole et Saut de page ; l'onglet
<strong>Mise en page</strong>, Retraits et Espacement (boîte Paragraphe).</p>
{shot("onglet-accueil", "Onglet Accueil du ruban")}
"""),
        ("volets", "Les volets", f"""
<p>À gauche, le volet de navigation a deux onglets : <strong>Pages</strong>, les miniatures de toutes les pages (un clic
sur une miniature affiche la page), et <strong>Titres</strong>, le plan du document. À droite, le volet
<strong>Mise en forme</strong> montre et modifie la mise en forme du point d'insertion : police, taille, gras,
italique, souligné, barré, alignement, interligne, et ouvre les boîtes Police et Paragraphe. Sous le ruban, la
<strong>règle</strong> montre les marges et les retraits du paragraphe : faites glisser les triangles pour les changer,
cliquez pour poser un taquet de tabulation, tirez un taquet hors de la règle pour le retirer. Les deux se masquent depuis l'onglet <strong>Affichage</strong>, qui règle aussi le zoom
({k('Ctrl')} + molette).</p>
{shot("onglet-affichage", "Onglet Affichage du ruban")}
"""),
    ]),
    ("documents.html", "Documents", "Ouvrir, enregistrer, annuler et récupérer vos documents.", "📄", [
        ("ouvrir", "Ouvrir un document", f"""
<p><strong>Fichier ▸ Ouvrir…</strong> ({k('Ctrl', 'O')}) accepte les documents Word <code>.docx</code>,
<code>.docm</code>, <code>.dotx</code> et <code>.dotm</code>. Vous pouvez aussi glisser un fichier sur la fenêtre ou
le choisir dans les documents récents. Les fichiers que Word signale comme endommagés s'ouvrent généralement.</p>
<p>Un document enregistré par une ancienne version de Word affiche « [Mode de compatibilité] » dans le titre, comme
dans Word.</p>
"""),
        ("enregistrer", "Enregistrer", f"""
<p>{k('Ctrl', 'S')} enregistre ; {k('Ctrl', 'Maj', 'S')} ou {k('F12')} enregistre sous un autre nom. Le format
reste celui du document : un <code>.docm</code> reste un <code>.docm</code>.</p>
<p>OpenWrite ne réécrit que les parties du fichier qu'il a modifiées ; tout le reste (images, objets, réglages
qu'il ne connaît pas encore) est recopié <strong>à l'identique</strong>. La version précédente est gardée à côté,
avec l'extension <code>.bak</code> (réglage dans le menu Fichier).</p>
"""),
        ("ecrire", "Écrire et mettre en forme", f"""
<p>Cliquez dans la page pour placer le point d'insertion, puis tapez. {k('Entrée')} commence un nouveau paragraphe
(après un titre, le paragraphe suivant prend le style prévu, en général « Normal »), {k('Maj', 'Entrée')} passe à la
ligne dans le même paragraphe, {k('Ctrl', 'Entrée')} insère un saut de page et {k('Tab')} une tabulation.
{k('Retour arrière')} et {k('Suppr')} effacent un caractère ; au début ou à la fin d'un paragraphe, ils le
réunissent au précédent ou au suivant.</p>
<p>Pour sélectionner : glissez avec la souris, double-cliquez sur un mot, triple-cliquez sur un paragraphe, ou
maintenez {k('Maj')} avec les flèches. Le texte tapé remplace la sélection.</p>
<p>{k('Ctrl', 'G')} met en gras, {k('Ctrl', 'I')} en italique, {k('Ctrl', 'U')} souligne — la sélection, ou le mot
où se trouve le point d'insertion ; entre deux mots, la mise en forme s'applique au texte que vous tapez ensuite.
{k('Ctrl', 'E')} centre le paragraphe, {k('Ctrl', 'J')} le justifie, {k('Ctrl', 'Maj', 'G')} et
{k('Ctrl', 'Maj', 'D')} l'alignent à gauche et à droite. Les raccourcis de Word en anglais ({k('Ctrl', 'B')},
{k('Ctrl', 'L')}, {k('Ctrl', 'R')}) fonctionnent aussi.</p>
<p>{k('Ctrl', 'D')} ouvre la boîte <strong>Police</strong> (police, style, taille, soulignement, couleur, barré,
exposant, indice, petites majuscules, masqué) ; la boîte <strong>Paragraphe</strong> (bouton du volet, double-clic sur
la règle) règle l'alignement, les retraits, les espacements, l'interligne et les enchaînements.
{k('Ctrl', 'Espace')} et {k('Ctrl', 'Q')} retirent la mise en forme ajoutée aux caractères et au paragraphe.</p>
<p>{k('Ctrl', 'Maj', '8')} affiche les marques de mise en forme : ¶ en fin de paragraphe, · pour les espaces, → pour les
tabulations, sauts de page.</p>
"""),
        ("presse-papiers", "Copier et coller", f"""
<p>{k('Ctrl', 'C')}, {k('Ctrl', 'X')} et {k('Ctrl', 'V')} copient, coupent et collent <strong>avec la mise en
forme</strong> : entre deux endroits du document (styles, tableaux, champs compris), depuis Word, LibreOffice ou un
navigateur, et vers eux (gras, italique, police, taille, couleur, alignement). {k('Ctrl', 'Maj', 'V')} colle le texte
seul. Un tableau compris dans la sélection est copié ou supprimé en entier.</p>
{warn("Pas encore modifiables : les en-têtes et pieds de page (version 0.6), les notes (version 1.2), et une sélection sur plusieurs cellules d'un même tableau (version 0.7). Les images d'un autre programme ne sont pas encore collées.")}
"""),
        ("styles", "Styles", f"""
<p>Un style donne d'un coup la police, la taille, les espacements… d'un paragraphe ; le modifier change tous les
paragraphes qui l'utilisent. Appliquez-le avec la <strong>galerie</strong> de l'onglet Accueil, la liste
<strong>Style</strong> du volet Mise en forme, ou les raccourcis de Word : {k('Ctrl', 'Alt', '1')}, {k('2')}, {k('3')}
pour Titre 1, 2, 3, {k('Ctrl', 'Maj', 'N')} pour Normal. Les styles intégrés de Word (Titre 1, Citation, Élevé…) sont
ajoutés au document la première fois qu'ils servent, comme dans Word.</p>
<p><strong>Accueil ▸ Styles</strong> liste tous les styles : appliquer, <strong>modifier</strong> (nom, style de base,
style suivant, boîtes Police et Paragraphe), <strong>mettre à jour selon la sélection</strong> (la mise en forme du
paragraphe passe dans son style), <strong>nouveau style</strong> à partir de la sélection. Le bouton
<strong>Modifier</strong> ouvre directement le style du paragraphe.</p>
"""),
        ("listes", "Puces et numéros", f"""
<p><strong>Puces</strong> ({k('Ctrl', 'Maj', 'L')}) et <strong>Numéros</strong> mettent les paragraphes sélectionnés
en liste, ou les en sortent. Au début d'un élément, {k('Tab')} et {k('Maj', 'Tab')} changent son niveau ;
{k('Entrée')} sur un élément vide et {k('Retour arrière')} au début d'un élément terminent la liste. Clic droit sur un
élément : <strong>Recommencer à 1</strong>, <strong>Continuer la numérotation</strong>.</p>
<p><strong>Titres 1.1</strong> numérote les titres du document : Titre 1 en « 1. », Titre 2 en « 1.1 », Titre 3 en
« 1.1.1 ». La numérotation suit les titres ajoutés ensuite.</p>
"""),
        ("themes", "Thèmes, polices et couleurs", """
<p>L'onglet <strong>Création</strong> change le thème du document : <strong>Polices</strong> (titres et corps du
texte), <strong>Couleurs</strong> (titres, accents, liens) ou les deux avec <strong>Thèmes</strong>. Seul le texte qui
utilise les polices et couleurs du thème change — les titres et citations notamment ; une police ou une couleur choisie
à la main reste. Dans la liste des polices du volet, « +Titres » et « +Corps » sont les polices du thème.</p>
"""),
        ("rechercher", "Rechercher et remplacer", f"""
<p>{k('Ctrl', 'F')} cherche un texte dans le document à partir du point d'insertion, {k('Ctrl', 'H')} le remplace.
Options : respecter la casse et les accents, mot entier. Le texte remplacé garde sa mise en forme ;
<strong>Remplacer tout</strong> s'annule d'un seul {k('Ctrl', 'Z')}.</p>
"""),
        ("liens", "Liens hypertextes", f"""
<p>{k('Ctrl', 'K')} transforme la sélection en lien (adresse web, adresse de courriel ou signet du document), ou
modifie et supprime le lien sous le point d'insertion. Survolez un lien pour voir son adresse ;
{k('Ctrl')} + clic l'ouvre. Les adresses qui ne sont ni web ni courriel demandent une confirmation.</p>
<p><strong>Insertion ▸ Symbole</strong> insère les caractères spéciaux : €, ©, tirets, espaces insécables
({k('Ctrl', 'Maj', 'Espace')}), flèches, lettres grecques, cases à cocher…</p>
"""),
        ("annuler", "Annuler et rétablir", f"""
<p>{k('Ctrl', 'Z')} annule la dernière modification, {k('Ctrl', 'Y')} la rétablit. Le titre de la fenêtre affiche
« • » tant que le document a des modifications non enregistrées ; annuler jusqu'à l'état enregistré le fait
disparaître.</p>
"""),
        ("cases", "Cases à cocher des formulaires", """
<p>Les cases à cocher des formulaires Word se cochent et se décochent d'un clic, directement dans le document.</p>
"""),
        ("secours", "Copies de secours", """
<p>Toutes les 5 minutes, OpenWrite garde une copie du document en cours de modification. Après une fermeture
imprévue, il propose de la récupérer au démarrage suivant. Si le fichier est modifié par un autre programme pendant
que vous travaillez, OpenWrite vous prévient avant d'écraser ces changements.</p>
"""),
        ("pages", "La vue des pages", f"""
<p>Le document s'affiche page par page, avec ses marges, ses en-têtes et pieds de page (numéros de page calculés),
ses notes de bas de page, ses tableaux (fond et bordures des cellules, rangées coupées entre deux pages, en-têtes
répétés), ses images en ligne et la numérotation de ses listes. La barre d'état indique la page affichée et le nombre
de pages ; le zoom se règle avec {k('Ctrl')} + molette ou le curseur en bas à droite.</p>
<p>Une police absente de l'ordinateur est remplacée, comme Word le fait, par la police de remplacement indiquée dans le
document ou par une police aux mêmes dimensions ; la barre d'état signale les remplacements.</p>
{warn("Pas encore affichés : les images et formes flottantes, les zones de texte et les colonnes (versions suivantes). Le texte qu'elles contiennent n'est pas perdu : il est conservé à l'enregistrement.")}
"""),
        ("informations", "Informations sur le document", """
<p><strong>Fichier ▸ Informations</strong> résume le document : nombre de paragraphes, tableaux, sections, images,
notes, commentaires, révisions, champs et contrôles de contenu, mode de compatibilité, langues et polices utilisées.</p>
"""),
    ]),
    ("feuille-de-route.html", "Feuille de route", "Ce que chaque version d'OpenWrite apporte.", "🗺️", [
        ("versions", "Les versions", table([
            ("0.1", "Socle : fenêtre, ruban, ouverture en lecture seule, mises à jour"),
            ("0.2", "Modèle de document sans perte, mise en forme fidèle, cases à cocher, annuler, enregistrer"),
            ("0.3", "Vue des pages : pagination comme dans Word, en-têtes et pieds, notes, miniatures"),
            ("0.4", "Saisie et mise en forme directe, presse-papiers"),
            ("0.5", "Styles, listes et thèmes : galerie, modifier et créer des styles, puces et numéros, polices et couleurs du thème"),
            ("0.6", "Mise en page, sections, en-têtes et pieds de page"),
            ("0.7", "Tableaux"),
            ("1.0", "Images, formes et zones de texte"),
            ("1.1", "Orthographe, commentaires, suivi des modifications"),
            ("1.2", "Notes, table des matières, renvois, index"),
            ("1.3", "Champs, blocs de construction, contrôles de contenu"),
            ("1.4", "Publipostage"),
            ("1.5", "PDF, impression, RTF, ODT, HTML"),
        ], head=("Version", "Contenu"))),
    ]),
    ("mises-a-jour.html", "Mises à jour", "Recevoir automatiquement les nouvelles versions.", "🔄", [
        ("automatiques", "Mises à jour automatiques", """
<p>Au démarrage, OpenWrite vérifie s'il existe une nouvelle version (réglage dans le menu Fichier). Si oui, un
bandeau propose de l'installer. L'installeur est <strong>signé</strong> : OpenWrite vérifie la signature et
l'empreinte du fichier avant de le lancer, et refuse tout fichier modifié.</p>
<p><strong>Fichier ▸ Rechercher des mises à jour…</strong> lance la vérification à la demande.</p>
"""),
        ("installation", "Installation", """
<p>L'installeur s'installe pour votre compte, sans droits administrateur, dans
<code>%LOCALAPPDATA%\\Programs\\OpenWrite</code>. Il ajoute OpenWrite au menu Démarrer et à « Ouvrir avec » pour les
fichiers Word, sans remplacer Word comme programme par défaut.</p>
"""),
    ]),
    ("raccourcis.html", "Raccourcis clavier", "Tous les raccourcis d'OpenWrite.", "⌨️", [
        ("general", "Général", table([
            ("Ouvrir", k("Ctrl", "O")), ("Enregistrer", k("Ctrl", "S")),
            ("Enregistrer sous", k("Ctrl", "Maj", "S") + " ou " + k("F12")),
            ("Annuler", k("Ctrl", "Z")), ("Rétablir", k("Ctrl", "Y")),
            ("Zoom", k("Ctrl") + " + molette"), ("Aide", k("F1")), ("Quitter", k("Alt", "F4")),
        ])),
        ("saisie", "Saisie et déplacements", table([
            ("Nouveau paragraphe", k("Entrée")), ("Saut de ligne", k("Maj", "Entrée")), ("Saut de page", k("Ctrl", "Entrée")),
            ("Mot précédent, suivant", k("Ctrl", "←") + " " + k("Ctrl", "→")),
            ("Début, fin de ligne", k("Origine") + " " + k("Fin")), ("Début, fin du document", k("Ctrl", "Origine") + " " + k("Ctrl", "Fin")),
            ("Page précédente, suivante", k("Pg.préc") + " " + k("Pg.suiv")),
            ("Effacer le mot précédent, suivant", k("Ctrl", "Retour arrière") + " " + k("Ctrl", "Suppr")),
            ("Étendre la sélection", k("Maj") + " + déplacement"), ("Tout sélectionner", k("Ctrl", "A")),
            ("Copier, couper, coller", k("Ctrl", "C") + " " + k("Ctrl", "X") + " " + k("Ctrl", "V")),
            ("Coller le texte seul", k("Ctrl", "Maj", "V")),
            ("Rechercher, remplacer", k("Ctrl", "F") + " " + k("Ctrl", "H")),
            ("Lien hypertexte", k("Ctrl", "K")), ("Suivre un lien", k("Ctrl") + " + clic"),
            ("Cellule suivante, précédente (tableau)", k("Tab") + " " + k("Maj", "Tab")),
            ("Espace insécable, trait d'union insécable", k("Ctrl", "Maj", "Espace") + " " + k("Ctrl", "Maj", "-")),
            ("Marques de mise en forme", k("Ctrl", "Maj", "8")),
        ])),
        ("styles-listes", "Styles et listes", table([
            ("Titre 1, 2, 3", k("Ctrl", "Alt", "1") + " " + k("Ctrl", "Alt", "2") + " " + k("Ctrl", "Alt", "3")),
            ("Normal", k("Ctrl", "Maj", "N")), ("Puces", k("Ctrl", "Maj", "L")),
            ("Niveau de liste suivant, précédent", k("Tab") + " " + k("Maj", "Tab") + " (au début d'un élément)"),
        ])),
        ("mise-en-forme", "Mise en forme", table([
            ("Gras", k("Ctrl", "G") + " ou " + k("Ctrl", "B")), ("Italique", k("Ctrl", "I")), ("Souligné", k("Ctrl", "U")),
            ("Centrer", k("Ctrl", "E")), ("Justifier", k("Ctrl", "J")),
            ("Aligner à gauche", k("Ctrl", "Maj", "G") + " ou " + k("Ctrl", "L")),
            ("Aligner à droite", k("Ctrl", "Maj", "D") + " ou " + k("Ctrl", "R")),
            ("Police…", k("Ctrl", "D")), ("Agrandir, réduire la police", k("Ctrl", "Maj", ">") + " " + k("Ctrl", "<")),
            ("Effacer la mise en forme des caractères, du paragraphe", k("Ctrl", "Espace") + " " + k("Ctrl", "Q")),
        ])),
    ]),
    ("faq.html", "Questions fréquentes", "Compatibilité avec Word, sauvegardes, confidentialité.", "❓", [
        ("word", "Mes documents restent-ils compatibles avec Word ?", """
<p>Oui. OpenWrite enregistre au format .docx et conserve à l'identique tout ce qu'il ne modifie pas. Un document
ouvert puis enregistré sans modification est recopié octet pour octet.</p>
"""),
        ("bak", "À quoi sert le fichier .bak ?", """
<p>C'est la version du document avant le dernier enregistrement. Renommez-le en .docx pour revenir en arrière.</p>
"""),
        ("confidentialite", "OpenWrite envoie-t-il mes documents ?", """
<p>Non. Les rapports d'erreur (désactivables dans le menu Fichier) sont anonymes : ni contenu, ni nom de fichier,
ni nom d'utilisateur. La vérification des mises à jour ne transmet rien d'autre qu'une demande du fichier de
version.</p>
"""),
    ]),
]

# ------------------------------------------------------------------ gabarit

CSS = """
.docs{display:grid;grid-template-columns:260px minmax(0,1fr);gap:40px;padding:36px 0 70px}
.side{position:sticky;top:20px;align-self:start}
.search{width:100%;padding:11px 14px;border:1px solid #d8dfee;border-radius:14px;font:inherit;background:#fff}
.search:focus{outline:3px solid #00c9d566;border-color:#087cff}
.results{list-style:none;margin:8px 0 0;padding:0;max-height:340px;overflow:auto}
.results a{display:block;padding:8px 10px;border-radius:10px;font-size:14px}.results a:hover{background:#eef5ff}
.results small{display:block;color:#52617e;font-size:12px}
.toc{list-style:none;margin:18px 0 0;padding:0}.toc a{display:flex;gap:8px;padding:9px 12px;border-radius:12px;font-weight:600;font-size:14px;color:#2b3a5c}
.toc a:hover{background:#eef5ff}.toc a.active{background:#e6f0ff;color:#0a63d8}
.toc .sub{list-style:none;padding:2px 0 6px 34px;margin:0}.toc .sub a{font-weight:500;padding:5px 10px;font-size:13px}
article.doc{min-width:0}
.doc h1{font:800 clamp(30px,3vw,42px)/1.15 'Plus Jakarta Sans',sans-serif;letter-spacing:-1.5px;margin:6px 0 8px}
.doc .intro{font-size:18px;color:#394b72;margin-bottom:26px}
.doc h2{font:800 24px/1.3 'Plus Jakarta Sans',sans-serif;letter-spacing:-.6px;margin:42px 0 12px;scroll-margin-top:20px}
.doc h2 a.anchor{opacity:0;margin-left:8px;color:#ff8712;font-size:18px}.doc h2:hover a.anchor{opacity:1}
.doc p,.doc li{font-size:16px;line-height:1.7;color:#24324f}.doc ul,.doc ol{padding-left:22px}
.doc a{color:#0a63d8;font-weight:600}.doc a:hover{text-decoration:underline}
kbd{display:inline-block;min-width:22px;padding:2px 7px;border:1px solid #c9d2e8;border-bottom-width:3px;border-radius:7px;background:#fff;font:600 13px/1.4 'DM Sans',sans-serif;color:#111a47;text-align:center}
code{background:#eef2fb;padding:2px 6px;border-radius:6px;font-size:14px}
.steps{counter-reset:s;list-style:none;padding:0}.steps li{counter-increment:s;position:relative;padding:4px 0 10px 46px}
.steps li:before{content:counter(s);position:absolute;left:0;top:4px;width:30px;height:30px;border-radius:50%;display:grid;place-items:center;color:#fff;font-weight:800;font-size:14px;background:linear-gradient(105deg,#027efb,#0a6be8 55%,#ff9a3c)}
.callout{border-radius:18px;padding:16px 20px;margin:20px 0;border:1px solid}.callout p{margin:4px 0 0}
.callout.tip{background:linear-gradient(125deg,#edffff,#f1f8ff);border-color:#bfeeee}.callout.tip strong{color:#05807c}
.callout.warn{background:#fff7ed;border-color:#ffd8a8}.callout.warn strong{color:#b45309}
figure.shot{margin:24px 0;background:rgba(255,255,255,.75);padding:10px;border:1px solid #fff;border-radius:22px;box-shadow:0 18px 50px rgba(52,88,171,.16)}
figure.shot img{width:100%;height:auto;border-radius:14px}figure.shot figcaption{text-align:center;color:#52617e;font-size:13px;padding:10px 6px 2px}
.cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:14px}
.card{display:block;padding:18px 18px;border-radius:20px;background:linear-gradient(125deg,#fff8f0,#eef4ff);border:1px solid #ffffffb3;box-shadow:0 8px 22px rgba(52,78,143,.08);transition:transform .15s}
.card:hover{transform:translateY(-2px);text-decoration:none!important}.card b{display:block;font:800 17px 'Plus Jakarta Sans',sans-serif;color:#111a47;margin-bottom:4px}.card span{color:#52617e;font-size:14px;font-weight:500}
.table-wrap{overflow-x:auto}table{border-collapse:collapse;width:100%;background:#fff;border-radius:16px;overflow:hidden;box-shadow:0 6px 18px rgba(52,78,143,.07)}
th,td{text-align:left;padding:11px 14px;border-bottom:1px solid #eef1f8;font-size:15px}th{background:#f3f6ff;font-weight:700;color:#2b3a5c}
.pager{display:flex;justify-content:space-between;gap:12px;margin-top:50px;padding-top:22px;border-top:1px solid #e2e8f4}
.pager a{padding:12px 16px;border-radius:14px;background:#fff;border:1px solid #e2e8f4;font-weight:700}
.doc-hero{background:linear-gradient(115deg,#fcfdff 0%,#eef7ff 58%,#fff3e6 100%);border-bottom:1px solid rgba(42,77,139,.10)}
.doc-hero .container{padding:26px 0}.crumbs{color:#52617e;font-size:14px}.crumbs a{color:#0a63d8;font-weight:600}
@media(max-width:900px){.docs{grid-template-columns:1fr;gap:20px}.side{position:static}}
"""

SEARCH_JS = """
(function(){var box=document.getElementById('q'),list=document.getElementById('results');if(!box)return;
function norm(s){return s.toLowerCase().normalize('NFD').replace(/[\\u0300-\\u036f]/g,'')}
box.addEventListener('input',function(){var q=norm(box.value.trim());list.innerHTML='';if(q.length<2)return;
var words=q.split(/\\s+/),hits=INDEX.filter(function(e){var t=norm(e.t+' '+e.x);return words.every(function(w){return t.indexOf(w)>=0})}).slice(0,12);
if(!hits.length){list.innerHTML='<li><a>Aucun résultat</a></li>';return}
hits.forEach(function(h){var li=document.createElement('li'),a=document.createElement('a');a.href=h.u;a.innerHTML=h.t+'<small>'+h.p+'</small>';li.appendChild(a);list.appendChild(li)})});
box.addEventListener('keydown',function(e){if(e.key==='Enter'){var a=list.querySelector('a[href]');if(a)location.href=a.href}});})();
"""

def layout(page_idx):
    fname, title, desc, icon, sections = PAGES[page_idx]
    toc = []
    for i, (f, t, _, ic, secs) in enumerate(PAGES):
        active = ' class="active"' if i == page_idx else ''
        sub = ""
        if i == page_idx:
            sub = '<ul class="sub">' + "".join(f'<li><a href="#{a}">{st}</a></li>' for a, st, _ in secs) + "</ul>"
        toc.append(f'<li><a href="{f}"{active}><span>{ic}</span>{t}</a>{sub}</li>')
    body = "".join(
        f'<h2 id="{a}">{st}<a class="anchor" href="#{a}" aria-label="Lien vers cette section">#</a></h2>{h}'
        for a, st, h in sections)
    prev_link = f'<a href="{PAGES[page_idx-1][0]}">← {PAGES[page_idx-1][1]}</a>' if page_idx > 0 else "<span></span>"
    next_link = f'<a href="{PAGES[page_idx+1][0]}">{PAGES[page_idx+1][1]} →</a>' if page_idx < len(PAGES) - 1 else "<span></span>"
    ver = f" · OpenWrite {html.escape(VERSION)}" if VERSION else ""
    return f"""<!doctype html>
<html lang="fr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{title} — Aide OpenWrite</title><meta name="description" content="{html.escape(desc)}">
<link rel="icon" href="../assets/openwrite.ico" sizes="any"><link rel="stylesheet" href="../styles.css"><link rel="stylesheet" href="aide.css"></head><body>
<header class="site-header"><div class="container nav"><a class="brand" href="../index.html" aria-label="OpenWrite, accueil"><img src="../assets/openwrite-128.png" alt="">OpenWrite</a>
<nav class="nav-links" aria-label="Navigation principale"><a href="../index.html">Accueil</a><a href="../fonctionnalites.html">Fonctionnalités</a><a class="active" href="index.html" aria-current="page">Aide</a></nav></div></header>
<div class="doc-hero"><div class="container"><div class="crumbs"><a href="index.html">Aide</a>{' · ' + html.escape(title) if page_idx else ''}{ver}</div></div></div>
<main class="container docs"><aside class="side"><input id="q" class="search" type="search" placeholder="Rechercher dans l'aide…" aria-label="Rechercher dans l'aide" autocomplete="off">
<ul id="results" class="results"></ul><ul class="toc">{''.join(toc)}</ul></aside>
<article class="doc"><h1>{title}</h1><p class="intro">{html.escape(desc)}</p>{body}<nav class="pager">{prev_link}{next_link}</nav></article></main>
<footer class="footer"><div class="container footer-inner"><span>© OpenWrite · Aide en ligne</span><span><a href="../index.html#telecharger">Télécharger OpenWrite</a> · <a href="../fonctionnalites.html">Fonctionnalités</a></span></div></footer>
<script src="recherche.js"></script></body></html>
"""


def main():
    out = os.path.join(OUT, "aide")
    os.makedirs(os.path.join(out, "img"), exist_ok=True)
    index = []
    for i, (fname, title, desc, icon, sections) in enumerate(PAGES):
        open(os.path.join(out, fname), "w", encoding="utf-8").write(layout(i))
        for a, st, h in sections:
            text = re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", h))).strip()
            index.append({"t": st, "p": title, "u": f"{fname}#{a}", "x": text[:600]})
    open(os.path.join(out, "aide.css"), "w", encoding="utf-8").write(CSS.strip() + "\n")
    open(os.path.join(out, "recherche.js"), "w", encoding="utf-8").write(
        "var INDEX=" + json.dumps(index, ensure_ascii=False) + ";\n" + SEARCH_JS.strip() + "\n")
    # verification : toutes les ancres utilisees par l'application existent
    anchors = {f"{f}#{a}" for f, _, _, _, s in PAGES for a, _, _ in s} | {f for f, *_ in PAGES}
    print(f"{len(PAGES)} pages, {len(index)} sections -> {out}")
    return anchors


if __name__ == "__main__":
    anchors = main()
    # liste des cibles valides (lue par le test des liens contextuels)
    open(os.path.join(OUT, "aide", "ancres.txt"), "w", encoding="utf-8").write("\n".join(sorted(anchors)) + "\n")
