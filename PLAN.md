# OpenWrite — plan d'implémentation

Traitement de texte de bureau (Windows), écrit en C# / .NET, de la famille d'OpenPub, qui remplace Microsoft Word pour un usage hors ligne et **ouvre et enregistre directement les `.docx`**.

*Rédigé le 2026-11-05, à partir de l'inventaire des fonctions de Word (`INVENTAIRE-WORD.md`), du code d'OpenPub 2.1.3 (`/mnt/c/dev/open-pub`) et de la charte graphique `assets/OpenWrite-package-complet`.*

---

## 1. Point de départ

### Ce qu'apporte OpenPub
OpenPub (≈ 20 000 lignes de C#, 2.1.3 publiée) a résolu tout ce qui entoure un éditeur de documents ; OpenWrite le reprend au lieu de le réécrire :

| Domaine d'OpenPub | Fichiers | Reprise dans OpenWrite |
|---|---|---|
| Mises à jour automatiques signées (manifeste, vérification, bandeau) | `Core/Update/*` | ✔ telle quelle (nouvelle clé de signature, nouveau flux) |
| Signalement des plantages vers les issues GitHub | `Core/Diagnostics/CrashReport.cs`, `App/Diagnostics/CrashReporter.cs`, `tools/site-api` | ✔ |
| Copie de secours, récupération, enregistrement sûr, fichier modifié ailleurs | `Core/Io/RecoveryStore.cs`, `MainWindow.Recovery.cs` | ✔ |
| Installeur MSI par utilisateur (WiX 5), exécutable autonome | `installer/`, `tools/build-installer.ps1` | ✔ (nouveaux GUID, association `.docx` facultative) |
| Pipelines GitHub Actions (version, portail), portail Azure Static Web Apps, aide en ligne, page Fonctionnalités | `.github/workflows/*`, `tools/build-site.py`, `tools/build-docs.py`, `fonctionnalites.json` | ✔ avec la charte OpenWrite (`assets/OpenWrite-package-complet/index.html` + `style.css` servent de maquette) |
| Système de design WPF (ruban, onglets contextuels, styles, polices embarquées, info-bulles, menus) | `App.xaml`, `MainWindow.xaml` | ◐ mêmes styles, couleurs OpenWrite |
| Fenêtre « À propos », documents récents, aide contextuelle F1 | `AboutWindow.cs`, `Help/HelpService.cs` | ✔ |
| Moteur de lignes : coupure, justification, tabulations, puces, lettrine, exposants, césure, mesure abstraite (`ITextMeasurer`) | `Layout/TextLayoutEngine.cs`, `Hyphenator.cs` | ◐ base du moteur de paragraphe (§3.3) ; motifs NL/EN/DE à ajouter |
| Habillage autour des obstacles | `Layout/PageLayouter.cs` | ◐ réutilisé pour les objets flottants |
| Orthographe (ISpellChecker de Windows), rechercher / remplacer, typographie française, symboles | `Editing/SpellService.cs`, `Core/Editing/TextSearch.cs`, `FrenchTypography` | ✔ |
| Formes (50, poignées de réglage), modifier les points, WordArt, effets, retouche d'images | `Core/Geometry/*`, `Rendering/DecorRenderer.cs`, `ImageEditorWindow.cs` | ◐ correspondance avec les formes prédéfinies DrawingML |
| Tableaux (cellules mises en page comme zones de texte) | `Core/Editing/TableOps.cs`, `MainWindow.Tables.cs` | ◐ les opérations ; la mise en page est à refaire (tableaux en flux) |
| Export PDF (PDFsharp : tout ce qui est dessiné à l'écran va dans le PDF), polices incorporées, liens, impression, livret | `Output/DrawingToPdf.cs`, `PdfExporter.cs`, `PrintService.cs`, `FontEmbedder.cs` | ✔ |
| Publipostage CSV, modèles, blocs de construction | `MergeOps`, `Templates` | ◐ |
| Auto-test de bout en bout (`--selftest`), modes en ligne de commande (`--render`, `--pdf`, `--screenshot`) | `SelfTest.cs`, `App.xaml.cs` | ◐ même principe |

**Ce qui ne se reprend pas** : le modèle de document (`Page → Shape`, pages fixes), le format `.opub`, l'annulation par photo complète du document (`UndoStack`), la mise en page par page (`PageLayouter.LayoutPage` avec cache par page).

### Ce qui change par rapport à OpenPub
1. **Texte en flux** : le document est une suite de blocs (paragraphes, tableaux) que l'application **pagine** elle-même. C'est le nouveau cœur à écrire (§3.3).
2. **Format natif `.docx`** : pas de format maison. Les documents arrivent de Word (collègues, administration, écoles) et y repartent. Exigence n° 1 : **ouvrir, modifier et réenregistrer un `.docx` sans rien perdre** de ce qu'OpenWrite ne sait pas modifier.
3. **Documents longs** (jusqu'à ≈ 500 pages) : mise en page incrémentale et annulation par opérations.
4. **Pas de document unique de référence** : OpenPub visait une newsletter ; OpenWrite vise des documents variés. Il faut donc un **corpus de test** (§5.2).

### Hypothèses (confirmées le 2026-11-05)
| # | Hypothèse | État |
|---|---|---|
| H1 | Usage bureautique courant, en **français, néerlandais et anglais** | ✔ confirmée |
| H2 | Un corpus de `.docx` réels | ✔ `samples/` (§1 bis) — **hors dépôt** (données personnelles) |
| H3 | Un poste avec **Word sous licence** pour les PDF de référence | ✔ ce poste ; script `tools/export-reference-pdf.ps1` (Word COM, lecture seule) |
| H4 | Dépôt GitHub **public**, comme OpenPub, portail sur Azure Static Web Apps | ✔ confirmée |
| H5 | Interface en français uniquement, comme OpenPub | ✔ (comme OpenPub) |

## 1 bis. Corpus (`samples/`, analysé le 2026-11-05)

25 `.docx` : **9 documents réels** et 16 petits fichiers de test unitaire (un cas par fichier : commentaires, notes, note avec lien, liste simple, tableau, zone de texte, image, image externe, souligné, barré, format strict, BOM UTF-8, document vide, fichier produit par LibreOffice). Un raccourci `.LNK` est sans objet.

| Document réel | Pages | Langue | Ce qu'il exerce |
|---|---|---|---|
| `Annexe-20-Recours` | 11 | fr | **Formulaire hérité** (7 cases `FORMCHECKBOX`), champ `USERADDRESS`, 2 tableaux, image, en-têtes, signets, liens |
| `FrameView SDK License` | 6 | en | Contrat : listes numérotées à plusieurs niveaux (26 paragraphes), en-têtes, 13 signets |
| `Syllabus TPs … histologie` | 36 | en (US, GB, CA) | **Notes de bas de page** (6), 4 tableaux, **191 paragraphes de liste**, puces Wingdings, en-têtes |
| `Horaires 2025-2026 WVG` | 2 | — | **6 sections**, **colonnes**, zones de texte et images flottantes ; mode de compatibilité Word 2010 (14) avec **anciennes options `w:compat`** |
| `convocation kazakhstan` (×2) | 1 | nl | Mise en page par **zones de texte flottantes** (6) et images ancrées |
| `pv réunion 8 août 2025` | 2 | fr | Compte rendu : listes (23 paragraphes) |
| `Proposition évènements Club` | 1 | fr | Liste courte |

Modes de compatibilité (relevés par Word) : 15 (Word 2013+) pour les documents récents, 14 (Word 2010) pour `Horaires` et les fichiers de test, 12 (Word 2007) pour deux fichiers de test.

**PDF de référence** (`tools/export-reference-pdf.ps1` → `work/reference/`, journal `references.csv` : pages, mots, mode) : 22 sur 25 produits par Word, ≈ 7 s par document. **Word refuse d'ouvrir 3 fichiers de test** qu'il juge endommagés (`endnotes`, `external-picture` — produit par LibreOffice —, `tiny-picture-target-base-relative`) : OpenWrite devra, lui, les ouvrir (lecture tolérante) — ils serviront de tests de robustesse sans référence visuelle.

**Conséquences sur le plan** :
- Les **formulaires hérités** (`FORMCHECKBOX`) passent de P3 à **P1** (lot 10) ; la case doit au minimum s'afficher, se cocher et se conserver.
- Les **zones de texte et objets flottants** (lot 7) sont centraux pour 3 documents sur 9 : à garder dans la 1.0.
- Les **listes** (lot 4) sont massives : la numérotation doit être juste dès la 0.5.
- Les **anciennes options de compatibilité** (`w:compat`) existent dans la réalité : une partie du lot 13 (options qui changent la mise en page) est à avancer pendant le lot 2.
- **Manques du corpus**, à combler par des documents fabriqués avec Word sur ce poste : suivi des modifications, table des matières, document long (≥ 300 pages, le plus long fait 36 pages), en-têtes « première page différente », contrôles de contenu, équations, graphiques, SmartArt.
- **Confidentialité** : les documents réels contiennent des noms (dont celui d'un enfant) ; `samples/` et `work/` sont exclus du dépôt public. Les 16 fichiers de test pourront rejoindre `tests/` après vérification de leur origine et de leur licence.

---

## 2. Choix techniques

| Sujet | Choix | Raison |
|---|---|---|
| Runtime, interface | **.NET 10 / WPF**, comme OpenPub | Même chaîne d'outils, même installeur, code d'OpenPub réutilisable ; rendu typographique et impression matures |
| Ruban | Ruban maison d'OpenPub (styles `RBig/RSmall/RGroup/RibbonTab`) | Déjà au point (défilement sur petits écrans, onglets contextuels) ; pas de dépendance |
| Lecture / écriture `.docx` | **DocumentFormat.OpenXml 3.x** (Open XML SDK, licence MIT) pour le package et les parties ; **mappage maison** vers le modèle d'OpenWrite | SDK officiel, typé, maintenu ; le modèle d'édition reste léger et indépendant du XML |
| Préservation de l'inconnu | Chaque nœud du modèle garde les **éléments et attributs XML qu'il ne comprend pas** (îlots opaques) ; les parties non modifiées du package sont recopiées octet pour octet | Aller-retour sans perte : SmartArt, graphiques, équations, macros, `customXml`, extensions `w14/w15/w16*` |
| Modèle de document | Calqué sur la sémantique WordprocessingML (sections, paragraphes, segments, propriétés à héritage) mais **sans dépendre du SDK** | Correspondance 1-pour-1 à l'import/export ; moteur et éditeur testables seuls |
| Annuler / rétablir | **Journal d'opérations inversibles** (insérer, supprimer, modifier des propriétés, opérations structurelles) | La photo complète d'OpenPub coûte trop cher sur 300 pages ; les mêmes opérations servent au suivi des modifications et au rendu incrémental |
| Moteur de texte | Moteur maison : moteur de lignes d'OpenPub + **paginateur** nouveau | Aucune bibliothèque .NET ne pagine un document Word fidèlement |
| Polices | Polices du système + substitution métriquement compatible embarquée : **Carlito** (Calibri), **Caladea** (Cambria), **Liberation** (Arial, Times, Courier) — licence OFL / SIL ; police par défaut des nouveaux documents : **Inter** (OFL, celle de la maquette) | Aptos (police par défaut de Word depuis 2024) est une police cloud Microsoft non redistribuable : on l'utilise si Office l'a déposée sur le poste, sinon substitution signalée |
| Césure | `Hyphenator` d'OpenPub + motifs LibreOffice **fr, nl, en-US, en-GB, de** (LGPL/MPL, via pyphen) | Belgique : français et néerlandais indispensables |
| Orthographe | `ISpellChecker` de Windows (repris d'OpenPub) | Dictionnaires des langues installées dans Windows, rien à embarquer |
| Grammaire (lot 8, facultatif) | **LanguageTool** en serveur **local** (LGPL, Java) détecté s'il est installé ; jamais l'API en ligne par défaut | Pas d'API de grammaire dans Windows ; le texte ne doit pas quitter le poste sans accord |
| Synonymes | Thésaurus LibreOffice (fr : LGPL, nl : OpenTaal, en : WordNet) | Libres, hors ligne |
| Chiffrement `.docx` | [MS-OFFCRYPTO] chiffrement Agile (AES, `System.Security.Cryptography`) + conteneur OLE **OpenMcdf** (MPL 2.0) | Spécification publique |
| RTF | Lecteur / écrivain maison (spécification RTF 1.9.1 publique) | Indispensable pour le presse-papiers avec Word et Outlook |
| ODT | Mappage maison (ODF 1.3, spécification OASIS) | Échange avec LibreOffice |
| `.doc` binaire | **LibreOffice en mode `--headless --convert-to docx`** s'il est installé ; sinon message explicite | [MS-DOC] est publique mais énorme ; usage résiduel |
| Export PDF | PDFsharp 6 (MIT), chaîne d'OpenPub (`DrawingToPdf`) + signets, PDF balisé, PDF/A-2b | Le PDF reste identique à l'écran |
| Synthèse vocale (lot 14) | `Windows.Media.SpeechSynthesis` (voix installées) | Hors ligne |
| Tests | xUnit + **aller-retour `.docx`** (XML normalisé avant / après) + **fidélité visuelle** (pages rendues comparées aux PDF de Word) + `--selftest` sur la vraie fenêtre | Mesurer objectivement l'absence de perte et la fidélité |

**Code commun avec OpenPub** : on **copie** au départ les fichiers d'infrastructure (en-tête « repris d'OpenPub 2.1.3 ») plutôt que de créer tout de suite une bibliothèque partagée : les deux applications gardent des cycles de version indépendants. Si une troisième application de la famille voit le jour, on extrait un paquet NuGet `OpenFamily.Common` (mises à jour, plantages, récupération, PDF, aide).

---

## 3. Architecture

```
OpenWrite.sln
├─ src/
│  ├─ OpenWrite.Core        Modèle de flux, opérations + annuler/rétablir, résolution des styles,
│  │                        listes, champs, révisions, commentaires ; mises à jour, plantages, récupération (repris)
│  ├─ OpenWrite.Docx        Import / export WordprocessingML (Open XML SDK), préservation de l'inconnu,
│  │                        thème, numérotation, styles, DrawingML (formes, images), chiffrement
│  ├─ OpenWrite.Formats     RTF, ODT, HTML, TXT, Markdown ; .doc via LibreOffice ; presse-papiers (HTML Format, RTF)
│  ├─ OpenWrite.Layout      Moteur de paragraphe (repris d'OpenPub), paginateur, tableaux, notes, objets flottants,
│  │                        mise en page incrémentale ; mesure abstraite (ITextMeasurer)
│  └─ OpenWrite.App         WPF : ruban, vue des pages virtualisée, édition, volets (navigation, styles,
│                           commentaires, mise en forme), dialogues, PDF, impression, auto-test
├─ tests/
│  ├─ OpenWrite.Tests          Unitaires (Core, Docx, Layout avec mesureur monospace)
│  └─ OpenWrite.Fidelity       Corpus .docx → aller-retour sans perte + écart visuel avec les PDF de Word
├─ tools/                   Portail, aide, installeur, publication (repris d'OpenPub), comparaison de PDF
└─ installer/               WiX 5 (MSI par utilisateur)
```

### 3.1 Modèle de document (`OpenWrite.Core`)
Unités internes : **twips** (1/20 de point) pour les longueurs, comme WordprocessingML, afin que l'aller-retour ne crée aucun arrondi ; le moteur calcule en points (`double`).

```
Document   { Sections[], Styles, Numbering, Theme, Settings (compat, langue, onglets par défaut…),
             Footnotes[], Endnotes[], Comments[], Properties, Package (parties conservées) }
Section    { PageSize, Margins, Orientation, Columns, HeaderFooters (défaut / première / paire),
             PageNumbering, LineNumbering, Borders, Type (page suivante, continu, paire, impaire), Blocks[] }
Block      = Paragraph | Table | ContentControl(Block[]) | Opaque(xml)
Paragraph  { StyleId, Props (alignement, retraits, espacements, enchaînements, bordures, trame, tabulations,
             NumId/Level), MarkProps, Inlines[], Revision? }
Inline     = Run { Text, StyleId, Props (police, taille, gras… ), Revision? }
           | Break (ligne, page, colonne) | Tab | Field { Code, Result Inline[] }
           | Drawing { Inline | Anchor (position, habillage, z), Shape | Picture | TextBox(Block[]) | Opaque }
           | NoteRef | CommentRange(start/end) | Bookmark(start/end) | Hyperlink(Inline[])
           | ContentControl(Inline[]) | Math(Opaque + rendu) | Opaque(xml)
Table      { StyleId, Props (largeur, mise en page fixe/auto, retrait, bordures, habillage), Grid[], Rows[] { Props, Cells[] { Props (fusion, largeur, trame, alignement), Blocks[] } } }
Revision   { Kind (ins, del, moveFrom, moveTo, format), Author, Date, Id }
```

Points clés :
- **Propriétés à héritage**, résolues par `StyleResolver` dans l'ordre de Word : valeurs par défaut du document → style de tableau (mise en forme conditionnelle) → numérotation → style de paragraphe (chaîne `basedOn`) → style de caractère → mise en forme directe. Le résultat est mis en cache par paragraphe et invalidé par les opérations.
- **Position dans le document** : `DocPos(chemin du bloc, index du segment, décalage)` ; les opérations s'expriment avec ces positions, ce qui rend le journal d'annulation, le suivi des modifications et la recherche indépendants de l'affichage.
- **Les révisions sont des données du modèle**, pas un mode d'affichage : une suppression suivie reste dans le texte, marquée `del`. Le moteur l'affiche barrée (toutes les marques) ou la saute (marques simples, original).
- **Opaque** : tout élément non compris est conservé tel quel, déplacé avec son paragraphe, réécrit à l'enregistrement. Il s'affiche, quand c'est possible, via son **dessin de repli** (`mc:Fallback`, dessin SmartArt, aperçu d'objet OLE).

### 3.2 Format `.docx`
- **Import** : package OPC → parties `document`, `styles`, `numbering`, `settings`, `theme`, `fontTable`, `footnotes`, `endnotes`, `comments` (+ `commentsExtended/Ids` pour les fils et l'état « résolu »), en-têtes / pieds, images, `docProps`. `mc:AlternateContent` : on choisit la branche comprise (`wps`, `wpg`, `w14`…) et on garde l'ensemble pour l'écriture.
- **Export** : on régénère les parties modifiées ; les autres sont recopiées. Les identifiants (`w14:paraId`, `rsid`) sont conservés ; les nouveaux sont créés comme Word le ferait.
- **Contrôle** : à chaque enregistrement en développement, validation par `OpenXmlValidator` du SDK ; dans les tests, ouverture du résultat par le SDK et comparaison XML normalisée avec l'original.
- **Mode de compatibilité** : `w:compat` / `compatSetting` (version 11, 12, 14, 15) est lu ; les options qui changent la mise en page (espacement en haut de page, largeur des tableaux, interligne des polices asiatiques…) sont prises en charge au fil des lots 2, 6 et 13.

### 3.3 Moteur de mise en page (`OpenWrite.Layout`)
Trois étages :
1. **Paragraphe → lignes** : `TextLayoutEngine` d'OpenPub, étendu (largeur variable par ligne autour des flottants, segments de révision, champs, notes, ancres d'objets, caractères masqués). Résultat mis en cache par paragraphe, clé = (version du paragraphe, largeur disponible, exclusions).
2. **Paginateur** : place les lignes des blocs dans les colonnes et les pages de chaque section ; gère sauts, veuves et orphelines, paragraphes et lignes solidaires, espacement en haut de page, **notes de bas de page** (réservation de place au pied, continuation sur la page suivante), **tableaux** (lignes fractionnables, répétition de l'en-tête), **objets flottants** (positionnement relatif au paragraphe / à la marge / à la page, exclusions via `PageLayouter` d'OpenPub, itération jusqu'à stabilité), en-têtes / pieds (place du corps ajustée si un en-tête est haut), colonnes équilibrées en fin de section continue.
3. **Mise en page incrémentale** : une modification invalide le paragraphe touché ; le paginateur repart de la page qui le contient et s'arrête dès qu'une page retrouve **le même point de départ** qu'avant (convergence). Les numéros de page et `NUMPAGES` sont mis à jour à part. Objectif : < 16 ms par frappe sur un document de 300 pages, mise en page complète en arrière-plan à l'ouverture (premières pages affichées tout de suite).

La vue (`App/Controls/DocumentView`) est **virtualisée** : seules les pages visibles sont dessinées (`PageRenderer` d'OpenPub adapté), ce qui garde aussi le PDF identique à l'écran.

---

## 4. Interface

D'après la maquette `assets/OpenWrite-package-complet/concept-openwrite-reference.webp` :
- **Ruban** aux onglets de Word, mêmes noms et même ordre (Fichier, Accueil, Insertion, Création, Mise en page, Références, Publipostage, Révision, Affichage) ; onglets contextuels Tableau, Image, Forme, En-tête et pied de page ; **raccourcis clavier de Word** à l'identique — le réapprentissage doit être nul.
- **Volet gauche** : miniatures des pages (comme dans la maquette) et, dans un second onglet, plan des titres (volet de navigation de Word).
- **Volet droit « Mise en forme »** (propre à OpenWrite, présent dans la maquette) : style, police, paragraphe, espacement de la sélection ; il accueille aussi les volets Styles, Commentaires, Vérifications.
- **Barre d'état** : page x sur y, nombre de mots, langue, suivi des modifications, modes d'affichage, zoom.

---

## 5. Qualité et mesure de la fidélité

### 5.1 Critères de réussite
- **Aller-retour sans perte** : ouvrir puis enregistrer un document du corpus sans le modifier ne change aucun élément XML significatif (comparaison normalisée : ordre des attributs, `rsid`, espaces de noms) ; après une modification locale, seul le paragraphe touché diffère.
- **Fidélité de pagination** : sur le corpus, **≥ 95 % des pages** ont les mêmes premiers et derniers mots que dans le PDF produit par Word, et l'écart en pixels par page est **< 3 %**.
- **Réactivité** : frappe < 16 ms et ouverture affichant la 1re page en < 1 s sur un document de 300 pages.
- Un `.docx` enregistré par OpenWrite s'ouvre **sans message de réparation** dans Word.

### 5.2 Corpus
`tests/OpenWrite.Fidelity/corpus` (hors dépôt pour les documents réels, comme `converted/` dans OpenPub ; des documents synthétiques versionnés couvrent les mêmes cas) : lettres, rapport avec table des matières et notes, document à sections et en-têtes différents, tableaux longs, documents avec révisions et commentaires, publipostage, formulaire à contrôles de contenu, document de 300 pages, documents NL et EN. Pour chacun : le PDF exporté par Word (hypothèse H3).

### 5.3 Règles communes à chaque lot (comme OpenPub)
- Une version par lot, publiée par la pipeline (étiquette `vX.Y.Z`).
- Tests unitaires de la logique pure (Core, Docx, Layout), vérifications ajoutées à `--selftest`, cas ajoutés au corpus de fidélité.
- Sections d'aide dans `tools/build-docs.py` (constantes `HelpService`), éléments `"depuis": "X.Y.Z"` dans `fonctionnalites.json`, section de ce journal.
- **Tant qu'un élément n'est pas éditable, il est conservé** : on n'enregistre jamais un `.docx` qui a perdu quelque chose sans l'avoir dit à l'utilisateur.

---

## 6. Risques

| Risque | Impact | Parade |
|---|---|---|
| **Algorithmes de mise en page de Word non documentés** (interligne exact des polices, largeur automatique des tableaux, espacement en haut de page, options de compatibilité) | Pages qui ne coupent pas aux mêmes endroits que Word | Banc de fidélité dès le lot 2 ; calibrations mesurées et consignées dans ce journal (comme `FontLineFactors` d'OpenPub) ; s'inspirer du comportement documenté par LibreOffice (code MPL, lecture seulement) |
| **Perte de données à l'enregistrement** | Document abîmé chez un collègue sous Word : perte de confiance immédiate | Préservation de l'inconnu dès le lot 1 ; test d'aller-retour sur tout le corpus à chaque version ; validation `OpenXmlValidator` ; enregistrement sûr + `.bak` (repris) |
| **Performances sur les documents longs** | Frappe saccadée | Mise en page incrémentale et vue virtualisée conçues au lot 2 et mesurées (document de 300 pages dans le corpus) ; prototype de 2 jours avant d'engager le lot |
| **Notes de bas de page et tableaux sur plusieurs pages** | Pagination fausse dans les rapports et mémoires | Prototypes dédiés dans les lots 6 et 9 ; cas simples d'abord (note qui tient sur sa page, tableau sans fusion verticale) |
| **Polices absentes** (Aptos, Calibri sans Office) | Texte qui ne passe plus aux mêmes endroits | Substitution métriquement compatible embarquée (Carlito, Caladea, Liberation) ; liste des substitutions affichée (Fichier ▸ Info) |
| **Presse-papiers avec Word / Outlook** | Collage qui perd la mise en forme : irritant quotidien | RTF et `HTML Format` en entrée et en sortie dès le lot 3 ; tests sur des extraits copiés depuis Word |
| **Périmètre de Word immense** | Projet sans fin | Priorités P0 → P3 de l'inventaire ; jalon 1.0 limité aux lots 0 à 7 ; lot 15 à confirmer fonction par fonction |
| Dépôt public et documents réels | Fuite de données personnelles | Corpus réel hors dépôt (`.gitignore`), captures du portail faites sur des documents de démonstration |

---

## 7. Phases

| Phase | Lots | Contenu | Livrable |
|---|---|---|---|
| **A — Socle** | 0 à 2 | Infrastructure reprise, modèle de flux, `.docx` sans perte, pagination | Visionneuse fidèle de `.docx`, enregistrement sans perte (0.1 → 0.3) |
| **B — Traitement de texte** | 3 à 7 | Édition, mise en forme, styles, listes, sections, en-têtes, tableaux, images | **OpenWrite 1.0** : lettres, rapports et comptes rendus de bout en bout |
| **C — Travail à plusieurs et documents longs** | 8 à 10 | Révision, références, champs | 1.1 → 1.3 |
| **D — Production** | 11 à 14 | Publipostage, sorties et formats, compatibilité, affichage et confort | 1.4 → 1.7 |
| **E — Fonctions rares** | 15 | À confirmer une par une | — |

---

## 8. Lots

Estimations en jours de développement pour une personne, tests et aide compris, comme dans OpenPub.

### 8.1 Lot 0 — Socle repris d'OpenPub (version 0.1)

| Tâche | Détail | Jours |
|---|---|---|
| Solution et projets | `OpenWrite.sln` (Core, Docx, Formats, Layout, App, Tests, Fidelity) ; `nuget.config` ; `CLAUDE.md` | 0,5 |
| Infrastructure reprise | Mises à jour signées (nouvelle clé), plantages, récupération, documents récents, « À propos », aide F1, modes en ligne de commande | 2 |
| Charte et fenêtre | `App.xaml` aux couleurs OpenWrite, icône `openwrite.ico`, ruban vide aux onglets de Word, volets gauche/droit, barre d'état | 1,5 |
| Publication | Installeur WiX (nouveaux GUID), pipelines GitHub Actions, portail à partir de `assets/OpenWrite-package-complet` (bouton « Installeur bientôt disponible » remplacé à la première version) | 2 |

### 8.2 Lot 1 — Modèle de flux et `.docx` sans perte (version 0.2) — critique

| Tâche | Détail | Jours |
|---|---|---|
| Prototype d'aller-retour | Sur 10 documents du corpus : import → export, comparaison XML ; valide la stratégie de préservation avant tout le reste | 2 |
| Modèle | Sections, blocs, segments, propriétés, `DocPos`, éléments opaques | 4 |
| Résolution des styles | `StyleResolver` : valeurs par défaut, `basedOn`, styles liés, styles de caractère, mise en forme directe, couleurs et polices du thème | 3 |
| Import `.docx` | `document`, `styles`, `numbering`, `settings`, `theme`, `fontTable`, images, `docProps`, `mc:AlternateContent` | 4 |
| Export `.docx` | Régénération des parties modifiées, copie des autres, identifiants, validation | 3 |
| Opérations et annulation | Insérer / supprimer / modifier des propriétés / scinder / fusionner des paragraphes ; journal inversible, regroupement de la frappe continue | 3 |
| Banc d'aller-retour | `OpenWrite.Fidelity` : normalisation XML, rapport des différences | 1,5 |

### 8.3 Lot 2 — Pagination et affichage (version 0.3) — critique

| Tâche | Détail | Jours |
|---|---|---|
| Prototype de performance | Document de 300 pages : mise en page complète, incrémentale, défilement ; décide de la structure des caches | 2 |
| Moteur de paragraphe | Reprise de `TextLayoutEngine` : propriétés résolues, interligne Word (simple = hauteur de ligne de la police, multiple, au moins, exactement), espacements avant/après avec « contextual spacing », tabulations par défaut | 4 |
| Paginateur | Sections, sauts de page/colonne/section, veuves et orphelines, paragraphes et lignes solidaires, saut de page avant, espacement en haut de page | 6 |
| Mise en page incrémentale | Invalidation, convergence, mise en page en arrière-plan | 3 |
| Vue des pages | `DocumentView` virtualisée, zoom (repris), miniatures, barre d'état (page x sur y, mots) | 3 |
| Polices | Substitution métriquement compatible, polices cloud d'Office si présentes, avertissement | 1,5 |
| Césure | Motifs nl, en, de ; options de coupure de Word (zone, coupures consécutives, majuscules) | 1 |
| Banc de fidélité visuelle | Rendu PNG des pages (`--render`) comparé aux PDF de Word (pdftoppm ou PDFium), score par page, premiers/derniers mots | 2 |

### 8.4 Lot 3 — Édition et mise en forme directe (version 0.4)

| Fonction Word | Mise en œuvre | Jours |
|---|---|---|
| Point d'insertion et sélection | Reprise et extension de `InPlaceTextEditor` : navigation clavier de Word, souris (mot, paragraphe, marge), sélection en colonne, sélection multiple, glisser-déposer | 5 |
| Saisie | Frappe, Entrée (nouveau paragraphe, style suivant), Maj+Entrée, Ctrl+Entrée, Suppr/Retour arrière aux limites de paragraphe et de liste, IME Windows | 2 |
| Onglet Accueil | Police, taille, gras… soulignements, casse, couleur, surlignage, effacer la mise en forme ; alignement, interligne, retraits, trame, bordures ; boîtes Police et Paragraphe (reprises de `ParagraphWindow`) | 4 |
| Règle | Retraits, tabulations, marges (repris de `.Rulers.cs`) | 1,5 |
| Presse-papiers | Interne (fragment de modèle), **RTF** et **HTML Format** en entrée et en sortie, texte ; options de collage ; images | 4 |
| Marques de mise en forme, caractères spéciaux, symboles | ¶, espaces, tabulations ; espaces insécables, tirets ; fenêtre Symboles reprise | 1,5 |
| Liens hypertextes | Insérer, modifier, Ctrl+clic | 1 |
| Rechercher / remplacer (base), statistiques, reproduire la mise en forme, mini barre | Repris d'OpenPub, adaptés au flux | 2 |

### 8.5 Lot 4 — Styles, listes et thèmes (version 0.5)

| Fonction Word | Mise en œuvre | Jours |
|---|---|---|
| Galerie et volet Styles | Styles rapides avec aperçu, appliquer, nouveau à partir de la sélection, modifier (met à jour le document), inspecteur de styles, gérer les styles | 5 |
| Puces et numérotation | `abstractNum` / `num` / remplacements de niveau ; compteurs le long du document ; recommencer / continuer ; Tab / Maj+Tab ; marque de liste dessinée (repris d'OpenPub) | 5 |
| Listes à plusieurs niveaux | Liaison aux styles de titre, formats « 1.1.2 », définir une nouvelle liste | 2 |
| Thèmes, jeux de styles, définir par défaut | `theme1.xml`, couleurs et polices du thème dans les sélecteurs, modèle Normal de l'utilisateur | 2 |

### 8.6 Lot 5 — Mise en page, sections, en-têtes et pieds (version 0.6)

| Fonction Word | Mise en œuvre | Jours |
|---|---|---|
| Onglet Mise en page | Marges, orientation, taille, par section ; boîte « Mise en page » ; insérer des sauts de section | 2 |
| En-têtes et pieds de page | Édition en place (corps grisé), première page / paires-impaires différentes, lier au précédent, position, navigation ; onglet contextuel | 5 |
| Numéros de page | Champ `PAGE`/`NUMPAGES`/`SECTIONPAGES`, format, début par section, galeries | 1,5 |
| Colonnes | Colonnes de section, largeurs, séparateur, équilibrage | 2,5 |
| Filigrane, couleur et bordures de page | Filigrane texte/image (forme dans l'en-tête, comme Word), bordures (BorderArt repris) | 2 |
| Numéros de ligne | | 1 |

### 8.7 Lot 6 — Tableaux (version 0.7) — critique

| Fonction Word | Mise en œuvre | Jours |
|---|---|---|
| Prototype de largeur automatique | Algorithme d'ajustement de Word (largeur préférée, contenu minimal/maximal, `tblLayout` fixe/auto) mesuré sur le corpus | 2 |
| Mise en page des tableaux en flux | Cellules = suites de blocs, hauteur des lignes, fusion horizontale et verticale, **fractionnement sur plusieurs pages**, répétition de l'en-tête, tableaux imbriqués | 6 |
| Édition | Insérer (grille, boîte), saisie, Tab, sélection de cellules / lignes / colonnes, insérer / supprimer, fusionner / fractionner, redimensionner à la souris et à la règle (repris de `TableOps`) | 5 |
| Mise en forme | Bordures, trame, marges de cellule, alignement, orientation du texte ; **styles de tableau** avec mise en forme conditionnelle | 4 |
| Divers | Convertir texte ↔ tableau, trier, formule simple, tableau flottant, coller depuis Excel | 3 |

### 8.8 Lot 7 — Images, formes et zones de texte (version 1.0)

| Fonction Word | Mise en œuvre | Jours |
|---|---|---|
| Images alignées sur le texte | Dans la ligne (hauteur de ligne), redimensionnement, rotation | 2 |
| Objets flottants | Ancrage au paragraphe, position relative (marge, page, colonne, paragraphe, caractère), habillages carré / rapproché / au travers / haut et bas / derrière / devant via `PageLayouter`, « déplacer avec le texte », icône d'ancre | 6 |
| Formes DrawingML | Correspondance des ~170 formes prédéfinies avec `ShapeLibrary` (50 reprises, les autres ajoutées), remplissages, contours, effets, groupes, modifier les points (repris) | 4 |
| Zones de texte | Contenu en blocs, mise en page par le moteur de paragraphe, zones liées | 2,5 |
| Format de l'image | Rogner, corrections, styles, compresser, texte de remplacement (repris de l'éditeur d'images) | 2 |
| Affichage des dessins de repli | SmartArt (dessin `dsp`), objets OLE (aperçu EMF/WMF), WordArt (repris), zone de dessin | 1,5 |

**Jalon OpenWrite 1.0** (lots 0 à 7 : ≈ 136 jours) : un document Word courant s'ouvre fidèlement, se modifie avec les mêmes gestes et se réenregistre sans perte ; export PDF et impression repris d'OpenPub (adaptés au flux dans le lot 2, sans signets ni balises).

### 8.9 Lot 8 — Révision et travail à plusieurs (version 1.1)

| Fonction Word | Mise en œuvre | Jours |
|---|---|---|
| Orthographe | `SpellService` repris : vérification en arrière-plan par paragraphe, langue par segment, F7, dictionnaire personnel, options | 2 |
| Commentaires | Plages ancrées, fils de réponses, résolu (`commentsExtended`), bulles dans une marge de révision ajoutée à droite de la page, volet Commentaires, navigation | 5 |
| Suivi des modifications | Opérations enregistrées comme révisions (insertions, suppressions, mise en forme, marque de paragraphe, lignes de tableau, déplacements) ; couleurs par auteur ; affichages Simple / Toutes / Aucune / Original ; accepter / refuser (une, toutes, par auteur) ; volet Vérifications | 9 |
| Comparer / combiner | Différence par paragraphes puis par mots (Myers) → révisions | 4 |
| Protection | Restreindre la modification (`documentProtection`, hachage SHA-512 avec itérations), plages d'exception ; chiffrement par mot de passe ([MS-OFFCRYPTO]) ; marquer comme final | 3 |
| Rechercher / remplacer avancé | Caractères génériques, mises en forme, caractères spéciaux (^p, ^t…) | 2 |
| Grammaire et synonymes (facultatifs) | LanguageTool local s'il est présent ; thésaurus LibreOffice | 3 |

### 8.10 Lot 9 — Références et documents longs (version 1.2)

| Fonction Word | Mise en œuvre | Jours |
|---|---|---|
| Notes de bas de page et de fin | Insertion, numérotation (formats, recommencer), séparateurs, **réservation de place au pied de la page et continuation**, conversion ; prototype de 1 jour avant | 7 |
| Table des matières | Champ `TOC` (niveaux, styles, `TC`), galeries, mise à jour, liens, « ajouter du texte » | 3 |
| Signets, renvois, légendes, table des illustrations | Champs `REF`, `PAGEREF`, `SEQ`, `TOC \c` | 4 |
| Index | `XE`, `INDEX`, concordance | 2 |
| Citations et bibliographie | `sources.xml`, styles APA / ISO 690, champs `CITATION`, `BIBLIOGRAPHY` (P3 : à confirmer) | 4 |

### 8.11 Lot 10 — Champs, blocs de construction et contrôles de contenu (version 1.3)

| Fonction Word | Mise en œuvre | Jours |
|---|---|---|
| Moteur de champs | Analyse des codes, commutateurs de format (`\@`, `\#`, `\*`), mise à jour (F9, avant impression), afficher les codes (Alt+F9), verrouiller ; DATE, AUTHOR, FILENAME, DOCPROPERTY, IF, =formule, STYLEREF… | 4 |
| Propriétés du document | Fichier ▸ Info, propriétés personnalisées, contrôles liés | 1,5 |
| QuickPart, blocs de construction, pages de garde | `glossaryDocument` et bibliothèque de l'utilisateur (repris de `Templates`) | 2,5 |
| Contrôles de contenu | Texte, case à cocher, listes, sélecteur de date, image, section répétitive ; espaces réservés, verrouillage, mode Création | 4 |

### 8.12 Lot 11 — Publipostage (version 1.4)

| Fonction Word | Mise en œuvre | Jours |
|---|---|---|
| Sources de données | CSV (repris de `MergeOps`), **Excel `.xlsx`** (Open XML SDK), saisie d'une nouvelle liste ; modifier la liste (cocher, trier, filtrer) | 3 |
| Champs et règles | MERGEFIELD, bloc d'adresse, formule d'appel, Si…Alors…Sinon, Ignorer, Numéro de fusion, Enregistrement suivant | 2,5 |
| Aperçu et fin | Navigation, vérification, nouveau document, impression, PDF (un ou un par enregistrement) | 2 |
| Enveloppes et étiquettes | Formats standard, planches Avery / Herma | 2 |

### 8.13 Lot 12 — Sorties et formats d'échange (version 1.5)

| Fonction Word | Mise en œuvre | Jours |
|---|---|---|
| PDF complet | Signets depuis les titres, liens internes (renvois, table des matières), propriétés, **PDF balisé** (structure, texte de remplacement), PDF/A-2b | 4 |
| Impression | Reprise de `PrintService` : plages, pages par feuille, recto verso manuel, livret ; aperçu dans Fichier ▸ Imprimer | 1,5 |
| RTF | Lecture / écriture complètes (au-delà du presse-papiers du lot 3) | 2,5 |
| ODT | Lecture / écriture (texte, styles, listes, tableaux, images, en-têtes) | 4 |
| HTML, TXT, Markdown | Import / export | 2 |

### 8.14 Lot 13 — Compatibilité et fidélité (version 1.6)

| Fonction Word | Mise en œuvre | Jours |
|---|---|---|
| Options de compatibilité | Prise en charge des options `w:compat` qui changent la mise en page, mode de compatibilité affiché dans le titre, « Convertir » | 4 |
| Calibrations | Passage sur tout le corpus, corrections du paginateur et des tableaux, consignées dans ce journal | 5 |
| Formats hérités | `.doc` via LibreOffice, `.docm` (macros conservées + avertissement), Flat OPC, ISO strict en lecture, insérer le texte d'un fichier | 3 |
| Mode protégé, vérifier la compatibilité | Fichiers venus d'Internet ouverts en lecture seule | 1 |

### 8.15 Lot 14 — Affichage, navigation et confort (version 1.7)

| Fonction Word | Mise en œuvre | Jours |
|---|---|---|
| Volet de navigation | Titres (glisser pour réorganiser, réduire), pages, résultats de recherche ; Atteindre (Ctrl+G) | 3 |
| Modes d'affichage | Web, Plan, Brouillon, Lecture ; déplacement côte à côte ; fractionner la fenêtre, plusieurs documents | 4 |
| Correction et mise en forme automatiques | Liste de remplacements FR/NL/EN, typographie française (reprise), listes automatiques, liens, ordinaux | 2,5 |
| Modèles | Nouveau à partir d'un modèle `.dotx`, modèles fournis, modèle Normal | 1,5 |
| Options, barre d'accès rapide, recherche de commandes, touches Alt | | 2,5 |
| Lire à voix haute, inspecteur de document, vérificateur d'accessibilité, thème sombre de l'interface | | 3 |

### 8.16 Lot 15 — Fonctions rares (à confirmer avant de les engager)

| Fonction Word | Mise en œuvre | Jours |
|---|---|---|
| Équations | Affichage OMML (moteur de mise en page mathématique, police Latin Modern Math / STIX Two, OFL) : 6 ; édition (structures, saisie linéaire) : 6 | 12 |
| Graphiques | Affichage depuis les valeurs en cache (barres, courbes, secteurs) : 4 ; édition des données : 4 | 8 |
| SmartArt éditable | | 6 |
| Encre | Affichage et dessin (crayon d'OpenPub) | 2 |
| Dictée | Reconnaissance vocale Windows | 2 |
| Signatures numériques | XML-DSig du package, vérification | 3 |
| Écriture de droite à gauche | Bidi dans le moteur de lignes | 6 |
| Divers | Supprimer l'arrière-plan d'une image, capture d'écran, icônes libres, combiner des documents, volet Presse-papiers | 4 |

### 8.17 Synthèse

| Lot | Version | Contenu | Jours |
|---|---|---|---|
| 0 | 0.1 | Socle repris d'OpenPub | 6 |
| 1 | 0.2 | Modèle de flux, `.docx` sans perte, annulation par opérations | 20,5 |
| 2 | 0.3 | Pagination incrémentale, affichage, polices, banc de fidélité | 22,5 |
| 3 | 0.4 | Édition, mise en forme directe, presse-papiers RTF/HTML | 21 |
| 4 | 0.5 | Styles, listes, thèmes | 14 |
| 5 | 0.6 | Mise en page, sections, en-têtes et pieds | 14 |
| 6 | 0.7 | Tableaux | 20 |
| 7 | **1.0** | Images, formes, zones de texte, objets flottants | 18 |
| 8 | 1.1 | Orthographe, commentaires, suivi des modifications, comparer, protection | 28 |
| 9 | 1.2 | Notes, table des matières, renvois, légendes, index, bibliographie | 20 |
| 10 | 1.3 | Champs, blocs de construction, contrôles de contenu | 12 |
| 11 | 1.4 | Publipostage | 9,5 |
| 12 | 1.5 | PDF complet, impression, RTF, ODT, HTML, Markdown | 14 |
| 13 | 1.6 | Compatibilité et calibrations | 13 |
| 14 | 1.7 | Navigation, modes d'affichage, correction automatique, modèles, options | 16,5 |
| 15 | — | Fonctions rares (à confirmer) | 43 |
| | | **Total** | **≈ 292 jours** (≈ 249 sans le lot 15) |

**Jalons** : après les lots 0 à 2 (≈ 49 jours), OpenWrite est une **visionneuse fidèle** qui réenregistre sans perte — c'est le moment de vérifier, sur le corpus, que l'approche tient. Après le lot 7 (≈ 136 jours), **OpenWrite 1.0** couvre l'usage bureautique courant. Les lots 8 à 10 (≈ 60 jours) ouvrent le travail à plusieurs et les documents longs ; les lots 11 à 14 complètent la production.

**Ordre ajustable** : si l'usage réel l'exige, le suivi des modifications et les commentaires (lot 8) peuvent passer avant les formes (lot 7) — ils ne dépendent que des lots 1 à 3.

**Comparaison avec OpenPub** : ≈ 160 jours pour la parité avec Publisher, sur un moteur plus simple (pages fixes) et sans contrainte d'aller-retour sur le format. Le surcoût d'OpenWrite vient du paginateur, des tableaux en flux et de l'exigence « `.docx` sans perte ».

---

## 9. Prochaines étapes proposées

1. Confirmer les hypothèses H1 à H5 (§1), en particulier **réunir le corpus** et un accès à Word pour les PDF de référence.
2. Lot 0 : créer le dépôt, reprendre l'infrastructure d'OpenPub, publier une 0.1 « visionneuse vide » par la pipeline pour valider la chaîne de bout en bout.
3. Prototype d'aller-retour `.docx` (lot 1, 2 jours) **avant** d'écrire le modèle définitif : c'est le risque n° 1.
4. Prototype de performance du paginateur (lot 2, 2 jours) sur le document de 300 pages.

---

## 10. Journal

### 10.1 Lot 0, première partie (2026-11-06) — version 0.1 en cours

**Fait**
- Solution `OpenWrite.sln` : `OpenWrite.Core`, `OpenWrite.Docx` (DocumentFormat.OpenXml 3.3.0), `OpenWrite.Layout`, `OpenWrite.App` (WPF), `tests/OpenWrite.Tests` ; `nuget.config`, `.gitattributes`, `.gitignore`, `CLAUDE.md`, `README.md`. Les projets `OpenWrite.Formats` et `OpenWrite.Fidelity` seront créés avec leurs lots (3 et 1), pour ne pas versionner de projets vides.
- Infrastructure reprise d'OpenPub 2.1.3 (en-tête « Repris d'OpenPub » dans chaque fichier) : manifeste de mise à jour signé (charge utile `openwrite-update/1`, distincte d'OpenPub : un manifeste d'OpenPub ne peut pas être accepté par OpenWrite), service de mise à jour et bandeau, rapports d'erreur anonymisés (signature sur les méthodes `OpenWrite.*`, noms de `.docx/.dotx/.docm/.rtf/.odt` masqués), réglages et documents récents (`%APPDATA%\OpenWrite`), fenêtre « À propos », aide contextuelle, césure française.
- `RecoveryStore` généralisé : l'appelant écrit la copie (rappel), le magasin ne dépend plus du format ; nouveau `SafeFile.Write` (fichier temporaire + `File.Replace`, `.bak` facultatif), base de l'enregistrement du lot 1.
- Charte OpenWrite dans `App.xaml` : styles d'OpenPub recolorés (encre `#111A47`, bleu `#087CFF`, cyan `#00C9D5`, orange `#FF8712`, dégradé bleu → orange des boutons principaux, pastille de version orange comme le portail) ; icônes 64/128/256 tirées de `openwrite-1024.png`.
- Fenêtre d'après la maquette : ruban aux 8 onglets de Word (mêmes noms, même ordre) dont **chaque groupe affiche la version qui l'activera** (table `Roadmap` de `MainWindow.Ribbon.cs`, issue du §8) ; zoom et volets déjà actifs (onglet Affichage) ; volet gauche Pages / Titres ; volet droit « Mise en forme » (inactif jusqu'à la 0.4) ; barre d'état (pages selon Word, mots, langue, zoom) ; glisser-déposer d'un fichier.
- **Ouverture en lecture seule** : `DocxInfo` (Docx) lit un résumé tolérant — comptes (paragraphes, tableaux, sections, images, zones de texte, notes, commentaires, révisions, signets, liens, contrôles de contenu, champs de formulaire hérités, équations), mode de compatibilité et anciennes options `w:compat`, titres (niveau de plan des styles, chaîne `basedOn`), langues, polices, liste des parties, texte brut des 60 premiers paragraphes. Affiché en « page » d'aperçu, dans le volet de droite et Fichier ▸ Informations. Titre de fenêtre « [Mode de compatibilité] » comme Word.
- Modes en ligne de commande : `--info`, `--screenshot` (scènes : onglets, `file-menu`, `about`), `--help-topics`, clés et manifestes de mise à jour.
- 28 tests unitaires (mises à jour, rapports d'erreur, copies de secours, enregistrement sûr, césure, `DocxInfo` sur des `.docx` fabriqués par le SDK dans les tests : rien du corpus personnel).

**Mesures sur le corpus** (`--info` sur les 25 fichiers de `samples/`)
- **Les 25 fichiers s'ouvrent**, y compris les 3 que Word juge endommagés (`endnotes`, `external-picture`, `tiny-picture-target-base-relative`).
- **Mode de compatibilité identique à Word sur les 22 références** (`references.csv`). Calibration : un fichier sans `compatibilityMode` est en mode **12** (Word 2007), pas 11 (`empty`, `utf8-bom`).
- Mots : écart de 0 à 2 % sur les documents réels (Annexe 3 119 / Word 3 113 ; Syllabus 8 699 / 8 803 ; FrameView 2 693 / 2 714), plus fort sur les petits fichiers de test parce que Word ne compte ni les notes ni les zones de texte (`footnotes` 8 / 2, `text-box` 4 / 0) et sur les convocations (101 / 91 : zones de texte). À aligner sur `ComputeStatistics` de Word avec la fonction Statistiques (lot 3) ; la barre d'état affichera alors le même nombre que Word.
- Les `pages`/`words` de `docProps/app.xml` sont absents ou périmés dans plusieurs fichiers (Syllabus enregistré par Word pour le web, Horaires) : la barre d'état n'en dépendra plus dès la pagination (lot 2).

**Reste pour terminer le lot 0**
- Publication : clé de signature OpenWrite (`--gen-update-key %APPDATA%\OpenWrite\update-signing-key.pem`), installeur WiX (nouveaux GUID, association `.docx` facultative), pipelines GitHub Actions, portail Azure Static Web Apps à partir de `assets/OpenWrite-package-complet`, aide (`tools/build-docs.py`) ; puis renseigner `UpdateFeed`, `CrashReportUrl`, `HelpBase`, `UpdatePublicKey` dans `OpenWrite.App.csproj`.
- Dépôt GitHub public (H4) et premier commit.
- Ensuite, conformément au §9 : **prototype d'aller-retour `.docx`** (lot 1, 2 jours) avant le modèle définitif.

### 10.2 Lot 1 — prototype d'aller-retour `.docx` (2026-11-06) : stratégie validée

**Construit**
- `Core/Model/Flow.cs` : modèle en flux prototype — `Story` → blocs (`Paragraph`, `Table`/`TableRow`/`TableCell`, `BlockContentControl`, `OpaqueBlock`) → éléments de ligne (`Run`, `InlineContainer` pour lien, révisions `ins/del/moveFrom/moveTo`, `smartTag`, `customXml`, `fldSimple` ; `InlineContentControl` ; `OpaqueInline`) → contenu de segment (`TextContent`, `TabContent`, `BreakContent`, `OpaqueRunContent`). Chaque nœud garde **tous les attributs d'origine** (rsid, `w14:paraId`…).
- `PropertyBag` : les propriétés (pPr, rPr, tblPr, trPr, tcPr, sectPr) **conservent l'élément XML d'origine** ; lecture typée, modification insérée dans l'ordre du schéma (ECMA-376), extensions inconnues gardées, historique (`*PrChange`) toujours dernier.
- `Docx/DocxPackage` : lecture directe du ZIP (tolérante), parties de texte (document, en-têtes, pieds, notes, commentaires) dans le modèle, **toutes les autres entrées gardées octet pour octet** ; à l'enregistrement, seules les parties modifiées sont régénérées ; ordre, noms et dates des entrées conservés. `StoryMapping` (XML ↔ modèle), `XmlCompare` (différence normalisée, étendue d'une modification).
- `FlowEdits` : insérer du texte, scinder un paragraphe (nouveaux `w14:paraId/textId` uniques, saut de section laissé sur le dernier morceau comme Word).
- `tests/OpenWrite.Fidelity` : banc sur le corpus local (rapport `work/roundtrip/rapport.md`), mode `--stress`. 8 tests unitaires sur des documents écrits à la main (36 tests au total).

**Résultats sur les 25 fichiers du corpus**
| Critère (§5.1) | Résultat |
|---|---|
| Aller-retour, **toutes** les parties de texte régénérées depuis le modèle | **25/25 identiques** (comparaison XML normalisée) |
| Autres parties du package | 100 % identiques octet pour octet |
| Texte atteint par le modèle | 100 % sauf zones de texte (dans les dessins, opaques) : convocations 90 %, Horaires 97,8 %, `text-box` 0 % → lot 7 |
| Modifications (texte, gras, scission) | **toujours locales** : 1 bloc modifié (2 pour la scission), le reste du corps identique |
| Validateur Open XML SDK | aucune erreur nouvelle sur les 100 fichiers produits |
| **Ouverture par Word** (COM, `export-reference-pdf.ps1` sur les sorties) | **50/50 ouverts** (aller-retour et scission) ; aller-retour : **mêmes pages et mêmes mots que l'original pour les 22 références** |
| Bonus | les 3 fichiers que Word refusait (`endnotes`, `external-picture`, `tiny-picture-target-base-relative`) **s'ouvrent dans Word après un aller-retour par OpenWrite** (archive ZIP réécrite proprement) |
| Performances (Debug) | Syllabus ×10 (≈ 360 pages, 6 090 paragraphes) : lecture ≈ 0,2 s, écriture ≈ 0,2 s |

Après scission, Word compte un mot de plus (le mot coupé en deux : attendu) ; 3 documents serrés passent à une page de plus (convocations : page 2 vide, la marque de fin repoussée sous les zones flottantes) — réaction de mise en page de Word à un paragraphe ajouté, pas un défaut du fichier.

**Calibrations et pièges relevés (à garder dans le modèle définitif)**
1. **Retours à la ligne dans les attributs** (texte de remplacement des images, `o:gfxdata` VML) : écrire avec `NewLineHandling.Entitize` (`&#xA;`), sinon la normalisation XML les change en espaces.
2. **`xml:space`** : Word 2010 écrit `<w:t> </w:t>` sans `xml:space` ; un texte non modifié est réécrit tel quel (`TextContent.LoadedValue`), `preserve` n'est ajouté qu'au texte modifié, et seulement pour les blancs XML (l'espace insécable n'en est pas un : «  :… » en tête de texte dans Annexe).
3. **Enregistrement** : `File.Replace` a échoué une fois (« Impossible de supprimer le fichier à remplacer », fichier juste écrit verrouillé par l'antivirus ou l'indexation) → `SafeFile` réessaie 5 fois.
4. Format strict ISO 29500 : même modèle, espace de noms lu sur la racine.

**Îlots opaques restants** (par fréquence) : `w:proofErr` (370, à jeter ou garder — Word les régénère), signets (37+37), champs complexes `fldChar`/`instrText` (32/11), VML `w:pict` (27), séparateurs de notes, `w:drawing` (10), références de notes, `w:lastRenderedPageBreak`, `mc:AlternateContent`, ancres de commentaires, `w:softHyphen`. Ordre de modélisation proposé pour le lot 1 : champs (FORMCHECKBOX du formulaire Annexe, PAGE, TOC) et signets ; références de notes/commentaires ; puis dessins et zones de texte (lot 7, contenu des `txbxContent` comme `Story`).

**Décision** : la stratégie « attributs d'origine + propriétés conservées en XML + îlots opaques + recopie des parties non modifiées » est **retenue** pour le modèle définitif. Prochaines étapes du lot 1 : `DocPos`, journal d'opérations inversibles (annuler/rétablir) au-dessus de `FlowEdits`, `StyleResolver`, modélisation des champs et signets, enregistrement depuis l'application (version 0.2).

### 10.3 Lot 1 — champs et signets (2026-11-06)

**Décision de modèle (révise le §3.1)** : un champ n'est **pas** un nœud `Field { Code, Result }` qui contiendrait ses segments. Comme dans Word, il reste une suite de **marqueurs dans le texte** — `FieldChar` (début / séparateur / fin, avec `w:ffData` conservé en XML), `FieldCode` (`w:instrText`) — et sa structure est **calculée** par `FieldIndex`. Raisons : un champ peut couvrir plusieurs paragraphes (table des matières), s'imbriquer dans le code ou le résultat d'un autre (IF { MERGEFIELD }, TOC + PAGEREF), et chaque marqueur porte la mise en forme de son segment ; restructurer les segments casserait l'aller-retour exact du §10.2. Les signets suivent le même principe : `BookmarkStart`/`BookmarkEnd` typés (dans le paragraphe, entre blocs via `InlineMarkerBlock`, entre lignes ou cellules de tableau) et `BookmarkIndex`.

**Construit**
- `Core/Model` : `Markers.cs` (marqueurs typés), `FlowWalker` (parcours dans l'ordre du document : paragraphes, cellules, contrôles de contenu, liens, révisions), `FieldIndex` (code avec les champs imbriqués entre accolades comme Alt+F9, résultat avec « \r » entre paragraphes, parent/enfants, segments du résultat, anomalies), `FormCheckBox` (état = `w:checked` sinon `w:default` ; revenir à l'état par défaut retire `w:checked`, le XML redevient l'original), `FieldInstruction` (type, arguments, commutateurs ; `\*`, `\@`, `\#` prennent toujours un argument, les autres seulement s'il est entre guillemets), `BookmarkIndex` (appariement par `w:id`, texte couvert, signets masqués `_…`, doublons), `MarkerOps` (remplacer le résultat d'un champ en gardant la mise en forme du premier segment, ajouter / renommer / supprimer un signet, règles de nom de Word, identifiant libre).
- `Docx` : lecture et écriture des marqueurs dans `StoryMapping` ; `DocxPackage.Relationships` et `HyperlinkTarget` (cible d'un lien, servira au Ctrl+clic du lot 3).
- Banc : modifications « case » (inverser la première case de formulaire) et « signet » (ajouter un signet), export de la vue d'OpenWrite (`modele/*.json`) ; `tools/word-inspect.ps1` relève la vue de Word par COM (lecture seule, une instance par fichier) ; `OpenWrite.Fidelity --compare-word`.
- 18 tests unitaires sur un document fabriqué qui réunit les cas absents du corpus : TOC sur trois paragraphes avec liens et PAGEREF imbriqués, IF avec MERGEFIELD dans son code, champ simple DATE, case à cocher, signet entre blocs et signet sur deux paragraphes, champs et signets mal formés (54 tests au total).

**Résultats sur le corpus**
| Critère | Résultat |
|---|---|
| Aller-retour | toujours **25/25 identiques** ; champs et signets ne sont plus des îlots opaques |
| Modifications « case » et « signet » | locales (un seul bloc), validateur SDK sans erreur nouvelle, **50/50 ouverts par Word** |
| Vue d'OpenWrite = vue de Word (champs, résultats, cases, liens, signets) | **22/22 documents** (les 3 fichiers que Word refuse n'ont pas de relevé) |
| Case cochée par OpenWrite | Word la voit cochée (Annexe : `[F,F,F,F,F,F,F]` → `[V,F,F,F,F,F,F]`) |
| Signet ajouté par OpenWrite | vu par Word dans les 19 documents concernés (les 6 autres n'ont pas de paragraphe de texte) |

**Calibrations — conventions du modèle objet de Word, absentes du fichier**
1. Word présente chaque `w:hyperlink` comme un champ `HYPERLINK "cible"` ; il ajoute `\h` quand `w:history` n'est pas activé (Horaires), `\l "ancre"` pour un lien interne.
2. Dans un document en mode de compatibilité 2010, chaque forme flottante apparaît comme un champ `SHAPE \* MERGEFORMAT` (Horaires : 6).
3. Le signet masqué `_GoBack` (dernière position d'édition) est retiré par Word à l'ouverture et recréé à l'enregistrement ; OpenWrite le conserve.
4. Le résultat de PAGE / NUMPAGES lu par Word dans un en-tête est recalculé hors contexte de page (0) ; le fichier garde la dernière valeur affichée (Syllabus : 2 et 24). Ce résultat est dynamique : la mise en page (lot 2) le calculera par page, sans le réécrire dans le fichier.

**Reste du lot 1** : `DocPos` et journal d'opérations inversibles (annuler / rétablir) au-dessus de `FlowEdits` et `MarkerOps` ; `StyleResolver` ; zones de texte (`w:txbxContent` comme `Story`, aujourd'hui dans les îlots `w:drawing`/`w:pict`) avancées si le lot 7 doit en dépendre ; enregistrement depuis l'application (version 0.2), avec affichage des cases ☐/☒ et des résultats de champs dans l'aperçu.

### 10.4 Lot 1 — positions et annuler / rétablir (2026-11-07)

**Positions (`DocPos`)** : chemin du paragraphe depuis la racine (un tableau consomme 3 index : bloc, ligne, cellule ; un contrôle de contenu 1) + décalage en caractères. Longueurs (`ParagraphAtoms`) : texte (supprimé compris) et code de champ = leur longueur ; tabulation, retour, caractère de champ, objet (dessin, appel de note, équation) = 1 ; repères (signets, commentaires, `proofErr`, `lastRenderedPageBreak`) = 0. Les positions ne référencent aucun objet : elles survivent aux annulations qui remplacent des blocs par leurs copies. Ordre du document = ordre lexicographique.

**Opérations (`Core/Editing/TextOps`)**
- `Insert` : le texte prend la mise en forme du caractère qui précède ; au début du paragraphe, celle du suivant ; dans un paragraphe vide, celle de la marque de paragraphe sans sa révision ; jamais dans une suppression suivie.
- `Delete` : dans un paragraphe ou sur plusieurs (même conteneur) ; blocs intermédiaires retirés (tableaux compris) ; segments, liens et révisions vidés retirés ; repères gardés. Le paragraphe fusionné **garde les propriétés du premier** et reçoit le saut de section du dernier — choix à confronter à Word (COM) au lot 3, avec la saisie (Word classique donne les propriétés du second paragraphe quand on supprime une marque de paragraphe).
- `Split` (Entrée) : nouveaux `w14:paraId/textId`, saut de section sur le second, un lien coupé est dupliqué ; repère au point de coupe à gauche. Remplace `FlowEdits.SplitParagraph` (gardé pour le banc).
- Suppression d'une cellule à une autre : refusée pour l'instant (sélection de cellules au lot 6).

**Annuler / rétablir (`UndoJournal`) — décision** : chaque entrée garde une **copie des seuls blocs touchés, avant et après**, désignés par leur chemin (pas d'objet). Plutôt que des opérations inverses écrites à la main : exact par construction (les copies gardent attributs, XML conservé, valeurs lues), coût local (un paragraphe par frappe ; OpenPub sérialisait tout le document), toute fonction d'édition devient annulable (`journal.Begin(...)` / `Commit`), et une édition qui échoue est défaite (transaction sans `Commit`). Frappe regroupée par mot (l'espace termine le mot, comme Word) ; état enregistré suivi (`MarkSaved`, `IsModified`, y compris quand l'état enregistré devient inatteignable) ; `Changed(racine)` pour marquer la partie à régénérer. Le suivi des modifications (lot 8) s'appuiera sur les mêmes positions, pas sur ces copies.

**Résultats**
| Critère | Résultat |
|---|---|
| Tests unitaires | 9 nouveaux (63 au total) : chemins dans tableaux et contrôles, héritage de mise en forme, suppression sur 7 blocs, scission dans un lien, tout annuler / tout rétablir, regroupement de la frappe, échec d'une opération, état enregistré |
| Corpus : 3 à 6 opérations par le journal (frappe, scission, suppression traversant des blocs, frappe en cellule, case à cocher) puis **tout annuler** | **19/19 documents identiques à l'original** (6 sans paragraphe de texte : sans objet) |
| **Tout rétablir** | 19/19 identiques à l'état modifié ; validateur SDK sans erreur nouvelle ; **25/25 ouverts par Word** |
| Performances (≈ 360 pages, Debug) | frappe journal compris **0,03 ms par caractère** (objectif §5.1 : 16 ms) ; annuler 1 000 frappes : 1 ms |

**Reste du lot 1** : `StyleResolver` (propriétés résolues : défauts du document, styles `basedOn`, styles de caractère, mise en forme directe, thème) ; enregistrement depuis l'application (0.2 : Fichier ▸ Enregistrer via `DocxPackage` + `SafeFile`, `.bak`, journal → partie `Dirty`) ; zones de texte si le lot 7 doit en dépendre.

### 10.5 Lot 1 — résolution des styles et version 0.2 (2026-11-07) : lot 1 terminé

**`StyleResolver` (`Core/Styles`)** : `StyleSheet` (styles, numérotation, thème, paramètres ; `DocxPackage.StyleSheet()`) et résolution vers `ParagraphFormat` (alignement, retraits, espacements, interligne, enchaînements, plan, liste ; twips) et `RunFormat` (polices des 4 emplacements, tailles, gras, italique, soulignement, barré, majuscules, masqué, couleur, surlignage, exposant, espacement, langue ; points), plus `ParagraphMark` et `ListMarker`. Fusion par propriété, par attribut pour `rFonts`/`ind`/`spacing`/`lang` (`firstLine`/`hanging` exclusifs, police explicite ⇄ police de thème par emplacement, `w:line` sans `w:lineRule` = auto), par côté pour les bordures, taquets cumulés avec `clear`. Propriétés bascule en OU exclusif entre niveaux de styles, la mise en forme directe impose. Polices et couleurs du thème (`clrSchemeMapping`, teinte et ombre en luminance TSL). Boucles `basedOn` coupées, style inconnu → Normal.

**Vérification contre Word** (`tools/word-format.ps1` par COM ↔ `OpenWrite.Fidelity --formats` / `--compare-formats`, 1 235 paragraphes appariés sur 22 documents) : **99,3 à 100 % d'accord par propriété** ; hors paragraphes vides (appariement ambigu entre paragraphes vides et cellules) **99,7 à 100 %** — alignement, retraits, première ligne, espace avant, police, italique, souligné : 100 %. **Couleurs de thème : 36/36 identiques** à la valeur calculée que Word enregistre dans `w:val`.

**Calibrations Word (contraires ou absentes de la norme)**
1. **Le niveau de liste s'applique après le style de paragraphe** (ECMA-376 §17.7.2 le place avant) : Annexe, « Paragraphe de liste » `firstLine=709` + liste directe `hanging=360` → Word −18 pt.
2. **Liste du style retirée par le paragraphe (`w:numId="0"`)** : ni le retrait de la liste ni celui du style ne s'appliquent (FrameView « Legal2_L1 » `firstLine=720` → 0 ; Syllabus « coupe », liste niveau 2 `left=1788` → 0). Une première hypothèse (« le retrait de la liste du style s'applique quand même ») corrigeait FrameView mais dégradait 66 paragraphes du Syllabus : le banc l'a montré, la règle a été corrigée.
3. **Format strict : mesures avec unité** (`w:after="8pt"`, `w:line="12.95pt"`, `in`, `cm`, `mm`, `pc`).
4. Relevé COM : une couleur de thème renvoyée par `Font.Color`/`TextColor.RGB` n'est pas fiable (`FFFF00` pour du noir) → vérifiée par `w:val` ; les marques de fin de ligne de tableau (`Information(31)`) sont écartées ; `ConvertTo-Json` de PowerShell 5.1 sur une liste .NET échoue (« types des arguments ») sans `-InputObject $list.ToArray()`.
5. Document sans valeurs par défaut ni style Normal (`empty.docx`) : Word applique celles de son modèle Normal.dotm, propres au poste — non reproductible, hors comparaison.

Restent hors du résolveur : mise en forme conditionnelle des styles de tableau (lot 6), styles de numérotation liés au-delà d'un niveau, `*Chars` (retraits en caractères), compteurs de liste (lot 4).

**Version 0.2 (application)**
- Le document ouvert est chargé dans le modèle (`DocxPackage`) ; l'aperçu est **construit sur le modèle et mis en forme par le résolveur** (page à la largeur et aux marges de la section, tableaux en grille, résultats de champs affichés et codes masqués, puces indicatives) — sans pagination (0.3). Le volet de droite décrit la mise en forme du paragraphe survolé.
- **Cases à cocher des formulaires** cliquables, **Annuler / Rétablir** (Ctrl+Z / Ctrl+Y, boutons d'accès rapide), **Enregistrer / Enregistrer sous** (Ctrl+S, Ctrl+Maj+S, F12 ; même famille de format : un `.docm` reste un `.docm`), version précédente en `.bak`, « • » et « modifié » dans les titres, question à la fermeture, avertissement si le fichier a changé sur le disque, **copies de secours** toutes les 5 minutes et récupération au démarrage (repris d'OpenPub, copies `.docx`).
- `DocxPackage.Save(chemin)` adopte le contenu enregistré (les parties régénérées deviennent la référence) : un second enregistrement cumule les modifications ; sans modification, il recopie le fichier à l'identique.
- **Auto-test de bout en bout** `--selftest <dossier>` (vraie fenêtre hors écran, document fabriqué) : 13 vérifications — ouverture, clic sur la case, titre modifié, Ctrl+Z ramène à l'état enregistré, Ctrl+Y, Ctrl+S, relecture (seule la case a changé), `.bak`.

**Bilan du lot 1** (estimé 20,5 jours) : prototype d'aller-retour, modèle en flux, champs et signets, positions, annuler/rétablir, résolution des styles, enregistrement depuis l'application. 73 tests unitaires, 13 vérifications de bout en bout, banc sur 25 documents (aller-retour, 6 modifications, annuler/rétablir), vues de Word et d'OpenWrite comparées pour les champs, signets et la mise en forme. Prochain lot : **lot 2 — pagination et affichage (0.3)**, en commençant par le prototype de performance sur un document de 300 pages (§8.3).

### 10.6 Lot 0 — publication : pipelines, installeur, portail Azure (2026-11-10)

- **Dépôt public** https://github.com/prolice/open-write (premier commit : lots 0 et 1, version 0.2.0 ; `samples/` et `work/` exclus, vérifiés).
- **Azure** (abonnement « Azure Noise2Quite », coût 0 €) : groupe de ressources `rg-openwrite-updates` (Europe de l'Ouest) et Static Web App **`openwrite-updates`, offre Free** : https://gentle-glacier-0e0efdd03.3.azurestaticapps.net. Quotas vérifiés : groupes de ressources gratuits, 980 par abonnement (4 utilisés) ; **10 Static Web Apps Free par abonnement** (3 utilisées : noise2quite, openpub, openwrite) ; 250 Mo par site, 100 Go de bande passante par mois.
- **Pipelines GitHub Actions** reprises d'OpenPub : `ci.yml` (nouveau : compilation et tests à chaque envoi), `release-app.yml` (étiquette `vX.Y.Z` → tests, exécutable autonome, MSI, `latest.json` signé, portail, release GitHub, Azure), `publish-portal.yml` (sans release : simple avertissement, rien n'est publié).
- **Installeur WiX** par utilisateur, GUID propres ; OpenWrite s'ajoute à « Ouvrir avec » des `.docx/.docm/.dotx/.dotm` (`OpenWithProgids`) **sans devenir le programme par défaut** (Word reste associé).
- **Portail** : charte OpenWrite (styles d'OpenPub recolorés bleu/cyan/orange), page de téléchargement, page Fonctionnalités (4 domaines, éléments à venir masqués jusqu'à leur version), **aide en ligne** (7 pages, 21 sections ; les 10 liens contextuels de l'application vérifiés). Captures sur un **document de démonstration fictif** fabriqué par `build-site.py` (jamais le corpus). Construction locale vérifiée : site de 59 Mo (MSI 58 Mo) pour 230 Mo autorisés.
- **Clé de signature** OpenWrite créée (`%APPDATA%\OpenWrite\update-signing-key.pem`) ; clé publique, flux, rapports d'erreur et aide renseignés dans `OpenWrite.App.csproj`.
- Reste : secrets GitHub `UPDATE_SIGNING_KEY` et `AZURE_STATIC_WEB_APPS_API_TOKEN`, réglage `GITHUB_TOKEN` du relais d'erreurs, puis première release `v0.2.0` par la pipeline.

### 10.7 Lot 2 — prototype de performance et premier moteur de pagination (2026-11-10)

**Moteur (`OpenWrite.Layout`)**, conçu pour Word et non repris tel quel d'OpenPub (son moteur dépend du modèle de Publisher : zones chaînées, réduction automatique) :
- `Fonts.cs` : `FontFace` (avances et métriques en em), `IFontProvider`, `FontVerticalMetrics` (tables `head`/`hhea`/`OS/2`, collections `.ttc`). **Hauteur de ligne simple de Word = formule GDI** : (usWinAscent + usWinDescent + tmExternalLeading) / unitsPerEm, avec tmExternalLeading = max(0, lineGap hhea − ((winAscent + winDescent) − (ascender − descender))). Times New Roman 12 pt → 13,8 pt, Calibri → 1,2207 em, **Aptos → 1,2847 em même si la police demande les métriques typo** (USE_TYPO_METRICS) : vérifié sur le PV (16,6 pt = 12 × 1,2847 × 1,079).
- `ParagraphLayout.cs` : texte affiché mis à plat (codes de champs masqués, résultats affichés, texte masqué et suppressions omis, objets en ligne à leur taille `wp:extent`, sauts de ligne / page / colonne), coupure aux espaces (espaces de fin suspendus), mot trop long coupé au caractère, taquets (`w:tabs` puis taquets par défaut `w:defaultTabStop`), puce ou numéro (retrait suspendu = taquet implicite ; compteurs au lot 4), interligne simple / multiple / au moins / exact, la marque de paragraphe compte dans la hauteur de la dernière ligne, alignements.
- `Paginator.cs` (`DocumentLayout`) : sections (format de page, page suivante / continue), saut de page avant, **veuves et orphelines**, **lignes solidaires**, **paragraphes solidaires** (chaîne jusqu'aux premières lignes du suivant), espacement contextuel, espace avant supprimé en haut d'une page commencée par le flux, **bordures de paragraphe groupées** (largeur + `w:space` en haut et en bas du groupe), **en-tête et pied qui repoussent le corps** (défaut, première page si `w:titlePg`), tableaux par rangées entières (grille, `w:gridSpan`, `w:vMerge`, `w:trHeight`).
- `App/Layout/WpfFontProvider.cs` : avances lues dans `GlyphTypeface` (sans crénage, comme Word par défaut ; tableau d'accès direct pour les 1 280 premiers caractères), polices **cloud d'Office** chargées depuis `%LOCALAPPDATA%\Microsoft\FontCache\4\CloudFonts` (Aptos…), noms à graisse (« Calibri Light ») et noms PostScript (« Calibri-Bold »), substitution compatible (Carlito, Caladea, Liberation) sinon Arial, remplacements relevés.

**Structure des caches (décision du prototype)** : mise en page des paragraphes en cache par **référence** de paragraphe (`ConditionalWeakTable`) et largeur ; le journal d'annulation remplace les blocs par des clones (nouvelle référence = recalcul naturel), une modification en place passe par `Update(blocs)`. Les pages mémorisent leur **point de départ** (élément de flux, ligne). Reprise : depuis la page qui précède la modification ; **convergence** dès qu'une nouvelle page commence au même endroit qu'une ancienne, après la dernière modification ; les pages suivantes sont réutilisées. Chemin rapide de la frappe (modification en place, pas de bloc ajouté ni retiré) : aucune reconstruction du flux, recherche dichotomique des pages.

**Mesures** (banc `OpenWrite --layout`, version Release, ce poste) :

| Document | Pages | Mise en page complète | Frappe (moyenne / pire) | Pages recalculées par frappe |
|---|---|---|---|---|
| Syllabus × 9 | 333 | 0,80 s | 0,4 ms / 3,0 ms | 2,3 |
| Syllabus × 20 | 740 | 0,79 s | 0,2 ms / 2,9 ms | 2,0 |
| Annexe × 28 | 281 | 0,69 s | 0,3 ms / 3,1 ms | 2,0 |

**Objectif atteint : < 16 ms par frappe**, indépendant de la longueur du document ; la mise en page complète (< 1 s pour 700 pages) se fera en arrière-plan à l'ouverture, premières pages d'abord. Le banc vérifie aussi que la reprise donne exactement les pages d'une mise en page complète (20 frappes par document, tout le corpus).

**Fidélité, premier relevé.** `tools/word-pages.ps1` (Word par COM, lecture seule) relève le début de chaque page et la page et la position verticale de chaque paragraphe ; `--word-pages` compare.
- Premier moteur, avant les calibrations : **nombre de pages identique à Word sur 21/22 documents**, **37/77 pages qui commencent au même endroit**.
- Avec les bordures de paragraphe, les en-têtes et pieds, et Aptos : 18/22 et 34/80. Ces règles sont justes une à une (Annexe : écart de 1,5 pt par groupe bordé résorbé ; Syllabus : corps de la première page à 99,7 pt contre 97,8 pt dans Word). Le recul vient de leur interaction avec ce qui n'est pas encore calibré :
  - pieds de page du Syllabus dont l'espacement fait remonter le bas du corps (hypothèse en cours : l'espacement après du dernier paragraphe d'un en-tête ou d'un pied ne compte pas) ;
  - PV et convocations en Aptos, plus bas que dans Word de +14 à +46 pt, à chercher dans la coupure des lignes ;
  - page de garde du Syllabus qui déborde d'une ligne ;
  - Horaires (colonnes, zones de texte, objets flottants).

**Suite du lot 2** : calibrer ces écarts un par un avec `--word-pages --detail` (et `OPENWRITE_BENCH_DUMP=<page>`) jusqu'à retrouver puis dépasser 21/22 et 37/77 ; colonnes ; lignes de tableau fractionnables et bordures de cellules ; puis la vue des pages virtualisée (`DocumentView`) dans l'application, la barre d'état (page x sur y) et le rendu PNG des pages comparé aux PDF de Word.

### 10.8 Lot 2 — calibration de la pagination sur Word (2026-11-10)

Méthode : `--layout --word-pages --detail` donne, document par document, les pages qui ne commencent pas au même endroit que dans Word et l'écart vertical paragraphe par paragraphe (appariement par plus longue sous-suite commune) ; `OPENWRITE_LAYOUT_TRACE=1` décrit chaque coupure de page et chaque cellule coupée, `OPENWRITE_TRACE_TEXT=<début d'un paragraphe>` sa coupure en lignes, `OPENWRITE_BENCH_DUMP=<page>` les positions d'OpenWrite. Pour la coupure des lignes, la position horizontale de chaque mot dans Word est relevée par COM (`Information(wdHorizontalPositionRelativeToPage)`). Les paramètres incertains sont réglables par variable d'environnement et calibrés sur tout le corpus, pas sur un cas.

**Résultat : 61/78 pages qui commencent au même endroit que dans Word** (37/77 au premier relevé). FrameView 6/6 et 76/76 paragraphes à 1 pt près ; Annexe 10/11 ; PV, convocations, tables : tous les paragraphes à 1 pt près ; Syllabus 27/37 (193 paragraphes à 1 pt près).

**Calibrations Word (mesurées)**
1. **Métriques typo quand la police les demande** (OS/2 USE_TYPO_METRICS) : Aptos 1,2207 em et non 1,2847 (formule GDI) ; vérifié sur le PV : 17,0 pt par ligne à 12 pt × 1,158. La formule GDI reste celle des autres polices. (Corrige le §10.7.)
2. **Interligne multiple** : l'espace ajouté au-delà de la hauteur naturelle peut déborder la marge du bas (ligne à 1,5 placée à 752,9 pt, bas du corps à 771 pt).
3. **En-têtes et pieds** : l'espacement après de leur dernier paragraphe compte, ainsi que leurs bordures (corps du Syllabus à 53 + 13,8 + 10 + 0,75 = 77,5 pt comme dans Word).
4. **Polices de remplacement du document** (`fontTable.xml`, `w:altName`) essayées en premier pour une police absente : ZapfChancery → Monotype Corsiva.
5. **Une tabulation ne donne pas sa hauteur à la ligne** : en-tête dont la ligne n'est faite que d'une tabulation en 12 pt → hauteur de la marque de paragraphe en 10 pt.
6. **Petites capitales** : minuscules en capitales à 70 % de la taille (largeurs) ; 70 à 80 % donnent le même résultat sur le corpus.
7. **Tableau flottant** (`w:tblpPr`) hors du flux ; ancré en bas de marge, il réduit le bas du corps de la page (page de garde du Syllabus).
8. **Une ligne qui ne contient que le saut de page reste en bas de page** (pas de page blanche) ; **la marque de paragraphe reste sur la ligne d'un saut de page final** (pas de ligne en tête de la page suivante).
9. **Objet en ligne** : la ligne prend la hauteur de l'objet (marges d'effet `wp:effectExtent` comprises) sans y ajouter le descendant de la police.
10. **Bordures de cellule** : chaque rangée ajoute l'épaisseur de sa bordure du bas (cellule, sinon `insideH` / `bottom` du tableau ou de son style) ; entre deux rangées, la plus épaisse des deux bordures partagées ; la première ajoute celle du haut.
11. **Justification : Word resserre les espaces pour faire tenir un mot de plus**, d'une quantité fixe par espace et non proportionnelle (22 % accepté à 10 pt, 18,8 % refusé à 12 pt) : **0,46 pt par espace** (débuts de page identiques entre 0,45 et 0,47).
12. **Rangées de tableau coupées entre deux pages** au niveau des lignes de chaque cellule (sauf `w:cantSplit`), **rangées d'en-tête répétées** (`w:tblHeader`) ; **une rangée à hauteur minimale (`w:trHeight`) ne se coupe pas** (rangée de 146 pt renvoyée entière par Word avec 120 pt libres ; écart ramené de −137 à −3 pt).
13. **Notes de bas de page** (premier pas du lot 9) : numéro d'appel affiché dans son style, place du texte de la note et du séparateur réservée en bas de la page de l'appel, la ligne de l'appel descend si sa note ne tient pas.

Tolérances essayées et écartées : dépassement en fin de ligne (sans effet), dépassement en bas de page (dégrade Annexe au-delà de 0,5 pt).

**Écarts restants** : pages commencées dans une cellule fusionnée verticalement (la clé de Word est le début du bloc fusionné : artefact de mesure) ; Syllabus pages 14 à 37 (prochaine cause à chercher) ; Horaires (sections continues à colonnes inégales avec équilibrage, saut de colonne, zone de texte, images flottantes) ; convocations (texte en zones de texte : lot 7). Performance inchangée : 0,2 à 0,3 ms par frappe (pire 3 ms) sur 333 et 740 pages, reprise incrémentale identique à la mise en page complète sur tout le corpus.

**Suite du lot 2** : colonnes (remplissage, saut de colonne, équilibrage en fin de section continue), écarts restants du Syllabus, puis la vue des pages dans l'application.

### 10.9 Version 0.3 — vue des pages (2026-11-10)

- **Vue des pages** (`App/Layout/PagesView.cs`) : les pages du moteur (§10.7, §10.8) les unes sous les autres, **seules les pages visibles sont dessinées** (à chaque défilement, d'après la zone visible du ScrollViewer) ; dessins gardés par page (`PageDrawings`, partagés avec les miniatures). Zoom (transformation de la zone de travail), **miniatures** dans le volet gauche (clic = aller à la page, page affichée en évidence), **« Page x sur y »** dans la barre d'état, durée de la mise en page et polices remplacées dans le message d'état.
- **Rendu** (`App/Layout/PageRenderer.cs`) : texte posé glyphe par glyphe (`GlyphRun`) avec **exactement les avances de la mise en page** — l'écran montre la coupure calculée ; police de secours pour les caractères absents (☐ ☒, symboles). Justification, couleurs, surlignage, soulignés et barrés, exposant et indice, petites capitales, puces et **numéros de liste** (compteurs par liste et niveau, valeurs de départ, formats décimal, lettres, romains : `Layout/ListNumbering.cs`, premier pas du lot 4), **images en ligne** (a:blip du paquet ; EMF/WMF : cadre gris), **fond et bordures des cellules**, **bordures et trame des paragraphes** (groupes comme Word), **en-têtes et pieds** (numéros PAGE / NUMPAGES calculés), **notes de bas de page** (séparateur et texte).
- **Interactions** : clic sur une case à cocher dans la page (repérage du caractère sous le pointeur), volet de droite qui décrit le paragraphe et le caractère survolés ; après une modification ou une annulation, nouvelle mise en page et position de défilement gardée.
- **Installeur** : OpenWrite se lance à la fin d'une installation faite à la main (double-clic : `UILevel = 5`) — l'installeur n'a pas d'écran de fin et, en 0.2, rien ne semblait se passer ; pas de lancement pendant les mises à jour automatiques (`/passive`), qui relancent l'application elles-mêmes.
- **Auto-test** : 15 vérifications (nombre de pages, case trouvée sur la page et cochée par un vrai clic sur sa position) ; scène de capture `page-N`.
- Pas encore affichés : objets flottants et zones de texte (lot 7), colonnes (suite du lot 2) ; mise en page sur le fil de l'interface (0,8 s pour 700 pages ; en arrière-plan plus tard).


### 10.10 Lot 3, première partie — saisie et mise en forme directe (2026-11-10)

- **Positions** : le texte affiché d'un paragraphe (`ParagraphContent.Text`) porte pour chaque caractère sa plage dans le modèle (`ModelStart`/`ModelEnd`, décalages de `DocPos`) : texte 1 pour 1, codes de champs masqués sautés, une case à cocher ou un numéro de page calculé couvre **tout le champ** (début → fin), un appel de note sa marque. Le point d'insertion (`Layout/Caret.cs`, `TextPoint` = paragraphe + caractère affiché) se convertit en `DocPos` à chaque modification : la frappe va **après le caractère qui précède** (sa mise en forme, comme Word), jamais dans le code d'un champ ; Retour arrière sur une case supprime le champ entier.
- **Géométrie** (`CaretGeometry`) : abscisse d'un point d'insertion avec justification et tabulations, point le plus proche d'une abscisse, ligne du point (fin d'une ligne coupée ou début de la suivante), mots (Ctrl+flèches, double-clic ; apostrophe typographique dans le mot). `DocumentLayout.LinesOf(paragraphe)` (recherche dichotomique de la page, cellules comprises) et `Step` (ligne suivante ou précédente, d'une page à l'autre) servent à la navigation et à la sélection.
- **Saisie** (`MainWindow.Typing.cs`) : frappe (regroupée par mot dans le journal), Entrée, Maj+Entrée, Ctrl+Entrée, Tab, Retour arrière / Suppr (et Ctrl+ : par mot ; fusion avec le paragraphe voisin du même conteneur), flèches, Origine/Fin, Ctrl+Origine/Fin, Pg.préc/Pg.suiv, Maj+déplacement, clic, glisser, double et triple clic, Ctrl+A ; copier / couper / coller **texte** (paragraphes par ligne, `\v` = retour à la ligne comme Word). Entrée en fin de paragraphe : **style suivant** (w:next) ; même style : la marque reprend la mise en forme du dernier caractère (on continue en gras). `TextOps.InsertSpecial`, `InsertMultiline`, `Replace` (sélection remplacée et collage en une seule opération annulable).
- **Mise en forme directe** (`Core/Editing/FormatOps.cs`, `MainWindow.Formatting.cs`) : segments coupés aux bornes de la plage (`SplitRunAt`), propriété posée sur chacun ; gras / italique (avec bCs / iCs) / souligné **retirés si tout le texte visé les a**, posés à « 0 » quand le style les donne encore ; sans sélection : le mot du point d'insertion, entre deux mots : la prochaine frappe (mise en forme en attente). Alignements des paragraphes (déjà aligné ainsi : retour à gauche). Raccourcis de Word **français et anglais** (Ctrl+G/B, I, U, E, J, Ctrl+Maj+G/L, Ctrl+Maj+D/R). Boutons du ruban actifs : Coller, Couper, Copier, Gras, Italique, Souligné, Gauche, Centrer, Droite, Sélectionner.
- **Remise en page** après chaque modification : `UndoJournal.LastChange` (conteneur, premier bloc, nombre de blocs après l'opération, aussi pour annuler / rétablir) → `DocumentLayout.Update` sur les blocs de premier niveau touchés ; dessins des pages gardés sauf pages refaites (tout redessiné si le nombre de pages change : numéros). Les paragraphes **renumérotés** (liste) comptent comme modifiés pour la convergence.
- **Mesures** (`--typing-bench`, vraie fenêtre hors écran, affichage compris) : médiane **1 à 5 ms par touche** sur le corpus, 4,6 ms sur 301 pages (Annexe ×30) ; remise en page seule 0,1–0,3 ms ; pire cas Entrée sur 301 pages ≈ 50 ms (reconstruction du flux). **Word ouvre les 25 documents du corpus** modifiés par OpenWrite (frappe, Entrée, Retour arrière, gras, centrage ; `--save`, `export-reference-pdf.ps1`).
- **Banc d'aller-retour** : la colonne des modifications locales lisait la partie d'origine *après* l'enregistrement (déjà mise à jour en mémoire) et affichait « ÉTENDUE 0→0 » depuis la 0.2 ; corrigé : toutes **locales**.
- **Auto-test** : 26 vérifications (frappe au clavier simulé, Entrée, fusion par Retour arrière, Ctrl+Z avec point d’insertion revenu, gras du mot, alignement, relecture du fichier). 101 tests unitaires.
- Reste du lot 3 : boîtes Police et Paragraphe, taille et police au ruban, règle, presse-papiers RTF / HTML et interne, marques de mise en forme, symboles, liens, rechercher / remplacer, Tab dans un tableau, sélection qui traverse un tableau, saisie dans les en-têtes, pieds et notes.

### 10.11 Lot 3, suite — mise en forme, presse-papiers, recherche, liens, règle (2026-11-10)

- **Volet Mise en forme actif** (`MainWindow.FormatPane.cs`) : police, taille, G / I / S / barré, alignements, interligne, boîtes Police et Paragraphe ; il montre l'état du point d'insertion. Mis à jour **après la rafale de touches** (priorité de fond) et seulement si les valeurs affichées changent : sinon +2 ms par touche et des pointes de 450 ms (zone de liste des polices), mesuré par `--typing-bench` (les formats résolus sont recréés à chaque remise en page : comparer les valeurs, pas les objets).
- **Boîtes Police** (Ctrl+D) et **Paragraphe** (`FontWindow`, `ParagraphWindow`, briques `DialogKit` reprises d'OpenPub) : seuls les réglages changés sont appliqués ; attribut retiré mais donné par le style → « 0 ». `FormatOps.SetIndent/SetSpacing/SetParagraphFlag/ClearCharacter/ClearParagraph` ; Ctrl+Espace, Ctrl+Q, agrandir / réduire par paliers de Word.
- **Marques de mise en forme** (Ctrl+Maj+8, `PageRenderer.ShowMarks`) aux positions exactes de la mise en page.
- **Rechercher / remplacer** (Ctrl+F, Ctrl+H, `MainWindow.Find.cs`) dans le texte affiché ; `TextOps.ReplaceText` : le nouveau texte prend la mise en forme du **premier caractère remplacé** (inséré après lui, puis l'ancien retiré) ; Remplacer tout = une transaction.
- **Presse-papiers avec mise en forme** (`MainWindow.Clipboard.cs`, `Core/Editing/Fragments.cs`, `Core/Interop/RtfReader.cs`, `HtmlReader.cs`) : copie = texte + RTF + HTML (Word relit le RTF : gras, taille, police, vérifié par COM) + morceau du modèle (jeton privé) pour le même document ; collage = morceau interne > RTF > HTML > texte, Ctrl+Maj+V texte seul. Un morceau garde sa marque finale ou non (`EndsWithParagraph`) : un dernier paragraphe partiel se fond dans le paragraphe d'arrivée et prend ses propriétés, un premier paragraphe entier remplace une moitié gauche vide. Signets, commentaires, appels de notes et sauts de section ne sont pas dupliqués ; nouveaux w14:paraId. **Piège** : juste après une écriture, le presse-papiers annonce un format dont `GetData` renvoie `null` ou **une chaîne vide** — lecture avec réessais (`ReadString`), vu en auto-test (1 échec sur 4 avant correction, 12/12 après).
- **Liens** (`LinkOps`, `MainWindow.Links.cs`, `HyperlinkWindow`) : Ctrl+K (sélection ou nouveau texte, modifier, supprimer), bulle d'adresse, Ctrl+clic (web et courriel directement, autres schémas après confirmation, signets) ; `DocxPackage.AddExternalRelationship` (relation reprise si identique) ; style « Hyperlink » du document sinon bleu souligné. Word relit le lien.
- **Tableaux** : Tab / Maj+Tab sélectionnent la cellule suivante / précédente ; **sélection qui sort d'un tableau** : `TextOps.AcrossTopLevel/DeleteAcross`, `FormatOps.Normalize` — le tableau touché est pris en entier (supprimé, copié, mis en forme), comme des rangées sélectionnées dans Word ; plusieurs cellules d'un même tableau : lot 6.
- **Symbole** (fenêtre non modale), Ctrl+Maj+Espace, Ctrl+Maj+- ; **règle** (`Layout/RulerBar.cs`, `MainWindow.Ruler.cs`) : retraits (première ligne, suspendu, gauche, droit) et taquets à la souris, colonne ou cellule du paragraphe, Affichage ▸ Règle ; taquets écrits avec « clear » pour ceux du style retirés.
- **Ruban** : boutons actifs déclarés par onglet (`HomeCommands`, `InsertCommands`, `LayoutCommands`).
- Mesures : médiane 1 à 5 ms par touche sur le corpus (affichage compris), 5 ms sur 301 pages ; Word ouvre les 25 documents modifiés. Auto-test **39 vérifications**, **113 tests**.
- Reportés à leur lot : saisie dans les en-têtes et pieds (lot 5, 0.6), les notes (lot 9), sélection sur plusieurs cellules (lot 6), images collées depuis un autre programme.

### 10.12 Version 0.4.0 (2026-11-10)

Publiée par la pipeline (tag `v0.4.0`) : lot 3 tel que décrit en §10.10 et §10.11 — saisie, sélection, mise en forme directe (volet, boîtes Police et Paragraphe, règle), presse-papiers RTF / HTML / interne, rechercher / remplacer, liens, symboles, marques de mise en forme, Tab dans les tableaux. Accueil et message d'ouverture mis à jour ; jalon 0.4 retiré de la liste « Prochaines versions ».

### 10.13 Lot 4 — styles, listes et thèmes (2026-11-10)

- **Parties XML modifiables** : `DocxPackage.XmlPartOf("styles" | "numbering" | "theme", create)` charge la partie une fois (les modifications s'y font en place ; `StyleSheet()` lit ces mêmes objets), la crée au besoin (entrée, relation, type de contenu ; jamais le thème) et ne la réécrit que si elle a changé — un document ouvert puis enregistré sans modification reste identique. **Journal** : `EditTransaction.TrackXml(root)` prend un instantané ; annuler / rétablir / échec restaurent le contenu **sans changer l'objet racine** (les références restent valables) ; `XmlChanged` (partie à réécrire), `LastXml`.
- **Styles** (`Core/Styles/StyleOps.cs`) : reconnus par leur **nom intégré** (« heading 1 ») et non leur identifiant (« Titre1 » dans un Word français) ; noms français affichés ; définitions des styles intégrés de Word 2013+ (titres, titre, sous-titre, citations, paragraphe de liste, sans interligne, élevé, accentuations, lien) **écrites au premier usage**, comme Word, avec les polices et couleurs du thème ; nouveau style / mise à jour depuis la sélection (`Capture`, `Merge`, `Absorb`), propriétés d'un style modifiables par les boîtes Police et Paragraphe (`Editable` : w:pPr / w:rPr du style vus comme ceux d'un paragraphe).
- **Listes** (`ListOps`) : puces et numéros écrits comme les définitions par défaut de Word (Symbol / Courier New / Wingdings ; décimal / lettres / romains), liste de titres 1. / 1.1 / 1.1.1 liée aux styles Titre 1–3, recommencer (`lvlOverride/startOverride`), continuer, niveaux ; un paragraphe Normal mis en liste prend « Paragraphe de liste » (Word).
- **Fenêtre** (`MainWindow.Styles.cs`, `.Theme.cs`, `StyleWindow`, `StylesWindow`) : galerie, volet Style, boîtes Styles et Modifier le style, Ctrl+Alt+1/2/3, Ctrl+Maj+N, Ctrl+Maj+L, Tab / Maj+Tab en tête d'élément, Entrée sur un élément vide et Retour arrière en tête : fin de liste, menu contextuel (Recommencer à 1, Continuer) ; Création ▸ Thèmes / Polices / Couleurs (`ThemeOps`), « +Titres » / « +Corps » dans la liste des polices.
- **Remise en page** : styles modifiés, thème → résolveur refait et document entier remis en page (≈ 1,4 s sur 301 pages) ; styles ou listes seulement **ajoutés** (`_xmlAdditive`) → `DocumentLayout.ReplaceStyles` et remise en page incrémentale (Titre 2 sur 301 pages : 1 316 → 89 ms ; test : identique à une mise en page complète) ; listes touchées → reconstruction du flux pour les numéros (`Update(rebuildFlow)`).
- Vérifié : Word relit Titre 1, Paragraphe de liste et puces à deux niveaux (COM) ; les 25 documents du corpus modifiés (Titre 2, numérotation, couleurs du thème en plus du lot 3) s'ouvrent dans Word. Auto-test **45 vérifications**, **121 tests**.
- Reporté : inspecteur de styles, jeux de styles, « définir par défaut » (modèle Normal de l'utilisateur), styles de tableau (lot 6), liste à plusieurs niveaux personnalisée (boîte « Définir une nouvelle liste »).

### 10.14 Version 0.5.0 (2026-11-10)

Publiée par la pipeline (tag `v0.5.0`) : lot 4 tel que décrit en §10.13 — galerie et boîtes de styles, styles intégrés écrits au premier usage, modifier / créer un style, puces, numéros et titres numérotés, niveaux et fin de liste, recommencer / continuer, thèmes (polices et couleurs). Jalon 0.5 retiré de la liste « Prochaines versions ».

### 10.15 Lot 5 — mise en page, sections, en-têtes et pieds (2026-11-10)

- **Colonnes** (moteur) : `SectionGeometry.Cols` (w:cols : égales, inégales w:col, trait w:sep) ; le curseur de pagination a une colonne courante (`ColX`, `ColW`, `ColTop`) ; un débordement ou un saut de colonne passe à la colonne suivante, un saut de page / de section à la page suivante (`NewPage(…, toColumn)`) ; une **section continue en colonnes est équilibrée** avant la suivante (`Balance` : hauteur la plus petite qui tient dans les N colonnes ; zone avec tableau laissée) ; traits entre colonnes (`Page.ColumnRules`). Horaires : 2 pages comme Word (5 avant) ; ses écarts verticaux viennent des objets flottants (lot 7). Débuts de page : 61/75.
- **En-têtes hérités** d'une section à l'autre (`SectionGeometry.HeaderRefs`, clés « header:default »…), **pages paires** (`w:evenAndOddHeaders`, settings.xml devenu partie modifiable), **numérotation par section** (`w:pgNumType` start / fmt → `Page.Number`, `NumberText`, `SectionPageCount` ; champ SECTIONPAGES), numéros de ligne (`w:lnNumType`, `NumberedLines` : corps hors tableaux, sans w:suppressLineNumbers), filigrane lu dans l'en-tête (`Watermark` : forme VML à v:textpath).
- **SectionOps** : section du point d'insertion, taille, **orientation** (marges tournées comme Word — vérifié par COM : haut ← gauche, bas ← droite, gauche ← bas, droite ← haut), marges, colonnes (dont « gauche » / « droite »), numérotation, bordures de page, numéros de ligne, **saut de section** (la section d'avant garde propriétés et en-têtes, celle d'après prend le type et hérite des en-têtes).
- **Édition des en-têtes et pieds** (`MainWindow.HeaderFooter.cs`) : double-clic dans la marge (créé au besoin : `DocxPackage.CreateHeaderPart`, style En-tête / Pied de page, référence dans la section), corps grisé et limites tracées (`PageRenderer.HeaderMode`), Échap / double-clic dans le corps pour sortir ; la saisie passe par une **source de lignes** (`ILineSource` : corps ou `StoryLines` de l'en-tête sur la page éditée) et une **racine d'édition** (`EditRoot`) — tout le code de saisie, mise en forme, presse-papiers, styles fonctionne dans l'en-tête ; hauteur inchangée → pages seulement redessinées, sinon remise en page. Onglet contextuel : N° de page (`FieldOps.Insert` : champ complexe PAGE, NUMPAGES, « Page X sur Y », format des numéros), Atteindre, 1re page différente, paires et impaires, **lier au précédent** (retire la référence propre ou copie l'en-tête hérité dans une nouvelle partie), position.
- **Journal** : `EditTransaction.TrackStory` (autres suites de blocs dans la même opération : un filigrane dans tous les en-têtes s'annule d'un seul Ctrl+Z).
- **Onglets** : Mise en page (Marges, Orientation, Taille, Colonnes, Sauts, Numéros de ligne, boîte Mise en page) ; Création (Filigrane VML écrit comme Word, Couleur de page w:background + w:displayBackgroundShape, Bordures de page).
- Vérifié : Word relit l'en-tête créé (texte, champ PAGE, style En-tête), le filigrane (forme BROUILLON tournée à 315°), la couleur de page ; les 25 documents du corpus modifiés (en-tête édité ou créé avec numéro de page, deux colonnes, en plus des lots 3 et 4) s'ouvrent dans Word. Auto-test **54 vérifications**, **125 tests**.
- Reporté : filigrane image, BorderArt (bordures artistiques), position exacte des colonnes dans les documents à objets flottants, notes continuées sur la page suivante.

### 10.16 Version 0.6.0 (2026-11-10)

Publiée par la pipeline (tag `v0.6.0`) : lot 5 tel que décrit en §10.15 — colonnes et sections, onglet Mise en page, en-têtes et pieds modifiables en place avec numéros de page, première page / paires et impaires / lier au précédent, filigrane, couleur et bordures de page, numéros de ligne. Jalon 0.6 retiré de la liste « Prochaines versions ».

### 10.17 Lot 6 — tableaux (2026-11-10)

- **Calibration** : le corpus n'a que 7 tableaux ; `tools/word-make-tables.ps1` fait écrire par Word (COM) des documents neutres de calibration — styles intégrés et options de style, fusions horizontales et verticales, alignement vertical, marges de cellule, tableau imbriqué, tableaux centré et en retrait, « ajuster au contenu » / « à la fenêtre », tableau de 90 rangées avec en-tête répété et rangée insécable, galerie des 36 styles. `tools/word-pages.ps1` relève aussi l'abscisse de chaque paragraphe ; le banc `--layout` donne l'écart horizontal (Δx) ; `tools/pdf-to-png.ps1` convertit les PDF de Word en images (lecteur PDF de Windows) pour comparer les rendus.
- **Styles de tableau conditionnels** (`Core/Styles/TableStyles.cs`, `StyleResolver.TableLevel`) : régions d'une cellule (ligne d'en-tête, ligne des totaux, première / dernière colonne, bandes, coins) d'après `w:tblLook` et la taille des bandes — **139/140 identiques au `w:cnfStyle` que Word écrit** (le dernier écart : ordre des bits des coins, ne/nw/se/sw) ; pPr / rPr / tcPr du style puis de ses `w:tblStylePr` dans l'ordre d'application (bandes < colonnes < lignes < coins).
- **Mise en page** (`Layout/Tables.cs`, partielle de `DocumentLayout`) : grille complétée par `w:tcW` ; tableau « ajusté au contenu » (`w:tblW` et tous les `w:tcW` explicitement automatiques) recalculé d'après les largeurs minimale (plus long mot) et maximale (ligne sans coupure) des cellules : colonnes qui tiennent dans leur part à leur largeur, le reste partagé ; `w:tblW` en % ou en points ; **retrait selon le mode de compatibilité** (Word 2013+ : `w:tblInd` jusqu'au bord du tableau ; avant : jusqu'au texte de la première cellule) ; `w:jc` ; marges de cellule (`w:tblCellMar`, `w:tcMar`, style) — **la marge haute la plus grande d'une rangée vaut pour toutes ses cellules** (mesuré) ; alignement vertical ; fusion verticale (la dernière rangée grandit) ; tableaux imbriqués (la marque de fin de cellule vide après un tableau imbriqué n'a pas de hauteur — mesuré) ; bordures résolues cellule > exceptions de rangée > tableau > conditionnel > style, avec couleur et style (double compté 3 fois l'épaisseur) ; veuves et orphelines dans les rangées coupées.
- **Veuves et orphelines actives par défaut** : `w:widowControl` absent = actif (Word, vérifié par COM : 286/286 paragraphes de l'Annexe ; ECMA dit l'inverse). Le tableau long passe de 71/176 à 182/182 positions à 1 pt.
- **Banc** (documents de calibration) : ajustement 28/28 positions à 1 pt (x et y), fusions 55/55, tableau long 182/182, styles 132/148 en y et 148/148 en x ; avant le lot, Δx systématique de −5,8 pt (retrait). Corpus : inchangé (63/76 débuts de page).
- **Opérations** (`Core/Editing/TableOps.cs`, `TableOps.Map` = grille avec `w:gridBefore`, `w:gridSpan`, fusions) : créer (comme Word : largeur du texte − 10 twips partagée, `w:tblLook` 04A0), insérer au point d'insertion (paragraphe coupé), rangées et colonnes (fusions prolongées, ramené à la largeur disponible), supprimer, **fusionner** (texte de toutes les cellules dans la première, `w:gridSpan` + `w:vMerge`), **fractionner** (nouvelles colonnes de grille, les autres rangées s'élargissent ; nouvelles rangées fusionnées verticalement ailleurs), fractionner le tableau, limites de colonnes (la voisine compense), hauteurs, distribuer, trame, bordures (menu de Word : toutes, extérieures, intérieures…), alignement 3 × 3, marges, style et options, en-tête répété, fractionnement des rangées, ajustement automatique, texte ↔ tableau, tri (nombres ou texte en ordre français), formules (`=SUM(ABOVE)`… calculées, champ inséré).
- **Styles de la galerie** (`TableStyleOps`, `TableStyleDefs`) : Grille du tableau + Grille 1 Clair, Grille 4, Grille 5 Foncé, Liste 3, Liste 4 (sans accent et accents 1–6) écrits au premier usage avec « Tableau Normal ». Couleurs recalculées pour le thème du document ; Word tronque (et n'arrondit pas) les nuances de ses styles de tableau, et 6 valeurs du thème Office sont des constantes qu'aucune formule ne retrouve : gardées telles quelles. **36/36 identiques à ceux de Word** (`OpenWrite.Fidelity --table-styles galerie.docx`).
- **Fenêtre** (`MainWindow.Tables.cs`, `TableWindows.cs`) : Insertion ▸ Tableau (grille 10 × 8, boîte, texte → tableau) ; **sélection de cellules** (glisser d'une cellule à l'autre : rectangle de la grille, fusions comprises ; mise en forme cellule par cellule, frappe et suppression qui vident les cellules, copie en tableau HTML et en morceau interne) ; Tab dans la dernière cellule ajoute une ligne ; Ctrl+Maj+Entrée fractionne le tableau ; onglets contextuels **Création de tableau** (options de style, galerie avec aperçus aux couleurs du thème, trame, bordures et trait) et **Disposition** (sélectionner, quadrillage, propriétés, insérer / supprimer, fusionner / fractionner, ajustement, distribuer, aligner, trier, répéter l'en-tête, convertir en texte, formule) ; menu contextuel ; **bordures tirées à la souris** (largeur de colonne, hauteur de ligne) ; quadrillage en tirets des cellules sans bordure ; collage d'un tableau HTML (Excel) en vrai tableau.
- **Frappe dans un tableau** : seuls les paragraphes touchés sont recoupés (`DocumentLayout.Update(…, dirty)`), la reprise commence à la page de la rangée modifiée et converge dans le tableau après elle (test : identique à une mise en page complète). Tableau de 90 rangées ×5 : 31 → 9,6 ms par touche (affichage compris). 301 pages : 5,5 ms.
- Vérifié : Word ouvre les 25 documents du corpus modifiés par le banc de saisie (tableau inséré et stylé, ligne fusionnée, ligne et colonne insérées, cellule fractionnée, en plus des lots 3 à 5) et le document de l'auto-test (tableau stylé à 3 colonnes). Auto-test **64 vérifications**, **141 tests**.
- **Piège** : les variables d'environnement de WSL n'arrivent pas au .NET de Windows ; `OPENWRITE_*` doivent être listées dans `WSLENV` (ex. `WSLENV=OPENWRITE_BENCH_DUMP`).
- Reporté : orientation verticale du texte dans une cellule, tableaux flottants habillés (lot 7, objets flottants), dessin de tableau et gomme, règle avec les colonnes du tableau, sélection d'une cellule / ligne / colonne par la marge à la souris (le menu Sélectionner le fait), espacement entre cellules (`w:tblCellSpacing`), nouveaux styles de tableau.

### 10.18 Version 0.7.0 (2026-11-10)

Publiée par la pipeline (tag `v0.7.0`) : lot 6 tel que décrit en §10.17 — tableaux insérés et modifiés comme dans Word (sélection de cellules, fusionner, fractionner, lignes et colonnes, largeurs à la souris), galerie des styles de tableau, trame et bordures, trier, formules, texte ↔ tableau, collage depuis Excel ; mise en page des tableaux recalibrée sur Word. Jalon 0.7 retiré de la liste « Prochaines versions ».

### 10.19 Collage depuis Excel (2026-11-10, après 0.7.0)

Retour utilisateur : une plage Excel collée arrivait en paragraphes (une valeur par ligne). Cause : Excel place `<table>` **avant** `<!--StartFragment-->`, le fragment ne contenait donc pas de tableau et le RTF (qui le met à plat) était choisi. `Core/Interop/HtmlTableReader` remplace l'ancien lecteur de l'application : tableau cherché dans tout le HTML, CSS des `<style>` (td, th, table, tr, `.classe`, `td.classe`) puis styles en ligne, classes des balises internes (`<font class=font5>`) recopiées en style. Calibré sur le collage de Word par COM (`Range.Paste` d'une plage Excel, `WordOpenXML`) : grille et `tblW` dxa depuis `<col>` (pt × 20), `trHeight` « au moins » depuis la hauteur de la rangée, `tblCellMar` gauche/droite depuis `cellpadding`, `tcMar` depuis `padding` (bas 0 quand absent), bordures par côté (0,5 pt → `sz` 4, 1 pt → 8, `none` → `nil`, `windowtext` → `auto`), `vAlign` (middle → center, bottom), `jc` (general ignoré), `noWrap`, trame, polices et couleurs ; marque de paragraphe à la mise en forme de la cellule. Écart volontaire : Word omet la taille 12 pt (texte retombe à la taille par défaut), OpenWrite garde celle d'Excel. Aussi colspan/rowspan (gridSpan/vMerge), `<th>` en ligne d'en-tête, paragraphes autour d'un tableau de page web gardés, grille simple seulement sans CSS de bordure et sans `border=0`. Banc : `OpenWrite.Fidelity --paste-html <docx> <html> <sortie>` ; le document obtenu s'ouvre dans Word avec le même rendu. Self-test 66 vérifications (collage d'un HTML de forme Excel par le vrai presse-papiers, annulation).

### 10.20 Version 0.7.1 (2026-11-10)

Publiée par la pipeline (tag `v0.7.1`) : correctif du collage depuis Excel (§10.19).

### 10.21 Lot 7 — images, formes et zones de texte (2026-11-10)

**Modèle.** `DrawingContent` (contenu de segment) remplace l'îlot opaque pour `w:drawing`, `w:pict`, `w:object` et `mc:AlternateContent` : XML d'origine dont les `w:txbxContent` sont vidés, leur texte devient des `Story` (`TextBoxes`, dans l'ordre du document ; la variante VML de repli a ses propres histoires, recopiées depuis la variante DrawingML après chaque modification). Aller-retour identique sur tout le corpus, texte modélisé à 100 % zones comprises.

**Lecture** (`Core/Drawing`). `DrawingReader` → `DrawingObject` (cadre, marges d'effet, repères et alignements ou décalages, habillage et côté, distances, polygone, derrière le texte, ordre, taille relative) et arbre de `Visual` (`PictureVisual` : rognage, découpe, corrections ; `ShapeVisual` : géométrie prédéfinie / libre / chemins VML, remplissage uni ou dégradé ou image, trait, tirets, flèches, trait composé, zone de texte et ses marges, ancrage du texte, WordArt ; `GroupVisual` : wpg, wpc, v:group avec repères enfants ; `FallbackVisual` : graphique, SmartArt, OLE). Couleurs du thème avec lumMod/lumOff/tint/shade/alpha, styles de forme (lnRef, fillRef) depuis le fmtScheme du thème. VML : style CSS (position, marges, repères mso-position-*), w10:wrap, z-index négatif = derrière, filets `o:hr`, types `o:spt` usuels, chemins sans formules. `ShapeGeometry` évalue les **187 formes prédéfinies officielles** (presetShapeDefinitions.xml d'ECMA-376, copie Apache POI) et les géométries libres : langage des guides complet, arcs en Béziers. `--drawings` (banc Fidelity) : 62 dessins du corpus lus, 0 illisible.

**Mise en page** (`Layout/Floats.cs`). Objets ancrés placés avant les lignes de leur paragraphe (page, marge, colonne, caractère, marges ; paragraphe, ligne) ; texte habillé ligne par ligne dans les segments libres (`LineSpace` → `WrapRow` : rangées de part et d'autre d'un objet, `LineBox.Gap`, `SameRow`, `BoxLeft`), haut et bas = la ligne descend ; un objet ancré plus bas qui recouvre du texte déjà placé fait refaire la page (au plus 3 fois) ; suite d'un paragraphe coupé habillée sur la page suivante (`Layout(..., resume)`). Objets des en-têtes et pieds sur chaque page (ils écartent aussi le corps), objets des cellules placés par rapport à la cellule, tableaux flottants `w:tblpPr` placés et habillés. Contenu des zones de texte mis en page dans leur cadre (`FloatBox.TextLines`).

**Calibrage par COM.** `tools/word-make-drawings.ps1` (documents neutres : carré à gauche / droite / milieu, un seul côté, haut et bas, rapproché, retrait, distances, ancre plus bas, page, devant, derrière, zone de texte, tableaux flottants) et `tools/word-lines.ps1` (Pages ▸ Rectangles ▸ Lines : haut, gauche, largeur, longueur de chaque ligne) ; `OpenWrite --lines` compare texte, longueur, page, haut et bord gauche : **456/480 lignes identiques**. Constats : bords d'objet arrondis au point (un mot de plus tient à droite d'un objet au bord à 372,3 pt) ; Word donne le bord de la boîte de ligne (marge, ou bord libre près d'un objet), pas le début du texte. Écart restant : une ligne pleine largeur après un tableau flottant où Word loge 2,5 pt de plus — une tolérance générale dégrade le reste (447 à 2 pt, 380 à 3 pt) : à étudier avec le lot 13.

**Corrections de mise en page trouvées en route** (Horaires) : échelle horizontale des caractères `w:w` ; avec un retrait négatif de première ligne, le retrait gauche sert de taquet ; ligne qui ne porte qu'un saut de colonne gardée en bas de colonne ; paragraphe vide qui porte un saut de section continu posé au bout de la ligne précédente, sans hauteur (mesuré : x = 658 pt) ; colonnes à saut explicite non équilibrées, section suivante sous la plus longue colonne. Horaires : 2 pages comme Word, 22/30 paragraphes à 1 pt (10/30 avant) ; corpus : 63/76 débuts de page (inchangé).

**Rendu** (`App/Layout/DrawingRenderer.cs`) : images (rognage, découpe, gris, transparence), formes, groupes, rotation et retournements (le texte des zones tourne avec), flèches, objets derrière / devant le texte, objets d'en-tête sous le corps ; images EMF et WMF jouées par GDI (`Metafiles.cs`) ; filets horizontaux.

**Édition** (`MainWindow.Objects.cs`, `Core/Drawing/DrawingOps.cs`). Sélection au clic (poignées, rotation, ancre ⚓), glisser pour déplacer (ancre au paragraphe le plus proche), redimensionner (proportions aux coins d'une image, Maj), pivoter (Maj : 15°) ; Suppr, flèches, Échap ; couper / copier / coller (identifiants `wp:docPr` renumérotés) ; onglet Format de l'image / de la forme (corrections, couleur, remplissage, contour, forme, texte de remplacement, position, habillage, avancer / reculer, rotation, rogner, taille) ; Insertion ▸ Images (fichier, presse-papiers ; image seule collée avec Ctrl+V), Formes (35), Zone de texte ; Mise en page ▸ Position, Habillage ; écrire dans une zone de texte (`BoxEdit` : lignes de la zone, remise en page du paragraphe d'ancrage). `DocxPackage.AddImage`, `EnsureDrawingNamespaces`. Formes et zones créées comme Word (wps dans mc:AlternateContent, style du thème) : Word ouvre le document du self-test et affiche la zone de texte d'OpenWrite habillée.

**Presse-papiers** : juste après une écriture, Windows peut rendre un tampon de la bonne taille rempli de zéros (collage Excel qui retombait parfois sur le RTF, test de copie instable) : relu.

**Tests** : `DrawingTests` (formes, lecture DrawingML et VML, segments libres, habillage, objet de cellule, modifications validées par le SDK) — 151 tests ; self-test 77 vérifications (image, habillage, déplacement, taille, annulation, forme, zone de texte, relecture).

**Reporté** : graphiques et SmartArt dessinés (cadre de repli ; la partie dessin `dsp` d'un SmartArt n'est pas encore lue), effets (ombre, lueur, 3D), habillage dans une cellule, cadres `w:framePr`, rognage à la souris, zones de texte liées, lettrine.

### 10.22 Version 1.0.0 (2026-11-10)

Publiée par la pipeline (tag `v1.0.0`) : lot 7 tel que décrit en §10.21 — images, formes et zones de texte affichées à leur place avec le texte habillé comme dans Word, insertion et modification des objets, tableaux flottants, images EMF/WMF. Jalon 1.0 atteint ; « Prochaines versions » de l'accueil : 1.1 à 1.3.
