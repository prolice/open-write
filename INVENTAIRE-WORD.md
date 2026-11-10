# Inventaire des fonctions de Microsoft Word (Microsoft 365, novembre 2026)

*Établi le 2026-11-05 comme base du plan d'OpenWrite (`PLAN.md`), sur le modèle du travail fait pour OpenPub (PLAN.md §8 d'OpenPub). Les fonctions sont classées par onglet du ruban, puis par thème transversal.*

## Légende

**Priorité** (pour un traitement de texte de bureau utilisé hors ligne, en français / néerlandais / anglais) :

| Code | Sens |
|---|---|
| **P0** | Indispensable : sans elle, OpenWrite ne remplace pas Word, même pour une lettre |
| **P1** | Usage courant (rapports, courriers, comptes rendus) : attendu pour la 1.0 ou juste après |
| **P2** | Usage avancé (documents longs, travail à plusieurs, mémoires) |
| **P3** | Rare : à confirmer avant de l'engager |
| **HP** | Hors périmètre : service en ligne Microsoft, IA, technologie propriétaire ou fermée. Le contenu correspondant est **conservé** dans le `.docx` quand c'est possible |

**Reprise OpenPub** : ✔ code repris presque tel quel · ◐ repris et adapté · ✖ à écrire · — sans objet.

**Lot** : numéro du lot dans `PLAN.md` §8.

---

## 1. Fichier (Backstage)

| Fonction | Détail | Prio | OpenPub | Lot |
|---|---|---|---|---|
| Nouveau document vierge | Modèle « Normal » de l'utilisateur (police, marges, styles par défaut) | P0 | ◐ | 3 |
| Nouveau à partir d'un modèle | Modèles `.dotx` de l'utilisateur et quelques modèles fournis (lettre, rapport, CV, compte rendu) | P1 | ◐ `Templates` | 14 |
| Modèles en ligne | Catalogue Microsoft | HP | — | — |
| Ouvrir | Récents, épinglés, parcourir ; glisser-déposer sur la fenêtre | P0 | ✔ | 0 |
| Enregistrer / Enregistrer sous | `.docx` natif, autres formats (§14) ; enregistrement sûr (temporaire + remplacement atomique + `.bak`) | P0 | ✔ | 1 |
| Enregistrement automatique local | Copie de secours périodique, récupération après plantage, « Récupérer les documents non enregistrés » | P0 | ✔ `RecoveryStore` | 0 |
| Enregistrement automatique cloud (AutoSave) | OneDrive / SharePoint | HP | — | — |
| Historique des versions | Versions OneDrive | HP | — | — |
| Fichier modifié ailleurs | Avertissement (OneDrive synchronisé, autre PC) | P1 | ✔ | 0 |
| Info ▸ Propriétés | Titre, auteur, sujet, mots-clés, commentaires, dates, statistiques ; propriétés personnalisées | P1 | ✖ | 10 |
| Info ▸ Protéger le document | Toujours ouvrir en lecture seule, chiffrer avec mot de passe, restreindre la modification, marquer comme final | P2 | ✖ | 8 |
| Info ▸ Signature numérique | Signature XML-DSig du package | P3 | ✖ | 15 |
| Info ▸ Inspecter le document | Supprimer commentaires, révisions, propriétés personnelles, texte masqué | P2 | ✖ | 14 |
| Info ▸ Vérifier l'accessibilité | Texte de remplacement manquant, titres, contraste, tableaux sans en-tête | P2 | ◐ | 14 |
| Info ▸ Vérifier la compatibilité | Fonctions perdues dans les anciens formats | P3 | ✖ | 13 |
| Imprimer | Aperçu, imprimante, pages (plages, sélection), copies, recto verso (manuel), assemblé, orientation, format, marges, pages par feuille, mise à l'échelle | P0 | ✔ `PrintService` | 12 |
| Partager / Envoyer par e-mail | Lien OneDrive, pièce jointe Outlook | HP | — | — |
| Exporter en PDF / XPS | PDF avec signets des titres, liens, PDF balisé, PDF/A ; XPS | P0 (PDF) / HP (XPS) | ◐ `DrawingToPdf` | 12 |
| Transformer (Sway), Publier sur un blog | | HP | — | — |
| Compte, Office Theme | Connexion Microsoft ; thème clair/sombre/système de l'interface | P2 (thème) | ◐ | 14 |
| Options ▸ Général | Nom et initiales de l'utilisateur (auteur des commentaires et révisions), mini barre, aperçu instantané | P1 | ✖ | 14 |
| Options ▸ Affichage | Marques de mise en forme, impression des couleurs d'arrière-plan, texte masqué | P1 | ✖ | 14 |
| Options ▸ Vérification | Options de correction automatique, vérification pendant la frappe, mots en majuscules, dictionnaires personnels | P1 | ◐ | 8 |
| Options ▸ Enregistrement | Format par défaut, intervalle de récupération, dossier par défaut, incorporer les polices | P1 | ◐ | 14 |
| Options ▸ Langue | Langue de l'interface (français uniquement), langues de vérification | P1 | ◐ | 8 |
| Options ▸ Options avancées | Coller (comportement par défaut), unités (cm), taille d'image par défaut, compatibilité de mise en page | P2 | ✖ | 14 |
| Personnaliser le ruban / Barre d'outils Accès rapide | | P2 (QAT) / P3 (ruban) | ✖ | 14 |
| Compléments (COM, Office JS), Centre de gestion de la confidentialité | | HP | — | — |

## 2. Accueil

| Fonction | Détail | Prio | OpenPub | Lot |
|---|---|---|---|---|
| Couper / Copier / Coller | Texte riche entre documents OpenWrite et **avec Word, Outlook, navigateurs** (formats HTML, RTF, texte) | P0 | ◐ | 3 |
| Options de collage | Conserver la mise en forme source, fusionner, texte seul ; collage spécial ; coller une image | P0 | ◐ | 3 |
| Volet Presse-papiers Office | 24 derniers éléments | P3 | ✖ | 14 |
| Reproduire la mise en forme | Simple clic / double clic (persistant) ; Ctrl+Maj+C / Ctrl+Maj+V | P1 | ◐ | 3 |
| Police, taille, augmenter / diminuer | Listes avec aperçu des polices, polices récentes, polices du thème | P0 | ✔ | 3 |
| Gras, italique, souligné | Styles de soulignement (simple, double, mots seuls, pointillé, ondulé…) et couleur du soulignement | P0 | ◐ | 3 |
| Barré, double barré, indice, exposant | | P0 | ✔ | 3 |
| Changer la casse | Phrase, minuscules, MAJUSCULES, 1re lettre des mots, inverser (Maj+F3) | P1 | ✖ | 3 |
| Effacer la mise en forme | Ctrl+Espace (caractère), Ctrl+Q (paragraphe) | P0 | ◐ | 3 |
| Surlignage | 15 couleurs Word | P0 | ✔ | 3 |
| Couleur de police | Couleurs du thème, standard, personnalisées, dégradé | P0 | ◐ | 3 |
| Effets de texte | Contour, ombre, reflet, lumière ; ligatures, chiffres, ensembles stylistiques OpenType | P2 | ◐ WordArt | 7 |
| Boîte Police | Petites majuscules, tout en majuscules, masqué ; espacement, échelle, position, crénage | P1 | ◐ | 3 |
| Puces, numérotation | Bibliothèques, définir une nouvelle puce / un nouveau format, recommencer / continuer la numérotation, définir la valeur | P0 | ◐ | 4 |
| Liste à plusieurs niveaux | Définitions `abstractNum` / `num`, liaison aux styles de titre (« 1.1.2 ») | P0 | ✖ | 4 |
| Retraits | Augmenter / diminuer, gauche, droite, première ligne, suspendu ; via la règle | P0 | ✔ | 3 |
| Trier | Paragraphes ou lignes de tableau, alphabétique / numérique / date | P2 | ✖ | 6 |
| Afficher les marques de mise en forme | ¶, espaces, tabulations, sauts, texte masqué, ancres | P0 | ✖ | 3 |
| Alignement | Gauche, centré, droite, justifié (+ distribué) | P0 | ✔ | 3 |
| Interligne et espacement | 1 / 1,15 / 1,5 / 2 / au moins / exactement / multiple ; avant / après ; « ne pas ajouter d'espace entre paragraphes du même style » | P0 | ✔ | 3 |
| Trame de fond, bordures de paragraphe | Bordures haut / bas / côtés / entre, ligne horizontale | P1 | ◐ | 3 |
| Boîte Paragraphe ▸ Enchaînements | Éviter veuves et orphelines, paragraphes solidaires, lignes solidaires, saut de page avant, supprimer les numéros de ligne, ne pas couper les mots | P0 | ✖ | 2 |
| Tabulations | Gauche, centrée, droite, décimale, barre ; points de suite ; tabulations par défaut | P0 | ✔ | 3 |
| Galerie de styles | Styles rapides, aperçu instantané | P0 | ◐ `StyleOps` | 4 |
| Volet Styles | Tous les styles, nouveau, modifier, inspecteur de styles, gérer les styles, organiseur | P1 | ◐ | 4 |
| Rechercher | Volet de navigation, surlignage des résultats, Ctrl+F | P0 | ✔ `TextSearch` | 3 |
| Remplacer | Respecter la casse, mot entier, caractères génériques, sons, formes de mots ; recherche de **mises en forme** et de **caractères spéciaux** (^p, ^t, ^m…) | P0 (base) / P1 (avancé) | ◐ | 3 / 8 |
| Atteindre | Page, section, ligne, signet, commentaire, note, champ, tableau, graphique, titre | P1 | ✖ | 14 |
| Sélectionner | Tout, objets, texte de mise en forme similaire ; volet Sélection | P1 | ◐ | 3 |
| Dicter | Reconnaissance vocale | P3 (Windows) / HP (cloud) | ✖ | 15 |
| Éditeur (Microsoft Editor) | Grammaire, clarté, concision, score | P2 (grammaire locale) / HP (score IA) | ✖ | 8 |
| Copilot, Réutiliser des fichiers, Sensibilité (étiquettes MIP) | | HP | — | — |

## 3. Insertion

| Fonction | Détail | Prio | OpenPub | Lot |
|---|---|---|---|---|
| Page de garde | Galerie de pages de garde (blocs de construction) | P2 | ◐ `Templates` | 10 |
| Page vierge, saut de page | Ctrl+Entrée | P0 | ✖ | 2 |
| Tableau | Grille, boîte « Insérer un tableau », dessiner, convertir texte ↔ tableau, tableaux rapides | P0 | ◐ `TableOps` | 6 |
| Feuille de calcul Excel incorporée | Objet OLE | HP (aperçu conservé) | — | 7 |
| Images de l'appareil | JPEG, PNG, GIF, BMP, TIFF, SVG, WebP, EMF/WMF | P0 | ✔ | 7 |
| Images en ligne, images stock, icônes Microsoft | Bing, bibliothèque Microsoft | HP | — | — |
| Icônes | Petite bibliothèque d'icônes libres (SVG) | P3 | ✖ | 15 |
| Formes | Environ 170 formes prédéfinies DrawingML ; dessin, modifier les points | P1 | ◐ `ShapeLibrary` (50) | 7 |
| Modèles 3D | | HP | — | — |
| SmartArt | **Affichage** du dessin de repli enregistré dans le `.docx` : P1 ; **édition** : P3 | P1 / P3 | ◐ | 7 / 15 |
| Graphique | **Affichage** à partir des valeurs en cache du `.docx` : P2 ; **édition** (feuille de données) : P3 | P2 / P3 | ✖ | 15 |
| Capture d'écran | Fenêtre ou zone | P3 | ✖ | 15 |
| Vidéo en ligne | | HP | — | — |
| Lien hypertexte | Vers URL, fichier, signet, titre, adresse e-mail ; info-bulle ; Ctrl+clic | P0 | ◐ | 3 |
| Signet | Ajouter, atteindre, masquer | P1 | ✖ | 9 |
| Renvoi | Vers titre, signet, note, légende ; numéro, texte, page, « au-dessus / au-dessous » | P2 | ✖ | 9 |
| Commentaire | Voir Révision | P1 | ✖ | 8 |
| En-tête / Pied de page | Galeries, édition par double clic (voir §11) | P0 | ✖ | 5 |
| Numéro de page | Haut, bas, marge, position actuelle ; format (1, i, A, -1-) ; numérotation par section | P0 | ◐ | 5 |
| Zone de texte | Flottante, liée à une autre zone (texte qui continue) | P1 | ◐ (chaînes) | 7 |
| QuickPart ▸ Champ | Boîte « Champ » (liste des champs Word) | P2 | ◐ | 10 |
| QuickPart ▸ Propriété du document | Contrôle lié aux propriétés (titre, auteur…) | P2 | ✖ | 10 |
| QuickPart ▸ Insertion automatique, Organisateur de blocs | Blocs de construction (`glossaryDocument`, `Building Blocks.dotx`) | P2 | ◐ `Templates` | 10 |
| WordArt | Texte décoratif dans une zone de texte | P2 | ✔ | 7 |
| Lettrine | Dans le texte / dans la marge, lignes, police, distance | P2 | ✔ | 3 |
| Ligne de signature | Bloc de signature (sans signature numérique) | P3 | ✖ | 15 |
| Date et heure | Formats, mise à jour automatique (champ DATE) | P1 | ✔ | 10 |
| Objet ▸ Texte d'un fichier | Insérer le contenu d'un autre document | P2 | ✖ | 13 |
| Objet OLE (Paint, Visio, Excel…) | Affichage de l'aperçu, double clic sans effet | P2 (aperçu) / HP (édition) | ✖ | 7 |
| Équation | **Affichage** OMML : P2 ; **édition** (structures, saisie linéaire UnicodeMath / LaTeX) : P3 | P2 / P3 | ✖ | 15 |
| Symbole, caractères spéciaux | Tableau de caractères, récents, raccourcis (tirets, espaces insécables, ©…) | P0 | ✔ | 3 |

## 4. Dessin

| Fonction | Détail | Prio | OpenPub | Lot |
|---|---|---|---|---|
| Stylets, surligneur, gomme, lasso | Encre (InkML dans le `.docx`) | P3 | ◐ (crayon à main levée) | 15 |
| Convertir l'encre en forme / en équation, rejouer l'encre | | HP | — | — |
| Zone de dessin (canevas) | Conteneur `wpc` | P3 (affichage) | ✖ | 7 |

## 5. Création (Design)

| Fonction | Détail | Prio | OpenPub | Lot |
|---|---|---|---|---|
| Thèmes | Thème `.thmx` / `theme1.xml` : couleurs, polices, effets ; changement de thème qui recolore le document | P1 | ◐ (jeux de couleurs/polices) | 4 |
| Jeux de styles | Ensembles de styles prédéfinis (« Ombré », « Minimaliste »…) | P2 | ✖ | 4 |
| Couleurs, polices du thème | | P1 | ◐ | 4 |
| Espacement de paragraphe | Préréglages du document | P2 | ✖ | 4 |
| Définir par défaut | Écrit dans le modèle Normal | P1 | ✖ | 4 |
| Filigrane | Texte (CONFIDENTIEL, BROUILLON…) ou image, en diagonale ou horizontal | P1 | ◐ | 5 |
| Couleur de page | Couleur, dégradé, texture, image (affichage ; impression optionnelle) | P2 | ✔ | 5 |
| Bordures de page | Trait ou motif (Art), par section, première page seulement | P2 | ◐ BorderArt | 5 |

## 6. Mise en page (Disposition)

| Fonction | Détail | Prio | OpenPub | Lot |
|---|---|---|---|---|
| Marges | Préréglages + personnalisées, reliure, pages en vis-à-vis, 2 pages par feuille, livre plié | P0 (base) / P2 (livre) | ◐ | 5 |
| Orientation | Par section | P0 | ✔ | 5 |
| Taille | A4, A5, Lettre, Légal… + personnalisée | P0 | ✔ | 5 |
| Colonnes | 1 à 13 colonnes, largeurs inégales, ligne séparatrice, à partir de ce point ; **équilibrage** de fin de section | P1 | ◐ | 5 |
| Sauts | Page, colonne, renvoi à la ligne (habillage), **sections** (page suivante, continu, page paire, page impaire) | P0 | ✖ | 2 / 5 |
| Numéros de ligne | Continu, par page, par section, intervalle | P3 | ✖ | 5 |
| Coupure de mots | Automatique / manuelle / aucune ; zone, nombre de coupures consécutives ; tiret conditionnel | P1 | ✔ `Hyphenator` (FR, à étendre NL/EN/DE) | 2 |
| Retrait et espacement | Raccourcis du paragraphe | P0 | ✔ | 3 |
| Position, habillage du texte | Aligné sur le texte, carré, rapproché, au travers, haut et bas, derrière, devant ; modifier les points d'habillage ; « déplacer avec le texte » / « fixer la position sur la page » | P1 | ◐ `PageLayouter` | 7 |
| Avancer, reculer, volet Sélection, aligner, grouper, faire pivoter | | P1 | ✔ `ShapeOps` | 7 |

## 7. Références

| Fonction | Détail | Prio | OpenPub | Lot |
|---|---|---|---|---|
| Table des matières | Galerie automatique, personnalisée (niveaux, points de suite, liens), mise à jour (numéros / tout), ajouter du texte (niveau) | P1 | ✖ | 9 |
| Notes de bas de page | Insérer (Ctrl+Alt+F), numérotation (par page / section / document, format), séparateur, **notes qui débordent sur la page suivante** | P1 | ✖ | 9 |
| Notes de fin | Fin de section / de document ; conversion notes de bas de page ↔ notes de fin | P2 | ✖ | 9 |
| Recherche (Researcher), Recherche intelligente | | HP | — | — |
| Citations et bibliographie | Gérer les sources (`sources.xml`), style (APA, Chicago, ISO 690…), insérer citation, bibliographie | P3 | ✖ | 9 |
| Légendes | Insérer une légende (Figure, Tableau, Équation ; numérotation avec chapitre), légende automatique | P2 | ✖ | 9 |
| Table des illustrations | Champ TOC \c | P2 | ✖ | 9 |
| Renvoi | Voir Insertion | P2 | ✖ | 9 |
| Index | Marquer les entrées (XE), insérer l'index (colonnes, format), mettre à jour, concordance | P3 | ✖ | 9 |
| Table des références juridiques | Champs TA / TOA | P3 | ✖ | 9 |

## 8. Publipostage

| Fonction | Détail | Prio | OpenPub | Lot |
|---|---|---|---|---|
| Démarrer la fusion | Lettres, messages électroniques, enveloppes, étiquettes, répertoire ; assistant pas à pas | P2 | ◐ `MergeOps` | 11 |
| Sélection des destinataires | Nouvelle liste (saisie), CSV, **Excel `.xlsx`**, Access (HP), contacts Outlook (HP) | P2 | ◐ (CSV) | 11 |
| Modifier la liste | Cocher, trier, filtrer, doublons | P2 | ✖ | 11 |
| Champs de fusion, bloc d'adresse, formule d'appel | Association des champs | P2 | ◐ | 11 |
| Règles | Si…Alors…Sinon, Demander, Remplir, Numéro de fusion, Enregistrement suivant, Ignorer l'enregistrement | P3 | ✖ | 11 |
| Aperçu des résultats | Navigation, rechercher un destinataire, vérifier les erreurs | P2 | ◐ | 11 |
| Terminer | Nouveau document, impression, **PDF (un fichier ou un par enregistrement)** | P2 | ◐ | 11 |
| Envoyer des messages électroniques | Via Outlook | HP | — | — |
| Enveloppes et étiquettes | Formats d'enveloppes (DL, C5…), planches d'étiquettes (Avery, Herma) | P2 | ✖ | 11 |

## 9. Révision

| Fonction | Détail | Prio | OpenPub | Lot |
|---|---|---|---|---|
| Orthographe pendant la frappe | Soulignement rouge, suggestions, ignorer, ajouter au dictionnaire, **langue par segment** | P0 | ✔ `SpellService` (ISpellChecker Windows) | 8 |
| Vérification de tout le document (F7) | | P0 | ✔ | 8 |
| Grammaire | Accords, ponctuation, espaces ; aucune API Windows : moteur local optionnel | P2 | ✖ | 8 |
| Synonymes (Maj+F7) | Dictionnaires de synonymes libres (FR, NL, EN) | P2 | ✖ | 8 |
| Statistiques | Pages, mots, caractères (avec/sans espaces), paragraphes, lignes ; inclure notes et zones de texte | P0 | ✔ | 3 |
| Lire à voix haute | Synthèse vocale Windows, surlignage du mot lu | P2 | ✖ | 14 |
| Vérifier l'accessibilité | Voir Fichier | P2 | ◐ | 14 |
| Traduire, Langue ▸ Détecter automatiquement | Traduction en ligne | HP (traduction) / P2 (choix de la langue de vérification) | ✖ | 8 |
| Commentaires | Nouveau (Ctrl+Alt+M), répondre (fils), résoudre / rouvrir, supprimer, précédent / suivant, volet Commentaires, bulles dans la marge ; auteur et date | P1 | ✖ | 8 |
| Mentions @ dans les commentaires | Notification par e-mail | HP (conservées à la lecture) | — | — |
| Suivi des modifications | Activer (Ctrl+Maj+E), insertions, suppressions, **mises en forme**, **déplacements**, lignes de tableau ; par auteur, couleurs par auteur | P1 | ✖ | 8 |
| Affichage pour révision | Marques simples, toutes les marques, aucune marque, original ; bulles / en ligne | P1 | ✖ | 8 |
| Volet Vérifications | Liste des révisions et commentaires | P2 | ✖ | 8 |
| Accepter / Refuser | Une, toutes, toutes celles affichées, par auteur ; suivante / précédente | P1 | ✖ | 8 |
| Verrouiller le suivi | Mot de passe | P3 | ✖ | 8 |
| Comparer | Deux versions → document des différences en révisions | P2 | ✖ | 8 |
| Combiner | Révisions de plusieurs auteurs | P3 | ✖ | 8 |
| Restreindre la modification | Lecture seule, commentaires seulement, révisions seulement, formulaires ; exceptions par plage ; mot de passe (hachage) | P2 | ✖ | 8 |
| Bloquer les auteurs | Co-édition SharePoint | HP | — | — |
| Encre ▸ Masquer l'encre | | P3 | ✖ | 15 |

## 10. Affichage

| Fonction | Détail | Prio | OpenPub | Lot |
|---|---|---|---|---|
| Page (mode normal) | Pages côte à côte ou défilement vertical, espaces entre pages masquables (double clic) | P0 | ◐ `PageView` | 2 |
| Mode Lecture | Colonnes adaptées à l'écran, sans ruban | P3 | ✖ | 14 |
| Web | Texte sans pagination, largeur de fenêtre | P2 | ✖ | 14 |
| Plan | Niveaux de titres, réduire / développer, déplacer des sections ; documents maîtres (HP) | P2 | ✖ | 14 |
| Brouillon | Sans mise en page, rapide | P3 | ✖ | 14 |
| Focus, Lecteur immersif | Largeur de colonne, couleur de page, focus sur la ligne, espacement, syllabes, lecture à voix haute | P3 | ✖ | 14 |
| Déplacement de page | Vertical / côte à côte (« feuilleter ») | P2 | ✖ | 14 |
| Règle | Horizontale (retraits, tabulations, marges, colonnes de tableau) et verticale | P0 | ◐ `Rulers` | 3 |
| Quadrillage | | P3 | ✔ | 14 |
| Volet de navigation | Titres (glisser pour réorganiser, réduire), **pages (miniatures)**, résultats de recherche | P0 (titres, pages) | ◐ (miniatures) | 14 |
| Zoom | Curseur, 100 %, une page, plusieurs pages, largeur de page, Ctrl+molette | P0 | ✔ | 2 |
| Nouvelle fenêtre, réorganiser, fractionner, côte à côte, défilement synchrone, changer de fenêtre | | P2 (fractionner, plusieurs documents) / P3 (reste) | ✖ | 14 |
| Macros ▸ Afficher, enregistrer | VBA | HP (macros conservées, jamais exécutées) | — | 13 |
| Propriétés SharePoint | | HP | — | — |

## 11. Onglets contextuels

### En-tête et pied de page
| Fonction | Détail | Prio | OpenPub | Lot |
|---|---|---|---|---|
| Édition en place | Double clic dans la marge, corps du document grisé | P0 | ◐ (pages maîtres) | 5 |
| Première page différente, pages paires et impaires différentes | Par section | P0 | ◐ | 5 |
| Lier au précédent | Héritage entre sections | P1 | ✖ | 5 |
| Position | Distance depuis le haut / le bas, tabulations d'alignement | P1 | ✖ | 5 |
| Insérer date, infos document, images, QuickPart | | P1 | ◐ | 5 |
| Atteindre l'en-tête / le pied, précédent / suivant | | P1 | ✖ | 5 |

### Création de tableau / Disposition de tableau
| Fonction | Détail | Prio | OpenPub | Lot |
|---|---|---|---|---|
| Styles de tableau | Galerie, **mise en forme conditionnelle** (ligne d'en-tête, ligne des totaux, lignes / colonnes à bandes, première / dernière colonne), nouveau style | P1 | ◐ | 6 |
| Trame, bordures, styles de bordure, reproduire les bordures | Par cellule et par côté | P0 | ✔ | 6 |
| Insérer / supprimer lignes et colonnes, sélectionner | | P0 | ✔ | 6 |
| Fusionner / fractionner les cellules, fractionner le tableau | | P0 / P1 | ◐ | 6 |
| Ajustement automatique | Au contenu, à la fenêtre, largeur fixe ; **algorithme de largeur automatique de Word** | P0 | ✖ | 6 |
| Hauteur, largeur, distribuer | | P0 | ✔ | 6 |
| Alignement dans la cellule, orientation du texte, marges de cellule | | P1 | ◐ | 6 |
| Répéter les lignes d'en-tête, autoriser le fractionnement des lignes | **Tableaux sur plusieurs pages** | P0 | ✖ | 6 |
| Tableaux imbriqués | | P1 | ✖ | 6 |
| Tableau flottant (habillage) | | P2 | ✖ | 6 |
| Trier, convertir en texte, formule (=SUM(ABOVE)) | | P2 | ✖ | 6 |
| Dessiner un tableau, gomme | | P3 | ✖ | 6 |

### Format de l'image
| Fonction | Détail | Prio | OpenPub | Lot |
|---|---|---|---|---|
| Rogner | Libre, selon une forme, selon un rapport, remplissage / ajustement | P1 | ✔ | 7 |
| Corrections, couleur, transparence | Luminosité, contraste, netteté, saturation, gris, sépia, recolorier | P2 | ✔ `ImageEditorWindow` | 7 |
| Effets artistiques | Flou, crayon, aquarelle… | P3 | ✖ | 7 |
| Supprimer l'arrière-plan | | P3 | ✖ | 15 |
| Styles d'image, bordure, effets (ombre, reflet, lumière, bords adoucis, 3D) | | P2 | ✔ | 7 |
| Compresser les images | | P2 | ✔ | 7 |
| Modifier / rétablir l'image | | P2 | ◐ | 7 |
| Texte de remplacement | Marquer comme décoratif | P1 | ✔ | 7 |
| Position, habillage, taille, rotation | | P0 | ◐ | 7 |

### Format de la forme
| Fonction | Détail | Prio | OpenPub | Lot |
|---|---|---|---|---|
| Remplissage, contour, effets, styles de forme | Couleurs du thème, dégradés, motifs, images | P1 | ✔ | 7 |
| Modifier la forme / les points | | P2 | ✔ `PathEditor` | 7 |
| Texte dans la forme | Direction, alignement, marges, ajustement | P1 | ◐ | 7 |
| Styles WordArt | | P2 | ✔ | 7 |
| Lier des zones de texte | | P2 | ◐ (chaînes) | 7 |

### Équation, SmartArt, Graphique
Voir Insertion (§3) : affichage P1–P2, édition P3 (lot 15).

## 12. Fonctions transversales de saisie et d'édition

| Fonction | Détail | Prio | OpenPub | Lot |
|---|---|---|---|---|
| Point d'insertion et sélection | Clavier complet (Ctrl+flèches, Maj+Fin, Ctrl+Maj+Fin…), double clic mot, triple clic paragraphe, Ctrl+clic phrase, clic dans la marge (ligne / paragraphe / tout), **sélection en colonne** (Alt+glisser), sélection multiple (Ctrl), F8 étendre | P0 | ◐ `InPlaceTextEditor` | 3 |
| Glisser-déposer du texte | Déplacer / copier (Ctrl) | P1 | ✖ | 3 |
| Cliquer-taper | Double clic n'importe où sur la page | P3 | ✖ | 3 |
| Annuler / Rétablir / Répéter (F4) | Liste déroulante de l'historique | P0 | ◐ (photo complète à remplacer) | 1 |
| Mode refrappe (Inser) | | P3 | ✖ | 3 |
| Correction automatique | Liste de remplacements (FR/NL/EN), majuscule en début de phrase, deux majuscules, Ctrl+Z annule la correction ; **typographie française** (guillemets « », espaces insécables avant ; : ! ?) | P1 | ✔ `FrenchTypography` | 14 |
| Mise en forme automatique à la frappe | Listes automatiques (« 1. », « - »), bordures (---), tableaux (+---+), liens hypertextes, 1er → 1<sup>er</sup>, fractions | P1 | ✖ | 14 |
| Saisie semi-automatique (prédictions de texte) | IA en ligne | HP | — | — |
| Mini barre d'outils, menus contextuels | Mise en forme rapide sur la sélection | P1 | ◐ | 3 |
| Rechercher une commande (Alt+Q) | Recherche dans les commandes du ruban | P2 | ✖ | 14 |
| Raccourcis clavier Word | Plusieurs centaines ; les 150 plus courants reproduits à l'identique | P0 | ◐ | 3 |
| Info-bulles (touches Alt) | Accès clavier au ruban | P2 | ✖ | 14 |
| Barre d'état | Page x sur y, section, mots, langue, suivi des modifications, modes d'affichage, zoom | P0 | ◐ | 2 |
| Texte masqué | Affiché / imprimé selon options | P2 | ✖ | 3 |
| Caractères spéciaux de mise en page | Espace insécable (Ctrl+Maj+Espace), trait d'union insécable, tiret conditionnel, saut de ligne (Maj+Entrée), espace fine | P0 | ◐ | 3 |

## 13. Contrôles de contenu, formulaires, automatisation (onglet Développeur)

| Fonction | Détail | Prio | OpenPub | Lot |
|---|---|---|---|---|
| Contrôles de contenu | Texte enrichi, texte brut, image, case à cocher, zone de liste déroulante, liste déroulante, sélecteur de date, bloc de construction, **section répétitive** ; texte d'espace réservé, verrouillage | P2 | ✖ | 10 |
| Mode Création des contrôles, propriétés | | P2 | ✖ | 10 |
| Liaison XML (Custom XML) | `dataBinding` vers une partie XML | P3 (conservée, valeurs affichées) | ✖ | 10 |
| Formulaires hérités | Champs FORMTEXT, FORMCHECKBOX, FORMDROPDOWN (présents dans le corpus : `Annexe-20-Recours`) | P1 | ✖ | 10 |
| Contrôles ActiveX | | HP (aperçu conservé) | — | — |
| Macros VBA, enregistreur | `.docm` : le projet VBA est **conservé**, jamais exécuté ; avertissement à l'ouverture | HP (exécution) / P2 (conservation) | ✖ | 13 |
| Modèle attaché, Normal.dotm | Modèle global de l'utilisateur, attacher un modèle, mettre à jour les styles | P1 | ✖ | 14 |
| Champs Word | PAGE, NUMPAGES, SECTIONPAGES, DATE, TIME, CREATEDATE, SAVEDATE, PRINTDATE, AUTHOR, TITLE, FILENAME, REF, PAGEREF, NOTEREF, SEQ, STYLEREF, TOC, INDEX, XE, TC, HYPERLINK, IF, COMPARE, =formule, MERGEFIELD, ASK, FILLIN, DOCPROPERTY, SYMBOL, QUOTE, INCLUDETEXT, INCLUDEPICTURE ; commutateurs de format (\* MERGEFORMAT, \@, \#) ; mettre à jour (F9), afficher les codes (Alt+F9), verrouiller (Ctrl+F11) | P1 (courants) / P2 (reste) | ◐ (`Run.Field`) | 10 |

## 14. Formats de fichiers

| Format | Lecture | Écriture | Prio | Lot |
|---|---|---|---|---|
| **`.docx`** (ECMA-376 / ISO 29500 transitionnel) | ✔ | ✔ **format natif**, aller-retour sans perte des parties inconnues | P0 | 1 |
| `.dotx` (modèle) | ✔ | ✔ | P1 | 14 |
| `.docm` / `.dotm` (avec macros) | ✔ | ✔ (macros conservées) | P2 | 13 |
| `.docx` strict (ISO 29500 strict) | ✔ | — | P3 | 13 |
| `.doc` (Word 97-2003, binaire [MS-DOC]) | via LibreOffice (s'il est installé) | — | P2 | 13 |
| `.rtf` | ✔ | ✔ | P1 | 12 |
| `.odt` (OpenDocument) | ✔ | ✔ | P2 | 12 |
| `.txt` (encodages, fins de ligne) | ✔ | ✔ | P0 | 12 |
| `.html` / `.htm` / `.mht` | ✔ (HTML) | ✔ (HTML filtré autonome) | P2 | 12 |
| `.md` (Markdown) | ✔ | ✔ | P3 | 12 |
| `.pdf` (export) | — | ✔ | P0 | 12 |
| `.pdf` (ouverture avec redistribution du texte) | — | — | HP | — |
| `.xml` (Word 2003 XML, Flat OPC) | ✔ Flat OPC | ✔ Flat OPC | P3 | 13 |
| `.xps` | — | — | HP | — |
| `.docx` chiffré (mot de passe, [MS-OFFCRYPTO] Agile) | ✔ | ✔ | P2 | 8 |

## 15. Langues, accessibilité, sécurité

| Fonction | Détail | Prio | OpenPub | Lot |
|---|---|---|---|---|
| Vérification multilingue | Langue par segment (`w:lang`), détection automatique, « ne pas vérifier » | P0 | ✔ | 8 |
| Césure FR, NL, EN, DE | Motifs Liang LibreOffice | P1 | ◐ (FR seul) | 2 |
| Écriture de droite à gauche (arabe, hébreu), scripts complexes | Bidi, mise en forme `w:rtl`, `w:cs` | P3 | ✖ | 15 |
| Langues d'Asie orientale | Grille de document, ruby, texte vertical | HP (contenu conservé) | — | — |
| Interface | Français uniquement (comme OpenPub) | P0 | ✔ | 0 |
| Accessibilité de l'application | UI Automation (lecteurs d'écran), contraste élevé, navigation au clavier | P1 | ◐ | 14 |
| PDF balisé, PDF/A | Structure, ordre de lecture, texte de remplacement | P2 | ◐ | 12 |
| Mode protégé (fichiers venus d'Internet) | Ouverture en lecture seule | P3 | ✖ | 13 |
| Étiquettes de confidentialité (MIP), IRM | | HP (le document protégé est refusé avec un message clair) | — | — |
| Co-édition en temps réel, présence | | HP | — | — |
| Copilot, Designer, Résumé, Réécrire | | HP | — | — |

---

## 16. Synthèse

| Priorité | Lignes de l'inventaire (environ) | Commentaire |
|---|---|---|
| **P0** | ≈ 65 | Cœur : modèle de flux, pagination, édition, mise en forme, styles, listes, sections, en-têtes, tableaux, images, `.docx`, PDF, impression, orthographe |
| **P1** | ≈ 60 | Usage professionnel courant : commentaires, suivi des modifications, table des matières, notes de bas de page, thèmes, champs, correction automatique |
| **P2** | ≈ 70 | Documents longs et collaboration : comparer, protection, légendes, contrôles de contenu, publipostage, formats secondaires |
| **P3** | ≈ 40 | Rare : à confirmer (index, bibliographie, équations éditables, encre, combiner…) |
| **HP** | ≈ 30 | Services Microsoft en ligne, IA, VBA, ActiveX, co-édition |

**Ce qui distingue fondamentalement Word de Publisher** (et donc OpenWrite d'OpenPub) :

1. **Texte en flux continu** : un seul récit qui se pagine tout seul, au lieu de zones posées sur la page. Le moteur de lignes d'OpenPub se réutilise ; le **paginateur** (sections, veuves et orphelines, notes de bas de page, tableaux sur plusieurs pages, objets ancrés) est entièrement nouveau.
2. **Le `.docx` est le format natif** : il est documenté (ECMA-376), donc lisible et réinscriptible. Les documents viennent de Word et y retournent : la **fidélité d'aller-retour** (ne rien perdre, ne rien déplacer) est l'exigence centrale, alors qu'OpenPub a converti un seul `.pub` une fois pour toutes.
3. **Le travail à plusieurs** (commentaires, suivi des modifications, comparaison) est attendu dès qu'un document est échangé.
4. **Les documents longs** (100 à 500 pages) imposent une mise en page incrémentale et un annuler / rétablir par opérations, pas par photo complète du document comme dans OpenPub.
