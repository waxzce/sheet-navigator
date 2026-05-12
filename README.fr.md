# Sheet Navigator

> Volet latéral Excel qui liste toutes les feuilles d'un classeur. Open source, cross-platform, sans compte.
>
> Démo & installation : **[sheet-navigator.waxzce.org](https://sheet-navigator.waxzce.org/)** · 🇬🇧 [README in English](README.md)

Add-in Office cross-platform (macOS, Windows, Web) qui ajoute à Excel un **volet latéral** listant verticalement toutes les feuilles du classeur actif. À la manière de Kutools, mais bâti avec la plateforme moderne Office.js, hébergé sur **Clever Cloud Static** et déployé via **GitHub Actions**.

## Fonctionnalités

- Liste compacte (32–36 px par ligne) des feuilles dans l'ordre du classeur
- Pastille colorée à gauche (`tabColor`), nom au centre, icônes d'action à droite
- Clic sur une ligne → active la feuille et sélectionne `A1`
- **Hiérarchie de dossiers via `>`** dans les noms : `Region > France` et `Region > Allemagne` apparaissent comme enfants d'un dossier pliable "Region". Profondeur illimitée. Drag entre dossiers = renommage auto. État expand/collapse persisté par classeur. (On utilise `>` parce qu'Excel interdit `/` dans les noms de feuilles — pareil pour `\ ? * [ ]`.)
- Recherche en haut, insensible aux accents et à la casse
- Toolbar : afficher/masquer les feuilles cachées · tri par classeur ou alpha · rafraîchir
- Bascule visible/masqué via l'icône œil
- Épinglage des feuilles fréquentes (icône punaise au survol), persisté **par classeur** dans `Office.context.document.settings`
- Glisser-déposer pour réordonner
- Synchronisation temps réel sur `onAdded`, `onDeleted`, `onActivated`, `onNameChanged`

## Stack

- **Office.js** (manifest XML classique — pas le JSON unifié)
- **TypeScript** strict
- **React 18** + **Fluent UI v9** (`@fluentui/react-components`)
- **Webpack 5** (HTTPS sur :3000 en dev)
- **Clever Cloud Static** (static-apache) avec build côté plateforme via `CC_PRE_BUILD_HOOK`

## Pré-requis

- macOS, Node.js LTS (≥ 20) — `brew install node` ou `nvm install --lts`
- Excel pour Mac récent (Microsoft 365)
- [`clever-tools`](https://www.clever.cloud/doc/clever-tools/getting_started/) pour le déploiement : `npm install -g clever-tools && clever login`

## Démarrage rapide

```bash
npm install
npm run dev              # webpack-dev-server HTTPS sur :3000
npm run sideload:mac     # installe manifest.local.xml dans le dossier wef d'Excel
```

À la première exécution de `npm run dev`, le template Office installe un **certificat HTTPS local** dans ton trousseau (mot de passe demandé) — obligatoire, Office refuse HTTP sauf `localhost`.

Ensuite :
1. Quitte complètement Excel (`Cmd+Q`)
2. Rouvre un classeur
3. **Insertion → Compléments → Mes compléments → Developer Add-ins → Sheet Navigator**
4. Le bouton « Navigateur de feuilles » apparaît dans l'onglet **Accueil**

Pour ouvrir les **devtools Safari** : clic droit dans le panneau → *Inspect Element*.

## Scripts npm

| Script | Action |
| --- | --- |
| `npm run dev` | Lance webpack-dev-server HTTPS sur `:3000`, hot reload |
| `npm run build` | Build de production dans `dist/`, bundles hashés |
| `npm run sideload:mac` | Copie `manifests/manifest.local.xml` dans le dossier wef |
| `npm run sideload:mac -- --prod` | Idem mais avec `manifest.prod.xml` (URL prod) |
| `npm run sideload:mac:clean` | Supprime tout manifest Sheet Navigator du dossier wef |
| `npm run deploy` | Déploie sur Clever Static (`clever deploy --force`) |
| `npm run validate` | Valide le manifest local |

## Hébergement Clever Static

L'app Clever Cloud est une **static-apache** configurée pour builder le projet côté plateforme.

Env vars de l'app :

| Variable | Valeur |
| --- | --- |
| `CC_PRE_BUILD_HOOK` | `npm ci && npm run build` |
| `CC_WEBROOT` | `/dist` |
| `CC_COMPOSER_VERSION` | `2` |

Le domaine custom `sheet-navigator.waxzce.org` est attaché à l'app, avec cert Let's Encrypt auto-provisionné par Clever Cloud. Le fichier `.clever.json` (versionné) référence l'`app_id`, donc `clever deploy` fonctionne sans `clever link` préalable.

## Déploiement local

```bash
clever login            # une fois
npm run deploy          # clever deploy --force
```

URL de prod : `https://sheet-navigator.waxzce.org/`

## Déploiement GitHub Actions

Le workflow `.github/workflows/deploy-clever.yml` se déclenche sur `push` vers `main` ou en `workflow_dispatch`.

Deux secrets à configurer dans **Settings → Secrets and variables → Actions** :

| Secret | Comment l'obtenir |
| --- | --- |
| `CLEVER_TOKEN` | `cat ~/.config/clever-cloud/clever-tools.json` après un `clever login` local, champ `token` |
| `CLEVER_SECRET` | Idem, champ `secret` |

## Distribution aux collègues

Page d'accueil pour les collègues : **https://sheet-navigator.waxzce.org/** (auto-détection Mac/Windows + instructions pas-à-pas).

### Mac

```bash
curl -o ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/sheet-navigator.manifest.xml \
  https://sheet-navigator.waxzce.org/manifest.xml
```

Puis relance Excel et va dans **Insertion → Compléments → Mes compléments → Developer Add-ins**.

### Windows

Le plus simple : aller sur https://sheet-navigator.waxzce.org/, télécharger `manifest.xml`, puis dans Excel **Insertion → Mes compléments → Charger mon complément** et sélectionner le fichier téléchargé.

Alternative pour multi-utilisateurs : poser `manifest.xml` sur un partage SMB ou OneDrive synchronisé, puis Excel → Fichier → Options → Centre de gestion de la confidentialité → Paramètres → **Catalogues de compléments approuvés** → colle l'URL UNC → coche *Afficher dans le menu* → relance Excel.

### Centralized Deployment Microsoft 365

Si Clever Cloud a un tenant M365 secondaire, un admin peut pousser le manifest depuis le **M365 admin center → Integrated apps → Upload custom apps**. L'add-in apparaît alors automatiquement dans l'Excel de tous les collègues ciblés, sans aucun sideload manuel.

## Architecture du code

```
src/
  taskpane/
    index.tsx              # Bootstrap React + FluentProvider
    App.tsx                # Composant racine, recherche + tri
    taskpane.html
    taskpane.css
    hooks/
      useWorksheets.ts     # Office.js encapsule : load + events + actions + epinglage
    components/
      SearchBar.tsx
      Toolbar.tsx
      SheetList.tsx
      SheetRow.tsx
  commands/
    commands.ts            # Function file referencee par le manifest (vide)
    commands.html
manifests/
  manifest.local.xml       # SourceLocation = https://localhost:3000/
  manifest.prod.xml        # SourceLocation = https://sheet-navigator.waxzce.org/
scripts/
  sideload-mac.mjs
.github/workflows/
  deploy-clever.yml
```

## Pièges connus

- **HTTPS obligatoire** : Office refuse HTTP sauf `localhost`. Clever Static + cert LE servent en HTTPS natif.
- **Cert wildcard refusé par Office Windows** : c'est pour ça qu'on est sur un domaine custom `sheet-navigator.waxzce.org` (cert dédié Let's Encrypt) plutôt que sur le wildcard `*.cleverapps.io` ou `*.cellar-c2.services.clever-cloud.com`.
- **Cache Office tenace sur Mac** : si une modification du manifest n'est pas prise en compte, fais `npm run sideload:mac:clean`, puis vide le cache Office : `rm -rf ~/Library/Containers/com.Microsoft.OsfWebHost/Data/*`, et relance Excel.
- **Le panneau est toujours à droite** : Microsoft impose une largeur minimum (~320 px) et n'autorise pas le dock à gauche.
- **Pas de manifest unifié JSON** : le format est en preview sur Mac, on reste sur XML. Migration possible en une commande plus tard : `npx office-addin-project convert`.

## Licence

MIT
