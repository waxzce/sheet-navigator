# Sheet Navigator

Add-in Office cross-platform (macOS, Windows, Web) qui ajoute à Excel un **volet latéral** listant verticalement toutes les feuilles du classeur actif. À la manière de Kutools, mais bâti avec la plateforme moderne Office.js, hébergé sur **Clever Cloud Cellar** et déployé via **GitHub Actions**.

## Fonctionnalités

- Liste compacte (32–36 px par ligne) des feuilles dans l'ordre du classeur
- Pastille colorée à gauche (`tabColor`), nom au centre, icônes d'action à droite
- Clic sur une ligne → active la feuille et sélectionne `A1`
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
- **AWS SDK v3** pour le déploiement Cellar (objet par objet, ACL `public-read`)

## Pré-requis

- macOS, Node.js LTS (≥ 20) — `brew install node` ou `nvm install --lts`
- Excel pour Mac récent (Microsoft 365)
- [`clever-tools`](https://www.clever-cloud.com/doc/clever-tools/getting_started/) pour créer le bucket Cellar : `npm install -g clever-tools && clever login`

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
| `npm run sideload:mac -- --prod` | Idem mais avec `manifest.prod.xml` (URL Cellar) |
| `npm run sideload:mac:clean` | Supprime tout manifest Sheet Navigator du dossier wef |
| `npm run deploy:cellar` | Build + push `dist/` vers Cellar via AWS SDK |
| `npm run validate` | Valide le manifest local |

## Configuration du bucket Cellar

Crée le bucket dédié au projet :

```bash
clever addon create cellar-addon sheet-navigator-storage
clever addon env sheet-navigator-storage      # affiche les credentials
```

Tu obtiens trois variables :
- `CELLAR_ADDON_KEY_ID`
- `CELLAR_ADDON_KEY_SECRET`
- `CELLAR_ADDON_HOST` (généralement `cellar-c2.services.clever-cloud.com`)

Puis crée le bucket `sheet-navigator` :

```bash
s3cmd \
  --access_key=$CELLAR_ADDON_KEY_ID \
  --secret_key=$CELLAR_ADDON_KEY_SECRET \
  --host=$CELLAR_ADDON_HOST \
  --host-bucket='%(bucket)s.'$CELLAR_ADDON_HOST \
  mb s3://sheet-navigator
```

> **Note importante** : Cellar n'expose **pas** un bucket entier en public par défaut. Chaque objet doit être poussé avec l'ACL `public-read` (`x-amz-acl: public-read`). Le script `scripts/deploy-cellar.mjs` et le workflow GitHub Actions le font automatiquement.

## Déploiement local

```bash
cp .env.example .env.local
# remplis CELLAR_KEY_ID, CELLAR_SECRET, CELLAR_BUCKET, CELLAR_ENDPOINT
source .env.local
npm run deploy:cellar
```

Le script :
1. lance `npm run build`
2. pousse tout `dist/` vers le bucket (HTML : cache 5 min, reste : 1 an immutable)
3. publie `manifests/manifest.prod.xml` à la racine du bucket sous le nom `manifest.xml`
4. supprime les objets orphelins (sauf `manifest.xml` à la racine)

URL résultante : `https://sheet-navigator.cellar-c2.services.clever-cloud.com/taskpane.html`

## Déploiement GitHub Actions

Le workflow `.github/workflows/deploy-cellar.yml` se déclenche sur `push` vers `main` ou en `workflow_dispatch`.

Trois secrets à configurer dans **Settings → Secrets and variables → Actions** :

| Secret | Valeur |
| --- | --- |
| `CELLAR_KEY_ID` | `CELLAR_ADDON_KEY_ID` obtenu via `clever addon env` |
| `CELLAR_SECRET` | `CELLAR_ADDON_KEY_SECRET` |
| `CELLAR_BUCKET` | `sheet-navigator` |

## Distribution aux collègues

### Mac

Chaque collègue télécharge le manifest et le place dans le dossier wef d'Excel :

```bash
curl -o ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/sheet-navigator.manifest.xml \
  https://sheet-navigator.cellar-c2.services.clever-cloud.com/manifest.xml
```

Puis relance Excel et va dans **Insertion → Compléments → Mes compléments**.

### Windows

Deux options.

**A — Dossier partagé.** Pose `manifest.xml` sur un SMB ou OneDrive synchronisé. Excel → Fichier → Options → Centre de gestion de la confidentialité → Paramètres → **Catalogues de compléments approuvés** → colle l'URL UNC → coche *Afficher dans le menu* → relance Excel. Insertion → Compléments → **Dossier partagé** fait apparaître Sheet Navigator.

**B — Téléchargement direct.** Récupère `manifest.xml` depuis Cellar, place-le dans `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` (chemin à confirmer selon la version d'Office).

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
  manifest.prod.xml        # SourceLocation = bucket Cellar
scripts/
  sideload-mac.mjs
  deploy-cellar.mjs
.github/workflows/
  deploy-cellar.yml
```

## Pièges connus

- **ACL `public-read` non automatique** : oubli classique. Si Excel charge mais affiche 403, vérifie que les objets sont bien publics.
- **HTTPS obligatoire** : Office refuse HTTP sauf `localhost`. Cellar sert nativement en HTTPS.
- **Cache Office tenace sur Mac** : si une modification du manifest n'est pas prise en compte, fais `npm run sideload:mac:clean`, puis vide le cache Office : `rm -rf ~/Library/Containers/com.Microsoft.OsfWebHost/Data/*`, et relance Excel.
- **Le panneau est toujours à droite** : Microsoft impose une largeur minimum (~320 px) et n'autorise pas le dock à gauche.
- **Pas de manifest unifié JSON** : le format est en preview sur Mac, on reste sur XML. Migration possible en une commande plus tard : `npx office-addin-project convert`.

## Licence

MIT
