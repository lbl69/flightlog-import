# Excel Twin Jet → FlightLog

Page web qui convertit le tableau Excel des vols (export Twin Jet) en CSV importable dans l'appli **FlightLog**.

**Utilisation :** ouvrir la page GitHub Pages, renseigner trigramme + nom, déposer l'Excel (et éventuellement l'export FlightLog actuel), cocher PF/PM, télécharger ou partager le CSV.

Tout est calculé localement dans le navigateur, aucun fichier n'est envoyé.

## Règles de conversion
| Excel | FlightLog |
|---|---|
| `Vol` numérique (5417) | `T75417` — `ITI2911` laissé tel quel |
| `Imm` FGRYL | `F-GRYL` |
| `Depblc/Arrblc/Dec/Att` | `Out/In/Off/On` (UTC) |
| `Tpsblc/Tpsvol` | `Total Block/Total Flight` |
| `Conso` (lbs) | `Fuel` (tonnes, 2 décimales) |
| `Paxfret` | `Pax` |
| PIC | `PIC Time`, `PIC Name` = nom saisi, `Function` PIC |
| SIC | `SIC Time`, `PIC Name` = trigramme du CDB, `Function` SIC |

La fonction se choisit en haut de page pour tout l'export (**PIC**, **SIC** ou **Auto** d'après la position du trigramme dans `Cdb`/`Opl`), puis peut être corrigée vol par vol. Pour une période mixte (SIC puis PIC), faire deux exports.

- **Nuit** : règle EASA (soleil < −6°), minute par minute sur le temps bloc, position interpolée entre les terrains.
- **Atterrissage jour/nuit** selon l'heure d'atterrissage ; PM = 0 décollage / 0 atterrissage.
- **Distance** orthodromique en NM (coordonnées OurAirports, `airports.js`).
- **Export FlightLog (facultatif)** : reprend `SchedOut/SchedIn` et classe chaque vol en *nouveau*, *remplace prévu* (vol planning sans immat) ou *doublon* (exclu par défaut).

## Fichiers
- `index.html` — interface
- `convert.js` — logique de conversion
- `airports.js` — coordonnées des aérodromes (OurAirports, domaine public)
- `xlsx.full.min.js` — SheetJS (Apache 2.0)
