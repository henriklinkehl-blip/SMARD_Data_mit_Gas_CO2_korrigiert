# Automatischer Download: SMARD, TTF-Gas und EUA-CO₂

Dieses Paket erweitert den bestehenden täglichen GitHub-Actions-Download um:

- **EEX TTF Next Day Index (NDI)** als Gaspreisreihe
- **EEX EUA Primary Market Auction Price** als CO₂-Preisreihe
- eine kalendertägliche, vorwärts fortgeschriebene EUA-Reihe für spätere Modelltests

Die bestehenden SMARD-Dateien bleiben unverändert im Ordner `Data`.

## Neue Dateien

Bei einem **privaten Repository** werden zusätzlich dauerhaft gespeichert:

```text
Data/TTF_NDI_aktuell.csv
Data/EUA_Auktionspreis_aktuell.csv
Data/EUA_Auktionspreis_taeglich.csv
```

Bei einem **öffentlichen Repository** werden die EEX-Dateien aus Lizenzgründen nicht committed. Sie liegen nach jedem Lauf im Workflow-Artefakt:

```text
SMARD-EEX-Daten-aktuell/EEX_Gas_CO2_aktuell.zip
```

## Einbau in dein bestehendes Repository

Folgende Elemente aus diesem Paket hochladen und vorhandene Dateien ersetzen:

```text
.github/workflows/smard-update.yml
Scripts/update_smard.py
Scripts/update_eex.py
requirements.txt
.gitignore
```

Die Ordnernamen sind absichtlich **`Data`** und **`Scripts`** mit großem Anfangsbuchstaben.

Danach unter **Actions → SMARD-, Gas- und CO2-Daten aktualisieren → Run workflow** einen manuellen Lauf starten.

## Abruflogik

### Gas

Quelle:

```text
https://gasandregistry.eex.com/Gas/NDI/NDI_45_Days.csv
```

Die EEX-Datei enthält ein rollierendes Fenster von ungefähr 45 Tagen. Bei einem privaten Repository werden die Werte täglich zusammengeführt, sodass die eigene Historie ab dem ersten erfolgreichen Lauf wächst. Mehrfaches Abrufen am selben Tag erzeugt keine ältere Historie.

Die standardisierte Datei enthält zusätzlich den abgeleiteten Zeitpunkt `Verfuegbar_ab_Europe_Berlin`. Der NDI wird für den Folgetag erst am Vorabend ermittelt. Für einen zeitlich sauberen Strompreis-Backtest darf daher nicht ohne Lag mit dem gleichen Liefertag gearbeitet werden.

### CO₂

Quellen:

- EEX EUA Primary Market Auction Report des aktuellen Jahres
- EEX-Archiv der Auktionsergebnisse früherer Jahre

Es werden zwei Dateien erzeugt:

- `EUA_Auktionspreis_aktuell.csv`: nur tatsächliche Auktionstage
- `EUA_Auktionspreis_taeglich.csv`: der jeweils zuletzt bekannte Auktionspreis wird bis zur nächsten Auktion vorwärts fortgeschrieben

Die zweite Datei ist als Modellmerkmal bequemer. Bei einem späteren Walk-forward-Test muss trotzdem sichergestellt werden, dass nur Informationen verwendet werden, die zum jeweiligen Prognosezeitpunkt bereits veröffentlicht waren.

## Lizenzhinweis

SMARD-Daten stehen unter CC BY 4.0.

Für EEX-Marktdaten gilt ein eigener Nutzungshinweis: Eine systematische öffentliche Weiterverbreitung erheblicher Datenmengen ist nur mit ausdrücklicher Genehmigung der EEX gestattet. Deshalb speichert der Workflow EEX-Daten in öffentlichen Repositories standardmäßig nur als nicht versioniertes Workflow-Artefakt. Für die dauerhafte automatische Gas-Historie ist ein privates Repository vorgesehen.

`FORCE_EEX_PUBLIC_COMMIT` sollte nur nach geklärter Nutzungsberechtigung auf `true` gesetzt werden.

## Noch nicht Bestandteil der Prognose-App

Dieses Paket sammelt die zusätzlichen Daten zunächst nur. Gas und CO₂ werden noch nicht ungeprüft als Modellmerkmale eingebaut. Der nächste Schritt ist ein zeitlich sauberer Vergleich:

1. bisheriges Strompreismodell
2. Modell plus verzögerter TTF-Wert
3. Modell plus TTF und EUA
4. Vergleich über viele Walk-forward-Referenztage
