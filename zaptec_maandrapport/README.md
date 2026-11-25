# Zaptec maandrapport voor Home Assistant OS

Deze lokale Home Assistant-app haalt op de eerste dag van iedere maand om 08:00 uur (Europe/Amsterdam) de gearchiveerde Zaptec-laadsessies van de voorgaande kalendermaand op, maakt een `.xlsx`-bestand en verstuurt dat via Gmail.

## Inhoud
- `config.yaml`: Home Assistant-app-configuratie
- `Dockerfile`: bouwt de Python-container
- `requirements.txt`: Python-afhankelijkheden
- `zaptec_report.py`: API-koppeling, Excel-export, Gmail-verzending en planning

## 1. Zaptec-gegevens verzamelen
1. Log in op het Zaptec Portal.
2. Open de installatie.
3. Ga naar **Settings / Instellingen → Details**.
4. Kopieer de **Installation ID**.
5. Zorg dat het Zaptec-account dat voor de API wordt gebruikt toegang heeft tot de installatie met de vereiste eigenaarstoegang voor het endpoint voor gearchiveerde sessies.
6. De Zaptec API kan voor API-toegang een serviceaccount/user group of aanvullende activering vereisen. Als aanmelden of ophalen HTTP 401/403 geeft, controleer dit bij Zaptec Support.

## 2. Gmail-app-wachtwoord maken
1. Schakel tweestapsverificatie in op het Google-account.
2. Maak in je Google-account een **app-wachtwoord** aan.
3. Gebruik dit app-wachtwoord in de add-on. Gebruik niet je normale Gmail-wachtwoord.
4. Sommige Google Workspace-accounts en accounts met Advanced Protection ondersteunen app-wachtwoorden niet.

## 3. Lokale app installeren
1. Maak op Home Assistant OS de map `/addons/local/zaptec_maandrapport`.
2. Kopieer alle bestanden uit deze map naar die map op Home Assistant.
3. Open **Instellingen → Apps → App Store**.
4. Open het menu rechtsboven en kies **Lokale apps opnieuw laden** (de precieze tekst kan per Home Assistant-versie iets verschillen).
5. Installeer **Zaptec maandrapport**.
6. Vul de Zaptec-, Gmail- en installatiegegevens in bij de configuratie.
7. Start de app.

## 4. Testen
Maak via de Terminal & SSH-app dit bestand aan:

```sh
mkdir -p /share/zaptec_maandrapport
touch /share/zaptec_maandrapport/nu-uitvoeren.txt
```

Binnen circa 30 seconden wordt een rapport over de voorgaande kalendermaand gemaakt en verstuurd. Het bestand blijft ook beschikbaar in `/share/zaptec_maandrapport/`.

## 5. Planning
De app controleert iedere 30 seconden de tijd. Op de eerste dag van de maand vanaf 08:00 uur (lokale tijd) wordt het rapport voor de voorgaande maand verstuurd. Een maandrapport wordt na succesvolle verzending niet automatisch dubbel verstuurd. Een handmatige test kan wel opnieuw versturen.

## Belangrijke aandachtspunten
- Het rapport groepeert sessies op de lokale startdatum in `Europe/Amsterdam`.
- Het endpoint `/api/sessions/archived` filtert technisch op eindtijd; de code haalt daarom vanaf de eerste dag van de rapportmaand tot het uitvoermoment op en filtert daarna op startdatum.
- Zaptec kan sessies die offline zijn opgenomen later archiveren. Een sessie die pas ná het maandrapport wordt gearchiveerd, kan daardoor ontbreken in de eerste verzending. Je kunt dan een handmatige heruitvoer doen.
- De code gebruikt de actuele Zaptec archived sessions API en cursor-paginering.
- De API-authenticatie die Zaptec documenteert gebruikt momenteel de OAuth password grant. Zaptec geeft aan dat deze grant in de toekomst wordt vervangen door een authorization-code-flow. Houd de API-documentatie in de gaten.
- De e-mailbijlage en het Excelbestand bevatten laadgegevens en kunnen persoonsgegevens bevatten. Beveilig je Home Assistant-installatie en deel het bestand alleen met de beoogde ontvanger.

## Logboeken
Bekijk de logboeken via **Instellingen → Apps → Zaptec maandrapport → Logboeken**.
