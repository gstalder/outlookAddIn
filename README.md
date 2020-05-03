# Outlook Add-In

Outlook Add-In, um Räume für Meetings vorzubereiten. Verschiedene Optionen stehen, z.B. die Art der Bestuhlung, Verpflegung etc.
Die Einstellung werden per HTTP-Request im JSON-Format an ein definiertes Umsystem gesendet.

## Add-In lokal testen

Benötigt [NodeJs](https://nodejs.org/en/). 

### Add-In in Outlook integrieren

Siehe [Querladen von Outlook-Add-Ins zu Testzwecken](https://docs.microsoft.com/de-de/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing).

### Dev-Server starten
Nur unter Windows möglich. Powershell öffnen, zu Ordner `...\RaumreservierungsTool` navigieren, Projekt bauen und lokalen Server unter https://localhost:3000/ hosten mit
```
npm run build
npm run dev-server
```
Das AddIn ist anschliessend verfügbar für Outlooktermine (nur für Neue).

## JSON-Receive lokal testen

In Powershell zu `...\SimpleServer` navigieren, Server lokal unter http://localhost:8080/ hosten mit
```
node jsServer.js
```

`index.html` mit Browser öffnen. Falls man über das Outlook Add-In an das Umsystem sendet, sollte (nach Nachladen von `index.html`) das gesendete JSON angezeigt werden. Funktioniert nicht mit Firefox Version neuer als 67.0.
