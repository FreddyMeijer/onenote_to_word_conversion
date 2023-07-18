## template.one
Alle logica in onderstaande code is gebaseerd op dit template. Dit template wordt door mij in Onenote als basis gebruikt wanneer ik een kennisitem maak.

## ufOnenoteImport.frm & ufOnenoteImport.frx
Dit is het Userform dat nodig is om het document uit OneNote om te zetten naar een kennisitem in docx. Van hieruit kan deze worden opgeslagen als PDF en gedistribueerd worden. 

### Paginaformaat
De gebruiker kan kiezen tussen A5, A4 en A3. Hierop worden de tabellen uiteindelijk gebasseerd. Als je een ander formaat wilt, kan dit uiteraard ook. In de basis zal het keuzeveld leeg zijn. Dan wordt het formaat gekozen as-is. Vermoedelijk is dit A4. Dit kan veranderd worden door een ander formaat te kiezen.

### Kolommen
Tabellen worden opgemaakt op basis van de paginabreedte. De paginabreedte wordt bepaald aan de hand van de huidige paginabreedte (deze is ingesteld door de gebruiker in *Me.CbPagina*). Deze paginabreedte wordt verminderd met de linker- en rechtermarge. **kolom 1** is net zo breed als 33% van de paginabreedte. **Kolom 2** is net zo breed als 66% van de paginabreedte. 

### Figuren
Figuren in het document krijgen automatisch een breedte. Hierbij moet je rekening houden met de vraag of een figuur in een tabel staat of niet. De breedte voor een figuur is:
 
- 75% van de kolombreedte van **kolom_2** indien het figuur in een tabel staat
- de kolombreedte van **kolom 1** indien het figuur niet in een tabel staat

## Start_Form.bas
Als je het formulier hebt ingelezen kan je dit bestand ook inlezen. Dit macro start simpelweg de userform op. Dit macro zou je als knop op kunnen nemen in jouw Word-lint. 