VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufOnenoteImport 
   Caption         =   "Onenote Import"
   ClientHeight    =   3315
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4800
   OleObjectBlob   =   "ufOnenoteImport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufOnenoteImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Ontwikkeld door: Freddy Meijer - Functioneel applicatiebeheerder VTH
'Organisatie: Gemeente Leiden
'Datum: 14-07-2023

Private Sub cmdbDoorvoeren_Click()

Dim intKleur As String
Dim intTeller As Integer
Dim strEersteCel As String
Dim strCheck As String
Dim strPad As String
Dim strLocatieLogos As String
Dim sngPaginaBreedte As Single

'De gebruiker kiest voor welke gemeente hij een kennisitem of handleiding wil schrijven. Dit doet de gebruiker via userform ufOnenoteImport.
'Hierin wordt een gemeente gekozen via de combobox cbGemeeente. Afhankelijk van de keuze wordt een basiskleur gekozen

If Me.cbGemeente = "Leiden" Then intKleur = 6
If Me.cbGemeente = "Leiderdorp" Then intKleur = 14
If Me.cbGemeente = "Oegstgeest" Then intKleur = 11
If Me.cbGemeente = "Zoeterwoude" Then intKleur = 9

'De gebruiker kiest of de orientatie van de pagina staand of liggend moet zijn. Dit doet hij via userform ufOnenoteImport. Hierin wordt in de combobox cbPagina aangegeven of de pagina's moeten liggen of staan.

If Me.cbPagina = "Staand" Then ActiveDocument.PageSetup.Orientation = 0
If Me.cbPagina = "Liggend" Then ActiveDocument.PageSetup.Orientation = 1

sngPaginaBreedte = ActiveDocument.PageSetup.pageWidth - (ActiveDocument.PageSetup.RightMargin + ActiveDocument.PageSetup.RightMargin)

dblKolom_1 = 0.33 * sngPaginaBreedte
dblKolom_2 = 0.66 * sngPaginaBreedte

'In onderstaande for-loop gebeuren twee dingen:
'- Als de paragraaftekst vetgedrukt is, moet deze tekst de kleur krijgen die overeenkomt met de gekozen gemeente (Me.cbGemeente)
'- Als het lettertype gelijk is aan consolas (het lettertype dat standaard uit Visual Studio Code komt) wordt de achtergrond grijs en het formaat 10,5. Dit zorgt ervoor dat een codeblok duidelijk opgemaakt wordt.

For intTeller = 1 To ActiveDocument.Paragraphs.Count

  If ActiveDocument.Paragraphs(intTeller).Range.Font.Bold = True Then
      With ActiveDocument.Paragraphs(intTeller).Range
        .Font.ColorIndex = intKleur
      End With
  End If
  If ActiveDocument.Paragraphs(intTeller).Range.Font.Name = "Consolas" Then
    With ActiveDocument.Paragraphs(intTeller).Range
        .Shading.BackgroundPatternColor = -603923969
        .Font.Size = 10.5
    End With
  End If
   
Next

  'Omdat de kennisitems beginnen met een tabel die 4 kolommen bevat en start met 'versie' (eerste cel, opgeslagen in strCheck) moet deze anders opgemaakt worden dan de meeste tabellen waarin toelichting staat.
  'In onderstaande for-loop wordt gecontroleerd of cel A1 gevuld is met het woord uit strCheck of dat de tabel meer dan 2 kolommen heeft. Indien zo, moet de eerste regel gevuld worden met de gekozen kleur.
  'Daarnaast moet de kleur van de tekst wit worden.

  strCheck = "Versie"
  
  For intTeller = 1 To ActiveDocument.Tables.Count
  
        strEersteCel = Left((ActiveDocument.Tables(intTeller).Cell(1, 1).Range.Text), (Len(ActiveDocument.Tables(intTeller).Cell(1, 1).Range.Text) - 2))

        If strEersteCel = strCheck Or ActiveDocument.Tables(intTeller).Columns.Count > 2 Then
            With ActiveDocument.Tables(intTeller).Rows(1)
                .Shading.Texture = wdTextureNone
                .Range.Font.Color = wdColorWhite
                .Shading.ForegroundPatternColorIndex = intKleur
                .Shading.BackgroundPatternColorIndex = intKleur
            End With
        End If
    Next

    'Met onderstaande if statements bepalen we waar de logo's van de gemeenten opgehaald kunnen worden.

    strLocatieLogos = "C:\Users\" & Environ("USERNAME") & "\OneDrive - Servicepunt 71\Playgrounds\Logo_gemeenten\"
    
    If Me.cbGemeente = "Leiden" Then strPad = strLocatieLogos & "Leiden.jpg"
    If Me.cbGemeente = "Leiderdorp" Then strPad = strLocatieLogos & "Leiderdorp.jpg"
    If Me.cbGemeente = "Oegstgeest" Then strPad = strLocatieLogos & "Oegstgeest.jpg"
    If Me.cbGemeente = "Zoeterwoude" Then strPad = strLocatieLogos & "Zoeterwoude.jpg"

    'Voordat het logo in de rechterbovenhoek geplaatst kan worden, dienen alle bestaande figuren verwijderd te worden.
    'Op deze manier zou je van een Leids kennisitem eenvoudig een Zoeterwouds kennisitem kunnen maken. Vervolgens wordt het figuur uit de map in de rechterbovenhoek geplaatst.

    With ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range
        .Delete
        .InlineShapes.AddPicture (strPad)
        .ParagraphFormat.Alignment = wdAlignParagraphRight
    End With

    'Op dezelfde manier worden onderaan de pagina paginanummers toegevoegd

    With ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary)
        .Range.Text = "Pagina "
        .PageNumbers.Add
        With .Range.Font
            .ColorIndex = intKleur
            .Bold = True
            .Name = "Calibri"
        End With
    End With
    
    'De totale breedte van een A4 is 16 cm. Door kolom 1 te definieren weet je ook de
    'breedte van kolom 2. Als kolom 1 5 cm breed is, is kolom 2 per definitie (16 - 5)
    '11 cm breed.

    strCheck = "Versie"

    For i = 1 To ActiveDocument.Tables.Count
        'De versie tabel moet niet geformateerd worden. Deze tabel start met Versie in cel 1. Als er dus Versie
        'in cel 1 staat, slaat de code deze tabel over. Als een tabel groter is dan 2 kolommen, wordt deze ook
        'overgeslagen.
        strEersteCel = Left((ActiveDocument.Tables(i).Cell(1, 1).Range.Text), (Len(ActiveDocument.Tables(i).Cell(1, 1).Range.Text) - 2))

        If strEersteCel = strCheck Or ActiveDocument.Tables(i).Columns.Count > 2 Then

        Else
            ActiveDocument.Tables(i).Columns(1).Width = dblKolom_1
            ActiveDocument.Tables(i).Columns(2).Width = dblKolom_2
        End If
    Next

   'Als een figuur in een tabel staat moet het figuur 75% van de celbreedte als breedte hebben.
   'Als het figuur niet in een tabel staat, wordt het 125% van de celbreedte van kolom 1 van een tabel.
   
    For i = 1 To ActiveDocument.InlineShapes.Count
    
        If ActiveDocument.InlineShapes(i).Range.Information(wdWithInTable) Then
            ActiveDocument.InlineShapes(i).Width = 0.75 * dblKolom_2
        Else
            ActiveDocument.InlineShapes(i).Width = 1.25 * dblKolom_1
        End If

    Next

End Sub

Private Sub UserForm_Initialize()

    With Me.cbGemeente
        .AddItem ("Leiden")
        .AddItem ("Leiderdorp")
        .AddItem ("Oegstgeest")
        .AddItem ("Zoeterwoude")
    End With

    With Me.cbPagina
        .AddItem ("Staand")
        .AddItem ("Liggend")
    End With

End Sub



