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

Dim intColor, strFirstCell, strCheckWord, strPath, strLocationLogos As String
Dim dblFirstColumn, dblSecondColumn as Double
Dim intCounter As Integer
Dim sngPageWidth As Single
Dim rngWord As Range
Dim doc as Document
Dim ilsLogo as InlineShape

Set doc = ActiveDocument

If Me.cbGemeente = "Leiden" Then intColor = 6
If Me.cbGemeente = "Leiderdorp" Then intColor = 14
If Me.cbGemeente = "Oegstgeest" Then intColor = 11
If Me.cbGemeente = "Zoeterwoude" Then intColor = 9
If Me.cbPagina = "Staand" Then doc.PageSetup.Orientation = 0
If Me.cbPagina = "Liggend" Then doc.PageSetup.Orientation = 1
If Me.cbFormaat = "A4" Then doc.PageSetup.PaperSize = wdPaperA4
If Me.cbFormaat = "A3" Then doc.PageSetup.PaperSize = wdPaperA3
If Me.cbFormaat = "A5" Then doc.PageSetup.PaperSize = wdPaperA5

sngPageWidth = doc.PageSetup.pageWidth - (doc.PageSetup.RightMargin + doc.PageSetup.RightMargin)

dblFirstColumn = 0.33 * sngPageWidth
dblSecondColumn = 0.66 * sngPageWidth

For intCounter = 1 To doc.Paragraphs.Count

  If doc.Paragraphs(intCounter).Range.Font.Bold = True Then
      With doc.Paragraphs(intCounter).Range
        .Font.ColorIndex = intColor
      End With
  End If

Next

For Each rngWord In doc.Words
    If rngWord.Font.Name = "Consolas" Then
        rngWord.Font.Size = 9.5
        rngWord.Shading.BackgroundPatternColor = -603923969
    End If
    rngWord.Collapse wdCollapseEnd
Next rngWord

  'Omdat de kennisitems beginnen met een tabel die 4 kolommen bevat en start met 'versie' (eerste cel, opgeslagen in strCheckWord) moet deze anders opgemaakt worden dan de meeste tabellen waarin toelichting staat.
  'In onderstaande for-loop wordt gecontroleerd of cel A1 gevuld is met het woord uit strCheckWord of dat de tabel meer dan 2 kolommen heeft. Indien zo, moet de eerste regel gevuld worden met de gekozen kleur.
  'Daarnaast moet de kleur van de tekst wit worden.

  strCheckWord = "Versie"
  
  For intCounter = 1 To doc.Tables.Count
  
        strFirstCell = Left((doc.Tables(intCounter).Cell(1, 1).Range.Text), (Len(doc.Tables(intCounter).Cell(1, 1).Range.Text) - 2))

        If strFirstCell = strCheckWord Or doc.Tables(intCounter).Columns.Count > 2 Then
            With doc.Tables(intCounter).Rows(1)
                .Shading.Texture = wdTextureNone
                .Range.Font.Color = wdColorWhite
                .Shading.ForegroundPatternColorIndex = intColor
                .Shading.BackgroundPatternColorIndex = intColor
            End With
        End If
    Next

    'Met onderstaande if statements bepalen we waar de logo's van de gemeenten opgehaald kunnen worden.

    strLocationLogos = "C:\Users\" & Environ("USERNAME") & "\OneDrive - Servicepunt 71\Playgrounds\Logo_gemeenten\"
    
    If Me.cbGemeente = "Leiden" Then strPath = strLocationLogos & "Leiden.jpg"
    If Me.cbGemeente = "Leiderdorp" Then strPath = strLocationLogos & "Leiderdorp.jpg"
    If Me.cbGemeente = "Oegstgeest" Then strPath = strLocationLogos & "Oegstgeest.jpg"
    If Me.cbGemeente = "Zoeterwoude" Then strPath = strLocationLogos & "Zoeterwoude.jpg"

    'Voordat het logo in de rechterbovenhoek geplaatst kan worden, dienen alle bestaande figuren verwijderd te worden.
    'Op deze manier zou je van een Leids kennisitem eenvoudig een Zoeterwouds kennisitem kunnen maken. Vervolgens wordt het figuur uit de map in de rechterbovenhoek geplaatst.

    With doc.Sections(1).Headers(wdHeaderFooterPrimary).Range
        .Delete
        If Me.cbFormaat = "A5" Then
            set ilsLogo =.InlineShapes.AddPicture (strPath)
            ilsLogo.lockaspectratio = msofalse
            islLogo.Width = islLogo.Width * 0,33
            islLogo.Height = islLogo.Height * 0,33
        End if
        .ParagraphFormat.Alignment = wdAlignParagraphRight
    End With

    'Op dezelfde manier worden onderaan de pagina paginanummers toegevoegd

    With doc.Sections(1).Footers(wdHeaderFooterPrimary)
        .Range.Text = "Pagina "
        .PageNumbers.Add
        With .Range.Font
            .ColorIndex = intColor
            .Bold = True
            .Name = "Calibri"
        End With
    End With
    
    'De totale breedte van een A4 is 16 cm. Door kolom 1 te definieren weet je ook de
    'breedte van kolom 2. Als kolom 1 5 cm breed is, is kolom 2 per definitie (16 - 5)
    '11 cm breed.

    strCheckWord = "Versie"

    For i = 1 To doc.Tables.Count
        'De versie tabel moet niet geformateerd worden. Deze tabel start met Versie in cel 1. Als er dus Versie
        'in cel 1 staat, slaat de code deze tabel over. Als een tabel groter is dan 2 kolommen, wordt deze ook
        'overgeslagen.
        strFirstCell = Left((doc.Tables(i).Cell(1, 1).Range.Text), (Len(doc.Tables(i).Cell(1, 1).Range.Text) - 2))

        If strFirstCell = strCheckWord Or doc.Tables(i).Columns.Count > 2 Then

        Else
            doc.Tables(i).Columns(1).Width = dblFirstColumn
            doc.Tables(i).Columns(2).Width = dblSecondColumn
        End If
    Next

   'Als een figuur in een tabel staat moet het figuur 75% van de celbreedte als breedte hebben.
   'Als het figuur niet in een tabel staat, wordt het 125% van de celbreedte van kolom 1 van een tabel.
   
    For i = 1 To doc.InlineShapes.Count
    
        If doc.InlineShapes(i).Range.Information(wdWithInTable) Then
            doc.InlineShapes(i).Width = 0.75 * dblSecondColumn
        Else
            doc.InlineShapes(i).Width = 1.5 * dblFirstColumn
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

    With Me.cbFormaat
        .AddItem ("")
        .AddItem ("A4")
        .AddItem ("A3")
        .AddItem ("A5")
    End With

End Sub



