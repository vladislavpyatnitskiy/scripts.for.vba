'Set Margins as 1.27
Sub SetMarginsTo1_27Cm()
    With ActiveDocument.PageSetup
        .LeftMargin = CentimetersToPoints(1.27)
        .RightMargin = CentimetersToPoints(1.27)
        .TopMargin = CentimetersToPoints(1.27)
        .BottomMargin = CentimetersToPoints(1.27)
    End With
End Sub

'Justify text
Sub JustifyText()
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
End Sub

'Make Heater & Footer .5 cm
Sub SetHeaderFooterMarginsTo0_5Cm()
    With ActiveDocument.PageSetup
        .TopMargin = CentimetersToPoints(0.5)
        .BottomMargin = CentimetersToPoints(0.5)
    End With
End Sub

Sub AddBottomBorderToHeader()
    Dim headerRange As Range
    Set headerRange = ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range

    With headerRange.Borders(wdBorderBottom)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth050pt
        .Color = wdColorAutomatic
    End With
End Sub

' Justify text in Heater & Footer
  
Sub JustifyTextInHeaderFooter()
    Dim sect As Section

    For Each sect In ActiveDocument.Sections
        With sect.Headers(wdHeaderFooterPrimary).Range
            .ParagraphFormat.Alignment = wdAlignParagraphJustify
        End With

        With sect.Footers(wdHeaderFooterPrimary).Range
            .ParagraphFormat.Alignment = wdAlignParagraphJustify
        End With
    Next sect
End Sub

'Make Font Calibri 10

Sub ChangeFontToCalibri()
    Dim doc As Document
    Set doc = ActiveDocument
    
    Dim rng As Range
    Set rng = doc.Content ' Change to the specific range you want if not the whole document
    
    ' Set the font name and size
    rng.Font.Name = "Calibri (Body)"
    rng.Font.Size = 10
End Sub
