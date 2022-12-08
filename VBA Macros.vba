' This macro change the proofing langauge on all slides
Sub changeProofingLanguage()

    Dim i As Integer, j As Integer, totalCount As Double
    totalCount = 0
    For i = 1 To ActivePresentation.Slides.Count
        For j = 1 To ActivePresentation.Slides(i).Shapes.Count
            If ActivePresentation.Slides(i).Shapes(j).HasTextFrame Then
                ActivePresentation.Slides(i).Shapes(j).TextFrame.TextRange.LanguageID = msoLanguageIDEnglishUS
                totalCount = totalCount + 1
            End If
        Next
    Next
    MsgBox "Done. Language updated on " & totalCount & " items.", vbInformation, "Done."

End Sub