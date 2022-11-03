Attribute VB_Name = "Módulo1"
Option Explicit

Sub CriarPagina(codigos As String, i As Integer)

    Dim fso As New Scripting.FileSystemObject
    Dim ts As Scripting.TextStream
   
    
    Set ts = fso.CreateTextFile( _
    Environ("UserProfile") & "\Desktop\" & Sheets(i).Name & ".txt", True, True)
    
    ts.Write codigos
    ts.Close

End Sub

Sub ExtrairCodigos()

    Dim qtdLinhas As Integer
    Dim i As Integer, j As Integer
    Dim codigos As String
    
    
    
    For i = 1 To Application.Sheets.Count
        codigos = ""
        qtdLinhas = Sheets(i).Range("A1").CurrentRegion.Rows.Count
        For j = 2 To qtdLinhas
            codigos = codigos & ", " & Sheets(i).Cells(j, 1).Value
        Next
   
        Call CriarPagina(codigos, i)
    Next

End Sub
