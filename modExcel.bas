Attribute VB_Name = "modExcel"
Option Explicit

Public Const xlSolid                As Long = 1
Public Const xlCenter               As Long = -4108
Public Const xlContinuous           As Long = 1
Public Const xlLeft                 As Long = -4131

Public Sub formatCell(ByVal objWSheet As Object, _
                       ByVal baris1 As Long, ByVal kolom1 As Integer, _
                       ByVal baris2 As Long, ByVal kolom2 As Integer, _
                       ByVal fontBold As Boolean, ByVal mergeCell As Boolean, ByVal horizontalAlign As Long, Optional ByVal setColorHeader As Boolean = False, _
                       Optional ByVal setBorder As Boolean = False, Optional isHoliday As Boolean = False)
                       
''    On Error GoTo errHandle
        
    With objWSheet
        .Range(.cells(baris1, kolom1), .cells(baris2, kolom2)).NumberFormat = "@"
        
        If isHoliday Then
            .Range(.cells(baris1, kolom1), .cells(baris2, kolom2)).Font.Color = vbRed
            .Range(.cells(baris1, kolom1), .cells(baris2, kolom2)).Font.Bold = True
            
        Else
            .Range(.cells(baris1, kolom1), .cells(baris2, kolom2)).Font.Size = 8
            .Range(.cells(baris1, kolom1), .cells(baris2, kolom2)).cells.HorizontalAlignment = horizontalAlign
            
            If fontBold Then .Range(.cells(baris1, kolom1), .cells(baris2, kolom2)).Font.Bold = fontBold
            If mergeCell Then .Range(.cells(baris1, kolom1), .cells(baris2, kolom2)).MergeCells = mergeCell
            
            If setColorHeader Then
                .Range(.cells(baris1, kolom1), .cells(baris2, kolom2)).Interior.ColorIndex = 15
                .Range(.cells(baris1, kolom1), .cells(baris2, kolom2)).Interior.Pattern = xlSolid
            End If
            
            If setBorder Then .Range(.cells(baris1, kolom1), .cells(baris2, kolom2)).Borders.LineStyle = xlContinuous
        End If
    End With
    
    Exit Sub
errHandle:
    Exit Sub
End Sub


