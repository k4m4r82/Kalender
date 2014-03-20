VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Demo membuat kalender sendiri"
   ClientHeight    =   720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbTahun 
      Height          =   315
      Left            =   1425
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdTampilkan 
      Caption         =   "Tampilkan"
      Height          =   315
      Left            =   2880
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Tahun"
      Height          =   195
      Left            =   840
      TabIndex        =   1
      Top             =   270
      Width           =   465
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
' MMMM  MMMMM  OMMM   MMMO    OMMM    OMMM    OMMMMO     OMMMMO    OMMMMO  '
'  MM    MM   MM MM    MMMO  OMMM    MM MM    MM   MO   OM    MO  OM    MO '
'  MM  MM    MM  MM    MM  OO  MM   MM  MM    MM   MO   OM    MO       OMO '
'  MMMM     MMMMMMMM   MM  MM  MM  MMMMMMMM   MMMMMO     OMMMMO      OMO   '
'  MM  MM        MM    MM      MM       MM    MM   MO   OM    MO   OMO     '
'  MM    MM      MM    MM      MM       MM    MM    MO  OM    MO  OM   MM  '
' MMMM  MMMM    MMMM  MMMM    MMMM     MMMM  MMMM  MMMM  OMMMMO   MMMMMMM  '
'                                                                          '
' K4m4r82's Laboratory                                                     '
' http://coding4ever.wordpress.com                                         '
'***************************************************************************

Option Explicit

Dim leap As Boolean

Private Function roundOff(ByVal num As Double) As Integer
    Dim str     As String
    Dim str2    As String
    Dim ctr     As Integer
    
    str = CStr(num)
    For ctr = 1 To Len(str)
        If Mid(str, ctr, 1) = "." Then
            roundOff = CInt(str2)
            Exit Function
        Else
            str2 = str2 & Mid(str, ctr, 1)
        End If
    Next
    
    roundOff = CInt(str2)
End Function

Private Function detrmMonth(ByVal bulan As Integer) As Integer
    Select Case bulan
        Case 1 'January
            If leap = True Then
                detrmMonth = 6
            Else
                detrmMonth = 0
            End If
            
        Case 2 'Febuary
            If leap = True Then
                detrmMonth = 2
            Else
                detrmMonth = 3
            End If
            
        Case 3 'March
            detrmMonth = 3
            
        Case 4 'April
            detrmMonth = 6
            
        Case 5 'May
            detrmMonth = 1
            
        Case 6 'June
            detrmMonth = 4
            
        Case 7 'July
            detrmMonth = 6
            
        Case 8 'August
            detrmMonth = 2
            
        Case 9 'September
            detrmMonth = 5
            
        Case 10 'October
            detrmMonth = 0
            
        Case 11 'November
            detrmMonth = 3
            
        Case 12 'December
            detrmMonth = 5
    End Select
End Function

Private Function getHariByAngka(ByVal hari As Integer) As String
    Select Case hari
        Case 0: getHariByAngka = "Minggu"
        Case 1: getHariByAngka = "Senin"
        Case 2: getHariByAngka = "Selasa"
        Case 3: getHariByAngka = "Rabu"
        Case 4: getHariByAngka = "Kamis"
        Case 5: getHariByAngka = "Jum'at"
        Case 6: getHariByAngka = "Sabtu"
    End Select
End Function

Private Function getJumlahHariByBulan(ByVal bulan As Integer, ByVal tahun As Long) As Integer
    getJumlahHariByBulan = Day(DateSerial(tahun, bulan + 1, 0))
End Function

Private Function DOTW(ByVal hari As Integer, ByVal bulan As Integer, ByVal tahun As Integer) As String
    Dim yr      As Double
    Dim result  As Integer
    
    yr = tahun / 4
    result = roundOff(yr) + tahun
    
    yr = tahun / 100
    result = result - roundOff(yr)
    
    yr = tahun / 400
    result = result + roundOff(yr)
    result = result + hari
    result = result + detrmMonth(bulan)
    result = result - 1
    result = result Mod 7
    
    DOTW = getHariByAngka(result)
End Function

Private Function getBulanIndonesia(ByVal bulan As Integer) As String
    Select Case bulan
        Case 1: getBulanIndonesia = "Januari"
        Case 2: getBulanIndonesia = "Februari"
        Case 3: getBulanIndonesia = "Maret"
        Case 4: getBulanIndonesia = "April"
        Case 5: getBulanIndonesia = "Mei"
        Case 6: getBulanIndonesia = "Juni"
        Case 7: getBulanIndonesia = "Juli"
        Case 8: getBulanIndonesia = "Agustus"
        Case 9: getBulanIndonesia = "September"
        Case 10: getBulanIndonesia = "Oktober"
        Case 11: getBulanIndonesia = "November"
        Case 12: getBulanIndonesia = "Desember"
    End Select
End Function

Private Function getAngkaByHari(ByVal hari As String) As Integer
    Select Case hari
        Case "Minggu": getAngkaByHari = 0
        Case "Senin": getAngkaByHari = 1
        Case "Selasa": getAngkaByHari = 2
        Case "Rabu": getAngkaByHari = 3
        Case "Kamis": getAngkaByHari = 4
        Case "Jum'at": getAngkaByHari = 5
        Case "Sabtu": getAngkaByHari = 6
    End Select
End Function

Private Function getRowByCell(ByVal cell As Integer) As Integer
    Select Case cell
        Case 1 To 7: getRowByCell = 1
        Case 8 To 14: getRowByCell = 2
        Case 15 To 21: getRowByCell = 3
        Case 22 To 28: getRowByCell = 4
        Case 29 To 35: getRowByCell = 5
        Case 36 To 42: getRowByCell = 6
        Case Else: getRowByCell = 1
    End Select
End Function

Private Function getColByCell(ByVal cell As Integer) As Integer
    Select Case cell
        Case 1, 8, 15, 22, 29, 36
            getColByCell = 0
            
        Case 2, 9, 16, 23, 30, 37
            getColByCell = 1
            
        Case 3, 10, 17, 24, 31, 38
            getColByCell = 2
            
        Case 4, 11, 18, 25, 32, 39
            getColByCell = 3
            
        Case 5, 12, 19, 26, 33, 40
            getColByCell = 4
            
        Case 6, 13, 20, 27, 34, 41
            getColByCell = 5
            
        Case 7, 14, 21, 28, 35, 42
            getColByCell = 6
    End Select
End Function

Private Sub initTahun()
    Dim i As Long
    
    For i = 2009 To Year(Now)
        cmbTahun.AddItem CStr(i)
    Next i
    If cmbTahun.ListCount > 0 Then cmbTahun.Text = Year(Now)
End Sub

Private Sub eksporKalender(ByVal tahun As Long)
    Dim objExcel            As Object
    Dim objWBook            As Object
    Dim objWSheet           As Object

    Dim initRow             As Long
    Dim initCol             As Long
    
    Dim i                   As Long
    Dim bulan               As Integer
    Dim kolom               As Integer
    
    Dim hari                As Integer
    Dim y                   As Integer
    Dim index               As Integer
    Dim cell                As Integer
    
    Dim baris               As Integer
    Dim ret                 As Integer
    
    Dim str                 As String
    
    Dim jumlahHariByBulan   As Integer
    Dim num                 As Integer
    Dim iBulan              As Integer
    
    Dim arrBulan(11)        As Integer
    
    Dim n                   As Long
    
    Dim mulai               As String
    Dim selesai             As String
    Dim lama                As String
    
    Dim ketHariLibur        As String
    
''    On Error GoTo errHandle
    
    num = tahun Mod 4
    If num = 0 Then
        leap = True
    Else
        leap = False
    End If
    
    mulai = Format(Now, "hh:mm:ss")
    
    Screen.MousePointer = vbHourglass
    DoEvents
    
    For i = 1 To 12
        arrBulan(n) = i
        n = n + 1
    Next
    
    'Create the Excel object
    Set objExcel = CreateObject("Excel.application") 'bikin object
    
    'Create the workbook
    Set objWBook = objExcel.Workbooks.Add
    
    Set objWSheet = objWBook.Worksheets(1)
    With objWSheet
        For i = 1 To 23
            .Columns(i).ColumnWidth = 2.86
        Next i
        
        Call formatCell(objWSheet, 1, 1, 1, 23, True, 11, xlCenter)
        .cells(1, 1) = "KALENDER NASIONAL"
        
        Call formatCell(objWSheet, 2, 1, 2, 23, True, 11, xlCenter)
        .cells(2, 1) = "TAHUN " & Year(Now)
        
        initRow = 5
        initCol = 1
        
        n = 1
        For iBulan = LBound(arrBulan) To UBound(arrBulan)
            bulan = arrBulan(iBulan)
            tahun = Year(Now)
            
            jumlahHariByBulan = getJumlahHariByBulan(bulan, tahun)
            
            cell = 0
            For hari = 1 To jumlahHariByBulan
                str = DOTW(hari, bulan, tahun)
                y = getAngkaByHari(str)
                        
                For index = cell To 41
                    baris = getRowByCell(cell)
                    kolom = getColByCell(cell)
                    
                    Call formatCell(objWSheet, initRow - 1, initCol, initRow - 1, initCol + 6, True, True, xlCenter, True, True)
                    .cells(initRow - 1, initCol) = getBulanIndonesia(bulan)
                    
                    Call formatCell(objWSheet, initRow, initCol, initRow, initCol + 6, True, False, xlCenter, False, True)
                    Call formatCell(objWSheet, initRow + 1, initCol, initRow + 1, initCol, False, False, xlCenter, False, True)
                    Select Case kolom
                        Case 0: .cells(initRow, kolom + initCol) = "M" 'Minggu
                        Case 1: .cells(initRow, kolom + initCol) = "S" 'Senin
                        Case 2: .cells(initRow, kolom + initCol) = "S" 'Selasa
                        Case 3: .cells(initRow, kolom + initCol) = "R" 'Rabu
                        Case 4: .cells(initRow, kolom + initCol) = "K" 'Kamis
                        Case 5: .cells(initRow, kolom + initCol) = "J" 'Jum'at
                        Case 6: .cells(initRow, kolom + initCol) = "S" 'Sabtu
                    End Select
                
                    If kolom = y Then
                        index = 41
                        
                        Call formatCell(objWSheet, baris + initRow, kolom + initCol, baris + initRow, kolom + initCol, False, False, xlCenter, False, True)
                        .cells(baris + initRow, kolom + initCol) = hari
                                                                            
                        If kolom = 0 Then 'hari minggu
                            Call formatCell(objWSheet, baris + initRow, kolom + initCol, baris + initRow, kolom + initCol, False, False, xlCenter, , , True)
                                                                            
                        Else
                            'libur nasional
                            strSql = "SELECT COUNT(*) FROM libur_nasional WHERE tanggal = " & hari & " AND bulan = " & bulan & ""
                            ret = CLng(dbGetValue(strSql, 0))
                            If ret > 0 Then Call formatCell(objWSheet, baris + initRow, kolom + initCol, baris + initRow, kolom + initCol, False, False, xlCenter, , , True)
                        End If
                        
                    Else
                        If baris > 0 And kolom > 0 Then
                            Call formatCell(objWSheet, baris + initRow, kolom + initCol, baris + initRow, kolom + initCol, False, False, xlCenter, , True)
                            .cells(baris + initRow, kolom + initCol) = ""
                        End If
                    End If
                    
                    cell = cell + 1
                Next index
            Next hari
            initCol = initCol + 8
            
            Select Case n
                Case 3, 6, 9
                    initCol = 1
                    initRow = initRow + 8
            End Select
            
            n = n + 1
        Next iBulan
        
        initRow = initRow + baris + 3
        
        Call formatCell(objWSheet, initRow, 1, initRow, 23, True, 8, xlCenter, True, True)
        .cells(initRow, 1) = "Keterangan Hari Libur"
        
        For iBulan = LBound(arrBulan) To UBound(arrBulan)
            bulan = arrBulan(iBulan)
            tahun = Year(Now)
            
            jumlahHariByBulan = getJumlahHariByBulan(bulan, tahun)
            
            For hari = 1 To jumlahHariByBulan
                strSql = "SELECT keterangan FROM libur_nasional WHERE tanggal = " & hari & " AND bulan = " & bulan & ""
                ketHariLibur = CStr(dbGetValue(strSql, ""))
                If Len(ketHariLibur) > 0 Then
                    Call formatCell(objWSheet, initRow + 1, 1, initRow + 1, 4, False, False, xlLeft)
                    .cells(initRow + 1, 1) = hari & " " & Left(getBulanIndonesia(bulan), 3) ' tglAwal & " - " & tglAkhir & " " & Left(getBulanIndonesia(bulan), 3)
                    .cells(initRow + 1, 4) = ketHariLibur
                    
                    initRow = initRow + 1
                End If
            Next hari
        Next iBulan

    End With
    
    objExcel.Visible = True
    
    If Not objWSheet Is Nothing Then Set objWSheet = Nothing
    If Not objWBook Is Nothing Then Set objWBook = Nothing
    If Not objExcel Is Nothing Then Set objExcel = Nothing
    
    Screen.MousePointer = vbDefault
    
    selesai = Format(Now, "hh:mm:ss")
    lama = TimeValue(selesai) - TimeValue(mulai)
    lama = Format(lama, "hh:mm:ss")
    Debug.Print "Lama : " & lama
    
    Exit Sub

errHandle:
''    Call prosesAkhir
    If Not objWSheet Is Nothing Then Set objWSheet = Nothing
    If Not objWBook Is Nothing Then Set objWBook = Nothing
    If Not objExcel Is Nothing Then Set objExcel = Nothing
End Sub

Private Sub cmdTampilkan_Click()
    If MsgBox("Apakah proses ingin dilanjutkan ?", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
        Call eksporKalender(Val(cmbTahun.Text))
    End If
End Sub

Private Sub Form_Load()
    Call initTahun
End Sub
