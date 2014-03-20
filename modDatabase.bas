Attribute VB_Name = "modDatabase"
Option Explicit

Public conn     As ADODB.Connection
Public strSql   As String

Public Function openDb() As Boolean
    Dim strCon As String
    
    On Error GoTo errHandle
    
    strCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\kalender.mdb"
    
    Set conn = New ADODB.Connection
    conn.ConnectionString = strCon
    conn.Open
    
    openDb = True
    
    Exit Function
errHandle:
    openDb = False
End Function

Public Function dbGetValue(ByVal query As String, ByVal defValue As Variant) As Variant
    Dim rsDbGetValue  As ADODB.Recordset
    
    On Error GoTo errHandle
    
    Set rsDbGetValue = New ADODB.Recordset
    rsDbGetValue.Open query, conn, adOpenForwardOnly, adLockReadOnly
    If Not rsDbGetValue.EOF Then
        If Not IsNull(rsDbGetValue(0).Value) Then
            dbGetValue = rsDbGetValue(0).Value
        Else
            dbGetValue = defValue
        End If
    Else
        dbGetValue = defValue
    End If
        
    rsDbGetValue.Close
    Set rsDbGetValue = Nothing
    
    Exit Function
errHandle:
    dbGetValue = defValue
End Function

