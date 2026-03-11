Attribute VB_Name = "Module4"
Option Explicit

Function IsValidEmail(strEmail As String) As Boolean
    Dim atPos As Long, dotPos As Long
    strEmail = Trim(strEmail)
    atPos = InStr(1, strEmail, "@")
    dotPos = InStrRev(strEmail, ".")
    If atPos = 0 Or dotPos < atPos + 2 Or InStr(1, strEmail, " ") > 0 Then
        IsValidEmail = False: Exit Function
    End If
    IsValidEmail = True
End Function

Sub Ertesites_Batch_Log()
    Dim ws As Worksheet, logWs As Worksheet
    Dim lastRow As Long, i As Long, batchSize As Long, batchCounter As Long
    Dim dataArr As Variant
    Dim statusTomb() As Variant, eredmenyTomb() As Variant, datumTomb() As Variant
    Dim emailApp As Object, emailItem As Object
    Dim nev As String, cimzett As String
    Dim statusz As String
    Dim sentEmails As Object
    Dim logRow As Long
    
    batchSize = 50 ' batch mérete
    
    Set ws = ThisWorkbook.ActiveSheet
    
    ' --- Log sheet létrehozása, ha nincs ---
    On Error Resume Next
    Set logWs = ThisWorkbook.Sheets("Log")
    If logWs Is Nothing Then
        Set logWs = ThisWorkbook.Sheets.Add
        logWs.Name = "Log"
        logWs.Range("A1:F1").Value = Array("Sor", "Név", "Email", "Státusz", "Eredmény", "Dátum")
    End If
    On Error GoTo 0
    
    logRow = logWs.Cells(logWs.Rows.Count, 1).End(xlUp).Row + 1
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    dataArr = ws.Range("A2:F" & lastRow).Value
    ReDim statusTomb(1 To UBound(dataArr), 1 To 1)
    ReDim eredmenyTomb(1 To UBound(dataArr), 1 To 1)
    ReDim datumTomb(1 To UBound(dataArr), 1 To 1)
    
    Set sentEmails = CreateObject("Scripting.Dictionary")
    
    ' Outlook csatlakozás
    On Error Resume Next
    Set emailApp = GetObject(, "Outlook.Application")
    If emailApp Is Nothing Then
        MsgBox "Az Outlook nem fut. Nyisd meg az asztali Outlookot.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    batchCounter = 0
    
    For i = 1 To UBound(dataArr)
        statusz = LCase(Trim(dataArr(i, 4)))
        
        Select Case statusz
            Case "értesítve"
                statusTomb(i, 1) = dataArr(i, 4)
                eredmenyTomb(i, 1) = dataArr(i, 5)
                datumTomb(i, 1) = dataArr(i, 6)
                
            Case "értesítendõ"
                nev = dataArr(i, 2)
                cimzett = Trim(dataArr(i, 3))
                
                If Not IsValidEmail(cimzett) Then
                    statusTomb(i, 1) = statusz
                    eredmenyTomb(i, 1) = "hiba: érvénytelen formátum"
                    datumTomb(i, 1) = Now
                    GoTo LogAndNext
                End If
                
                If sentEmails.Exists(LCase(cimzett)) Then
                    statusTomb(i, 1) = statusz
                    eredmenyTomb(i, 1) = "hiba: duplikált email"
                    datumTomb(i, 1) = Now
                    GoTo LogAndNext
                End If
                
                On Error GoTo OutlookHiba
                Set emailItem = emailApp.CreateItem(0)
                With emailItem
                    .To = cimzett
                    .Subject = "Értesítés a vezetéképítési projektrõl"
                    .Body = "Kedves " & nev & "," & vbCrLf & vbCrLf & _
                            "Ez egy automatikus értesítés a projektrõl." & vbCrLf & _
                            "Kérjük vegye fel velünk a kapcsolatot szükség esetén." & vbCrLf & vbCrLf & _
                            "Üdvözlettel," & vbCrLf & _
                            "Energetikai Projekt Csapat"
                    .Send
                End With
                
                sentEmails.Add LCase(cimzett), True
                statusTomb(i, 1) = "értesítve"
                eredmenyTomb(i, 1) = "sikeres"
                datumTomb(i, 1) = Now
                batchCounter = batchCounter + 1
                
                ' --- Batch visszaírás és log ---
                If batchCounter >= batchSize Then
                    ws.Range("D2").Resize(UBound(statusTomb), 1).Value = statusTomb
                    ws.Range("E2").Resize(UBound(eredmenyTomb), 1).Value = eredmenyTomb
                    ws.Range("F2").Resize(UBound(datumTomb), 1).Value = datumTomb
                    batchCounter = 0
                End If
                
LogAndNext:
                ' --- Log sheet-re írás ---
                logWs.Cells(logRow, 1).Value = i + 1
                logWs.Cells(logRow, 2).Value = nev
                logWs.Cells(logRow, 3).Value = cimzett
                logWs.Cells(logRow, 4).Value = statusTomb(i, 1)
                logWs.Cells(logRow, 5).Value = eredmenyTomb(i, 1)
                logWs.Cells(logRow, 6).Value = datumTomb(i, 1)
                logRow = logRow + 1
                
            Case Else
                statusTomb(i, 1) = dataArr(i, 4)
                eredmenyTomb(i, 1) = dataArr(i, 5)
                datumTomb(i, 1) = dataArr(i, 6)
        End Select
NextRow:
    Next i
    
    ' --- Végsõ visszaírás ---
    ws.Range("D2").Resize(UBound(statusTomb), 1).Value = statusTomb
    ws.Range("E2").Resize(UBound(eredmenyTomb), 1).Value = eredmenyTomb
    ws.Range("F2").Resize(UBound(datumTomb), 1).Value = datumTomb
    
    MsgBox "Küldés és logolás befejezve!", vbInformation
    Exit Sub

OutlookHiba:
    statusTomb(i, 1) = statusz
    eredmenyTomb(i, 1) = "hiba: Outlook hiba"
    datumTomb(i, 1) = Now
    Resume LogAndNext
End Sub

