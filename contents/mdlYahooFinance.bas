Attribute VB_Name = "mdlYahooFinance"
' ===============================================
' ���W���[����: mdlYahooFinance
' ����: Yahoo Finance���犔�����n��f�[�^���擾�i���S����Łj
' ===============================================

Option Explicit

' ===============================================
' ���C�������F�c�^���C�A�E�g�i�����j
' ===============================================
Public Sub GetYahooFinanceData_Vertical(companyName As String, stockCode As String, _
                                        startDate As Date, endDate As Date, timeFrame As String)
    
    Dim ws As Worksheet
    Dim pageNum As Integer
    Dim hasMorePages As Boolean
    Dim pageData As Variant
    Dim outputRow As Long
    Dim i As Long
    Dim totalCount As Long
    
    ' �����̃V�[�g���N���A�܂��͐V�K�쐬
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Yahoo�����f�[�^")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Yahoo�����f�[�^"
    Else
        ws.cells.Clear
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    Application.StatusBar = "�f�[�^�擾��..."
    
    ' �w�b�_�[�s
    ws.cells(1, 1).Value = "��Ж�"
    ws.cells(1, 2).Value = "�،��R�[�h"
    ws.cells(1, 3).Value = "���t"
    ws.cells(1, 4).Value = "������I�l"
    
    outputRow = 2
    totalCount = 0
    
    ' ���ׂẴy�[�W����f�[�^���擾
    pageNum = 1
    hasMorePages = True
    
    Do While hasMorePages
        Application.StatusBar = "�f�[�^�擾��... �y�[�W " & pageNum
        
        pageData = GetPageData(stockCode, startDate, endDate, timeFrame, pageNum)
        
        If Not IsEmpty(pageData) Then
            ' �y�[�W�f�[�^�𒼐ڃV�[�g�ɏ�������
            For i = LBound(pageData, 1) To UBound(pageData, 1)
                ws.cells(outputRow, 1).Value = companyName
                ws.cells(outputRow, 2).Value = stockCode
                ws.cells(outputRow, 3).Value = pageData(i, 1)
                ws.cells(outputRow, 4).Value = pageData(i, 2)
                outputRow = outputRow + 1
                totalCount = totalCount + 1
            Next i
            
            ' ���y�[�W�̊m�F
            Dim pageRows As Long
            pageRows = UBound(pageData, 1) - LBound(pageData, 1) + 1
            If pageRows < 100 Then
                hasMorePages = False
            Else
                pageNum = pageNum + 1
            End If
        Else
            hasMorePages = False
        End If
        
        ' ���S���u�F�ő�20�y�[�W�܂�
        If pageNum > 20 Then
            hasMorePages = False
        End If
    Loop
    
    ' �����ݒ�
    If totalCount > 0 Then
        With ws
            .rows(1).Font.Bold = True
            .rows(1).Interior.Color = RGB(200, 220, 240)
            .Columns("A:D").AutoFit
            .Range("C2:C" & (totalCount + 1)).NumberFormat = "yyyy/mm/dd"
            .Range("D2:D" & (totalCount + 1)).NumberFormat = "#,##0.00"
        End With
    End If
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    If totalCount > 0 Then
        MsgBox "�f�[�^�擾�����I" & vbCrLf & _
               "�擾����: " & totalCount & "��" & vbCrLf & _
               "�擾�y�[�W��: " & pageNum & "�y�[�W", vbInformation
    Else
        MsgBox "�f�[�^���擾�ł��܂���ł����B" & vbCrLf & _
               "�،��R�[�h����t�͈͂��m�F���Ă��������B", vbExclamation
    End If
    
End Sub

' ===============================================
' ���C�������F���^���C�A�E�g
' ===============================================
Public Sub GetYahooFinanceData_Horizontal(companyName As String, stockCode As String, _
                                         startDate As Date, endDate As Date, timeFrame As String)
    
    Dim ws As Worksheet
    Dim pageNum As Integer
    Dim hasMorePages As Boolean
    Dim pageData As Variant
    Dim outputCol As Long
    Dim i As Long
    Dim totalCount As Long
    Dim maxCols As Long
    
    maxCols = 16000
    
    ' �����̃V�[�g���N���A�܂��͐V�K�쐬
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Yahoo�����f�[�^")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Yahoo�����f�[�^"
    Else
        ws.cells.Clear
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    Application.StatusBar = "�f�[�^�擾��..."
    
    ' �w�b�_�[
    ws.cells(1, 1).Value = "��Ж�"
    ws.cells(1, 2).Value = "�،��R�[�h"
    ws.cells(2, 1).Value = companyName
    ws.cells(2, 2).Value = stockCode
    
    outputCol = 3
    totalCount = 0
    
    ' ���ׂẴy�[�W����f�[�^���擾
    pageNum = 1
    hasMorePages = True
    
    Do While hasMorePages
        Application.StatusBar = "�f�[�^�擾��... �y�[�W " & pageNum
        
        pageData = GetPageData(stockCode, startDate, endDate, timeFrame, pageNum)
        
        If Not IsEmpty(pageData) Then
            ' �񐔐����`�F�b�N
            If outputCol + UBound(pageData, 1) - LBound(pageData, 1) > maxCols Then
                MsgBox "�x��: �f�[�^�������������邽�߁A" & totalCount & "���őł��؂�܂��B", vbExclamation
                Exit Do
            End If
            
            ' �y�[�W�f�[�^�𒼐ڃV�[�g�ɏ�������
            For i = LBound(pageData, 1) To UBound(pageData, 1)
                ws.cells(1, outputCol).Value = pageData(i, 1)
                ws.cells(2, outputCol).Value = pageData(i, 2)
                outputCol = outputCol + 1
                totalCount = totalCount + 1
            Next i
            
            ' ���y�[�W�̊m�F
            Dim pageRows As Long
            pageRows = UBound(pageData, 1) - LBound(pageData, 1) + 1
            If pageRows < 100 Then
                hasMorePages = False
            Else
                pageNum = pageNum + 1
            End If
        Else
            hasMorePages = False
        End If
        
        If pageNum > 20 Then
            hasMorePages = False
        End If
    Loop
    
    ' �����ݒ�
    If totalCount > 0 Then
        With ws
            .rows(1).Font.Bold = True
            .rows(1).Interior.Color = RGB(200, 220, 240)
            .Columns("A:B").AutoFit
            .Range(.cells(1, 3), .cells(1, totalCount + 2)).NumberFormat = "yyyy/mm/dd"
            .Range(.cells(2, 3), .cells(2, totalCount + 2)).NumberFormat = "#,##0.00"
            .Range(.cells(1, 3), .cells(1, totalCount + 2)).AutoFit
        End With
    End If
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    If totalCount > 0 Then
        MsgBox "�f�[�^�擾�����I" & vbCrLf & _
               "�擾����: " & totalCount & "��" & vbCrLf & _
               "�擾�y�[�W��: " & pageNum & "�y�[�W", vbInformation
    Else
        MsgBox "�f�[�^���擾�ł��܂���ł����B", vbExclamation
    End If
    
End Sub

' ===============================================
' 1�y�[�W���̃f�[�^���擾
' ===============================================
Private Function GetPageData(stockCode As String, startDate As Date, endDate As Date, _
                             timeFrame As String, pageNum As Integer) As Variant
    
    On Error GoTo ErrorHandler
    
    Dim url As String
    Dim httpReq As Object
    Dim htmlText As String
    Dim timeFrameCode As String
    
    ' �^�C���t���[���R�[�h�̕ϊ�
    Select Case timeFrame
        Case "����"
            timeFrameCode = "d"
        Case "�T��"
            timeFrameCode = "w"
        Case "����"
            timeFrameCode = "m"
        Case Else
            timeFrameCode = "d"
    End Select
    
    ' URL����
    url = "https://finance.yahoo.co.jp/quote/" & stockCode & ".T/history?" & _
          "styl=stock&from=" & Format(startDate, "yyyymmdd") & _
          "&to=" & Format(endDate, "yyyymmdd") & _
          "&timeFrame=" & timeFrameCode & _
          "&page=" & pageNum
    
    ' HTTP�v��
    Set httpReq = CreateObject("MSXML2.XMLHTTP")
    httpReq.Open "GET", url, False
    httpReq.setRequestHeader "User-Agent", "Mozilla/5.0"
    httpReq.send
    
    If httpReq.Status <> 200 Then
        Debug.Print "HTTP Error: " & httpReq.Status
        GetPageData = Empty
        Exit Function
    End If
    
    htmlText = httpReq.responseText
    GetPageData = ParseStockData(htmlText)
    
    Exit Function
    
ErrorHandler:
    Debug.Print "GetPageData Error: " & Err.Description
    GetPageData = Empty
End Function

' ===============================================
' HTML���犔���f�[�^�𒊏o
' ===============================================
Private Function ParseStockData(htmlText As String) As Variant
    
    On Error GoTo ErrorHandler
    
    Dim htmlDoc As Object
    Dim tables As Object
    Dim table As Object
    Dim rows As Object
    Dim row As Object
    Dim cells As Object
    Dim tempData() As Variant
    Dim resultData() As Variant
    Dim maxRows As Long
    Dim rowCount As Long
    Dim dateStr As String
    Dim priceStr As String
    Dim dateVal As Date
    Dim priceVal As Double
    Dim i As Long
    Dim foundTable As Boolean
    Dim dateColIdx As Long
    Dim priceColIdx As Long
    Dim headerText As String

    Set htmlDoc = CreateObject("HTMLFile")
    htmlDoc.body.innerHTML = htmlText
    Set tables = htmlDoc.getElementsByTagName("table")
    
    ' �ő�s����z��
    Debug.Print "ParseStockData: Found " & tables.Length & " tables in HTML"
    maxRows = 200
    ReDim tempData(1 To maxRows, 1 To 2)
    rowCount = 0
    
    ' �e�[�u��������
    For Each table In tables
        Set rows = table.getElementsByTagName("tr")
        
        If rows.Length > 1 Then
            Set cells = rows(0).getElementsByTagName("th")
            foundTable = False
            dateColIdx = -1
            priceColIdx = -1

            ' �w�b�_�[�s�Łu���t�v�Ɓu������I�l�v�̗񂪂��
            Debug.Print "  Checking table with " & rows.Length & " rows, " & cells.Length & " header cells"
            For i = 0 To cells.Length - 1
                headerText = TrimAll(cells(i).innerText)
                Debug.Print "    Header " & i & ": [" & headerText & "]"

                ' ���t������
                If InStr(headerText, "���t") > 0 Then
                    dateColIdx = i
                    foundTable = True
                    Debug.Print "    -> Found date header at column " & i
                End If

                ' ������I�l������i�D��j
                If InStr(headerText, "������I�l") > 0 Or InStr(headerText, "������") > 0 Then
                    priceColIdx = i
                    Debug.Print "    -> Found adjusted close price header at column " & i
                End If
            Next i

            ' ������I�l���������Ȃ��ꍇ�͏I�l�����
            If priceColIdx = -1 Then
                For i = 0 To cells.Length - 1
                    headerText = TrimAll(cells(i).innerText)
                    If InStr(headerText, "I�l") > 0 Then
                        priceColIdx = i
                        Debug.Print "    -> Found close price header at column " & i
                        Exit For
                    End If
                Next i
            End If

            If foundTable And dateColIdx >= 0 And priceColIdx >= 0 Then
                ' �f�[�^�s������
                Debug.Print "ParseStockData: Found data table with " & rows.Length & " rows"
                Debug.Print "ParseStockData: Using dateColIdx=" & dateColIdx & ", priceColIdx=" & priceColIdx
                For i = 1 To rows.Length - 1
                    Set row = rows(i)
                    Set cells = row.getElementsByTagName("td")

                    If cells.Length > dateColIdx And cells.Length > priceColIdx Then
                        dateStr = TrimAll(cells(dateColIdx).innerText)
                        priceStr = TrimAll(cells(priceColIdx).innerText)
                        priceStr = Replace(priceStr, ",", "")
                        
                        Debug.Print "Row " & i & ": dateStr=[" & dateStr & "] priceStr=[" & priceStr & "]"
                        
                        If dateStr <> "" And priceStr <> "" And IsNumeric(priceStr) Then
                            On Error Resume Next
                            dateVal = ConvertJapaneseDate(dateStr)
'                            priceVal = CDbl(priceStr)
                            
                            If Err.Number = 0 Then
'                                rowCount = rowCount + 1
'                                If rowCount > maxRows Then
'                                    ' �z����g��
'                                    maxRows = maxRows + 100
'                                    ReDim Preserve tempData(1 To maxRows, 1 To 2)
                                If Year(dateVal) >= 1900 And Year(dateVal) <= 2100 Then
                                    priceVal = CDbl(priceStr)

                                    If Err.Number = 0 Then
                                        rowCount = rowCount + 1
                                        If rowCount > maxRows Then

                                            maxRows = maxRows + 100
                                            ReDim Preserve tempData(1 To maxRows, 1 To 2)
                                        End If
                                        tempData(rowCount, 1) = dateVal
                                        tempData(rowCount, 2) = priceVal
                                        Debug.Print "  -> Added: Date=" & dateVal & " Price=" & priceVal
                                    Else
                                        Debug.Print "  -> Price conversion error: " & Err.Description
                                    End If
                                Else
                                    Debug.Print "  -> Date out of range: " & dateVal
                                End If
'                                tempData(rowCount, 1) = dateVal
'                                tempData(rowCount, 2) = priceVal
                            Else
                                Debug.Print "  -> Date conversion error: " & Err.Description
                            End If
                            Err.Clear
                            On Error GoTo ErrorHandler
                        End If
                    End If
                Next i
                
                Debug.Print "ParseStockData: Total rows parsed: " & rowCount
                Exit For
            End If
        End If
    Next table
    
    ' �L���ȃf�[�^�݂̂�Ԃ�
    If rowCount > 0 Then
        ReDim resultData(1 To rowCount, 1 To 2)
        For i = 1 To rowCount
            resultData(i, 1) = tempData(i, 1)
            resultData(i, 2) = tempData(i, 2)
        Next i
        ParseStockData = resultData
    Else
        ParseStockData = Empty
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "ParseStockData Error: " & Err.Description
    ParseStockData = Empty
End Function

' ===============================================
' ���{��`���̓��t��ϊ�
' ===============================================
Private Function TrimAll(text As String) As String

    Dim result As String
    Dim i As Long
    Dim char As String
    Dim charCode As Integer

    result = ""

    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        charCode = Asc(char)

        If charCode <> 32 And charCode <> 9 And charCode <> 10 And charCode <> 13 And charCode <> 160 Then
            result = result & char
        End If
    Next i

    TrimAll = result

End Function

Private Function ConvertJapaneseDate(dateStr As String) As Date
    
'    On Error Resume Next
    On Error GoTo ErrorHandler
    
    Dim pos1 As Integer, pos2 As Integer, pos3 As Integer
    Dim yearStr As String, monthStr As String, dayStr As String
    
'    pos1 = InStr(dateStr, "�N")
'    pos2 = InStr(dateStr, "��")
'    pos3 = InStr(dateStr, "��")

    Dim yearVal As Integer, monthVal As Integer, dayVal As Integer
    Dim resultDate As Date
    Dim cleanStr As String

    Debug.Print "ConvertJapaneseDate Input: [" & dateStr & "]"

    If Len(Trim(dateStr)) = 0 Then
        Debug.Print "  -> Empty string"
        Err.Raise 1001, , "Empty date string"
    End If

    cleanStr = Trim(dateStr)

    pos1 = InStr(cleanStr, "�N")
    pos2 = InStr(cleanStr, "��")
    pos3 = InStr(cleanStr, "��")
    
    If pos1 > 0 And pos2 > 0 And pos3 > 0 Then
'        yearStr = Left(dateStr, pos1 - 1)
'        monthStr = Mid(dateStr, pos1 + 1, pos2 - pos1 - 1)
'        dayStr = Mid(dateStr, pos2 + 1, pos3 - pos2 - 1)
'
'        ConvertJapaneseDate = DateSerial(CInt(yearStr), CInt(monthStr), CInt(dayStr))
'    Else
'        ConvertJapaneseDate = CDate(dateStr)

    End If
        yearStr = Trim(Left(cleanStr, pos1 - 1))
        monthStr = Trim(Mid(cleanStr, pos1 + 1, pos2 - pos1 - 1))
        dayStr = Trim(Mid(cleanStr, pos2 + 1, pos3 - pos2 - 1))

        Debug.Print "  -> Japanese format: Year=[" & yearStr & "] Month=[" & monthStr & "] Day=[" & dayStr & "]"

        If IsNumeric(yearStr) And IsNumeric(monthStr) And IsNumeric(dayStr) Then
            yearVal = CInt(yearStr)
            monthVal = CInt(monthStr)
            dayVal = CInt(dayStr)

            If yearVal >= 1900 And yearVal <= 2100 And _
               monthVal >= 1 And monthVal <= 12 And _
               dayVal >= 1 And dayVal <= 31 Then
                resultDate = DateSerial(yearVal, monthVal, dayVal)
                Debug.Print "  -> Success: " & resultDate
                ConvertJapaneseDate = resultDate
                Exit Function
            Else
                Debug.Print "  -> Invalid range: Year=" & yearVal & " Month=" & monthVal & " Day=" & dayVal
            End If
        End If
'    If Err.Number <> 0 Then
'        ConvertJapaneseDate = Date
'        Err.Clear
    cleanStr = Replace(cleanStr, "-", "/")
    If InStr(cleanStr, "/") > 0 Then
        On Error Resume Next
        resultDate = CDate(cleanStr)
        If Err.Number = 0 Then
            If Year(resultDate) >= 1900 And Year(resultDate) <= 2100 Then
                Debug.Print "  -> Slash format success: " & resultDate
                ConvertJapaneseDate = resultDate
                On Error GoTo ErrorHandler
                Exit Function
            Else
                Debug.Print "  -> Slash format invalid year: " & Year(resultDate)
            End If
        End If
        On Error GoTo ErrorHandler
    End If
    
    On Error Resume Next
    resultDate = CDate(cleanStr)
    If Err.Number = 0 Then
        If Year(resultDate) >= 1900 And Year(resultDate) <= 2100 Then
            Debug.Print "  -> CDate success: " & resultDate
            ConvertJapaneseDate = resultDate
            On Error GoTo ErrorHandler
            Exit Function
        Else
            Debug.Print "  -> CDate invalid year: " & Year(resultDate)
        End If
    End If
    On Error GoTo ErrorHandler

ErrorHandler:
    Debug.Print "  -> Conversion failed for: [" & dateStr & "]"
    Err.Raise 1002, , "Failed to convert date: " & dateStr
    
End Function

' ===============================================
' �f�o�b�O�p�R�[�h
' ===============================================

Public Sub DebugYahooFinance()
    
    Dim stockCode As String
    Dim startDate As Date
    Dim endDate As Date
    Dim timeFrame As String
    Dim pageData As Variant
    Dim i As Long
    
    ' �e�X�g�p�p�����[�^
    stockCode = "8306"
    startDate = DateSerial(2024, 10, 1)
    endDate = DateSerial(2024, 10, 31)
    timeFrame = "����"
    
    Debug.Print "========== �f�o�b�O�J�n =========="
    
    ' 1�y�[�W�ڂ̃f�[�^���擾
    pageData = GetPageData(stockCode, startDate, endDate, timeFrame, 1)
    
    ' pageData�̌^�Ɠ��e���m�F
    Debug.Print "IsEmpty(pageData): " & IsEmpty(pageData)
    Debug.Print "TypeName(pageData): " & TypeName(pageData)
    
    If Not IsEmpty(pageData) Then
        Debug.Print "IsArray(pageData): " & IsArray(pageData)
        
        If IsArray(pageData) Then
            On Error Resume Next
            Debug.Print "LBound(pageData, 1): " & LBound(pageData, 1)
            Debug.Print "UBound(pageData, 1): " & UBound(pageData, 1)
            Debug.Print "LBound(pageData, 2): " & LBound(pageData, 2)
            Debug.Print "UBound(pageData, 2): " & UBound(pageData, 2)
            
            If Err.Number = 0 Then
                ' �ŏ���3�s�̃f�[�^��\��
                For i = LBound(pageData, 1) To LBound(pageData, 1) + 2
                    If i <= UBound(pageData, 1) Then
                        Debug.Print "Row " & i & ": Date=" & TypeName(pageData(i, 1)) & " (" & pageData(i, 1) & "), Price=" & TypeName(pageData(i, 2)) & " (" & pageData(i, 2) & ")"
                    End If
                Next i
            Else
                Debug.Print "�z��A�N�Z�X�G���[: " & Err.Description
            End If
            On Error GoTo 0
        End If
    Else
        Debug.Print "�f�[�^����ł�"
    End If
    
    Debug.Print "========== �f�o�b�O�I�� =========="
    
    MsgBox "�C�~�f�B�G�C�g�E�B���h�E�iCtrl+G�j���m�F���Ă�������", vbInformation
    
End Sub
' ===============================================
' HTML�\���m�F�p�R�[�h
' ===============================================

Public Sub DumpHTML()
    
    Dim url As String
    Dim httpReq As Object
    Dim htmlText As String
    Dim fso As Object
    Dim txtFile As Object
    Dim filePath As String
    
    ' �e�X�g�pURL
    url = "https://finance.yahoo.co.jp/quote/8306.T/history?styl=stock&from=20241001&to=20241031&timeFrame=d&page=1"
    
    ' HTTP�v��
    Set httpReq = CreateObject("MSXML2.XMLHTTP")
    httpReq.Open "GET", url, False
    httpReq.setRequestHeader "User-Agent", "Mozilla/5.0"
    httpReq.send
    
    If httpReq.Status = 200 Then
        htmlText = httpReq.responseText
        
        ' �f�X�N�g�b�v��HTML�t�@�C���Ƃ��ĕۑ�
        filePath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\yahoo_finance_dump.html"
        
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set txtFile = fso.CreateTextFile(filePath, True, True)
        txtFile.Write htmlText
        txtFile.Close
        
        MsgBox "HTML�t�@�C����ۑ����܂���:" & vbCrLf & filePath, vbInformation
        
        ' HTMLDocument�I�u�W�F�N�g�ŉ�͂��ăe�[�u������\��
        Dim htmlDoc As Object
        Dim tables As Object
        Dim table As Object
        Dim tableNum As Long
        
        Set htmlDoc = CreateObject("HTMLFile")
        htmlDoc.body.innerHTML = htmlText
        Set tables = htmlDoc.getElementsByTagName("table")
        
        Debug.Print "========== �e�[�u����� =========="
        Debug.Print "�e�[�u����: " & tables.Length
        
        tableNum = 0
        For Each table In tables
            tableNum = tableNum + 1
            Debug.Print "--- �e�[�u�� " & tableNum & " ---"
            Debug.Print "�s��: " & table.getElementsByTagName("tr").Length
            
            If table.getElementsByTagName("tr").Length > 0 Then
                Dim firstRow As Object
                Set firstRow = table.getElementsByTagName("tr")(0)
                Debug.Print "�ŏ��̍s�̃Z����: " & (firstRow.getElementsByTagName("th").Length + firstRow.getElementsByTagName("td").Length)
                
                ' �ŏ��̍s�̓��e��\��
                Dim cells As Object
                Set cells = firstRow.getElementsByTagName("th")
                If cells.Length > 0 Then
                    Debug.Print "�w�b�_�[: "
                    Dim j As Long
                    For j = 0 To cells.Length - 1
                        Debug.Print "  " & j & ": " & cells(j).innerText
                    Next j
                End If
            End If
            Debug.Print ""
        Next table
        
        MsgBox "�C�~�f�B�G�C�g�E�B���h�E�iCtrl+G�j���m�F���Ă�������", vbInformation
    Else
        MsgBox "HTTP Error: " & httpReq.Status, vbCritical
    End If
    
End Sub
