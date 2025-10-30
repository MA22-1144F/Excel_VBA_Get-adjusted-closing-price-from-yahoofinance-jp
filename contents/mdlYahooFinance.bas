Attribute VB_Name = "mdlYahooFinance"
' ===============================================
' モジュール名: mdlYahooFinance
' 説明: Yahoo Financeから株価時系列データを取得（完全動作版）
' ===============================================

Option Explicit

' ===============================================
' メイン処理：縦型レイアウト（推奨）
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
    
    ' 既存のシートをクリアまたは新規作成
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Yahoo株価データ")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Yahoo株価データ"
    Else
        ws.cells.Clear
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    Application.StatusBar = "データ取得中..."
    
    ' ヘッダー行
    ws.cells(1, 1).Value = "会社名"
    ws.cells(1, 2).Value = "証券コード"
    ws.cells(1, 3).Value = "日付"
    ws.cells(1, 4).Value = "調整後終値"
    
    outputRow = 2
    totalCount = 0
    
    ' すべてのページからデータを取得
    pageNum = 1
    hasMorePages = True
    
    Do While hasMorePages
        Application.StatusBar = "データ取得中... ページ " & pageNum
        
        pageData = GetPageData(stockCode, startDate, endDate, timeFrame, pageNum)
        
        If Not IsEmpty(pageData) Then
            ' ページデータを直接シートに書き込み
            For i = LBound(pageData, 1) To UBound(pageData, 1)
                ws.cells(outputRow, 1).Value = companyName
                ws.cells(outputRow, 2).Value = stockCode
                ws.cells(outputRow, 3).Value = pageData(i, 1)
                ws.cells(outputRow, 4).Value = pageData(i, 2)
                outputRow = outputRow + 1
                totalCount = totalCount + 1
            Next i
            
            ' 次ページの確認
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
        
        ' 安全装置：最大20ページまで
        If pageNum > 20 Then
            hasMorePages = False
        End If
    Loop
    
    ' 書式設定
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
        MsgBox "データ取得完了！" & vbCrLf & _
               "取得件数: " & totalCount & "件" & vbCrLf & _
               "取得ページ数: " & pageNum & "ページ", vbInformation
    Else
        MsgBox "データが取得できませんでした。" & vbCrLf & _
               "証券コードや日付範囲を確認してください。", vbExclamation
    End If
    
End Sub

' ===============================================
' メイン処理：横型レイアウト
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
    
    ' 既存のシートをクリアまたは新規作成
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Yahoo株価データ")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Yahoo株価データ"
    Else
        ws.cells.Clear
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    Application.StatusBar = "データ取得中..."
    
    ' ヘッダー
    ws.cells(1, 1).Value = "会社名"
    ws.cells(1, 2).Value = "証券コード"
    ws.cells(2, 1).Value = companyName
    ws.cells(2, 2).Value = stockCode
    
    outputCol = 3
    totalCount = 0
    
    ' すべてのページからデータを取得
    pageNum = 1
    hasMorePages = True
    
    Do While hasMorePages
        Application.StatusBar = "データ取得中... ページ " & pageNum
        
        pageData = GetPageData(stockCode, startDate, endDate, timeFrame, pageNum)
        
        If Not IsEmpty(pageData) Then
            ' 列数制限チェック
            If outputCol + UBound(pageData, 1) - LBound(pageData, 1) > maxCols Then
                MsgBox "警告: データ件数が多すぎるため、" & totalCount & "件で打ち切ります。", vbExclamation
                Exit Do
            End If
            
            ' ページデータを直接シートに書き込み
            For i = LBound(pageData, 1) To UBound(pageData, 1)
                ws.cells(1, outputCol).Value = pageData(i, 1)
                ws.cells(2, outputCol).Value = pageData(i, 2)
                outputCol = outputCol + 1
                totalCount = totalCount + 1
            Next i
            
            ' 次ページの確認
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
    
    ' 書式設定
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
        MsgBox "データ取得完了！" & vbCrLf & _
               "取得件数: " & totalCount & "件" & vbCrLf & _
               "取得ページ数: " & pageNum & "ページ", vbInformation
    Else
        MsgBox "データが取得できませんでした。", vbExclamation
    End If
    
End Sub

' ===============================================
' 1ページ分のデータを取得
' ===============================================
Private Function GetPageData(stockCode As String, startDate As Date, endDate As Date, _
                             timeFrame As String, pageNum As Integer) As Variant
    
    On Error GoTo ErrorHandler
    
    Dim url As String
    Dim httpReq As Object
    Dim htmlText As String
    Dim timeFrameCode As String
    
    ' タイムフレームコードの変換
    Select Case timeFrame
        Case "日間"
            timeFrameCode = "d"
        Case "週間"
            timeFrameCode = "w"
        Case "月間"
            timeFrameCode = "m"
        Case Else
            timeFrameCode = "d"
    End Select
    
    ' URL生成
    url = "https://finance.yahoo.co.jp/quote/" & stockCode & ".T/history?" & _
          "styl=stock&from=" & Format(startDate, "yyyymmdd") & _
          "&to=" & Format(endDate, "yyyymmdd") & _
          "&timeFrame=" & timeFrameCode & _
          "&page=" & pageNum
    
    ' HTTP要求
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
' HTMLから株価データを抽出
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
    
    Set htmlDoc = CreateObject("HTMLFile")
    htmlDoc.body.innerHTML = htmlText
    Set tables = htmlDoc.getElementsByTagName("table")
    
    ' 最大行数を想定
    maxRows = 200
    ReDim tempData(1 To maxRows, 1 To 2)
    rowCount = 0
    
    ' テーブルを検索
    For Each table In tables
        Set rows = table.getElementsByTagName("tr")
        
        If rows.Length > 1 Then
            Set cells = rows(0).getElementsByTagName("th")
            foundTable = False
            
            ' ヘッダー行で「日付」列を確認
            For i = 0 To cells.Length - 1
                If InStr(cells(i).innerText, "日付") > 0 Then
                    foundTable = True
                    Exit For
                End If
            Next i
            
            If foundTable Then
                ' データ行を処理
                For i = 1 To rows.Length - 1
                    Set row = rows(i)
                    Set cells = row.getElementsByTagName("td")
                    
                    If cells.Length >= 6 Then
                        dateStr = Trim(cells(0).innerText)
                        priceStr = Trim(cells(5).innerText)
                        priceStr = Replace(priceStr, ",", "")
                        
                        If dateStr <> "" And priceStr <> "" And IsNumeric(priceStr) Then
                            On Error Resume Next
                            dateVal = ConvertJapaneseDate(dateStr)
                            priceVal = CDbl(priceStr)
                            
                            If Err.Number = 0 Then
                                rowCount = rowCount + 1
                                If rowCount > maxRows Then
                                    ' 配列を拡張
                                    maxRows = maxRows + 100
                                    ReDim Preserve tempData(1 To maxRows, 1 To 2)
                                End If
                                tempData(rowCount, 1) = dateVal
                                tempData(rowCount, 2) = priceVal
                            End If
                            On Error GoTo ErrorHandler
                        End If
                    End If
                Next i
                
                Exit For
            End If
        End If
    Next table
    
    ' 有効なデータのみを返す
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
' 日本語形式の日付を変換
' ===============================================
Private Function ConvertJapaneseDate(dateStr As String) As Date
    
    On Error Resume Next
    
    Dim pos1 As Integer, pos2 As Integer, pos3 As Integer
    Dim yearStr As String, monthStr As String, dayStr As String
    
    pos1 = InStr(dateStr, "年")
    pos2 = InStr(dateStr, "月")
    pos3 = InStr(dateStr, "日")
    
    If pos1 > 0 And pos2 > 0 And pos3 > 0 Then
        yearStr = Left(dateStr, pos1 - 1)
        monthStr = Mid(dateStr, pos1 + 1, pos2 - pos1 - 1)
        dayStr = Mid(dateStr, pos2 + 1, pos3 - pos2 - 1)
        
        ConvertJapaneseDate = DateSerial(CInt(yearStr), CInt(monthStr), CInt(dayStr))
    Else
        ConvertJapaneseDate = CDate(dateStr)
    End If
    
    If Err.Number <> 0 Then
        ConvertJapaneseDate = Date
        Err.Clear
    End If
    
End Function

' ===============================================
' デバッグ用コード
' ===============================================

Public Sub DebugYahooFinance()
    
    Dim stockCode As String
    Dim startDate As Date
    Dim endDate As Date
    Dim timeFrame As String
    Dim pageData As Variant
    Dim i As Long
    
    ' テスト用パラメータ
    stockCode = "8306"
    startDate = DateSerial(2024, 10, 1)
    endDate = DateSerial(2024, 10, 31)
    timeFrame = "日間"
    
    Debug.Print "========== デバッグ開始 =========="
    
    ' 1ページ目のデータを取得
    pageData = GetPageData(stockCode, startDate, endDate, timeFrame, 1)
    
    ' pageDataの型と内容を確認
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
                ' 最初の3行のデータを表示
                For i = LBound(pageData, 1) To LBound(pageData, 1) + 2
                    If i <= UBound(pageData, 1) Then
                        Debug.Print "Row " & i & ": Date=" & TypeName(pageData(i, 1)) & " (" & pageData(i, 1) & "), Price=" & TypeName(pageData(i, 2)) & " (" & pageData(i, 2) & ")"
                    End If
                Next i
            Else
                Debug.Print "配列アクセスエラー: " & Err.Description
            End If
            On Error GoTo 0
        End If
    Else
        Debug.Print "データが空です"
    End If
    
    Debug.Print "========== デバッグ終了 =========="
    
    MsgBox "イミディエイトウィンドウ（Ctrl+G）を確認してください", vbInformation
    
End Sub
' ===============================================
' HTML構造確認用コード
' ===============================================

Public Sub DumpHTML()
    
    Dim url As String
    Dim httpReq As Object
    Dim htmlText As String
    Dim fso As Object
    Dim txtFile As Object
    Dim filePath As String
    
    ' テスト用URL
    url = "https://finance.yahoo.co.jp/quote/8306.T/history?styl=stock&from=20241001&to=20241031&timeFrame=d&page=1"
    
    ' HTTP要求
    Set httpReq = CreateObject("MSXML2.XMLHTTP")
    httpReq.Open "GET", url, False
    httpReq.setRequestHeader "User-Agent", "Mozilla/5.0"
    httpReq.send
    
    If httpReq.Status = 200 Then
        htmlText = httpReq.responseText
        
        ' デスクトップにHTMLファイルとして保存
        filePath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\yahoo_finance_dump.html"
        
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set txtFile = fso.CreateTextFile(filePath, True, True)
        txtFile.Write htmlText
        txtFile.Close
        
        MsgBox "HTMLファイルを保存しました:" & vbCrLf & filePath, vbInformation
        
        ' HTMLDocumentオブジェクトで解析してテーブル情報を表示
        Dim htmlDoc As Object
        Dim tables As Object
        Dim table As Object
        Dim tableNum As Long
        
        Set htmlDoc = CreateObject("HTMLFile")
        htmlDoc.body.innerHTML = htmlText
        Set tables = htmlDoc.getElementsByTagName("table")
        
        Debug.Print "========== テーブル情報 =========="
        Debug.Print "テーブル数: " & tables.Length
        
        tableNum = 0
        For Each table In tables
            tableNum = tableNum + 1
            Debug.Print "--- テーブル " & tableNum & " ---"
            Debug.Print "行数: " & table.getElementsByTagName("tr").Length
            
            If table.getElementsByTagName("tr").Length > 0 Then
                Dim firstRow As Object
                Set firstRow = table.getElementsByTagName("tr")(0)
                Debug.Print "最初の行のセル数: " & (firstRow.getElementsByTagName("th").Length + firstRow.getElementsByTagName("td").Length)
                
                ' 最初の行の内容を表示
                Dim cells As Object
                Set cells = firstRow.getElementsByTagName("th")
                If cells.Length > 0 Then
                    Debug.Print "ヘッダー: "
                    Dim j As Long
                    For j = 0 To cells.Length - 1
                        Debug.Print "  " & j & ": " & cells(j).innerText
                    Next j
                End If
            End If
            Debug.Print ""
        Next table
        
        MsgBox "イミディエイトウィンドウ（Ctrl+G）を確認してください", vbInformation
    Else
        MsgBox "HTTP Error: " & httpReq.Status, vbCritical
    End If
    
End Sub
