VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmYahooFinance 
   Caption         =   "Yahoo Finance データ取得"
   ClientHeight    =   3330
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6375
   OleObjectBlob   =   "frmYahooFinance.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmYahooFinance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===============================================
' フォーム名: frmYahooFinance
' 説明: Yahoo Finance データ取得用ユーザーフォーム（最終版）
' ===============================================

Option Explicit

' ===============================================
' フォーム初期化
' ===============================================
Private Sub UserForm_Initialize()
    
    ' タイムフレームのコンボボックスを設定
    With Me.cmbTimeFrame
        .Clear
        .AddItem "日間"
        .AddItem "週間"
        .AddItem "月間"
        .ListIndex = 0
    End With
    
    ' レイアウトのコンボボックスを設定
    With Me.cmbLayout
        .Clear
        .AddItem "横型（日付を列に配置）"
        .AddItem "縦型（日付を行に配置）"
        .ListIndex = 1 ' デフォルトは縦型
    End With
    
    ' デフォルト値の設定
    Me.txtEndDate.Value = Format(Date, "yyyy/mm/dd")
    Me.txtStartDate.Value = Format(DateAdd("yyyy", -1, Date), "yyyy/mm/dd")
    Me.txtCompanyName.Value = ""
    Me.txtStockCode.Value = ""
    
End Sub

' ===============================================
' 実行ボタンクリック
' ===============================================
Private Sub btnExecute_Click()
    
    Dim companyName As String
    Dim stockCode As String
    Dim startDate As Date
    Dim endDate As Date
    Dim timeFrame As String
    Dim layoutType As Integer
    
    ' 入力チェック
    If Not ValidateInput() Then
        Exit Sub
    End If
    
    ' 入力値の取得
    companyName = Trim(Me.txtCompanyName.Value)
    stockCode = Trim(Me.txtStockCode.Value)
    startDate = CDate(Me.txtStartDate.Value)
    endDate = CDate(Me.txtEndDate.Value)
    timeFrame = Me.cmbTimeFrame.Value
    layoutType = Me.cmbLayout.ListIndex
    
    ' フォームを非表示
    Me.Hide
    
    ' データ取得処理を実行
    If layoutType = 0 Then
        ' 横型レイアウト
        Call GetYahooFinanceData_Horizontal(companyName, stockCode, startDate, endDate, timeFrame)
    Else
        ' 縦型レイアウト
        Call GetYahooFinanceData_Vertical(companyName, stockCode, startDate, endDate, timeFrame)
    End If
    
    ' フォームを閉じる
    Unload Me
    
End Sub

' ===============================================
' キャンセルボタンクリック
' ===============================================
Private Sub btnCancel_Click()
    Unload Me
End Sub

' ===============================================
' 入力値の検証
' ===============================================
Private Function ValidateInput() As Boolean
    
    ValidateInput = False
    
    ' 会社名チェック
    If Trim(Me.txtCompanyName.Value) = "" Then
        MsgBox "会社名を入力してください。", vbExclamation
        Me.txtCompanyName.SetFocus
        Exit Function
    End If
    
    ' 証券コードチェック
    If Trim(Me.txtStockCode.Value) = "" Then
        MsgBox "証券コードを入力してください。", vbExclamation
        Me.txtStockCode.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(Me.txtStockCode.Value) Then
        MsgBox "証券コードは数値で入力してください。", vbExclamation
        Me.txtStockCode.SetFocus
        Exit Function
    End If
    
    ' 開始日チェック
    If Not IsDate(Me.txtStartDate.Value) Then
        MsgBox "開始日が正しい日付形式ではありません。(YYYY/MM/DD)", vbExclamation
        Me.txtStartDate.SetFocus
        Exit Function
    End If
    
    ' 終了日チェック
    If Not IsDate(Me.txtEndDate.Value) Then
        MsgBox "終了日が正しい日付形式ではありません。(YYYY/MM/DD)", vbExclamation
        Me.txtEndDate.SetFocus
        Exit Function
    End If
    
    ' 日付の前後関係チェック
    If CDate(Me.txtStartDate.Value) > CDate(Me.txtEndDate.Value) Then
        MsgBox "開始日は終了日より前の日付を指定してください。", vbExclamation
        Me.txtStartDate.SetFocus
        Exit Function
    End If
    
    ' タイムフレームチェック
    If Me.cmbTimeFrame.ListIndex = -1 Then
        MsgBox "スパンを選択してください。", vbExclamation
        Me.cmbTimeFrame.SetFocus
        Exit Function
    End If
    
    ' レイアウトチェック
    If Me.cmbLayout.ListIndex = -1 Then
        MsgBox "レイアウトを選択してください。", vbExclamation
        Me.cmbLayout.SetFocus
        Exit Function
    End If
    
    ValidateInput = True
    
End Function

' ===============================================
' 証券コードフィールド：数値のみ入力
' ===============================================
Private Sub txtStockCode_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

