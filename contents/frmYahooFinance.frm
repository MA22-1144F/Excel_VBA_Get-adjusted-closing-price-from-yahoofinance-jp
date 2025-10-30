VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmYahooFinance 
   Caption         =   "Yahoo Finance �f�[�^�擾"
   ClientHeight    =   3330
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6375
   OleObjectBlob   =   "frmYahooFinance.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmYahooFinance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===============================================
' �t�H�[����: frmYahooFinance
' ����: Yahoo Finance �f�[�^�擾�p���[�U�[�t�H�[���i�ŏI�Łj
' ===============================================

Option Explicit

' ===============================================
' �t�H�[��������
' ===============================================
Private Sub UserForm_Initialize()
    
    ' �^�C���t���[���̃R���{�{�b�N�X��ݒ�
    With Me.cmbTimeFrame
        .Clear
        .AddItem "����"
        .AddItem "�T��"
        .AddItem "����"
        .ListIndex = 0
    End With
    
    ' ���C�A�E�g�̃R���{�{�b�N�X��ݒ�
    With Me.cmbLayout
        .Clear
        .AddItem "���^�i���t���ɔz�u�j"
        .AddItem "�c�^�i���t���s�ɔz�u�j"
        .ListIndex = 1 ' �f�t�H���g�͏c�^
    End With
    
    ' �f�t�H���g�l�̐ݒ�
    Me.txtEndDate.Value = Format(Date, "yyyy/mm/dd")
    Me.txtStartDate.Value = Format(DateAdd("yyyy", -1, Date), "yyyy/mm/dd")
    Me.txtCompanyName.Value = ""
    Me.txtStockCode.Value = ""
    
End Sub

' ===============================================
' ���s�{�^���N���b�N
' ===============================================
Private Sub btnExecute_Click()
    
    Dim companyName As String
    Dim stockCode As String
    Dim startDate As Date
    Dim endDate As Date
    Dim timeFrame As String
    Dim layoutType As Integer
    
    ' ���̓`�F�b�N
    If Not ValidateInput() Then
        Exit Sub
    End If
    
    ' ���͒l�̎擾
    companyName = Trim(Me.txtCompanyName.Value)
    stockCode = Trim(Me.txtStockCode.Value)
    startDate = CDate(Me.txtStartDate.Value)
    endDate = CDate(Me.txtEndDate.Value)
    timeFrame = Me.cmbTimeFrame.Value
    layoutType = Me.cmbLayout.ListIndex
    
    ' �t�H�[�����\��
    Me.Hide
    
    ' �f�[�^�擾���������s
    If layoutType = 0 Then
        ' ���^���C�A�E�g
        Call GetYahooFinanceData_Horizontal(companyName, stockCode, startDate, endDate, timeFrame)
    Else
        ' �c�^���C�A�E�g
        Call GetYahooFinanceData_Vertical(companyName, stockCode, startDate, endDate, timeFrame)
    End If
    
    ' �t�H�[�������
    Unload Me
    
End Sub

' ===============================================
' �L�����Z���{�^���N���b�N
' ===============================================
Private Sub btnCancel_Click()
    Unload Me
End Sub

' ===============================================
' ���͒l�̌���
' ===============================================
Private Function ValidateInput() As Boolean
    
    ValidateInput = False
    
    ' ��Ж��`�F�b�N
    If Trim(Me.txtCompanyName.Value) = "" Then
        MsgBox "��Ж�����͂��Ă��������B", vbExclamation
        Me.txtCompanyName.SetFocus
        Exit Function
    End If
    
    ' �،��R�[�h�`�F�b�N
    If Trim(Me.txtStockCode.Value) = "" Then
        MsgBox "�،��R�[�h����͂��Ă��������B", vbExclamation
        Me.txtStockCode.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(Me.txtStockCode.Value) Then
        MsgBox "�،��R�[�h�͐��l�œ��͂��Ă��������B", vbExclamation
        Me.txtStockCode.SetFocus
        Exit Function
    End If
    
    ' �J�n���`�F�b�N
    If Not IsDate(Me.txtStartDate.Value) Then
        MsgBox "�J�n�������������t�`���ł͂���܂���B(YYYY/MM/DD)", vbExclamation
        Me.txtStartDate.SetFocus
        Exit Function
    End If
    
    ' �I�����`�F�b�N
    If Not IsDate(Me.txtEndDate.Value) Then
        MsgBox "�I���������������t�`���ł͂���܂���B(YYYY/MM/DD)", vbExclamation
        Me.txtEndDate.SetFocus
        Exit Function
    End If
    
    ' ���t�̑O��֌W�`�F�b�N
    If CDate(Me.txtStartDate.Value) > CDate(Me.txtEndDate.Value) Then
        MsgBox "�J�n���͏I�������O�̓��t���w�肵�Ă��������B", vbExclamation
        Me.txtStartDate.SetFocus
        Exit Function
    End If
    
    ' �^�C���t���[���`�F�b�N
    If Me.cmbTimeFrame.ListIndex = -1 Then
        MsgBox "�X�p����I�����Ă��������B", vbExclamation
        Me.cmbTimeFrame.SetFocus
        Exit Function
    End If
    
    ' ���C�A�E�g�`�F�b�N
    If Me.cmbLayout.ListIndex = -1 Then
        MsgBox "���C�A�E�g��I�����Ă��������B", vbExclamation
        Me.cmbLayout.SetFocus
        Exit Function
    End If
    
    ValidateInput = True
    
End Function

' ===============================================
' �،��R�[�h�t�B�[���h�F���l�̂ݓ���
' ===============================================
Private Sub txtStockCode_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

