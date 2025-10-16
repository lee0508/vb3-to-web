VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm출력물관리 
   BorderStyle     =   1  '단일 고정
   Caption         =   "출력물관리"
   ClientHeight    =   9690
   ClientLeft      =   90
   ClientTop       =   1125
   ClientWidth     =   7845
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "출력물관리.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   10060
   ScaleMode       =   0  '사용자
   ScaleWidth      =   7845
   Begin VB.Frame Frame2 
      Caption         =   "(조  건)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2119
      Left            =   120
      TabIndex        =   16
      Top             =   6983
      Width           =   7575
      Begin VB.TextBox txtSupName 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1440
         MaxLength       =   8
         TabIndex        =   8
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txtBuyName 
         Enabled         =   0   'False
         Height          =   270
         Left            =   4800
         MaxLength       =   8
         TabIndex        =   9
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox txtMtName 
         Enabled         =   0   'False
         Height          =   270
         Left            =   4800
         MaxLength       =   18
         TabIndex        =   5
         Top             =   960
         Width           =   2535
      End
      Begin VB.ComboBox cboKind 
         Height          =   300
         Left            =   1440
         Style           =   2  '드롭다운 목록
         TabIndex        =   3
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtMt 
         Height          =   270
         Left            =   4800
         MaxLength       =   18
         TabIndex        =   4
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtSup 
         Height          =   270
         Left            =   1440
         MaxLength       =   8
         TabIndex        =   6
         Top             =   1325
         Width           =   1095
      End
      Begin VB.TextBox txtBuy 
         Height          =   270
         Left            =   4800
         MaxLength       =   8
         TabIndex        =   7
         Top             =   1325
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtpF_Date 
         Height          =   270
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   56754177
         CurrentDate     =   38190
      End
      Begin MSComCtl2.DTPicker dtpT_Date 
         Height          =   270
         Left            =   3480
         TabIndex        =   2
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   56754177
         CurrentDate     =   38190
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "매입처명 :"
         Height          =   255
         Index           =   9
         Left            =   3360
         TabIndex        =   28
         Top             =   1750
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "매입처명 :"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   27
         Top             =   1750
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "품목명 :"
         Height          =   255
         Index           =   7
         Left            =   3360
         TabIndex        =   26
         Top             =   990
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "기준일자 :"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   285
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "-"
         Height          =   255
         Left            =   3000
         TabIndex        =   24
         Top             =   285
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "자재분류 :"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "품목코드 :"
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   22
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "매입처코드 :"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   21
         Top             =   1350
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "매출처코드 :"
         Height          =   255
         Index           =   4
         Left            =   3360
         TabIndex        =   20
         Top             =   1350
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "[Home]"
         Height          =   240
         Index           =   15
         Left            =   2640
         TabIndex        =   19
         Top             =   1350
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "[Home]"
         Height          =   240
         Index           =   5
         Left            =   6000
         TabIndex        =   18
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "[Home]"
         Height          =   240
         Index           =   6
         Left            =   6840
         TabIndex        =   17
         Top             =   630
         Width           =   615
      End
   End
   Begin VB.ListBox lstPort 
      Height          =   240
      Left            =   1200
      TabIndex        =   15
      Top             =   9255
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   300
      Left            =   1800
      Style           =   2  '드롭다운 목록
      TabIndex        =   14
      Top             =   9270
      Visible         =   0   'False
      Width           =   3135
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   720
      Top             =   9240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdExit 
      Height          =   390
      Left            =   6600
      Picture         =   "출력물관리.frx":014A
      Style           =   1  '그래픽
      TabIndex        =   11
      Top             =   9240
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   375
      Left            =   5280
      Picture         =   "출력물관리.frx":0A98
      Style           =   1  '그래픽
      TabIndex        =   10
      Top             =   9240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "(제  목)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   7575
      Begin VB.Label lblSelect 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "출력물을 선택하세요."
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   210
         Width           =   7095
      End
   End
   Begin MSComctlLib.TreeView TrView 
      Height          =   6097
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   10742
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageL"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageL 
      Left            =   120
      Top             =   9120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "출력물관리.frx":13FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "출력물관리.frx":1555
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm출력물관리"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' 프로그램 제 목 : 출력물관리
' 사용된 Control :
' 참조된 Table   :
' 업  무  설  명 :
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived         As Boolean
Private P_adoRec             As New ADODB.Recordset
Private P_intButton          As Integer
Private P_strFindString2     As String
Private KeySet               As String
Private prn_Select           As Integer

'+--------------------------------+
'/// LOAD FORM ( 한번만 실행 ) ///
'+--------------------------------+
Private Sub Form_Load()
    P_blnActived = False
End Sub

'+-------------------------------------------+
'/// ACTIVATE FORM 활성화 ( 한번만 실행 ) ///
'+-------------------------------------------+
Private Sub Form_Activate()
Dim strSQL             As String
Dim inti               As Integer

Dim P                  As Printer
Dim strDefaultPrinter  As String
Dim aryPrinter()       As String
Dim strBuffer          As String

    frmMain.SBar.Panels(4).Text = ""
    If P_blnActived = False Then
       Screen.MousePointer = vbHourglass
       On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
       
       'Printer Setting and Seaching(API)
       strBuffer = Space(1024)
       inti = GetProfileString("windows", "Device", "", strBuffer, Len(strBuffer))
       aryPrinter = Split(strBuffer, ",")
       strDefaultPrinter = Trim(Trim(aryPrinter(0)))
       For Each P In Printers
           cboPrinter.AddItem Trim(P.DeviceName)
           lstPort.AddItem P.Port
       Next
       For inti = 0 To cboPrinter.ListCount - 1
           cboPrinter.ListIndex = inti
           If UCase(Trim(cboPrinter.Text)) = UCase(Trim(strDefaultPrinter)) Then
              Exit For
           End If
       Next inti
       '---
       SubOther_FILL
       dtpF_Date.Value = Now: dtpT_Date.Value = Now:
       SubTreeAdd
       P_blnActived = True
       Screen.MousePointer = vbDefault
    End If
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "출력물관리(서버와의 연결 실패)"
    Unload Me
    Exit Sub
End Sub

'+----------+
'--- 입력 ---
'+----------+
Private Sub dtpF_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{Tab}"
    End If
End Sub

Private Sub dtpT_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{Tab}"
    End If
End Sub

'+--------+
' 자재분류
'+--------+
Private Sub cboKind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{Tab}"
    End If
End Sub

'+--------+
' 자재코드
'+--------+
Private Sub txtMt_GotFocus()
    With txtMt
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtMt_Keydown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn) And Len(Trim(txtMt.Text)) > 0 Then '자재검색
       PB_strCallFormName = "frm출력물관리"
       PB_strMaterialsCode = Trim(txtMt.Text)
       PB_strMaterialsName = ""
       frm자재검색.Show vbModal
       If (Len(PB_strMaterialsCode) + Len(PB_strMaterialsName)) = 0 Then '검색에서 취소(ESC)
       Else
          txtMt.Text = Mid(PB_strMaterialsCode, 3)
          txtMtName.Text = PB_strMaterialsName
       End If
       If PB_strMaterialsCode <> "" Then
          SendKeys "{tab}"
       End If
       PB_strCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
    Else
       If KeyCode = vbKeyReturn Then
          SendKeys "{tab}"
       End If
    End If
    Exit Sub
End Sub
Private Sub txtMt_LostFocus()
    With txtMt
         .Text = Trim(.Text)
         If Len(.Text) = 0 Then
            txtMtName.Text = ""
         End If
    End With
End Sub
'+----------+
' 매입처코드
'+----------+
Private Sub txtSup_GotFocus()
    With txtSup
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtSup_Keydown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn) And Len(Trim(txtSup.Text)) > 0 Then  '매입처검색
       PB_strSupplierCode = UPPER(Trim(txtSup.Text))
       PB_strSupplierName = "" 'Trim(Text1(Index + 1).Text)
       frm매입처검색.Show vbModal
       If (Len(PB_strSupplierCode) + Len(PB_strSupplierName)) = 0 Then '검색에서 취소(ESC)
       Else
          txtSup.Text = PB_strSupplierCode
          txtSupName.Text = PB_strSupplierName
       End If
       If PB_strSupplierCode <> "" Then
          SendKeys "{TAB}"
       End If
       PB_strSupplierCode = "": PB_strSupplierName = ""
    Else
       If KeyCode = vbKeyReturn Then
          SendKeys "{tab}"
       End If
    End If
    Exit Sub
End Sub
Private Sub txtSup_LostFocus()
    With txtSup
         .Text = Trim(.Text)
         If Len(.Text) = 0 Then
            txtSupName.Text = ""
         End If
    End With
End Sub
'+----------+
' 매출처코드
'+----------+
Private Sub txtBuy_GotFocus()
    With txtBuy
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtBuy_Keydown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn) And Len(Trim(txtBuy.Text)) > 0 Then  '매출처검색
       PB_strSupplierCode = UPPER(Trim(txtBuy.Text))
       PB_strSupplierName = "" 'Trim(Text1(Index + 1).Text)
       frm매출처검색.Show vbModal
       If (Len(PB_strSupplierCode) + Len(PB_strSupplierName)) = 0 Then '검색에서 취소(ESC)
       Else
          txtBuy.Text = PB_strSupplierCode
          txtBuyName.Text = PB_strSupplierName
       End If
       If PB_strSupplierCode <> "" Then
          SendKeys "{TAB}"
       End If
       PB_strSupplierCode = "": PB_strSupplierName = ""
    Else
       If KeyCode = vbKeyReturn Then
          SendKeys "{tab}"
       End If
    End If
    Exit Sub
End Sub
Private Sub txtBuy_LostFocus()
    With txtBuy
         .Text = Trim(.Text)
         If Len(.Text) = 0 Then
            txtBuyName.Text = ""
         End If
    End With
End Sub

'+--------------------+
'--- Select Printer ---
'+--------------------+
Private Sub cboPrinter_Click()
    lstPort.ListIndex = cboPrinter.ListIndex
End Sub

'+---------------+
'--- 출력 선택 ---
' 2,업체별 미지급금 현황
'+---------------+
'+--------+
' 출력물
'+--------+
Private Sub TrView_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
           Select Case prn_Select
           Case 1, 3 To 4, 6 To 11, 13 To 14, 16 To 22, 24 To 27
                SendKeys "{Tab}"
           Case Else
    End Select
    End If
End Sub

Private Sub cmdPrint_Click()
    If dtpF_Date.Enabled = True Then
       If dtpF_Date.Value > dtpT_Date Then
          dtpF_Date.SetFocus
          Exit Sub
       End If
    End If
    Select Case prn_Select
           Case 0
                MsgBox "출력할 항목을 선택 하세요.", vbCritical, "판매관리(출력물관리)"
                Exit Sub
           Case 2, 5, 12, 15, 23
                MsgBox "출력할 세부항목을 선택 하세요.", vbCritical, "판매관리시스템(출력물관리)"
                Exit Sub
           Case 1, 3 To 4, 6 To 11, 13 To 14, 16 To 22, 24 To 27
                SubPrintCrystalReports
           Case Else
    End Select
End Sub

'+-----------+
'/// 종료 ///
'+-----------+
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
    If P_adoRec.State <> adStateClosed Then
       P_adoRec.Close
    End If
    Set P_adoRec = Nothing
    Set frm출력물관리 = Nothing
    frmMain.SBar.Panels(4).Text = ""
End Sub
'+---------------+
'--- Add Tree ---
'+---------------+
Sub SubTreeAdd()
Dim nodX     As Node
Dim tRee_PRN As Integer

    Set nodX = TrView.Nodes.Add(, , "A", "일계표", 1)                                     '== 1
    
    Set nodX = TrView.Nodes.Add(, , "B", "미지급금", 2)                                 '== 2
    Set nodX = TrView.Nodes.Add("B", tvwChild, "B1", "업체별 미지급금 현황(집계표)", 1)   '== 3
    Set nodX = TrView.Nodes.Add("B", tvwChild, "B2", "업체별 지급 현황(일자별)", 1)       '== 4
    
    Set nodX = TrView.Nodes.Add(, , "C", "매입현황", 2)                                 '== 5
    Set nodX = TrView.Nodes.Add("C", tvwChild, "C1", "품목별 매입 현황", 1)               '== 6
    Set nodX = TrView.Nodes.Add("C", tvwChild, "C2", "업체별 매입 현황(품목별)", 1)       '== 7
    Set nodX = TrView.Nodes.Add("C", tvwChild, "C3", "업체별 매입 현황(일자별)", 1)       '== 8
    Set nodX = TrView.Nodes.Add("C", tvwChild, "C4", "업체별 매입 현황", 1)               '== 9
    Set nodX = TrView.Nodes.Add("C", tvwChild, "C5", "매입세금계산서 장부 현황", 1)       '== 10
    Set nodX = TrView.Nodes.Add("C", tvwChild, "C6", "업체별 매입 현황(3개월간)", 1)      '== 11

    Set nodX = TrView.Nodes.Add(, , "D", "미수금", 2)                                   '== 12
    Set nodX = TrView.Nodes.Add("D", tvwChild, "D1", "업체별 미수금 현황(집계표)", 1)     '== 13
    Set nodX = TrView.Nodes.Add("D", tvwChild, "D2", "업체별 수금 현황(일자별)", 1)       '== 14
    
    Set nodX = TrView.Nodes.Add(, , "E", "매출현황", 2)                                 '== 15
    Set nodX = TrView.Nodes.Add("E", tvwChild, "E1", "품목별 매출 현황", 1)               '== 16
    Set nodX = TrView.Nodes.Add("E", tvwChild, "E2", "업체별 매출 현황(품목별)", 1)       '== 17
    Set nodX = TrView.Nodes.Add("E", tvwChild, "E3", "업체별 매출 현황(일자별)", 1)       '== 18
    Set nodX = TrView.Nodes.Add("E", tvwChild, "E4", "업체별 매출 현황", 1)               '== 19
    Set nodX = TrView.Nodes.Add("E", tvwChild, "E5", "매출세금계산서 장부 현황", 1)       '== 20
    Set nodX = TrView.Nodes.Add("E", tvwChild, "E6", "매출세금계산서 현황", 1)            '== 21
    Set nodX = TrView.Nodes.Add("E", tvwChild, "E7", "업체별 매출 현황(3개월간)", 1)      '== 22
    
    Set nodX = TrView.Nodes.Add(, , "F", "회계관리", 2)                                 '== 23
    Set nodX = TrView.Nodes.Add("F", tvwChild, "F1", "출납관리(일자별)", 1)               '== 24
    Set nodX = TrView.Nodes.Add("F", tvwChild, "F2", "출납관리(계정과목별)", 1)           '== 25
    Set nodX = TrView.Nodes.Add("F", tvwChild, "F3", "출납관리(계정별집계)", 1)           '== 26
    Set nodX = TrView.Nodes.Add("F", tvwChild, "F4", "합계시산표", 1)                     '== 27
    
    If tRee_PRN <> 0 Then
       Set nodX = TrView.Nodes.Item(tRee_PRN)
       TrView_NodeClick nodX
    End If
End Sub

Private Sub TrView_NodeClick(ByVal Node As MSComctlLib.Node)
Dim CpText    As String
    TrView.Nodes(Node.Index).EnsureVisible
    KeySet = Node.Key
    lblSelect.Caption = Node.Text
    If Left(KeySet, 1) = "A" And Len(KeySet) > 0 Then         '일계표
       cboKind.Enabled = False: txtMt.Enabled = False
       txtSup.Enabled = False: txtBuy.Enabled = False
       cboPrinter.Visible = False: cboPrinter.Enabled = False
       dtpF_Date.Enabled = False                              '시작일자 Not Used)
    ElseIf _
       Left(KeySet, 1) = "B" And Len(KeySet) > 1 Then         '미지급금
       cboKind.Enabled = False: txtMt.Enabled = False
       txtSup.Enabled = True: txtBuy.Enabled = False
       cboPrinter.Visible = False: cboPrinter.Enabled = False
       Select Case KeySet
              Case "B1"
                   dtpF_Date.Enabled = False
              Case "B2"
                   dtpF_Date.Enabled = True
              Case Else
                   dtpF_Date.Enabled = True
       End Select
    ElseIf _
       Left(KeySet, 1) = "C" And Len(KeySet) > 1 Then         '매입현황
       cboKind.Enabled = True: txtMt.Enabled = True
       txtSup.Enabled = True: txtBuy.Enabled = False
       cboPrinter.Visible = False: cboPrinter.Enabled = False
       Select Case KeySet
              Case "C1", "C2", "C3", "C4", "C5"
                   dtpF_Date.Enabled = True
              Case "C6"
                   dtpF_Date.Enabled = False
              Case Else
                   dtpF_Date.Enabled = True
       End Select
    ElseIf _
       Left(KeySet, 1) = "D" And Len(KeySet) > 1 Then         '미수금
       cboKind.Enabled = False: txtMt.Enabled = False
       txtSup.Enabled = False: txtBuy.Enabled = True
       cboPrinter.Visible = False: cboPrinter.Enabled = False
       Select Case KeySet
              Case "D1"
                   dtpF_Date.Enabled = False
              Case "D2"
                   dtpF_Date.Enabled = True
              Case Else
                   dtpF_Date.Enabled = True
       End Select
    ElseIf _
       Left(KeySet, 1) = "E" And Len(KeySet) > 1 Then         '매출현황
       cboKind.Enabled = True: txtMt.Enabled = True
       txtSup.Enabled = False: txtBuy.Enabled = True
       cboPrinter.Visible = False: cboPrinter.Enabled = False
       Select Case KeySet
              Case "E1", "E2", "E3", "E4", "E5", "E6"
                   dtpF_Date.Enabled = True
              Case "E7"
                   dtpF_Date.Enabled = False
              Case Else
                   dtpF_Date.Enabled = True
       End Select
    ElseIf _
       Left(KeySet, 1) = "F" And Len(KeySet) > 1 Then         '회계관리
       cboKind.Enabled = False: txtMt.Enabled = False
       txtSup.Enabled = False: txtBuy.Enabled = False
       cboPrinter.Visible = False: cboPrinter.Enabled = False
       Select Case KeySet
              Case "F1", "F2", "F3"
                   dtpF_Date.Enabled = True
              Case "F4"
                   dtpF_Date.Enabled = False
              Case Else
                   dtpF_Date.Enabled = True
       End Select
    Else
       cboKind.Enabled = False: txtMt.Enabled = False
       txtSup.Enabled = False: txtBuy.Enabled = False
       cboPrinter.Visible = False: cboPrinter.Enabled = False
    End If
    prn_Select = Node.Index
End Sub
 
'+----------+
'--- FILL ---
'+-----------+
Private Sub SubOther_FILL()
Dim strSQL        As String
Dim intIndex      As Integer
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.분류코드 AS 분류코드, ISNULL(T1.분류명,'') AS 분류명 " _
             & "FROM 자재분류 T1 " _
            & "ORDER BY T1.분류코드 "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       cboKind.ListIndex = -1
       Exit Sub
    Else
       cboKind.AddItem "00. 전체"
       Do Until P_adoRec.EOF
          cboKind.AddItem P_adoRec("분류코드") & ". " & P_adoRec("분류명")
          P_adoRec.MoveNext
       Loop
       P_adoRec.Close
       cboKind.ListIndex = 0
    End If
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "자재분류 실패"
    Unload Me
    Exit Sub
End Sub

'+---------------------------+
'/// 크리스탈 리포터 출력 ///
'+---------------------------+
Private Sub SubPrintCrystalReports()
Dim strSQL                 As String
Dim strWhere               As String
Dim strOrderBy             As String

Dim varRetVal              As Variant '리포터 파일
Dim strExeFile             As String
Dim strExeMode             As String
Dim intRetCHK              As Integer '실행여부

Dim lngR                   As Long
Dim lngC                   As Long
Dim strForPrtDateTime      As String  '출력일시           (Formula)
    
    Screen.MousePointer = vbHourglass
    '서버일시(출력일시)
    strSQL = "SELECT CONVERT(VARCHAR(19),GETDATE(), 120) AS 서버시간 "
    On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    strForPrtDateTime = Format(PB_regUserinfoU.UserServerDate, "0000-00-00") & Space(1) _
                      & Format(Right(P_adoRec("서버시간"), 8), "hh:mm:ss")
    P_adoRec.Close
    
    intRetCHK = 99
    With CrystalReport1
         Select Case prn_Select
                Case 1 '일계표
                     strExeFile = App.Path & ".\일계표.rpt"
                     varRetVal = Dir(strExeFile)
                     If Len(varRetVal) = 0 Then
                        intRetCHK = 0
                     Else
                        .ReportFileName = strExeFile
                        On Error GoTo ERROR_CRYSTAL_REPORTS
                        '--- Formula Fields ---
                        .Formulas(0) = "ForPgDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'" '프로그램실행일자
                        .Formulas(1) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                    '사업장명
                        .Formulas(2) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                               '출력일시
                        '기준일자
                        .Formulas(3) = "ForAppDate = '기준일자 : ' & '" & Format(DTOS(dtpT_Date.Value), "0000-00-00") & "' "
                        '--- Parameter Fields ---
                        '사업장코드
                        .StoredProcParam(0) = PB_regUserinfoU.UserBranchCode
                        '적용(기준)일자종료
                        .StoredProcParam(1) = DTOS(dtpT_Date.Value)
                        '--- 속성 ---
                        .WindowShowGroupTree = False
                        .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "일계표"
                     End If
                Case 3 '업체별 미지급금 현황(집계표)
                     strExeFile = App.Path & ".\업체별미지급금현황(집계표).rpt"
                     varRetVal = Dir(strExeFile)
                     If Len(varRetVal) = 0 Then
                        intRetCHK = 0
                     Else
                        .ReportFileName = strExeFile
                        On Error GoTo ERROR_CRYSTAL_REPORTS
                        '--- Formula Fields ---
                        .Formulas(0) = "ForPgDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'" '프로그램실행일자
                        .Formulas(1) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                    '사업장명
                        .Formulas(2) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                               '출력일시
                        '기준일자
                        .Formulas(3) = "ForAppDate = '기준일자 : ' & '" & Format(DTOS(dtpT_Date.Value), "0000-00-00") & "' "
                        '--- Parameter Fields ---
                        '사업장코드
                        .StoredProcParam(0) = PB_regUserinfoU.UserBranchCode
                        '적용(기준)일자종료
                        .StoredProcParam(1) = DTOS(dtpT_Date.Value)
                        '매입처코드
                        If (Len(txtSup.Text) = 0) Then
                          .StoredProcParam(2) = " "
                        Else
                          .StoredProcParam(2) = Trim(txtSup.Text)
                        End If
                        '미지급금발생구분 1.전표, 2.계산서
                        .StoredProcParam(3) = PB_regUserinfoU.UserMJGbn
                        '--- 속성 ---
                        .WindowShowGroupTree = False
                        .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "업체별미지급금현황(집계표)"
                     End If
                Case 4 '업체별 지급 현황(일자별)
                     strExeFile = App.Path & ".\업체별지급현황(일자별).rpt"
                     varRetVal = Dir(strExeFile)
                     If Len(varRetVal) = 0 Then
                        intRetCHK = 0
                     Else
                        .ReportFileName = strExeFile
                        On Error GoTo ERROR_CRYSTAL_REPORTS
                        '--- Formula Fields ---
                        .Formulas(0) = "ForPgDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'" '프로그램실행일자
                        .Formulas(1) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                    '사업장명
                        .Formulas(2) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                               '출력일시
                        '기준일자
                        .Formulas(3) = "ForAppDate = '기준일자 : ' & '" & Format(DTOS(dtpF_Date.Value), "0000-00-00") & "' & ' 부터 ' & '" & Format(DTOS(dtpT_Date.Value), "0000-00-00") & "' & ' 까지' "
                        '--- Parameter Fields ---
                        '사업장코드
                        .StoredProcParam(0) = PB_regUserinfoU.UserBranchCode
                        '적용(기준)일자시작
                        .StoredProcParam(1) = DTOS(dtpF_Date.Value)
                        '적용(기준)일자종료
                        .StoredProcParam(2) = DTOS(dtpT_Date.Value)
                        '매입처코드
                        If (Len(txtSup.Text) = 0) Then
                          .StoredProcParam(3) = " "
                        Else
                          .StoredProcParam(3) = Trim(txtSup.Text)
                        End If
                        '--- 속성 ---
                        .WindowShowGroupTree = True
                        .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "업체별지급현황(일자별)"
                     End If
                Case 6 '품목별 매입 현황
                     strExeFile = App.Path & ".\품목별매입현황.rpt"
                     varRetVal = Dir(strExeFile)
                     If Len(varRetVal) = 0 Then
                        intRetCHK = 0
                     Else
                        .ReportFileName = strExeFile
                        On Error GoTo ERROR_CRYSTAL_REPORTS
                        '--- Formula Fields ---
                        .Formulas(0) = "ForPgDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'" '프로그램실행일자
                        .Formulas(1) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                    '사업장명
                        .Formulas(2) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                               '출력일시
                        '기준일자
                        .Formulas(3) = "ForAppDate = '기준일자 : ' & '" & Format(DTOS(dtpF_Date.Value), "0000-00-00") & "' & ' 부터 ' & '" & Format(DTOS(dtpT_Date.Value), "0000-00-00") & "' & ' 까지' "
                        '--- Parameter Fields ---
                        '사업장코드
                        .StoredProcParam(0) = PB_regUserinfoU.UserBranchCode
                        '적용(기준)일자시작
                        .StoredProcParam(1) = DTOS(dtpF_Date.Value)
                        '적용(기준)일자종료
                        .StoredProcParam(2) = DTOS(dtpT_Date.Value)
                        '분류코드
                        .StoredProcParam(3) = Mid(cboKind.Text, 1, 2)
                        '세부코드
                        .StoredProcParam(4) = IIf(Len(Trim(txtMt.Text)) = 0, " ", Trim(txtMt.Text))
                        '--- 속성 ---
                        .WindowShowGroupTree = True
                        .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "품목별매입현황"
                     End If
                Case 7 '업체별 매입 현황(품목별)
                     strExeFile = App.Path & ".\업체별매입현황(품목별).rpt"
                     varRetVal = Dir(strExeFile)
                     If Len(varRetVal) = 0 Then
                        intRetCHK = 0
                     Else
                        .ReportFileName = strExeFile
                        On Error GoTo ERROR_CRYSTAL_REPORTS
                        '--- Formula Fields ---
                        .Formulas(0) = "ForPgDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'" '프로그램실행일자
                        .Formulas(1) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                    '사업장명
                        .Formulas(2) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                               '출력일시
                        '기준일자
                        .Formulas(3) = "ForAppDate = '기준일자 : ' & '" & Format(DTOS(dtpF_Date.Value), "0000-00-00") & "' & ' 부터 ' & '" & Format(DTOS(dtpT_Date.Value), "0000-00-00") & "' & ' 까지' "
                        '--- Parameter Fields ---
                        '사업장코드
                        .StoredProcParam(0) = PB_regUserinfoU.UserBranchCode
                        '적용(기준)일자시작
                        .StoredProcParam(1) = DTOS(dtpF_Date.Value)
                        '적용(기준)일자종료
                        .StoredProcParam(2) = DTOS(dtpT_Date.Value)
                        '매입처코드
                        If (Len(txtSup.Text) = 0) Then
                          .StoredProcParam(3) = " "
                        Else
                          .StoredProcParam(3) = Trim(txtSup.Text)
                        End If
                        '분류코드
                        .StoredProcParam(4) = Mid(cboKind.Text, 1, 2)
                        '세부코드
                        .StoredProcParam(5) = IIf(Len(Trim(txtMt.Text)) = 0, " ", Trim(txtMt.Text))
                        '--- 속성 ---
                        .WindowShowGroupTree = True
                        .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "업체별매입현황(품목별)"
                     End If
                Case 8 '업체별 매입 현황(일자별)
                     strExeFile = App.Path & ".\업체별매입현황(일자별).rpt"
                     varRetVal = Dir(strExeFile)
                     If Len(varRetVal) = 0 Then
                        intRetCHK = 0
                     Else
                        .ReportFileName = strExeFile
                        On Error GoTo ERROR_CRYSTAL_REPORTS
                        '--- Formula Fields ---
                        .Formulas(0) = "ForPgDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'" '프로그램실행일자
                        .Formulas(1) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                    '사업장명
                        .Formulas(2) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                               '출력일시
                        '기준일자
                        .Formulas(3) = "ForAppDate = '기준일자 : ' & '" & Format(DTOS(dtpF_Date.Value), "0000-00-00") & "' & ' 부터 ' & '" & Format(DTOS(dtpT_Date.Value), "0000-00-00") & "' & ' 까지' "
                        '--- Parameter Fields ---
                        '사업장코드
                        .StoredProcParam(0) = PB_regUserinfoU.UserBranchCode
                        '적용(기준)일자시작
                        .StoredProcParam(1) = DTOS(dtpF_Date.Value)
                        '적용(기준)일자종료
                        .StoredProcParam(2) = DTOS(dtpT_Date.Value)
                        '매입처코드
                        If (Len(txtSup.Text) = 0) Then
                          .StoredProcParam(3) = " "
                        Else
                          .StoredProcParam(3) = Trim(txtSup.Text)
                        End If
                        '분류코드
                        .StoredProcParam(4) = Mid(cboKind.Text, 1, 2)
                        '세부코드
                        .StoredProcParam(5) = IIf(Len(Trim(txtMt.Text)) = 0, " ", Trim(txtMt.Text))
                        '--- 속성 ---
                        .WindowShowGroupTree = True
                        .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "업체별매입현황(일자별)"
                     End If
                Case 9 '업체별 매입 현황
                     strExeFile = App.Path & ".\업체별매입현황.rpt"
                     varRetVal = Dir(strExeFile)
                     If Len(varRetVal) = 0 Then
                        intRetCHK = 0
                     Else
                        .ReportFileName = strExeFile
                        On Error GoTo ERROR_CRYSTAL_REPORTS
                        '--- Formula Fields ---
                        .Formulas(0) = "ForPgDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'" '프로그램실행일자
                        .Formulas(1) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                    '사업장명
                        .Formulas(2) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                               '출력일시
                        '기준일자
                        .Formulas(3) = "ForAppDate = '기준일자 : ' & '" & Format(DTOS(dtpF_Date.Value), "0000-00-00") & "' & ' 부터 ' & '" & Format(DTOS(dtpT_Date.Value), "0000-00-00") & "' & ' 까지' "
                        '--- Parameter Fields ---
                        '사업장코드
                        .StoredProcParam(0) = PB_regUserinfoU.UserBranchCode
                        '적용(기준)일자시작
                        .StoredProcParam(1) = DTOS(dtpF_Date.Value)
                        '적용(기준)일자종료
                        .StoredProcParam(2) = DTOS(dtpT_Date.Value)
                        '매입처명
                        If (Len(txtSup.Text) = 0) Then
                          .StoredProcParam(3) = " "
                        Else
                          .StoredProcParam(3) = Trim(txtSup.Text)
                        End If
                        '분류코드
                        .StoredProcParam(4) = Mid(cboKind.Text, 1, 2)
                        '세부코드
                        .StoredProcParam(5) = IIf(Len(Trim(txtMt.Text)) = 0, " ", Trim(txtMt.Text))
                        '--- 속성 ---
                        .WindowShowGroupTree = False
                        .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "업체별매입현황"
                     End If
                Case 10 '매입세금계산서 장부 현황
                     strExeFile = App.Path & ".\매입세금계산서장부현황.rpt"
                     varRetVal = Dir(strExeFile)
                     If Len(varRetVal) = 0 Then
                        intRetCHK = 0
                     Else
                        .ReportFileName = strExeFile
                        On Error GoTo ERROR_CRYSTAL_REPORTS
                        '--- Formula Fields ---
                        '.Formulas(0) = "ForPgDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'" '프로그램실행일자
                        .Formulas(0) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                    '사업장명
                        .Formulas(1) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                               '출력일시
                        '기준일자
                        .Formulas(2) = "ForAppDate = '기준일자 : ' & '" & Format(DTOS(dtpF_Date.Value), "0000-00-00") & "' & ' 부터 ' & '" & Format(DTOS(dtpT_Date.Value), "0000-00-00") & "' & ' 까지' "
                        '--- Parameter Fields ---
                        '사업장코드
                        .StoredProcParam(0) = PB_regUserinfoU.UserBranchCode
                        '적용(기준)일자시작
                        .StoredProcParam(1) = DTOS(dtpF_Date.Value)
                        '적용(기준)일자종료
                        .StoredProcParam(2) = DTOS(dtpT_Date.Value)
                        '매출처코드
                        If (Len(txtBuy.Text) = 0) Then
                          .StoredProcParam(3) = " "
                        Else
                          .StoredProcParam(3) = Trim(txtSup.Text)
                        End If
                        '--- 속성 ---
                        .WindowShowGroupTree = False
                        .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "매입세금계산서장부현황"
                     End If
                Case 11 '업체별 매입 현황(3개월간)
                     strExeFile = App.Path & ".\업체별매입현황(3개월간).rpt"
                     varRetVal = Dir(strExeFile)
                     If Len(varRetVal) = 0 Then
                        intRetCHK = 0
                     Else
                        .ReportFileName = strExeFile
                        On Error GoTo ERROR_CRYSTAL_REPORTS
                        '--- Formula Fields ---
                        .Formulas(0) = "ForPgDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'" '프로그램실행일자
                        .Formulas(1) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                    '사업장명
                        .Formulas(2) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                               '출력일시
                        '기준일자
                        .Formulas(3) = "ForAppDate = '기준일자 : ' & '" & Format(DTOS(dtpT_Date.Value), "0000-00-00") & "' "
                        '--- Parameter Fields ---
                        '사업장코드
                        .StoredProcParam(0) = PB_regUserinfoU.UserBranchCode
                        '적용(기준)일자종료
                        .StoredProcParam(1) = DTOS(dtpT_Date.Value)
                        '삼개월전년월, 첫번째년월, 두번째년월, 세번째년월
                        .StoredProcParam(2) = " ": .StoredProcParam(3) = " ": .StoredProcParam(4) = " ": .StoredProcParam(5) = " "
                        '--- 속성 ---
                        .WindowShowGroupTree = False
                        .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "업체별매입현황(3개월간)"
                     End If
                Case 13 '업체별 미수금 현황(집계표)
                     strExeFile = App.Path & ".\업체별미수금현황(집계표).rpt"
                     varRetVal = Dir(strExeFile)
                     If Len(varRetVal) = 0 Then
                        intRetCHK = 0
                     Else
                        .ReportFileName = strExeFile
                        On Error GoTo ERROR_CRYSTAL_REPORTS
                        '--- Formula Fields ---
                        .Formulas(0) = "ForPgDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'" '프로그램실행일자
                        .Formulas(1) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                    '사업장명
                        .Formulas(2) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                               '출력일시
                        '기준일자
                        .Formulas(3) = "ForAppDate = '기준일자 : ' & '" & Format(DTOS(dtpT_Date.Value), "0000-00-00") & "' "
                        '--- Parameter Fields ---
                        '사업장코드
                        .StoredProcParam(0) = PB_regUserinfoU.UserBranchCode
                        '적용(기준)일자종료
                        .StoredProcParam(1) = DTOS(dtpT_Date.Value)
                        '매출처코드
                        If (Len(txtBuy.Text) = 0) Then
                          .StoredProcParam(2) = " "
                        Else
                          .StoredProcParam(2) = Trim(txtBuy.Text)
                        End If
                        '미수금발생구분 1.전표, 2.계산서
                        .StoredProcParam(3) = PB_regUserinfoU.UserMSGbn
                        '--- 속성 ---
                        .WindowShowGroupTree = False
                        .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "업체별미수금현황(집계표)"
                     End If
                Case 14 '업체별 수금 현황(일자별)
                     strExeFile = App.Path & ".\업체별수금현황(일자별).rpt"
                     varRetVal = Dir(strExeFile)
                     If Len(varRetVal) = 0 Then
                        intRetCHK = 0
                     Else
                        .ReportFileName = strExeFile
                        On Error GoTo ERROR_CRYSTAL_REPORTS
                        '--- Formula Fields ---
                        .Formulas(0) = "ForPgDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'" '프로그램실행일자
                        .Formulas(1) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                    '사업장명
                        .Formulas(2) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                               '출력일시
                        '기준일자
                        .Formulas(3) = "ForAppDate = '기준일자 : ' & '" & Format(DTOS(dtpF_Date.Value), "0000-00-00") & "' & ' 부터 ' & '" & Format(DTOS(dtpT_Date.Value), "0000-00-00") & "' & ' 까지' "
                        '--- Parameter Fields ---
                        '사업장코드
                        .StoredProcParam(0) = PB_regUserinfoU.UserBranchCode
                        '적용(기준)일자시작
                        .StoredProcParam(1) = DTOS(dtpF_Date.Value)
                        '적용(기준)일자종료
                        .StoredProcParam(2) = DTOS(dtpT_Date.Value)
                        '매출처코드
                        If (Len(txtBuy.Text) = 0) Then
                          .StoredProcParam(3) = " "
                        Else
                          .StoredProcParam(3) = Trim(txtBuy.Text)
                        End If
                        '--- 속성 ---
                        .WindowShowGroupTree = True
                        .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "업체별수금현황(일자별)"
                     End If
                Case 16 '품목별 매출 현황
                     strExeFile = App.Path & ".\품목별매출현황.rpt"
                     varRetVal = Dir(strExeFile)
                     If Len(varRetVal) = 0 Then
                        intRetCHK = 0
                     Else
                        .ReportFileName = strExeFile
                        On Error GoTo ERROR_CRYSTAL_REPORTS
                        '--- Formula Fields ---
                        .Formulas(0) = "ForPgDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'" '프로그램실행일자
                        .Formulas(1) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                    '사업장명
                        .Formulas(2) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                               '출력일시
                        '기준일자
                        .Formulas(3) = "ForAppDate = '기준일자 : ' & '" & Format(DTOS(dtpF_Date.Value), "0000-00-00") & "' & ' 부터 ' & '" & Format(DTOS(dtpT_Date.Value), "0000-00-00") & "' & ' 까지' "
                        '--- Parameter Fields ---
                        '사업장코드
                        .StoredProcParam(0) = PB_regUserinfoU.UserBranchCode
                        '적용(기준)일자시작
                        .StoredProcParam(1) = DTOS(dtpF_Date.Value)
                        '적용(기준)일자종료
                        .StoredProcParam(2) = DTOS(dtpT_Date.Value)
                        '분류코드
                        .StoredProcParam(3) = Mid(cboKind.Text, 1, 2)
                        '세부코드
                        .StoredProcParam(4) = IIf(Len(Trim(txtMt.Text)) = 0, " ", Trim(txtMt.Text))
                        '--- 속성 ---
                        .WindowShowGroupTree = True
                        .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "품목별매출현황"
                     End If
                Case 17 '업체별 매출 현황(품목별)
                     strExeFile = App.Path & ".\업체별매출현황(품목별).rpt"
                     varRetVal = Dir(strExeFile)
                     If Len(varRetVal) = 0 Then
                        intRetCHK = 0
                     Else
                        .ReportFileName = strExeFile
                        On Error GoTo ERROR_CRYSTAL_REPORTS
                        '--- Formula Fields ---
                        .Formulas(0) = "ForPgDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'" '프로그램실행일자
                        .Formulas(1) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                    '사업장명
                        .Formulas(2) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                               '출력일시
                        '기준일자
                        .Formulas(3) = "ForAppDate = '기준일자 : ' & '" & Format(DTOS(dtpF_Date.Value), "0000-00-00") & "' & ' 부터 ' & '" & Format(DTOS(dtpT_Date.Value), "0000-00-00") & "' & ' 까지' "
                        '--- Parameter Fields ---
                        '사업장코드
                        .StoredProcParam(0) = PB_regUserinfoU.UserBranchCode
                        '적용(기준)일자시작
                        .StoredProcParam(1) = DTOS(dtpF_Date.Value)
                        '적용(기준)일자종료
                        .StoredProcParam(2) = DTOS(dtpT_Date.Value)
                        '매출처코드
                        If (Len(txtBuy.Text) = 0) Then
                          .StoredProcParam(3) = " "
                        Else
                          .StoredProcParam(3) = Trim(txtBuy.Text)
                        End If
                        '분류코드
                        .StoredProcParam(4) = Mid(cboKind.Text, 1, 2)
                        '세부코드
                        .StoredProcParam(5) = IIf(Len(Trim(txtMt.Text)) = 0, " ", Trim(txtMt.Text))
                        '--- 속성 ---
                        .WindowShowGroupTree = True
                        .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "업체별매출현황(품목별)"
                     End If
                Case 18 '업체별 매출 현황(일자별)
                     strExeFile = App.Path & ".\업체별매출현황(일자별).rpt"
                     varRetVal = Dir(strExeFile)
                     If Len(varRetVal) = 0 Then
                        intRetCHK = 0
                     Else
                        .ReportFileName = strExeFile
                        On Error GoTo ERROR_CRYSTAL_REPORTS
                        '--- Formula Fields ---
                        .Formulas(0) = "ForPgDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'" '프로그램실행일자
                        .Formulas(1) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                    '사업장명
                        .Formulas(2) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                               '출력일시
                        '기준일자
                        .Formulas(3) = "ForAppDate = '기준일자 : ' & '" & Format(DTOS(dtpF_Date.Value), "0000-00-00") & "' & ' 부터 ' & '" & Format(DTOS(dtpT_Date.Value), "0000-00-00") & "' & ' 까지' "
                        '--- Parameter Fields ---
                        '사업장코드
                        .StoredProcParam(0) = PB_regUserinfoU.UserBranchCode
                        '적용(기준)일자시작
                        .StoredProcParam(1) = DTOS(dtpF_Date.Value)
                        '적용(기준)일자종료
                        .StoredProcParam(2) = DTOS(dtpT_Date.Value)
                        '매출처코드
                        If (Len(txtBuy.Text) = 0) Then
                          .StoredProcParam(3) = " "
                        Else
                          .StoredProcParam(3) = Trim(txtBuy.Text)
                        End If
                        '분류코드
                        .StoredProcParam(4) = Mid(cboKind.Text, 1, 2)
                        '세부코드
                        .StoredProcParam(5) = IIf(Len(Trim(txtMt.Text)) = 0, " ", Mid(Trim(txtMt.Text), 3))
                        '--- 속성 ---
                        .WindowShowGroupTree = True
                        .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "업체별매출현황(일자별)"
                     End If
                Case 19 '업체별 매출 현황
                     strExeFile = App.Path & ".\업체별매출현황.rpt"
                     varRetVal = Dir(strExeFile)
                     If Len(varRetVal) = 0 Then
                        intRetCHK = 0
                     Else
                        .ReportFileName = strExeFile
                        On Error GoTo ERROR_CRYSTAL_REPORTS
                        '--- Formula Fields ---
                        .Formulas(0) = "ForPgDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'" '프로그램실행일자
                        .Formulas(1) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                    '사업장명
                        .Formulas(2) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                               '출력일시
                        '기준일자
                        .Formulas(3) = "ForAppDate = '기준일자 : ' & '" & Format(DTOS(dtpF_Date.Value), "0000-00-00") & "' & ' 부터 ' & '" & Format(DTOS(dtpT_Date.Value), "0000-00-00") & "' & ' 까지' "
                        '--- Parameter Fields ---
                        '사업장코드
                        .StoredProcParam(0) = PB_regUserinfoU.UserBranchCode
                        '적용(기준)일자시작
                        .StoredProcParam(1) = DTOS(dtpF_Date.Value)
                        '적용(기준)일자종료
                        .StoredProcParam(2) = DTOS(dtpT_Date.Value)
                        '매출처명
                        If (Len(txtBuy.Text) = 0) Then
                          .StoredProcParam(3) = " "
                        Else
                          .StoredProcParam(3) = Trim(txtBuy.Text)
                        End If
                        '분류코드
                        .StoredProcParam(4) = Mid(cboKind.Text, 1, 2)
                        '세부코드
                        .StoredProcParam(5) = IIf(Len(Trim(txtMt.Text)) = 0, " ", Trim(txtMt.Text))
                        '--- 속성 ---
                        .WindowShowGroupTree = False
                        .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "업체별매출현황"
                     End If
                Case 20 '매출세금계산서 장부 현황
                     strExeFile = App.Path & ".\매출세금계산서장부현황.rpt"
                     varRetVal = Dir(strExeFile)
                     If Len(varRetVal) = 0 Then
                        intRetCHK = 0
                     Else
                        .ReportFileName = strExeFile
                        On Error GoTo ERROR_CRYSTAL_REPORTS
                        '--- Formula Fields ---
                        '.Formulas(0) = "ForPgDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'" '프로그램실행일자
                        .Formulas(0) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                    '사업장명
                        .Formulas(1) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                               '출력일시
                        '기준일자
                        .Formulas(2) = "ForAppDate = '기준일자 : ' & '" & Format(DTOS(dtpF_Date.Value), "0000-00-00") & "' & ' 부터 ' & '" & Format(DTOS(dtpT_Date.Value), "0000-00-00") & "' & ' 까지' "
                        '--- Parameter Fields ---
                        '사업장코드
                        .StoredProcParam(0) = PB_regUserinfoU.UserBranchCode
                        '적용(기준)일자시작
                        .StoredProcParam(1) = DTOS(dtpF_Date.Value)
                        '적용(기준)일자종료
                        .StoredProcParam(2) = DTOS(dtpT_Date.Value)
                        '매출처코드
                        If (Len(txtBuy.Text) = 0) Then
                          .StoredProcParam(3) = " "
                        Else
                          .StoredProcParam(3) = Trim(txtBuy.Text)
                        End If
                        '--- 속성 ---
                        .WindowShowGroupTree = False
                        .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "매출세금계산서장부현황"
                     End If
                Case 21 '매출세금계산서 현황
                     strExeFile = App.Path & ".\매출세금계산서현황.rpt"
                     varRetVal = Dir(strExeFile)
                     If Len(varRetVal) = 0 Then
                        intRetCHK = 0
                     Else
                        .ReportFileName = strExeFile
                        On Error GoTo ERROR_CRYSTAL_REPORTS
                        '--- Formula Fields ---
                        '.Formulas(0) = "ForPgDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'" '프로그램실행일자
                        .Formulas(0) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                    '사업장명
                        .Formulas(1) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                               '출력일시
                        '기준일자
                        .Formulas(2) = "ForAppDate = '기준일자 : ' & '" & Format(DTOS(dtpF_Date.Value), "0000-00-00") & "' & ' 부터 ' & '" & Format(DTOS(dtpT_Date.Value), "0000-00-00") & "' & ' 까지' "
                        '--- Parameter Fields ---
                        '사업장코드
                        .StoredProcParam(0) = PB_regUserinfoU.UserBranchCode
                        '적용(기준)일자시작
                        .StoredProcParam(1) = DTOS(dtpF_Date.Value)
                        '적용(기준)일자종료
                        .StoredProcParam(2) = DTOS(dtpT_Date.Value)
                        '매출처코드
                        If (Len(txtBuy.Text) = 0) Then
                          .StoredProcParam(3) = " "
                        Else
                          .StoredProcParam(3) = Trim(txtBuy.Text)
                        End If
                        '--- 속성 ---
                        .WindowShowGroupTree = False
                        .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "매출세금계산서현황"
                     End If
                Case 22 '업체별 매출 현황(3개월간)
                     strExeFile = App.Path & ".\업체별매출현황(3개월간).rpt"
                     varRetVal = Dir(strExeFile)
                     If Len(varRetVal) = 0 Then
                        intRetCHK = 0
                     Else
                        .ReportFileName = strExeFile
                        On Error GoTo ERROR_CRYSTAL_REPORTS
                        '--- Formula Fields ---
                        .Formulas(0) = "ForPgDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'" '프로그램실행일자
                        .Formulas(1) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                    '사업장명
                        .Formulas(2) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                               '출력일시
                        '기준일자
                        .Formulas(3) = "ForAppDate = '기준일자 : ' & '" & Format(DTOS(dtpT_Date.Value), "0000-00-00") & "' "
                        '--- Parameter Fields ---
                        '사업장코드
                        .StoredProcParam(0) = PB_regUserinfoU.UserBranchCode
                        '적용(기준)일자종료
                        .StoredProcParam(1) = DTOS(dtpT_Date.Value)
                        '삼개월전년월, 첫번째년월, 두번째년월, 세번째년월
                        .StoredProcParam(2) = " ": .StoredProcParam(3) = " ": .StoredProcParam(4) = " ": .StoredProcParam(5) = " "
                        '--- 속성 ---
                        .WindowShowGroupTree = False
                        .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "업체별매출현황(3개월간)"
                     End If
                Case 24 '출납관리(일자별)
                     strExeFile = App.Path & ".\출납관리(일자별).rpt"
                     varRetVal = Dir(strExeFile)
                     If Len(varRetVal) = 0 Then
                        intRetCHK = 0
                     Else
                        .ReportFileName = strExeFile
                        On Error GoTo ERROR_CRYSTAL_REPORTS
                        '--- Formula Fields ---
                        '.Formulas(0) = "ForPgDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'" '프로그램실행일자
                        .Formulas(0) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                    '사업장명
                        .Formulas(1) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                               '출력일시
                        '기준일자
                        .Formulas(2) = "ForAppDate = '기준일자 : ' & '" & Format(DTOS(dtpF_Date.Value), "0000-00-00") & "' & ' 부터 ' & '" & Format(DTOS(dtpT_Date.Value), "0000-00-00") & "' & ' 까지' "
                        '--- Parameter Fields ---
                        '사업장코드
                        .StoredProcParam(0) = PB_regUserinfoU.UserBranchCode
                        '적용(기준)일자시작
                        .StoredProcParam(1) = DTOS(dtpF_Date.Value)
                        '적용(기준)일자종료
                        .StoredProcParam(2) = DTOS(dtpT_Date.Value)
                        '--- 속성 ---
                        .WindowShowGroupTree = True
                        .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "출납관리(일자별)"
                     End If
                Case 25 '출납관리(계정별)
                     strExeFile = App.Path & ".\출납관리(계정과목별).rpt"
                     varRetVal = Dir(strExeFile)
                     If Len(varRetVal) = 0 Then
                        intRetCHK = 0
                     Else
                        .ReportFileName = strExeFile
                        On Error GoTo ERROR_CRYSTAL_REPORTS
                        '--- Formula Fields ---
                        '.Formulas(0) = "ForPgDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'" '프로그램실행일자
                        .Formulas(0) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                    '사업장명
                        .Formulas(1) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                               '출력일시
                        '기준일자
                        .Formulas(2) = "ForAppDate = '기준일자 : ' & '" & Format(DTOS(dtpF_Date.Value), "0000-00-00") & "' & ' 부터 ' & '" & Format(DTOS(dtpT_Date.Value), "0000-00-00") & "' & ' 까지' "
                        '--- Parameter Fields ---
                        '사업장코드
                        .StoredProcParam(0) = PB_regUserinfoU.UserBranchCode
                        '적용(기준)일자시작
                        .StoredProcParam(1) = DTOS(dtpF_Date.Value)
                        '적용(기준)일자종료
                        .StoredProcParam(2) = DTOS(dtpT_Date.Value)
                        '--- 속성 ---
                        .WindowShowGroupTree = True
                        .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "출납관리(계정과목별)"
                     End If
                Case 26 '출납관리(계정별집계)
                     strExeFile = App.Path & ".\출납관리(계정별집계).rpt"
                     varRetVal = Dir(strExeFile)
                     If Len(varRetVal) = 0 Then
                        intRetCHK = 0
                     Else
                        .ReportFileName = strExeFile
                        On Error GoTo ERROR_CRYSTAL_REPORTS
                        '--- Formula Fields ---
                        '.Formulas(0) = "ForPgDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'" '프로그램실행일자
                        .Formulas(0) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                    '사업장명
                        .Formulas(1) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                               '출력일시
                        '기준일자
                        .Formulas(2) = "ForAppDate = '기준일자 : ' & '" & Format(DTOS(dtpF_Date.Value), "0000-00-00") & "' & ' 부터 ' & '" & Format(DTOS(dtpT_Date.Value), "0000-00-00") & "' & ' 까지' "
                        '--- Parameter Fields ---
                        '사업장코드
                        .StoredProcParam(0) = PB_regUserinfoU.UserBranchCode
                        '적용(기준)일자시작
                        .StoredProcParam(1) = DTOS(dtpF_Date.Value)
                        '적용(기준)일자종료
                        .StoredProcParam(2) = DTOS(dtpT_Date.Value)
                        '--- 속성 ---
                        .WindowShowGroupTree = False
                        .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "출납관리(계정집계)"
                     End If
                Case 27 '합계시산표
                     strExeFile = App.Path & ".\합계시산표.rpt"
                     varRetVal = Dir(strExeFile)
                     If Len(varRetVal) = 0 Then
                        intRetCHK = 0
                     Else
                        .ReportFileName = strExeFile
                        On Error GoTo ERROR_CRYSTAL_REPORTS
                        '--- Formula Fields ---
                        .Formulas(0) = "ForPgDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'" '프로그램실행일자
                        .Formulas(1) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                    '사업장명
                        .Formulas(2) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                               '출력일시
                        '기준일자
                        .Formulas(3) = "ForAppDate = '기준일자 : ' & '" & Format(DTOS(dtpT_Date.Value), "0000-00-00") & "' "
                        '--- Parameter Fields ---
                        '사업장코드
                        .StoredProcParam(0) = PB_regUserinfoU.UserBranchCode
                        '적용(기준)일자종료
                        .StoredProcParam(1) = DTOS(dtpT_Date.Value)
                        '--- 속성 ---
                        .WindowShowGroupTree = False
                        .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "합계시산표"
                     End If
                Case Else
                     intRetCHK = 0
                     
         End Select
         If intRetCHK = 99 Then
            .Connect = PB_adoCnnSQL.ConnectionString
            .Destination = crptToWindow
            .DiscardSavedData = True
            .ProgressDialog = True
            .ReportSource = crptReport
            .WindowAllowDrillDown = False
            .WindowShowProgressCtls = True
            .WindowShowCloseBtn = True
            .WindowShowExportBtn = False
            '.WindowShowGroupTree = True
            .WindowShowPrintSetupBtn = True
            .WindowTop = 0: .WindowTop = 0: .WindowHeight = 11100: .WindowWidth = 15405
            .WindowState = crptMaximized
            '.WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "WindowTitle"
            .Action = 1
            .Reset
         End If
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "(서버와의 연결 실패)"
    Unload Me
    Exit Sub
ERROR_CRYSTAL_REPORTS:
    MsgBox Err.Number & Space(1) & Err.Description
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

