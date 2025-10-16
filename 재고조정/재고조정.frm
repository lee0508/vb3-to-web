VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm재고조정 
   BorderStyle     =   0  '없음
   Caption         =   "재고조정"
   ClientHeight    =   10095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15405
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10100
   ScaleMode       =   0  '사용자
   ScaleWidth      =   15405
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  '평면
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   60
      TabIndex        =   8
      Top             =   0
      Width           =   15195
      Begin VB.CommandButton cmdPrint 
         Height          =   375
         Left            =   7980
         Picture         =   "재고조정.frx":0000
         Style           =   1  '그래픽
         TabIndex        =   17
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   9120
         Picture         =   "재고조정.frx":0963
         Style           =   1  '그래픽
         TabIndex        =   13
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   13635
         Picture         =   "재고조정.frx":1308
         Style           =   1  '그래픽
         TabIndex        =   12
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   390
         Left            =   12510
         Picture         =   "재고조정.frx":1C56
         Style           =   1  '그래픽
         TabIndex        =   11
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   11385
         Picture         =   "재고조정.frx":25DA
         Style           =   1  '그래픽
         TabIndex        =   10
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Height          =   390
         Left            =   10260
         Picture         =   "재고조정.frx":2E61
         Style           =   1  '그래픽
         TabIndex        =   5
         Top             =   195
         Width           =   1095
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   4320
         Top             =   195
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin MSComCtl2.DTPicker dtpFDate 
         Height          =   270
         Left            =   4920
         TabIndex        =   19
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   57540609
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin MSComCtl2.DTPicker dtpTDate 
         Height          =   270
         Left            =   6600
         TabIndex        =   21
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   57540609
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "-"
         Height          =   240
         Index           =   11
         Left            =   6240
         TabIndex        =   20
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00008000&
         BorderStyle     =   1  '단일 고정
         Caption         =   "재 고 조 정"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   18
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   165
         TabIndex        =   9
         Top             =   180
         Width           =   4650
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid vsfg1 
      Height          =   8310
      Left            =   60
      TabIndex        =   6
      Top             =   1665
      Width           =   15195
      _cx             =   26802
      _cy             =   14658
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   60
      TabIndex        =   7
      Top             =   630
      Width           =   15195
      Begin VB.TextBox txtFindNM 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '영문
         Left            =   7800
         MaxLength       =   50
         TabIndex        =   4
         Top             =   600
         Width           =   3525
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Left            =   7800
         MaxLength       =   18
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox cboTaxGbn 
         Height          =   300
         Left            =   2400
         Style           =   2  '드롭다운 목록
         TabIndex        =   1
         Top             =   600
         Width           =   1110
      End
      Begin VB.ComboBox cboMt 
         Height          =   300
         Left            =   2400
         Style           =   2  '드롭다운 목록
         TabIndex        =   0
         Top             =   195
         Width           =   3735
      End
      Begin VB.ComboBox cboState 
         Height          =   300
         Left            =   5050
         Style           =   2  '드롭다운 목록
         TabIndex        =   2
         Top             =   600
         Width           =   1080
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "품명"
         Height          =   240
         Index           =   13
         Left            =   6960
         TabIndex        =   24
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "품목코드"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   6600
         TabIndex        =   23
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "[Home]"
         Height          =   240
         Index           =   1
         Left            =   10245
         TabIndex        =   22
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "과세구분"
         Height          =   240
         Index           =   0
         Left            =   1245
         TabIndex        =   18
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "검색조건)"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   35
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "분류"
         Height          =   240
         Index           =   34
         Left            =   1245
         TabIndex        =   15
         Top             =   250
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "사용구분"
         Height          =   240
         Index           =   25
         Left            =   4000
         TabIndex        =   14
         Top             =   660
         Width           =   855
      End
   End
End
Attribute VB_Name = "frm재고조정"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' 프로그램 제 목 : 재고조정
' 사용된 Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' 참조된 Table   : 자재원장, 자재, 자재분류
'                  자재원장마감, 자재입출고내역
' 업  무  설  명 :
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived       As Boolean
Private P_adoRec           As New ADODB.Recordset
Private P_intButton        As Integer
Private P_strFindString1   As String
Private P_strFindString2   As String
Private P_strSortM(1000)   As String
Private P_strSortS(1000)   As String
Private Const PC_intRowCnt As Integer = 22  '그리드 한 페이지 당 행수(FixedRows 포함)

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
Dim SQL     As String
    If P_blnActived = False Then
       Screen.MousePointer = vbHourglass
       On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
       Subvsfg1_INIT
       Select Case Val(PB_regUserinfoU.UserAuthority)
              Case Is <= 10 '조회
                   cmdPrint.Enabled = False: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 20 '인쇄, 조회
                   cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 40 '추가, 저장, 인쇄, 조회
                   cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = True: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 50 '삭제, 추가, 저장, 조회, 인쇄
                   cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = True: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 99 '삭제, 추가, 저장, 조회, 인쇄
                   cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = True: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Else
                   cmdClear.Enabled = False: cmdFind.Enabled = False: cmdSave.Enabled = False
                   cmdDelete.Enabled = False: cmdExit.Enabled = True
       End Select
       dtpFDate.Value = Format(Mid(PB_regUserinfoU.UserClientDate, 1, 6) & "01", "0000-00-00")
       dtpTDate.Value = DateAdd("d", -1, DateAdd("m", 1, dtpFDate.Value))
       SubOther_FILL
       txtCode.SetFocus
       P_blnActived = True
       Screen.MousePointer = vbDefault
    End If
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "(서버와의 연결 실패)"
    Unload Me
    Exit Sub
End Sub

'+---------------+
'/// 인쇄조건 ///
'+---------------+
Private Sub dtpFDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub dtpTDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdPrint.SetFocus
End Sub

'+----------------+
'/// cboMt() ///
'+----------------+
Private Sub cboMt_GotFocus()
Dim strSQL As String
Dim nRet   As Long
    '자동 펼침
    'SendKeys "{F4}"
    '자동 펼침
    'nRet = SendMessage(cboFdMtGp(Index).hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
    'ListIndex의 값을 바꾸어도 Click 이벤트가 발생하지 않도록 함.
    'SendMessage cboFdMtGp(index).hwnd, &H14E&, 0, ByVal 0&
End Sub
Private Sub cboMt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    End If
End Sub
'+----------------+
'/// cboTaxGbn ///
'+----------------+
Private Sub cboTaxGbn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    End If
End Sub
'+---------------+
'/// cboState ///
'+---------------+
Private Sub cboState_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    End If
End Sub

'+-------------+
'/// txtCode ///
'+-------------+
Private Sub txtCode_GotFocus()
    With txtCode
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL     As String
Dim strExeFile As String
Dim varRetVal  As Variant
    If (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn) Then '자재검색
       PB_strCallFormName = "frm재고조정"
       PB_strMaterialsCode = Trim(txtCode.Text)
       PB_strMaterialsName = txtFindNM.Text
       frm자재검색.Show vbModal
       If (Len(PB_strMaterialsCode) + Len(PB_strMaterialsName)) = 0 Then '검색에서 취소(ESC)
       Else
          txtCode.Text = PB_strMaterialsCode
          txtFindNM.Text = PB_strMaterialsName
       End If
       If PB_strMaterialsCode <> "" Then
          'SendKeys "{tab}"
       End If
       PB_strCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
       cmdFind_Click
    End If
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "자재원장 읽기 실패"
    Unload Me
    Exit Sub
End Sub

Private Sub txtCode_LostFocus()
Dim strSQL As String
Dim lngR   As Long
    With txtCode
         .Text = Trim(.Text)
         If Len(Trim(.Text)) = 0 Then
            txtCode.Text = ""
            txtFindNM.Text = ""
         End If
    End With
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "재고조정 실패"
    Unload Me
    Exit Sub
End Sub

'+----------------------------+
'/// txtFindNM(자재명검색) ///
'+----------------------------+
Private Sub txtFindNM_GotFocus()
    With txtFindNM
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtFindNM_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL     As String
Dim strExeFile As String
Dim varRetVal  As Variant
    If KeyCode = vbKeyReturn Then
       cmdFind_Click
    End If
End Sub

'+-----------+
'/// Grid ///
'+-----------+
Private Sub vsfg1_BeforeSort(ByVal Col As Long, Order As Integer)
    With vsfg1
         'P_strFindString2 = Trim(.Cell(flexcpData, .Row, 0)) 'Not Used
    End With
End Sub
Private Sub vsfg1_AfterSort(ByVal Col As Long, Order As Integer)
    With vsfg1
         'If .FindRow(P_strFindString2, , 0) > 0 Then
         '   .Row = .FindRow(P_strFindString2, , 0) 'Not Used
         'End If
         'If PC_intRowCnt < .Rows Then
         '   .TopRow = .Row
         'End If
    End With
End Sub
Private Sub vsfg1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    With vsfg1
         .ToolTipText = ""
         If .MouseRow < .FixedRows Or .MouseCol < 0 Then
            Exit Sub
         End If
         .ToolTipText = .TextMatrix(.MouseRow, .MouseCol)
    End With
End Sub
Private Sub vsfg1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    P_intButton = Button
End Sub
Private Sub vsfg1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    With vsfg1
         If .MouseRow >= .FixedRows Then
            If Len(.TextMatrix(.Row, 0)) <> 0 Then '15.재고조정량, 18.적요
               If (.MouseCol = 15 Or .MouseCol = 18) Then
                  If Button = vbLeftButton Then
                     .Select .MouseRow, .MouseCol
                     .EditCell
                  End If
               End If
            End If
         End If
    End With
End Sub
Sub vsfg1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfg1
         If Row >= .FixedRows Then
            If Len(.TextMatrix(Row, 0)) <> 0 And (Col = 15 Or Col = 18) Then
               If (Col = 15) Then         '재고조정량
                  If .TextMatrix(Row, Col) <> .EditText Then
                     'If IsNumeric(.EditText) = False Or Vals(.EditText) < 0 Or _
                     '   fix(Vals(.EditText)) < Vals(.EditText) Then
                     If IsNumeric(.EditText) = False Or Fix(Vals(.EditText)) < Vals(.EditText) Then
                        Beep
                        Cancel = True
                     Else
                        '.TextMatrix(Row, 7) = Vals(.EditText) * (.ValueMatrix(Row, 5) + .ValueMatrix(Row, 6))
                     End If
                  End If
               ElseIf _
                  (Col = 18) Then '적요 길이 검사
                  If .TextMatrix(Row, Col) <> .EditText Then
                     If Not (LenH(Trim(.EditText)) <= 50) Then
                        Beep
                        Cancel = True
                     End If
                  End If
               End If
            End If
         End If
    End With
End Sub
Private Sub vsfg1_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsfg1
         If .Row >= .FixedRows Then
         End If
    End With
End Sub
Private Sub vsfg1_Click()
Dim strData As String
    With vsfg1
         If (.MouseRow >= 0 And .MouseRow < .FixedRows) Then
            .Col = .MouseCol
            '.Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlack
            '.Cell(flexcpForeColor, 0, .MouseCol, 0, .MouseCol) = vbRed
            'strData = Trim(.Cell(flexcpData, .Row, 4))
            'Select Case .MouseCol
            '       Case 0, 2
            '            .ColSel = 2
            '            .ColSort(0) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .ColSort(1) = flexSortNone
            '            .ColSort(2) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .Sort = flexSortUseColSort
            '       Case 1
            '            .ColSel = 2
            '            .ColSort(0) = flexSortNone
            '            .ColSort(1) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .ColSort(2) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .Sort = flexSortUseColSort
            '       Case Else
            '            .Sort = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            'End Select
            'If .FindRow(strData, , 4) > 0 Then
            '   .Row = .FindRow(strData, , 4)
            'End If
            'If PC_intRowCnt < .Rows Then
            '   .TopRow = .Row
            'End If
         End If
    End With
End Sub
Private Sub vsfg1_EnterCell()
Dim lngR As Long
Dim lngC As Long
    With vsfg1
         If .Row >= .FixedRows Then
         End If
    End With
End Sub

'+-----------+
'/// 추가 ///
'+-----------+
Private Sub cmdClear_Click()
Dim strSQL As String
    '
End Sub
'+-----------+
'/// 조회 ///
'+-----------+
Private Sub cmdFind_Click()
    If Len(txtCode.Text) = 0 Then
       txtCode.SetFocus
       Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    cmdFind.Enabled = False
    P_strFindString1 = Trim(txtCode.Text)     '조회할 경우 검색할 자재코드
    'P_strFindString2 = Trim(txtFindNM.Text)  '조회할 경우 검색할 자재명 보관
    Subvsfg1_FILL
    cmdFind.Enabled = True
    vsfg1.SetFocus
    Screen.MousePointer = vbDefault
End Sub
'+-----------+
'/// 저장 ///
'+-----------+
Private Sub cmdSave_Click()
Dim strSQL         As String
Dim lngR           As Long
Dim lngRR          As Long
Dim lngRRR         As Long
Dim lngC           As Long
Dim lngLogCnt      As Long
Dim blnOK          As Boolean
Dim intRetVal      As Integer
Dim strServerTime  As String
Dim strTime        As String
Dim CurInputMny    As Currency '입고단가
Dim CurInputVat    As Currency '입고부가
Dim CurOutPutMny   As Currency '출고단가
Dim CurOutPutVat   As Currency '출고부가

    If (vsfg1.Rows) = 1 Then
       Exit Sub
    End If
    With vsfg1
         If .ValueMatrix(.Row, 15) = 0 Then
            Exit Sub
         End If
    End With
    intRetVal = MsgBox("개별로 재고조정된 현재 행을 저장하시겠습니까 ?", vbQuestion + vbYesNo + vbDefaultButton1, "자료 저장")
    If intRetVal = vbNo Then
       vsfg1.SetFocus
       Exit Sub
    End If
    '서버시간 구하기
    Screen.MousePointer = vbHourglass
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT RIGHT(CONVERT(VARCHAR(23),GETDATE(), 121),12) AS 서버시간 "
    On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    strServerTime = Mid(P_adoRec("서버시간"), 1, 2) + Mid(P_adoRec("서버시간"), 4, 2) + Mid(P_adoRec("서버시간"), 7, 2) _
                  + Mid(P_adoRec("서버시간"), 10)
    P_adoRec.Close
    strTime = strServerTime
    
    If vsfg1.ValueMatrix(vsfg1.Row, 15) > 0 Then '입고(+)
       strSQL = "SELECT TOP 1 " _
                  & "ISNULL(T1.입고단가1, 0) AS 입고단가, ISNULL(ROUND(T1.입고단가1 * (" & PB_curVatRate & "), 0, 1), 0) AS 입고부가, " _
                  & "ISNULL(T1.출고단가1, 0) AS 출고단가, ISNULL(ROUND(T1.출고단가1 * (" & PB_curVatRate & "), 0, 1), 0) AS 출고부가 " _
                & "FROM 자재원장 T1 " _
               & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                 & "AND T1.분류코드 = '" & Mid(vsfg1.TextMatrix(vsfg1.Row, 4), 1, 2) & "' " _
                 & "AND T1.세부코드 = '" & Mid(vsfg1.TextMatrix(vsfg1.Row, 4), 3) & "' "
    Else
       strSQL = "SELECT TOP 1 " _
                  & "ISNULL(T1.입고단가1, 0) AS 입고단가, ISNULL(ROUND(T1.입고단가1 * (" & PB_curVatRate & "), 0, 1), 0) AS 입고부가, " _
                  & "ISNULL(T1.출고단가1, 0) AS 출고단가, ISNULL(ROUND(T1.출고단가1 * (" & PB_curVatRate & "), 0, 1), 0) AS 출고부가 " _
                & "FROM 자재원장 T1 " _
               & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                 & "AND T1.분류코드 = '" & Mid(vsfg1.TextMatrix(vsfg1.Row, 4), 1, 2) & "' " _
                 & "AND T1.세부코드 = '" & Mid(vsfg1.TextMatrix(vsfg1.Row, 4), 3) & "' "
    End If
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount <> 0 Then
       CurInputMny = P_adoRec("입고단가"): CurInputVat = P_adoRec("입고부가")
       CurOutPutMny = P_adoRec("출고단가"): CurOutPutVat = P_adoRec("출고부가")
    End If
    P_adoRec.Close
    PB_adoCnnSQL.BeginTrans
    P_adoRec.CursorLocation = adUseClient
    With vsfg1
         If .ValueMatrix(.Row, 15) > 0 Then '입고(+)
            '거래번호 구하기
            P_adoRec.CursorLocation = adUseClient
            strSQL = "spLogCounter '자재입출내역', '" & PB_regUserinfoU.UserBranchCode + PB_regUserinfoU.UserClientDate + "5" & "', " _
                                 & "0, 0, '" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' "
            On Error GoTo ERROR_STORED_PROCEDURE
            P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
            lngLogCnt = P_adoRec(0)
            P_adoRec.Close
            '자재입출내역
            strSQL = "INSERT INTO 자재입출내역(사업장코드, 분류코드, " _
                                            & "세부코드, 입출고구분, " _
                                            & "입출고일자, 입출고시간, " _
                                            & "입고수량, 입고단가, " _
                                            & "입고부가, 출고수량, " _
                                            & "출고단가, 출고부가, " _
                                            & "매입처코드, 매출처코드, " _
                                            & "원래입출고일자, 직송구분, " _
                                            & "발견일자, 발견번호, 거래일자, 거래번호, " _
                                            & "계산서발행여부, 현금구분, 감가구분, 적요, 작성년도, 책번호, 일련번호, " _
                                            & "사용구분, 수정일자, " _
                                            & "사용자코드, 재고이동사업장코드) VALUES( " _
                      & "'" & PB_regUserinfoU.UserBranchCode & "', '" & Mid(.TextMatrix(.Row, 4), 1, 2) & "', " _
                      & "'" & Mid(.TextMatrix(.Row, 4), 3) & "', 5, " _
                      & "'" & PB_regUserinfoU.UserClientDate & "', '" & strServerTime & "', " _
                      & "" & .ValueMatrix(.Row, 15) & ", " & CurInputMny & ", " _
                      & "" & CurInputVat & ", 0, " _
                      & "0, 0, " _
                      & "'', '', " _
                      & "'" & PB_regUserinfoU.UserClientDate & "', 0, " _
                      & "'', 0, '" & PB_regUserinfoU.UserClientDate & "', " & lngLogCnt & ", " _
                      & "0, 0, 0, '" & .TextMatrix(.Row, 18) & "', '', 0, 0, " _
                      & "0, '" & PB_regUserinfoU.UserServerDate & "', " _
                      & "'" & PB_regUserinfoU.UserCode & "', '' ) "
            On Error GoTo ERROR_TABLE_INSERT
            PB_adoCnnSQL.Execute strSQL
         Else                            '출고(-)
            '거래번호 구하기
            P_adoRec.CursorLocation = adUseClient
            strSQL = "spLogCounter '자재입출내역', '" & PB_regUserinfoU.UserBranchCode + PB_regUserinfoU.UserClientDate + "6" & "', " _
                                 & "0, 0, '" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' "
            On Error GoTo ERROR_STORED_PROCEDURE
            P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
            lngLogCnt = P_adoRec(0)
            P_adoRec.Close
            strSQL = "INSERT INTO 자재입출내역(사업장코드, 분류코드, " _
                                            & "세부코드, 입출고구분, " _
                                            & "입출고일자, 입출고시간, " _
                                            & "입고수량, 입고단가, " _
                                            & "입고부가, 출고수량, " _
                                            & "출고단가, 출고부가, " _
                                            & "매입처코드, 매출처코드, " _
                                            & "원래입출고일자, 직송구분, " _
                                            & "발견일자, 발견번호, 거래일자, 거래번호, " _
                                            & "계산서발행여부, 현금구분, 감가구분, 적요, 작성년도, 책번호, 일련번호, " _
                                            & "사용구분, 수정일자, " _
                                            & "사용자코드, 재고이동사업장코드) VALUES( " _
                      & "'" & PB_regUserinfoU.UserBranchCode & "', '" & Mid(.TextMatrix(.Row, 4), 1, 2) & "', " _
                      & "'" & Mid(.TextMatrix(.Row, 4), 3) & "', 6, " _
                      & "'" & PB_regUserinfoU.UserClientDate & "', '" & strServerTime & "', " _
                      & "0, 0, " _
                      & "0, " & (.ValueMatrix(.Row, 15) * -1) & ", " _
                      & "" & CurOutPutMny & ", " & CurOutPutVat & ", " _
                      & "'', '', " _
                      & "'" & PB_regUserinfoU.UserClientDate & "', 0, " _
                      & "'', 0, '" & PB_regUserinfoU.UserClientDate & "', " & lngLogCnt & ", " _
                      & "0, 0, 0, '" & .TextMatrix(.Row, 18) & "', '', 0, 0, " _
                      & "0, '" & PB_regUserinfoU.UserServerDate & "', " _
                      & "'" & PB_regUserinfoU.UserCode & "', '' ) "
            On Error GoTo ERROR_TABLE_INSERT
            PB_adoCnnSQL.Execute strSQL
         End If
         .RemoveItem .Row
    End With
    PB_adoCnnSQL.CommitTrans
    Screen.MousePointer = vbDefault
    cmdSave.Enabled = True
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "(서버와의 연결 실패)"
    Unload Me
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "재고조정 추가 실패"
    Unload Me
    Exit Sub
ERROR_TABLE_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "재고조정 추가 실패"
    Unload Me
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "재고조정 저장 실패"
    Unload Me
    Exit Sub
ERROR_STORED_PROCEDURE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "로그 변경 실패"
    Unload Me
    Exit Sub
End Sub

'+-----------+
'/// 삭제 ///
'+-----------+
Private Sub cmdDelete_Click()
    '
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
    Set frm재고조정 = Nothing
End Sub

'+--------------------+
'/// Sub Procedure ///
'+--------------------+
'+----------------------------------+
'/// VsFlexGrid(vsfg1) 초기화 ///
'+----------------------------------+
Private Sub Subvsfg1_INIT()
Dim lngR    As Long
Dim lngC    As Long
    With vsfg1                 'Rows 1, Cols 19, RowHeightMax(Min) 300
         .AllowBigSelection = False
         .AllowSelection = False
         .AllowUserResizing = flexResizeColumns
         .BackColorBkg = &H8000000F
         .BackColorSel = &H8000&
         .Ellipsis = flexEllipsisEnd
         '.ExplorerBar = flexExSortShow
         .ExtendLastCol = True
         .FocusRect = flexFocusHeavy
         .ScrollBars = flexScrollBarHorizontal
         .ScrollTrack = True
         .SelectionMode = flexSelectionByRow
         .FixedRows = 1
         .FixedCols = 4
         .Rows = 1             'Subvsfg1_Fill수행시에 설정
         .Cols = 19
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1000   '분류코드       'H
         .ColWidth(1) = 1200   '분류명
         .ColWidth(2) = 1900   '품목코드(분류코드+세부코드)
         .ColWidth(3) = 2500   '품목명
         .ColWidth(4) = 1900   '분류코드 + 세부코드 'H
         .ColWidth(5) = 2500   '규격
         .ColWidth(6) = 1000   '단위
         .ColWidth(7) = 1000   '폐기율         'H
         .ColWidth(8) = 1000   '과세구분       'H
         .ColWidth(9) = 1000   '사용구분       'H
         .ColWidth(10) = 1200  '적정재고
         .ColWidth(11) = 1200  '이월재고       'H
         .ColWidth(12) = 1200  '입고수량       'H
         .ColWidth(13) = 1200  '출고수량       'H
         .ColWidth(14) = 1200  '현재재고
         .ColWidth(15) = 1200  '조정수량
         .ColWidth(16) = 1200  '최종입고일자   'H
         .ColWidth(17) = 1200  '최종출고일자   'H
         .ColWidth(18) = 5000  '적요
         .Cell(flexcpFontBold, 0, 0, .FixedRows - 1, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "분류코드"         'H
         .TextMatrix(0, 1) = "분류명"
         .TextMatrix(0, 2) = "품목코드"
         .TextMatrix(0, 3) = "품명"
         .TextMatrix(0, 4) = "(분류+세부)코드"  'H
         .TextMatrix(0, 5) = "규격"
         .TextMatrix(0, 6) = "단위"
         .TextMatrix(0, 7) = "폐기율"           'H
         .TextMatrix(0, 8) = "과세구분"         'H
         .TextMatrix(0, 9) = "사용구분"         'H
         .TextMatrix(0, 10) = "적정재고"
         .TextMatrix(0, 11) = "이월재고"        'H
         .TextMatrix(0, 12) = "매입수량"        'H
         .TextMatrix(0, 13) = "매출수량"        'H
         .TextMatrix(0, 14) = "현재재고"
         .TextMatrix(0, 15) = "조정량(+/-)"
         .TextMatrix(0, 16) = "최종입고일자"    'H
         .TextMatrix(0, 17) = "최종출고일자"    'H
         .TextMatrix(0, 18) = "적요"
         .ColFormat(7) = "#,#.00"
         For lngC = 10 To 15
             .ColFormat(lngC) = "#,#.00"
         Next lngC
         .ColHidden(0) = True: .ColHidden(4) = True
         .ColHidden(7) = True: .ColHidden(8) = True: .ColHidden(9) = True
         .ColHidden(11) = True: .ColHidden(12) = True: .ColHidden(13) = True
         .ColHidden(16) = True: .ColHidden(17) = True
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 0, 1, 2, 3, 4, 5, 6, 18
                         .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 8, 9, 16, 17
                         .ColAlignment(lngC) = flexAlignCenterCenter
                    Case Else
                         .ColAlignment(lngC) = flexAlignRightCenter
             End Select
         Next lngC
         .MergeCells = flexMergeRestrictColumns
         For lngC = 0 To 1
             .MergeCol(lngC) = True
         Next lngC
    End With
End Sub

Private Sub SubOther_FILL()
Dim strSQL        As String
Dim lngI          As Long
Dim intIndex      As Integer
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.분류코드 AS 분류코드, " _
                  & "ISNULL(T1.분류명,'') AS 분류명 " _
             & "FROM 자재분류 T1 " _
            & "ORDER BY T1.분류코드 "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       cboMt.ListIndex = -1
       Exit Sub
    Else
       cboMt.AddItem "00. " & "전체"
       Do Until P_adoRec.EOF
          cboMt.AddItem P_adoRec("분류코드") & ". " & P_adoRec("분류명")
          P_adoRec.MoveNext
       Loop
       P_adoRec.Close
       cboMt.ListIndex = 0
    End If
    With cboState
         .AddItem "전    체"
         .AddItem "정    상"
         .AddItem "사용불가"
         .AddItem "기    타"
         .ListIndex = 1
    End With
    With cboTaxGbn
         .AddItem "전    체"
         .AddItem "비 과 세"
         .AddItem "과    세"
         .ListIndex = 0
    End With
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "자재분류 읽기 실패"
    Unload Me
    Exit Sub
End Sub

'+---------------------------------+
'/// VsFlexGrid(vsfg1) 채우기///
'+---------------------------------+
Private Sub Subvsfg1_FILL()
Dim strSQL     As String
Dim strGroupBy As String
Dim strHaving  As String
Dim strWhere   As String
Dim strOrderBy As String
Dim lngR       As Long
Dim lngC       As Long
Dim lngRR      As Long
Dim lngRRR     As Long
    vsfg1.Rows = 1
    With vsfg1
         '검색조건 자재분류
         Select Case Mid(Trim(cboMt.Text), 1, 2)
                Case "00"      '분류 전체
                     strWhere = ""
                Case Else      '분류 전체 아니면
                     strWhere = "WHERE T1.분류코드 = '" & Mid(Trim(cboMt.Text), 1, 2) & "' "
         End Select
         strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                  & "T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' "
    End With
    '검색조건 과세구분
    Select Case cboTaxGbn.ListIndex
           Case 0 '전체
                strWhere = strWhere
           Case 1 '비과세
                strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") + "T3.과세구분 = 0 "
           Case 2 '과세
                strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") + "T3.과세구분 = 1 "
    End Select
    '검색조건 사용구분
    Select Case cboState.ListIndex
           Case 0 '전체
                strWhere = strWhere
           Case 1 '정상
                strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") + "T1.사용구분 = 0 "
           Case 2 '사용불가
                strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") + "T1.사용구분 = 9 "
           Case Else
                strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") + "NOT(T1.사용구분 = 0 OR T1.사용구분 = 9) "
    End Select
    If Len(P_strFindString1) = 0 Then            '정상적인 조회
       strOrderBy = "ORDER BY T1.사업장코드, T1.분류코드, T1.세부코드 "
    Else
       'strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                           & "T1.세부코드 LIKE '%" & P_strFindString1 & "%' " _
                           & "AND T3.자재명 LIKE '%" & P_strFindString2 & "%' "
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                           & "T1.분류코드 = '" & Mid(P_strFindString1, 1, 2) & "' AND T1.세부코드 = '" & Mid(P_strFindString1, 3) & "' "
                           
       strOrderBy = "ORDER BY T3.자재명 "
    End If
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.사업장코드 AS 사업장코드, T2.사업장명 AS 사업장명, " _
                  & "ISNULL(T1.분류코드,'') AS 분류코드, ISNULL(T4.분류명,'') AS 분류명, " _
                  & "ISNULL(T1.세부코드,'') AS 세부코드, T3.자재명 AS 자재명, " _
                  & "T3.규격 AS 규격, T3.단위 AS 단위, T3.폐기율 AS 폐기율, T3.과세구분 AS 과세구분, " _
                  & "T1.사용구분 AS 사용구분, T1.적정재고 AS 적정재고, " _
                  & "ISNULL(T1.최종입고일자,'') AS 최종입고일자, ISNULL(T1.최종출고일자,'') AS 최종출고일자, " _
                  & "(SELECT ISNULL(SUM(입고누계수량-출고누계수량),0) " _
                     & "FROM 자재원장마감 " _
                    & "WHERE 분류코드 = T1.분류코드 AND 세부코드 = T1.세부코드 " _
                      & "AND 사업장코드 = T1.사업장코드 " _
                      & "AND 마감년월 >= '" & Left(PB_regUserinfoU.UserClientDate, 4) + "00" & "' " _
                      & "AND 마감년월 < '" & Left(PB_regUserinfoU.UserClientDate, 6) & "') AS 이월재고, "
    strSQL = strSQL _
                  & "(SELECT ISNULL(SUM(입고수량),0) " _
                     & "FROM 자재입출내역 " _
                    & "WHERE 분류코드 = T1.분류코드 AND 세부코드 = T1.세부코드 " _
                      & "AND 사업장코드 = T1.사업장코드 AND 사용구분 = 0 " _
                      & "AND 입출고일자 BETWEEN '" & Left(PB_regUserinfoU.UserClientDate, 6) + "01" & "' " _
                                         & "AND '" & PB_regUserinfoU.UserClientDate & "') AS 입고수량, "
    strSQL = strSQL _
                  & "(SELECT ISNULL(SUM(출고수량),0) " _
                     & "FROM 자재입출내역 " _
                    & "WHERE 분류코드 = T1.분류코드 AND 세부코드 = T1.세부코드 " _
                      & "AND 사업장코드 = T1.사업장코드 AND 사용구분 = 0 " _
                      & "AND 입출고일자 BETWEEN '" & Left(PB_regUserinfoU.UserClientDate, 6) + "01" & "' " _
                                         & "AND '" & PB_regUserinfoU.UserClientDate & "') AS 출고수량 "
    strSQL = strSQL _
             & "FROM 자재원장 T1 " _
             & "LEFT JOIN 사업장 T2 ON T2.사업장코드 = T1.사업장코드 " _
             & "LEFT JOIN 자재 T3 " _
                    & "ON T3.분류코드 = T1.분류코드 AND T3.세부코드 = T1.세부코드 " _
             & "LEFT JOIN 자재분류 T4 ON T4.분류코드 = T1.분류코드 "
    strSQL = strSQL _
           & "" & strWhere & " " _
           & "" & strOrderBy & " "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    vsfg1.Rows = P_adoRec.RecordCount + vsfg1.FixedRows
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       Exit Sub
    Else
       With vsfg1
            .Cell(flexcpForeColor, 0, 0, .FixedRows - 1, .Cols - 1) = vbBlack
            If .Rows <= PC_intRowCnt Then
               .ScrollBars = flexScrollBarHorizontal
            Else
               .ScrollBars = flexScrollBarBoth
            End If
            Do Until P_adoRec.EOF
               lngR = lngR + 1
               .TextMatrix(lngR, 0) = P_adoRec("분류코드")
               .TextMatrix(lngR, 1) = IIf(IsNull(P_adoRec("분류명")), "", P_adoRec("분류명"))
               .TextMatrix(lngR, 2) = P_adoRec("분류코드") & P_adoRec("세부코드")
               .TextMatrix(lngR, 3) = IIf(IsNull(P_adoRec("자재명")), "", P_adoRec("자재명"))
               'FindRow 사용을 위해
               .TextMatrix(lngR, 4) = .TextMatrix(lngR, 0) & P_adoRec("세부코드")
               .Cell(flexcpData, lngR, 4, lngR, 4) = .TextMatrix(lngR, 4)
               .TextMatrix(lngR, 5) = IIf(IsNull(P_adoRec("규격")), "", P_adoRec("규격"))
               .TextMatrix(lngR, 6) = IIf(IsNull(P_adoRec("단위")), "", P_adoRec("단위"))
               .TextMatrix(lngR, 7) = IIf(IsNull(P_adoRec("폐기율")), "", P_adoRec("폐기율"))
               If P_adoRec("과세구분") = 0 Then
                  .TextMatrix(lngR, 8) = "비 과 세"
               Else
                  .TextMatrix(lngR, 8) = "과    세"
               End If
               If P_adoRec("사용구분") = 0 Then
                  .TextMatrix(lngR, 9) = "정    상"
               ElseIf _
                  P_adoRec("사용구분") = 9 Then
                  .TextMatrix(lngR, 9) = "사용불가"
               Else
                  .TextMatrix(lngR, 9) = "코드오류"
               End If
               .TextMatrix(lngR, 10) = IIf(IsNull(P_adoRec("적정재고")), "", P_adoRec("적정재고"))
               .TextMatrix(lngR, 11) = IIf(IsNull(P_adoRec("이월재고")), "", P_adoRec("이월재고"))
               .TextMatrix(lngR, 12) = IIf(IsNull(P_adoRec("입고수량")), "", P_adoRec("입고수량"))
               .TextMatrix(lngR, 13) = IIf(IsNull(P_adoRec("출고수량")), "", P_adoRec("출고수량"))
               .TextMatrix(lngR, 14) = .ValueMatrix(lngR, 11) + .ValueMatrix(lngR, 12) - .ValueMatrix(lngR, 13) '현재재고
               .TextMatrix(lngR, 15) = ""
               'If .ValueMatrix(lngR, 14) < .ValueMatrix(lngR, 10) Then
               '   .Cell(flexcpForeColor, lngR, 15, lngR, 15) = vbRed
               'End If
               If Len(P_adoRec("최종입고일자")) = 8 Then
                  .TextMatrix(lngR, 16) = Format(P_adoRec("최종입고일자"), "0000-00-00")
               End If
               If Len(P_adoRec("최종출고일자")) = 8 Then
                  .TextMatrix(lngR, 17) = Format(P_adoRec("최종출고일자"), "0000-00-00")
               End If
               If .TextMatrix(lngR, 2) = P_strFindString1 Then
                  lngRR = lngR
               End If
               P_adoRec.MoveNext
            Loop
            P_adoRec.Close
            If lngRR = 0 Then
               .Row = lngRR       'vsfg1_EnterCell 자동실행(만약 한건 일때는 자동실행 안함)
               If .Rows > PC_intRowCnt Then
                  '.TopRow = .Rows - PC_intRowCnt + .FixedRows
                  .TopRow = 1
               End If
            Else
               .Row = lngRR       'vsfg1_EnterCell 자동실행(만약 한건 일때는 자동실행 안함)
               If .Rows > PC_intRowCnt Then
                  .TopRow = .Row
               End If
            End If
            vsfg1_EnterCell       'vsfg1_EnterCell 자동실행(만약 한건 일때도 강제로 자동실행)
            .SetFocus
       End With
    End If
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "자재원장 읽기 실패"
    Unload Me
    Exit Sub
End Sub

'+-------------------------+
'/// Clear text1(index) ///
'+-------------------------+
Private Sub SubClearText()
Dim lngC As Long
    'For lngC = Text1.LBound To Text1.UBound
    '    Text1(lngC).Text = ""
    'Next lngC
End Sub

'+---------------------------+
'/// 크리스탈 리포터 출력 ///
'+---------------------------+
Private Sub cmdPrint_Click()
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
         If PB_Test = 0 Then
            strExeFile = App.Path & ".\재고조정내역.rpt"
         Else
            strExeFile = App.Path & ".\재고조정내역T.rpt"
         End If
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
            .Formulas(3) = "ForAppDate = '기준일자 : ' & '" & Format(DTOS(dtpFDate.Value), "0000-00-00") & "' & ' 부터 ' & '" & Format(DTOS(dtpTDate.Value), "0000-00-00") & "' & ' 까지' "
            '--- Formula Fields(Select Record) ---
            .Formulas(4) = "ForSelKindCode = '" & Mid(cboMt.Text, 1, 2) & "'"                           '분류코드
            If cboTaxGbn.ListIndex = 1 Then       '비과세
               .Formulas(5) = "ForSelTaxGbn = 0"
            ElseIf _
               cboTaxGbn.ListIndex = 2 Then       '과  세
               .Formulas(5) = "ForSelTaxGbn = 1"
            Else
               .Formulas(5) = "ForSelTaxGbn = 2"  '전  체
            End If
            If cboState.ListIndex = 1 Then         '정    상
               .Formulas(6) = "ForSelUsageGbn = 0"
            ElseIf _
               cboState.ListIndex = 2 Then         '사용불가
               .Formulas(6) = "ForSelUsageGbn = 9"
            Else
               .Formulas(6) = "ForSelUsageGbn = 2"  '전   체
            End If
            '--- Parameter Fields ---
            .StoredProcParam(0) = PB_regUserinfoU.UserBranchCode  '지점코드
            .StoredProcParam(1) = DTOS(dtpFDate.Value)           '기준일자(시작일자)
            .StoredProcParam(2) = DTOS(dtpTDate.Value)           '기준일자(종료일자)
         End If
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
            .WindowShowGroupTree = True
            .WindowShowPrintSetupBtn = True
            .WindowTop = 0: .WindowTop = 0: .WindowHeight = 11100: .WindowWidth = 15405
            .WindowState = crptMaximized
            .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "재고조정내역"
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

