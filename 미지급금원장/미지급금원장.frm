VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm미지급금원장 
   BorderStyle     =   0  '없음
   Caption         =   "미지급금원장"
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
      TabIndex        =   16
      Top             =   0
      Width           =   15195
      Begin VB.OptionButton optPrtChk1 
         Caption         =   "이름순"
         Height          =   255
         Left            =   6840
         TabIndex        =   33
         Top             =   390
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optPrtChk0 
         Caption         =   "코드순"
         Height          =   255
         Left            =   6840
         TabIndex        =   32
         Top             =   150
         Width           =   975
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   4920
         Top             =   200
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   375
         Left            =   7980
         Picture         =   "미지급금원장.frx":0000
         Style           =   1  '그래픽
         TabIndex        =   22
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   9120
         Picture         =   "미지급금원장.frx":0963
         Style           =   1  '그래픽
         TabIndex        =   20
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   13635
         Picture         =   "미지급금원장.frx":1308
         Style           =   1  '그래픽
         TabIndex        =   19
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   390
         Left            =   12510
         Picture         =   "미지급금원장.frx":1C56
         Style           =   1  '그래픽
         TabIndex        =   18
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   11385
         Picture         =   "미지급금원장.frx":25DA
         Style           =   1  '그래픽
         TabIndex        =   7
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Height          =   390
         Left            =   10260
         Picture         =   "미지급금원장.frx":2E61
         Style           =   1  '그래픽
         TabIndex        =   12
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00008000&
         BorderStyle     =   1  '단일 고정
         Caption         =   "매 입 처 지 급 처 리"
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
         TabIndex        =   17
         Top             =   180
         Width           =   4650
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid vsfg1 
      Height          =   8116
      Left            =   60
      TabIndex        =   8
      Top             =   1979
      Width           =   15195
      _cx             =   26802
      _cy             =   14316
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
      Height          =   1299
      Left            =   60
      TabIndex        =   14
      Top             =   630
      Width           =   15195
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   10  '한글 
         Index           =   2
         Left            =   9240
         MaxLength       =   20
         TabIndex        =   5
         Top             =   560
         Width           =   3255
      End
      Begin MSComCtl2.DTPicker dtpExpired_Date 
         Height          =   285
         Left            =   5040
         TabIndex        =   4
         Top             =   560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   57540609
         CurrentDate     =   38217
      End
      Begin VB.ComboBox cboSactionWay 
         Height          =   300
         Left            =   2475
         Style           =   2  '드롭다운 목록
         TabIndex        =   2
         Top             =   560
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   3
         Left            =   2475
         MaxLength       =   14
         TabIndex        =   3
         Top             =   915
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   10  '한글 
         Index           =   4
         Left            =   9240
         MaxLength       =   50
         TabIndex        =   6
         Top             =   915
         Width           =   5535
      End
      Begin VB.CheckBox chkTotal 
         Caption         =   "전체 매입처"
         Height          =   255
         Left            =   7785
         TabIndex        =   9
         Top             =   200
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   10  '한글 
         Index           =   1
         Left            =   4515
         MaxLength       =   50
         TabIndex        =   1
         Top             =   185
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   0
         Left            =   2475
         MaxLength       =   8
         TabIndex        =   0
         Top             =   185
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpF_Date 
         Height          =   270
         Left            =   10440
         TabIndex        =   10
         Top             =   200
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   57540609
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin MSComCtl2.DTPicker dtpT_Date 
         Height          =   270
         Left            =   12480
         TabIndex        =   11
         Top             =   200
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   57540609
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "어음번호"
         Height          =   240
         Index           =   5
         Left            =   7920
         TabIndex        =   31
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "지급)"
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
         Index           =   4
         Left            =   120
         TabIndex        =   30
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "만기일자"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   3
         Left            =   4155
         TabIndex        =   29
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "금액"
         Height          =   240
         Index           =   2
         Left            =   1275
         TabIndex        =   28
         Top             =   945
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "적요"
         Height          =   240
         Index           =   8
         Left            =   7920
         TabIndex        =   27
         Top             =   945
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "까지"
         Height          =   240
         Index           =   12
         Left            =   13800
         TabIndex        =   26
         Top             =   245
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "부터"
         Height          =   240
         Index           =   11
         Left            =   11760
         TabIndex        =   25
         Top             =   245
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "기준일자"
         Height          =   240
         Index           =   10
         Left            =   9360
         TabIndex        =   24
         Top             =   245
         Width           =   855
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
         TabIndex        =   23
         Top             =   245
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "[Home]"
         Height          =   240
         Index           =   15
         Left            =   3720
         TabIndex        =   21
         Top             =   245
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "결제방법"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   1275
         TabIndex        =   15
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "매입처코드"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   1275
         TabIndex        =   13
         Top             =   245
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frm미지급금원장"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' 프로그램 제 목 : 미지급금원장
' 사용된 Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' 참조된 Table   :
' 업  무  설  명 : 매입장부이력(거래(전표)내역)조회 + 매입처지급처리
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived       As Boolean
Private P_adoRec           As New ADODB.Recordset
Private P_intButton        As Integer
Private Const PC_intRowCnt As Integer = 26  '그리드 한 페이지 당 행수(FixedRows 포함)

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
       dtpF_Date.Value = Format(Mid(PB_regUserinfoU.UserClientDate, 1, 6) & "01", "0000-00-00")
       dtpT_Date.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
       'dtpT_Date.Value = DateAdd("d", -1, DateAdd("m", 1, dtpF_Date.Value))
       SubOther_FILL
       P_blnActived = True
       Screen.MousePointer = vbDefault
    End If
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "사업장 정보(서버와의 연결 실패)"
    Unload Me
    Exit Sub
End Sub

'+---------------+
'/// 검색조건 ///
'+---------------+
Private Sub chkTotal_Click()
    If chkTotal.Value = 1 Then
       cboSactionWay.Enabled = False: dtpExpired_Date.Enabled = False: Text1(2).Enabled = False
       Text1(3).Enabled = False: Text1(4).Enabled = False
    Else
       cboSactionWay.Enabled = True: Text1(2).Enabled = True: Text1(3).Enabled = True: Text1(4).Enabled = True
       If cboSactionWay.ListIndex = 1 Then dtpExpired_Date.Enabled = True
    End If
End Sub
Private Sub chkTotal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub dtpF_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub dtpT_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdFind.SetFocus
End Sub

'+---------------+
'/// 결제방법 ///
'+---------------+
Private Sub cboSactionWay_Click()
    With cboSactionWay
        If .ListIndex = 0 Or .ListIndex = 1 Then '현금 또는 수표
            dtpExpired_Date.Enabled = False
            Text1(2).Enabled = False
         Else
            dtpExpired_Date.Enabled = True
            Text1(2).Enabled = True
         End If
    End With
End Sub
Private Sub cboSactionWay_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    End If
End Sub
'+---------------+
'/// 만기일자 ///
'+---------------+
Private Sub dtpExpired_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       Text1(2).SetFocus
    End If
End Sub
'+-------------------+
'/// Text1(index) ///
'+-------------------+
Private Sub Text1_GotFocus(Index As Integer)
    With Text1(Index)
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim strSQL     As String
Dim strExeFile As String
Dim varRetVal  As Variant
    If (Index = 0 And (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn)) Then  '매입처검색
       PB_strSupplierCode = UPPER(Trim(Text1(Index).Text))
       PB_strSupplierName = ""  'Trim(Text1(Index + 1).Text)
       frm매입처검색.Show vbModal
       If (Len(PB_strSupplierCode) + Len(PB_strSupplierName)) = 0 Then '검색에서 취소(ESC)
       Else
          Text1(Index).Text = PB_strSupplierCode
          Text1(Index + 1).Text = PB_strSupplierName
       End If
       If PB_strSupplierCode <> "" Then
          SendKeys "{TAB}"
       End If
       PB_strSupplierCode = "": PB_strSupplierName = ""
    ElseIf _
       KeyCode = vbKeyDelete Then
       If Len(Text1(0).Text) = 0 Then
          Text1(1).Text = ""
          Exit Sub
       End If
    Else
       If KeyCode = vbKeyReturn Then
          Select Case Index
                 Case Text1.UBound
                      If chkTotal.Value = 0 And cmdSave.Enabled = True And vsfg1.Rows > 1 Then
                         cmdSave.SetFocus
                         Exit Sub
                      End If
           End Select
           SendKeys "{tab}"
       End If
    End If
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "매입처 읽기 실패"
    Unload Me
    Exit Sub
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim strSQL As String
    With Text1(Index)
         .Text = Trim(.Text)
         Select Case Index
                Case 0 '매입처
                     .Text = UPPER(Trim(.Text))
                     If Len(.Text) < 1 Then
                        Text1(1).Text = ""
                     End If
                Case 3
                     .Text = Format(Vals(Trim(.Text)), "#,#")
                Case Else
                     .Text = Trim(.Text)
         End Select
    End With
    Exit Sub
End Sub

'+-----------+
'/// Grid ///
'+-----------+
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
Private Sub vsfg1_Click()
Dim strData As String
    With vsfg1
         If (.MouseRow >= 0 And .MouseRow < .FixedRows) Then
         End If
    End With
End Sub
Private Sub vsfg1_EnterCell()
Dim lngR As Long
Dim lngC As Long
    With vsfg1
         'Text1(0).Enabled = False: Text1(2).Enabled = False
         If .Row >= .FixedRows Then
         End If
    End With
End Sub
Private Sub vsfg1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL       As String
Dim intRetVal    As Integer
Dim lngCnt       As Long
    With vsfg1
         'If KeyCode = vbKeyInsert Then
         '   SubClearText
         '   .Row = 0
         '   Text1(Text1.LBound).SetFocus
         '   Exit Sub
         'End If
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
    Screen.MousePointer = vbHourglass
    cmdFind.Enabled = False
    'SubClearText
    Subvsfg1_FILL
    cmdFind.Enabled = True
    Screen.MousePointer = vbDefault
    'If chkTotal.Value = 0 And cmdSave.Enabled = True And vsfg1.Rows > 1 Then
    '   cboSactionWay.Enabled = True: Text1(3).Enabled = True: Text1(4).Enabled = True
    '   cboSactionWay.SetFocus
    '   Exit Sub
    'End If
End Sub

'+-----------+
'/// 저장 ///
'+-----------+
Private Sub cmdSave_Click()
Dim strSQL        As String
Dim lngR          As Long
Dim lngC          As Long
Dim lngCnt        As Long
Dim blnOK         As Boolean
Dim intRetVal     As Integer
Dim strServerTime As String
Dim strTime       As String
    '입력내역 검사
    blnOK = False
    FncCheckTextBox lngC, blnOK
    If blnOK = False Then
       Select Case lngC
              Case -1
                   chkTotal.SetFocus
                   Exit Sub
              Case 0, 2, 3, 4
                   If Text1(lngC).Enabled = False Then
                      Text1(0).Enabled = True: Text1(2).Enabled = True: Text1(3).Enabled = True: Text1(4).Enabled = True
                   End If
       End Select
       Text1(lngC).SetFocus
       Exit Sub
    End If
    If blnOK = True Then
       intRetVal = MsgBox("입력된 자료를 저장하시겠습니까 ?", vbQuestion + vbYesNo + vbDefaultButton1, "자료 저장")
       If intRetVal = vbNo Then
          vsfg1.SetFocus
          Exit Sub
       End If
       cmdSave.Enabled = False
    End If
    Screen.MousePointer = vbHourglass
    '서버시간 구하기
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT RIGHT(CONVERT(VARCHAR(23),GETDATE(), 121),12) AS 서버시간 "
    On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    strServerTime = Mid(P_adoRec("서버시간"), 1, 2) + Mid(P_adoRec("서버시간"), 4, 2) + Mid(P_adoRec("서버시간"), 7, 2) _
                  + Mid(P_adoRec("서버시간"), 10)
    P_adoRec.Close
    strTime = strServerTime
    '미지급금내역
    PB_adoCnnSQL.BeginTrans
    strSQL = "INSERT INTO 미지급금내역(사업장코드, 매입처코드, " _
                                    & "미지급금지급일자, 미지급금지급시간," _
                                    & "미지급금지급금액, 결제방법, " _
                                    & "만기일자, 어음번호, " _
                                    & "적요, 수정일자, " _
                                    & "사용자코드) VALUES(" _
                        & "'" & PB_regUserinfoU.UserBranchCode & "', '" & Text1(0).Text & "', " _
                        & "'" & PB_regUserinfoU.UserClientDate & "', '" & strServerTime & "', " _
                        & "" & Vals(Text1(3).Text) & ", " & cboSactionWay.ListIndex & ", " _
                        & "'" & IIf(cboSactionWay.ListIndex = 2, DTOS(dtpExpired_Date.Value), "") & "', '" & Text1(2).Text & "', " _
                        & "'" & Text1(4).Text & "', '" & PB_regUserinfoU.UserServerDate & "', " _
                        & "'" & PB_regUserinfoU.UserCode & "' )"
    On Error GoTo ERROR_TABLE_INSERT
    PB_adoCnnSQL.Execute strSQL
    PB_adoCnnSQL.CommitTrans
    cmdSave.Enabled = True
    cmdFind.Enabled = False
    Text1(2).Text = "": Text1(3).Text = "": Text1(4).Text = ""
    Subvsfg1_FILL
    cmdFind.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "(서버와의 연결 실패)"
    Unload Me
    Exit Sub
ERROR_TABLE_SELECT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "미지급금 추가 실패"
    Unload Me
    Exit Sub
ERROR_TABLE_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "미지급금 추가 실패"
    Unload Me
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "미지급금 저장 실패"
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
    Set frm미지급금원장 = Nothing
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
    With vsfg1              'Rows 1, Cols 13, RowHeightMax(Min) 300
         .AllowBigSelection = False
         .AllowSelection = False
         .AllowUserResizing = flexResizeColumns
         .BackColorBkg = &H8000000F
         .BackColorSel = &H8000&
         .ExtendLastCol = True
         .ScrollBars = flexScrollBarHorizontal
         .ScrollTrack = True
         .SelectionMode = flexSelectionByRow
         .FixedRows = 1
         .FixedCols = 5
         '.FrozenCols = 5
         .Rows = 1             'Subvsfg1_Fill수행시에 설정
         .Cols = 13
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1000   '사업장코드
         .ColWidth(1) = 1000   '매입처코드
         .ColWidth(2) = 2000   '매입처명
         .ColWidth(3) = 1200   '일자
         .ColWidth(4) = 3000   '적요
         .ColWidth(5) = 1000   '입출고구분(구분)
         .ColWidth(6) = 1000   '입출고구분명(구분)
         .ColWidth(7) = 1000   '만기일자(0000-00-00)
         .ColWidth(8) = 2000   '어음번호
         .ColWidth(9) = 2000   '미지급금액
         .ColWidth(10) = 2000  '지급액
         .ColWidth(11) = 2000  '잔액
         .ColWidth(12) = 4000  '비고
         
         .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "사업장코드"  'H
         .TextMatrix(0, 1) = "매입처코드"  'H
         .TextMatrix(0, 2) = "매입처명"    'H
         .TextMatrix(0, 3) = "날짜"
         .TextMatrix(0, 4) = "적요"
         .TextMatrix(0, 5) = "구분"        'H
         .TextMatrix(0, 6) = "구분"
         .TextMatrix(0, 7) = "만기일자"
         .TextMatrix(0, 8) = "어음번호"
         .TextMatrix(0, 9) = "금액"
         .TextMatrix(0, 10) = "지급액"
         .TextMatrix(0, 11) = "잔액"
         .TextMatrix(0, 12) = "비고"
         
         .ColHidden(0) = True: .ColHidden(1) = True: .ColHidden(2) = True
         .ColHidden(5) = True
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 2, 4, 8, 12
                         .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 0, 1, 3, 5, 6, 7
                         .ColAlignment(lngC) = flexAlignCenterCenter
                    Case Else
                         .ColAlignment(lngC) = flexAlignRightCenter
             End Select
         Next lngC
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 9 To 11
                         .ColFormat(lngC) = "#,#.00"
             End Select
         Next lngC
         '.MergeCells = flexMergeFixedOnly
         '.MergeRow(0) = True: .MergeRow(1) = True
         'For lngC = 0 To .Cols - 1
         '    .MergeCol(lngC) = True
         'Next lngC
    End With
End Sub

'+--------------------------------------------------------------------+
'/// VsFlexGrid(vsfg1) 채우기(미지급금발생구분 : 1.전표, 2.계산서) ///
'+--------------------------------------------------------------------+
Private Sub Subvsfg1_FILL()
Dim strSQL      As String
Dim strJoin     As String
Dim strGroupBy  As String
Dim strHaving   As String
Dim strWhere    As String
Dim strOrderBy  As String
Dim lngR        As Long
Dim lngC        As Long
Dim lngRR       As Long
Dim lngRRR      As Long
Dim StrDate     As String    '해당일자
Dim curMonIMny  As Currency  '해당월누계금액(입고)
Dim curMonOMny  As Currency  '해당월누계금액(지급)
Dim curMonTMny  As Currency  '해당월누계금액(잔액)
Dim curTotIMny  As Currency  '해당누계금액(입고)
Dim curTotOMny  As Currency  '해당누계금액(지급)
Dim curTotTMny  As Currency  '해당누계금액(잔액)
Dim curTotTIMny As Currency  '전체누계금액(입고)
Dim curTotTOMny As Currency  '전체누계금액(지급)
Dim curTotTTMny As Currency  '전체누계금액(잔액)
    vsfg1.Rows = 1
    With vsfg1
         '검색조건 매입처
         If chkTotal.Value = 0 Then '건별 조회
            If Len(Text1(0).Text) > 0 Then
               strWhere = "WHERE T1.매입처코드 = '" & Trim(Text1(0).Text) & "' "
            Else
               Text1(0).SetFocus
               Exit Sub
            End If
         End If
         strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                  & "T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' "
    End With
    strOrderBy = "ORDER BY T1.사업장코드, " & IIf(optPrtChk0.Value = True, "T1.매입처코드, T3.매입처명 ", "T3.매입처명, T1.매입처코드 ") & ", 일자, 시간, 구분 "
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.사업장코드 AS 사업장코드, T2.사업장명 AS 사업장명, " _
                  & "T1.매입처코드 AS 매입처코드, T3.매입처명 AS 매입처명, " _
                  & "(T1.마감년월 + '00') AS 일자, '(년 이 월)' AS 적요, " _
                  & "0 AS 구분, (T1.미지급금누계금액) AS 입고금액, " _
                  & "(T1.미지급금지급누계금액) AS 지급금액, " _
                  & "'' AS 결제방법, '' AS 만기일자, '' AS 어음번호, '' AS 시간 " _
             & "FROM 미지급금원장마감 T1 " _
             & "LEFT JOIN 사업장 T2 ON T2.사업장코드 = T1.사업장코드 " _
             & "LEFT JOIN 매입처 T3 ON T3.사업장코드 = T1.사업장코드 AND T3.매입처코드 = T1.매입처코드 " _
             & "" & strWhere & " AND SUBSTRING(T1.마감년월,5,2) = '00' AND (T1.미지급금누계금액 <> 0 OR T1.미지급금지급누계금액 <> 0) " _
              & "AND T1.마감년월 BETWEEN '" & (Mid(DTOS(dtpF_Date.Value), 1, 4) + "00") & "' " _
                                  & "AND '" & (Mid(DTOS(dtpT_Date.Value), 1, 4) + "00") & "' "
    strSQL = strSQL & "UNION ALL " _
           & "SELECT T1.사업장코드 AS 사업장코드, T2.사업장명 AS 사업장명, " _
                  & "T1.매입처코드 AS 매입처코드, T3.매입처명 AS 매입처명, " _
                  & "(T1.마감년월 + '00') AS 일자, '(월    계)' AS 적요, " _
                  & "0 AS 구분, (T1.미지급금누계금액) AS 입고금액, " _
                  & "(T1.미지급금지급누계금액) AS 지급금액, " _
                  & "'' AS 결제방법, '' AS 만기일자, '' AS 어음번호, '' AS 시간 " _
             & "FROM 미지급금원장마감 T1 " _
             & "LEFT JOIN 사업장 T2 ON T2.사업장코드 = T1.사업장코드 " _
             & "LEFT JOIN 매입처 T3 ON T3.사업장코드 = T1.사업장코드 AND T3.매입처코드 = T1.매입처코드 " _
             & "" & strWhere & " AND SUBSTRING(T1.마감년월,5,2) <> '00' AND (T1.미지급금누계금액 <> 0 OR T1.미지급금지급누계금액 <> 0) " _
              & "AND T1.마감년월 > '" & (Mid(DTOS(dtpF_Date.Value), 1, 4) + "00") & "' " _
              & "AND T1.마감년월 < '" & Mid(DTOS(dtpF_Date.Value), 1, 6) & "' "
    If PB_regUserinfoU.UserMJGbn = "1" Then
       strSQL = strSQL & "UNION ALL " _
              & "SELECT T1.사업장코드 AS 사업장코드, T2.사업장명 AS 사업장명, " _
                  & "T1.매입처코드 AS 매입처코드, T3.매입처명 AS 매입처명, " _
                  & "(T1.입출고일자) AS 일자, (T1.적요) AS 적요, " _
                  & "T1.입출고구분 AS 구분, (SUM(T1.입고수량 * T1.입고단가) * " & (PB_curVatRate + 1) & ") AS 입고금액, " _
                  & "0 AS 지급금액, " _
                  & "결제방법 = CASE WHEN T1.현금구분 = 1 THEN '현금' ELSE '외상' END,  '' AS 만기일자, '' AS 어음번호, " _
                  & "T1.입출고시간 AS 시간 " _
             & "FROM 자재입출내역 T1 " _
             & "LEFT JOIN 사업장 T2 ON T2.사업장코드 = T1.사업장코드 " _
             & "LEFT JOIN 매입처 T3 ON T3.사업장코드 = T1.사업장코드 AND T3.매입처코드 = T1.매입처코드 " _
             & "" & strWhere & " AND T1.입출고구분 = 1 " _
              & "AND (T1.사용구분 = 0) " _
              & "AND T1.입출고일자 BETWEEN '" & (Mid(DTOS(dtpF_Date.Value), 1, 6) + "01") & "' AND '" & DTOS(dtpT_Date.Value) & "' " _
            & "GROUP BY T1.사업장코드, T2.사업장명, T1.매입처코드, T3.매입처명, " _
                     & "T1.입출고일자, T1.입출고시간, T1.현금구분, " _
                     & "T1.입출고구분 "
    Else
       strSQL = strSQL & "UNION ALL " _
           & "SELECT T1.사업장코드 AS 사업장코드, T2.사업장명 AS 사업장명, " _
                  & "T1.매입처코드 AS 매입처코드, T3.매입처명 AS 매입처명, " _
                  & "(T1.작성일자) AS 일자, T1.적요 AS 적요, " _
                  & "1 AS 구분, (SUM(T1.공급가액 + T1.세액)) AS 입고금액, " _
                  & "0 AS 지급금액, " _
                  & "결제방법 = CASE WHEN T1.금액구분 = 0 THEN '현금' WHEN T1.금액구분 = 1 THEN '수표' " _
                                  & "WHEN T1.금액구분 = 2 THEN '어음' ELSE '외상' END,  '' AS 만기일자, '' AS 어음번호, " _
                  & "T1.작성시간 AS 시간 " _
             & "FROM 매입세금계산서장부 T1 " _
             & "LEFT JOIN 사업장 T2 ON T2.사업장코드 = T1.사업장코드 " _
             & "LEFT JOIN 매입처 T3 ON T3.사업장코드 = T1.사업장코드 AND T3.매입처코드 = T1.매입처코드 " _
             & "" & strWhere & " " _
              & "AND (T1.사용구분 = 0) AND T1.미지급구분 = 1 " _
              & "AND T1.작성일자 BETWEEN '" & (Mid(DTOS(dtpF_Date.Value), 1, 6) + "01") & "' AND '" & DTOS(dtpT_Date.Value) & "' " _
            & "GROUP BY T1.사업장코드, T2.사업장명, T1.매입처코드, T3.매입처명, " _
                     & "T1.작성일자, T1.작성시간, T1.적요, T1.금액구분 "
    End If
    strSQL = strSQL & "UNION ALL " _
           & "SELECT T1.사업장코드 AS 사업장코드, T2.사업장명 AS 사업장명, " _
                  & "T1.매입처코드 AS 매입처코드, T3.매입처명 AS 매입처명, " _
                  & "(T1.미지급금지급일자) AS 일자, '' AS 적요, " _
                  & "0 AS 구분, 0 AS 입고금액, " _
                  & "ISNULL(SUM(T1.미지급금지급금액), 0) As 지급금액, " _
                  & "결제방법 = CASE WHEN T1.결제방법 = 0 THEN '현금' WHEN T1.결제방법 = 1 THEN '수표' " _
                                  & "WHEN T1.결제방법 = 2 THEN '어음' ELSE '기타' END, " _
                  & "T1.만기일자, T1.어음번호, T1.미지급금지급시간 AS 시간 " _
             & "FROM 미지급금내역 T1 " _
             & "LEFT JOIN 사업장 T2 ON T2.사업장코드 = T1.사업장코드 " _
             & "LEFT JOIN 매입처 T3 ON T3.사업장코드 = T1.사업장코드 AND T3.매입처코드 = T1.매입처코드 " _
            & "" & strWhere & " AND T1.미지급금지급일자 " _
                         & "BETWEEN '" & (Mid(DTOS(dtpF_Date.Value), 1, 6) + "01") & "' AND '" & DTOS(dtpT_Date.Value) & "' " _
            & "GROUP BY T1.사업장코드, T2.사업장명, T1.매입처코드, T3.매입처명, " _
                     & "T1.미지급금지급일자, T1.미지급금지급시간, T1.적요, T1.결제방법, T1.만기일자, T1.어음번호 "
    strSQL = strSQL _
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
            lngR = 0
            .AddItem ""
            lngR = lngR + 1
            .TextMatrix(lngR, 3) = P_adoRec("매입처코드"): .TextMatrix(lngR, 4) = P_adoRec("매입처명")
            .Cell(flexcpBackColor, lngR, 0, lngR, .Cols - 1) = vbYellow
            Do Until P_adoRec.EOF
               lngR = lngR + 1
               If lngR <> 2 Then    '처음 레코드 아니면
                  If .TextMatrix(lngR - 1, 1) <> P_adoRec("매입처코드") Then '매입처코드가 다르면
                     If .TextMatrix(lngR - 1, 3) <> "" Then '일 집계면
                        '해당월누계수량
                         .AddItem ""
                         '.TextMatrix(lngR, 4) = "(" & Format(Mid(StrDate, 1, 6), "0000-00") & Space(1) & "월계)"
                         .TextMatrix(lngR, 4) = "(월    계)"
                         .TextMatrix(lngR, 9) = curMonIMny   '월누계금액(입고)
                         .TextMatrix(lngR, 10) = curMonOMny  '월누계금액(지급)
                         curMonIMny = 0: curMonOMny = 0: curMonTMny = 0
                         lngR = lngR + 1
                     End If
                     '해당입(출)고누계
                     .TextMatrix(lngR, 4) = "(누    계)"
                     .TextMatrix(lngR, 9) = curTotIMny       '매입처누계금액(입고)
                     .TextMatrix(lngR, 10) = curTotOMny      '매입처누계금액(지급)
                     curTotIMny = 0: curTotOMny = 0: curTotTMny = 0
                     .AddItem ""
                     lngR = lngR + 1
                     .AddItem ""
                     lngR = lngR + 1
                     '매입처명
                     .TextMatrix(lngR, 3) = P_adoRec("매입처코드"): .TextMatrix(lngR, 4) = P_adoRec("매입처명")
                     .Cell(flexcpBackColor, lngR, 0, lngR, .Cols - 1) = vbYellow
                     .AddItem ""
                     lngR = lngR + 1
                  Else
                     If .TextMatrix(lngR - 1, 3) <> "" And _
                         Mid(StrDate, 1, 6) <> Mid(P_adoRec("일자"), 1, 6) Then '일 집계면 And 월이 다르면
                         '해당월입(출)고누계
                         .AddItem ""
                         '.TextMatrix(lngR, 4) = "(" & Format(Mid(StrDate, 1, 6), "0000-00") & Space(1) & "월계)"
                         .TextMatrix(lngR, 4) = "(월    계)"
                         .TextMatrix(lngR, 9) = curMonIMny   '월누계금액(입고)
                         .TextMatrix(lngR, 10) = curMonOMny  '월누계금액(지급)
                         curMonIMny = 0: curMonOMny = 0: curMonTMny = 0
                         lngR = lngR + 1
                     End If
                  End If
               End If
               .TextMatrix(lngR, 0) = P_adoRec("사업장코드")
               .TextMatrix(lngR, 1) = P_adoRec("매입처코드")
               .TextMatrix(lngR, 2) = P_adoRec("매입처명")
               '3. 일자
               If Mid(P_adoRec("일자"), 7, 2) = "00" Then
                  .TextMatrix(lngR, 3) = ""
               Else
                  .TextMatrix(lngR, 3) = Format(P_adoRec("일자"), "0000-00-00")
               End If
               '4. 적요
               If Mid(P_adoRec("일자"), 7, 2) = "00" Then
                  If Mid(P_adoRec("일자"), 5, 2) = "00" Then
                     .TextMatrix(lngR, 4) = "(" & Mid(P_adoRec("일자"), 1, 4) & " 년이월)"
                  Else
                     .TextMatrix(lngR, 4) = "(" & Format(Mid(P_adoRec("일자"), 1, 6), "0000-00") & " 월계)"
                  End If
               End If
               '5. 구분코드
               .TextMatrix(lngR, 5) = P_adoRec("구분")
               '6. 구분
               If PB_regUserinfoU.UserMJGbn = "1" Then
                  If P_adoRec("구분") = 0 Then
                     .TextMatrix(lngR, 6) = IIf(Mid(P_adoRec("일자"), 7, 2) <> "00", "지급", "") + IIf(P_adoRec("결제방법") = "", "", "(" + P_adoRec("결제방법") + ")")
                  ElseIf _
                     P_adoRec("구분") = 1 Then
                     .TextMatrix(lngR, 6) = "매입" + IIf(P_adoRec("결제방법") = "", "", "(" + P_adoRec("결제방법") + ")")
                  End If
               Else
                  If P_adoRec("구분") = 0 Then
                     .TextMatrix(lngR, 6) = IIf(Mid(P_adoRec("일자"), 7, 2) <> "00", "지급", "") + IIf(P_adoRec("결제방법") = "", "", "(" + P_adoRec("결제방법") + ")")
                  ElseIf _
                     P_adoRec("구분") = 2 Then
                     .TextMatrix(lngR, 6) = "매출" + IIf(P_adoRec("결제방법") = "", "", "(" + P_adoRec("결제방법") + ")")
                  End If
               End If
               '7.만기일자, 8.어음번호
               If Mid(P_adoRec("일자"), 7, 2) = "00" Then
               Else
                  .TextMatrix(lngR, 7) = IIf(Len(P_adoRec("만기일자")) > 0, Format(P_adoRec("만기일자"), "0000-00-00"), "")
                  .TextMatrix(lngR, 8) = IIf(Len(P_adoRec("어음번호")) > 0, P_adoRec("어음번호"), "")
               End If
               '9. 입고금액
               .TextMatrix(lngR, 9) = P_adoRec("입고금액")
               '10. 지급액
               .TextMatrix(lngR, 10) = P_adoRec("지급금액")
               '11. 잔액
               .TextMatrix(lngR, 11) = curTotTMny + (P_adoRec("입고금액") - P_adoRec("지급금액"))
               '12. 적요
               If Mid(P_adoRec("일자"), 7, 2) = "00" Then
               Else
                  .TextMatrix(lngR, 12) = P_adoRec("적요")
               End If
               If Mid(P_adoRec("일자"), 7, 2) <> 0 Then
                  curMonIMny = curMonIMny + P_adoRec("입고금액")
                  curMonOMny = curMonOMny + P_adoRec("지급금액")
                  curMonTMny = curMonTMny + (P_adoRec("입고금액") - P_adoRec("지급금액"))
               End If
               '해당누계금액
               curTotIMny = curTotIMny + P_adoRec("입고금액")
               curTotOMny = curTotOMny + P_adoRec("지급금액")
               curTotTMny = curTotTMny + (P_adoRec("입고금액") - P_adoRec("지급금액"))
               '전체누계금액
               curTotTIMny = curTotTIMny + P_adoRec("입고금액")
               curTotTOMny = curTotTOMny + P_adoRec("지급금액")
               curTotTTMny = curTotTTMny + (P_adoRec("입고금액") - P_adoRec("지급금액"))
               StrDate = P_adoRec("일자")
               
               'FindRow 사용을 위해
               '.TextMatrix(lngR, 4) = .TextMatrix(lngR, 0) & P_adoRec("세분류코드")
               '.Cell(flexcpData, lngR, 4, lngR, 4) = .TextMatrix(lngR, 4)
               P_adoRec.MoveNext
               If P_adoRec.EOF = True Then '마지막 레코드면
                  If .TextMatrix(lngR, 3) <> "" Then     '일 집계면
                     lngR = lngR + 1
                     '해당월누계
                     .AddItem ""
                     .TextMatrix(lngR, 4) = "(월    계)"
                     .TextMatrix(lngR, 9) = curMonIMny    '월누계금액(입고)
                     .TextMatrix(lngR, 10) = curMonOMny   '월누계금액(지급)
                     curMonIMny = 0: curMonOMny = 0: curMonTMny = 0
                  End If
                  '해당누계
                  lngR = lngR + 1
                  .AddItem ""
                  .TextMatrix(lngR, 4) = "(누    계)"
                  .TextMatrix(lngR, 9) = curTotIMny       '매입처누계금액(입고)
                  .TextMatrix(lngR, 10) = curTotOMny      '매입처누계금액(지급)
                  curTotIMny = 0: curTotOMny = 0: curTotTMny = 0
               End If
            Loop
            P_adoRec.Close
            '전체 합계
            If chkTotal.Value = 1 Then
               lngR = lngR + 1
               .AddItem ""
               lngR = lngR + 1
               .AddItem ""
               .TextMatrix(lngR, 4) = "(전체누계)"
               .TextMatrix(lngR, 9) = curTotTIMny        '전체누계금액(입고)
               .TextMatrix(lngR, 10) = curTotTOMny       '전체누계금액(지급)
               .TextMatrix(lngR, 11) = curTotTTMny       '전체누계금액(잔액)
               .Cell(flexcpBackColor, lngR, 0, lngR, .Cols - 1) = vbYellow
            End If
            If .Rows <= PC_intRowCnt Then
               .ScrollBars = flexScrollBarHorizontal
            Else
               .ScrollBars = flexScrollBarBoth
            End If
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
            If chkTotal.Value = 1 Then
               .TopRow = 1
            End If
            vsfg1_EnterCell       'vsfg1_EnterCell 자동실행(만약 한건 일때도 강제로 자동실행)
            .SetFocus
       End With
    End If
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "미지급금 읽기 실패"
    Unload Me
    Exit Sub
End Sub

Private Sub SubOther_FILL()
Dim strSQL        As String
Dim intIndex      As Integer
    cboSactionWay.AddItem "0. 현 금"
    cboSactionWay.AddItem "1. 수 표"
    cboSactionWay.AddItem "2. 어 음"
    cboSactionWay.ListIndex = 0
    dtpExpired_Date.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "TABLE 읽기 실패"
    Unload Me
    Exit Sub
End Sub

'+-------------------------+
'/// Clear text1(index) ///
'+-------------------------+
Private Sub SubClearText()
Dim lngC As Long
    For lngC = Text1.LBound To Text1.UBound
        Text1(lngC).Text = ""
    Next lngC
End Sub

'+-------------------------+
'/// Check text1(index) ///
'+-------------------------+
Private Function FncCheckTextBox(lngC As Long, blnOK As Boolean)
Dim lngR      As Long
Dim curJanMny As Currency '지급가능금액
    'If Not (chkTotal.Value = 0) Then '전체로는 지급 불가
    '   lngC = -1
    '   Exit Function
    'End If
    For lngC = Text1.LBound To Text1.UBound
        Select Case lngC
               Case 0  '매입처코드
                    If Not (Len(Text1(lngC).Text) > 0) Then
                       Exit Function
                    End If
               Case 1  '매입처명
                    If Len(Trim(Text1(lngC).Text)) = 0 Then
                       lngC = 0
                       Exit Function
                    End If
               Case 2  '어음번호
                    If Not (LenH(Text1(lngC).Text) <= 20) Then
                       Exit Function
                    End If
               Case 3  '지급금액
                    If Not (Vals(Text1(lngC).Text) <> 0) Then
                       Exit Function
                    Else
                       For lngR = vsfg1.Rows - 1 To 1 Step -1
                           If vsfg1.TextMatrix(lngR, 1) = Text1(0).Text Then
                              curJanMny = vsfg1.ValueMatrix(lngR, 9)
                              Exit For
                           End If
                       Next lngR
                       'If Not (curJanMny > 0) Then
                       '   Exit Function
                       'End If
                       'If Not (Vals(Text1(lngC).Text) <= curJanMny) Then
                       '   Exit Function
                       'End If
                    End If
               Case 4  '적요
                    If Not (LenH(Text1(lngC).Text) <= 50) Then
                       Exit Function
                    End If
               Case Else
        End Select
    Next lngC
    blnOK = True
End Function

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
    
    If DTOS(dtpF_Date.Value) > DTOS(dtpT_Date.Value) Then
       Exit Sub
    End If
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
            strExeFile = App.Path & ".\미지급금원장.rpt"
         Else
            strExeFile = App.Path & ".\미지급금원장T.rpt"
         End If
         varRetVal = Dir(strExeFile)
         If Len(varRetVal) = 0 Then
            intRetCHK = 0
         Else
            .ReportFileName = strExeFile
            On Error GoTo ERROR_CRYSTAL_REPORTS
            '--- Grid Size = 0.101 ---
            '--- Formula Fields ---
            .Formulas(0) = "ForAppPgDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'" '프로그램실행일자
            .Formulas(1) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                       '사업장명
            .Formulas(2) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                                  '출력일시
            '적용기준일자
            .Formulas(3) = "ForAppDate = '기준일자 : ' & '" & Format(DTOS(dtpF_Date.Value), "0000-00-00") & "' & ' 부터 ' & '" & Format(DTOS(dtpT_Date.Value), "0000-00-00") & "' & ' 까지' "
            'DECLARE @ParAppPgDate VarChar(8), @ParAppFDate VarChar(8),  @ParAppTDate VarChar(8), @ParSupplierCode VarChar(10)
            .Formulas(4) = "ForMJGbn = '" & PB_regUserinfoU.UserMJGbn & "'"                                '미지급금발생구분
            '프로그램실행일자
            '--- Parameter Fields ---
            .StoredProcParam(0) = PB_regUserinfoU.UserBranchCode
            .StoredProcParam(1) = PB_regUserinfoU.UserClientDate
            '적용(기준)일자시작
            .StoredProcParam(2) = DTOS(dtpF_Date.Value)
            '적용(기준)일자종료
            .StoredProcParam(3) = DTOS(dtpT_Date.Value)
            '매입처코드
            If chkTotal.Value = 0 Then
               If Len(Text1(0).Text) = 0 Then
                  .StoredProcParam(4) = " "
               Else
                  .StoredProcParam(4) = Trim(Text1(0).Text)
               End If
            Else
               .StoredProcParam(4) = " "
            End If
            .StoredProcParam(5) = CInt(PB_regUserinfoU.UserMJGbn)                 '미지급금발생구분(1.전표, 2.계산서)
            '.StoredProcParam(6) = IIf(optPrtChk0.Value = True, 0, 1)             '정렬순서(0.매입처코드, 1.매입처명)
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
            .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "미지급금원장"
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

