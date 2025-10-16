VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm금전출납관리 
   BorderStyle     =   0  '없음
   Caption         =   "제경비관리"
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
      TabIndex        =   18
      Top             =   0
      Width           =   15195
      Begin VB.CommandButton cmdPrint 
         Height          =   375
         Left            =   7980
         Picture         =   "금전출납관리.frx":0000
         Style           =   1  '그래픽
         TabIndex        =   22
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   9120
         Picture         =   "금전출납관리.frx":0963
         Style           =   1  '그래픽
         TabIndex        =   7
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   13635
         Picture         =   "금전출납관리.frx":1308
         Style           =   1  '그래픽
         TabIndex        =   21
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   390
         Left            =   12510
         Picture         =   "금전출납관리.frx":1C56
         Style           =   1  '그래픽
         TabIndex        =   20
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   11385
         Picture         =   "금전출납관리.frx":25DA
         Style           =   1  '그래픽
         TabIndex        =   6
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Height          =   390
         Left            =   10260
         Picture         =   "금전출납관리.frx":2E61
         Style           =   1  '그래픽
         TabIndex        =   12
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
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00008000&
         BorderStyle     =   1  '단일 고정
         Caption         =   "금 전 출 납 관 리"
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
         TabIndex        =   19
         Top             =   180
         Width           =   4650
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid vsfg1 
      Height          =   7896
      Left            =   60
      TabIndex        =   13
      Top             =   2055
      Width           =   15195
      _cx             =   26802
      _cy             =   13928
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
      Height          =   1395
      Left            =   60
      TabIndex        =   14
      Top             =   630
      Width           =   15195
      Begin VB.ComboBox cboIOGbn 
         Height          =   300
         Left            =   915
         Style           =   2  '드롭다운 목록
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtAccName 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         Left            =   4995
         MaxLength       =   30
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox txtFindAccName 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         Left            =   12480
         MaxLength       =   30
         TabIndex        =   9
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtFindAccCode 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Left            =   10440
         MaxLength       =   4
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtAccCode 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Left            =   3315
         MaxLength       =   4
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   10  '한글 
         Left            =   4995
         MaxLength       =   60
         TabIndex        =   5
         Top             =   1000
         Width           =   1575
      End
      Begin VB.TextBox txtJukyo 
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   10  '한글 
         Left            =   915
         MaxLength       =   60
         TabIndex        =   3
         Top             =   620
         Width           =   5655
      End
      Begin MSComCtl2.DTPicker dtpFDate 
         Height          =   270
         Left            =   10440
         TabIndex        =   10
         Top             =   600
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
         Left            =   12480
         TabIndex        =   11
         Top             =   600
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   57540609
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin MSComCtl2.DTPicker dtpCDate 
         Height          =   270
         Left            =   915
         TabIndex        =   0
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
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "금액"
         Height          =   240
         Index           =   6
         Left            =   4080
         TabIndex        =   34
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "계정코드"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   5
         Left            =   2040
         TabIndex        =   33
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "총(입/출)합계"
         Height          =   240
         Index           =   4
         Left            =   8880
         TabIndex        =   32
         Top             =   1020
         Width           =   1335
      End
      Begin VB.Label lblTotOut 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "0"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   12480
         TabIndex        =   31
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblTotIn 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "0"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   10440
         TabIndex        =   30
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "[Home]"
         Height          =   240
         Index           =   3
         Left            =   11760
         TabIndex        =   29
         Top             =   265
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "[Home]"
         Height          =   240
         Index           =   15
         Left            =   4320
         TabIndex        =   28
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "계정코드"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   9120
         TabIndex        =   27
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "까지"
         Height          =   240
         Index           =   12
         Left            =   13800
         TabIndex        =   26
         Top             =   660
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "부터"
         Height          =   240
         Index           =   11
         Left            =   11760
         TabIndex        =   25
         Top             =   660
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "일자"
         Height          =   240
         Index           =   10
         Left            =   9360
         TabIndex        =   24
         Top             =   660
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
         Left            =   8040
         TabIndex        =   23
         Top             =   285
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "구분"
         Height          =   240
         Index           =   8
         Left            =   75
         TabIndex        =   17
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "적요"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   75
         TabIndex        =   16
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "일자"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   75
         TabIndex        =   15
         Top             =   285
         Width           =   735
      End
   End
End
Attribute VB_Name = "frm금전출납관리"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' 프로그램 제 목 : 제경비관리
' 사용된 Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' 참조된 Table   :
' 업  무  설  명 :
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived       As Boolean
Private P_adoRec           As New ADODB.Recordset
Private P_intButton        As Integer
Private Const PC_intRowCnt As Integer = 24  '그리드 한 페이지 당 행수(FixedRows 포함)

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
                   cmdPrint.Enabled = False: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 40 '추가, 저장, 인쇄, 조회
                   cmdPrint.Enabled = False: cmdClear.Enabled = True: cmdFind.Enabled = True
                   cmdSave.Enabled = True: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 50 '삭제, 추가, 저장, 조회, 인쇄
                   cmdPrint.Enabled = False: cmdClear.Enabled = True: cmdFind.Enabled = True
                   cmdSave.Enabled = True: cmdDelete.Enabled = True: cmdExit.Enabled = True
              Case Is <= 99 '삭제, 추가, 저장, 조회, 인쇄
                   cmdPrint.Enabled = False: cmdClear.Enabled = True: cmdFind.Enabled = True
                   cmdSave.Enabled = True: cmdDelete.Enabled = True: cmdExit.Enabled = True
              Case Else
                   cmdClear.Enabled = False: cmdFind.Enabled = False: cmdSave.Enabled = False
                   cmdDelete.Enabled = False: cmdExit.Enabled = True
       End Select
       dtpCDate.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
       dtpFDate.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
       dtpTDate.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
       'dtpFDate.Value = Format(Mid(PB_regUserinfoU.UserClientDate, 1, 6) & "01", "0000-00-00")
       'dtpTDate.Value = DateAdd("d", -1, DateAdd("m", 1, dtpFDate.Value))
       SubOther_FILL
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

'+--------------------+
'/// 입력/수정조건 ///
'+--------------------+
'작성일자
Private Sub dtpCDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
'계정코드
Private Sub txtAccCode_GotFocus()
    With txtAccCode
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtAccCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn) Then  '계정코드검색
       PB_strAccCode = UPPER(Trim(txtAccCode.Text))
       PB_strAccName = ""  'Trim(Text1(Index + 1).Text)
       frm계정코드검색.Show vbModal
       If (Len(PB_strAccCode) + Len(PB_strAccName)) = 0 Then '검색에서 취소(ESC)
       Else
          txtAccCode.Text = PB_strAccCode
          txtAccName.Text = PB_strAccName
       End If
       If PB_strAccCode <> "" Then
          SendKeys "{TAB}"
       End If
       PB_strAccCode = "": PB_strAccName = ""
    ElseIf _
       KeyCode = vbKeyDelete Then
       If Len(txtAccCode) = 0 Then
          txtAccName.Text = ""
          Exit Sub
       End If
    Else
       If KeyCode = vbKeyReturn Then
          SendKeys "{tab}"
       End If
    End If
    Exit Sub
End Sub
Private Sub txtAccCode_LostFocus()
    With txtAccCode
         .Text = Trim(.Text)
         If Len(.Text) < 1 Then
            txtAccName.Text = ""
         End If
    End With
End Sub

Private Sub txtJukyo_GotFocus()
    With txtJukyo
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtJukyo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    End If
    Exit Sub
End Sub
'입출구분
Private Sub cboIOGbn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    End If
    Exit Sub
End Sub
'금액
Private Sub txtMoney_GotFocus()
    With txtMoney
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtMoney_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    End If
    Exit Sub
End Sub
Private Sub txtMoney_LostFocus()
    With txtMoney
         .Text = Format(Vals(Trim(.Text)), "#,0")
    End With
End Sub
'+---------------+
'/// 검색조건 ///
'+---------------+
Private Sub txtFindAccCode_GotFocus()
    With txtFindAccCode
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtFindAccCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn) Then  '계정코드검색
       PB_strAccCode = UPPER(Trim(txtFindAccCode.Text))
       PB_strAccName = ""  'Trim(Text1(Index + 1).Text)
       frm계정코드검색.Show vbModal
       If (Len(PB_strAccCode) + Len(PB_strAccName)) = 0 Then '검색에서 취소(ESC)
       Else
          txtFindAccCode.Text = PB_strAccCode
          txtFindAccName.Text = PB_strAccName
       End If
       If PB_strAccCode <> "" Then
          SendKeys "{TAB}"
       Else
          SendKeys "{TAB}"
       End If
       PB_strAccCode = "": PB_strAccName = ""
    ElseIf _
       KeyCode = vbKeyDelete Then
       If Len(txtFindAccCode) = 0 Then
          txtFindAccName.Text = ""
          Exit Sub
       End If
    Else
       If KeyCode = vbKeyReturn Then
          SendKeys "{tab}"
       End If
    End If
    Exit Sub
End Sub
Private Sub txtFindAccCode_LostFocus()
    With txtFindAccCode
         .Text = Trim(.Text)
         If Len(.Text) < 1 Then
            txtFindAccName.Text = ""
         End If
    End With
End Sub
Private Sub dtpFDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub dtpTDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdFind.SetFocus
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
            .Col = .MouseCol
            .Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlack
            .Cell(flexcpForeColor, 0, .MouseCol, 0, .MouseCol) = vbRed
            strData = Trim(.Cell(flexcpData, .Row, 0))
            Select Case .MouseCol
                   Case 1 '작성일자
                        .ColSel = 2
                        .ColSort(0) = flexSortNone
                        .ColSort(1) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .ColSort(2) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .Sort = flexSortUseColSort
                   Case 2 '작성시간
                        .ColSel = 3
                        .ColSort(0) = flexSortNone
                        .ColSort(1) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .ColSort(2) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .Sort = flexSortUseColSort
                   Case 3 '계정코드
                        .ColSel = 4
                        .ColSort(0) = flexSortNone
                        .ColSort(1) = flexSortNone
                        .ColSort(2) = flexSortNone
                        .ColSort(3) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .Sort = flexSortUseColSort
                   Case 4 '계정명
                        .ColSel = 5
                        .ColSort(0) = flexSortNone
                        .ColSort(1) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .ColSort(2) = flexSortNone
                        .ColSort(3) = flexSortNone
                        .ColSort(4) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .Sort = flexSortUseColSort
                   Case Else
                        .Sort = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            End Select
            If .FindRow(strData, , 0) > 0 Then
               .Row = .FindRow(strData, , 0)
            End If
            If PC_intRowCnt < .Rows Then
               .TopRow = .Row
            End If
         End If
    End With
End Sub
Private Sub vsfg1_EnterCell()
Dim lngR As Long
Dim lngC As Long
    With vsfg1
         If .Row >= .FixedRows Then
            dtpCDate.Value = Format(DTOS(.TextMatrix(.Row, 1)), "0000-00-00") '일자
            txtAccCode.Text = .TextMatrix(.Row, 3): txtAccName.Text = .TextMatrix(.Row, 4)  '계정코드, 계정명
            txtJukyo.Text = .TextMatrix(.Row, 5) '적요
            If .ValueMatrix(.Row, 6) > 0 Then
               cboIOGbn.ListIndex = 0
               cboIOGbn.Text = "1. 입금"
               txtMoney.Text = Format(.ValueMatrix(.Row, 6), "#,0")
            Else
               cboIOGbn.ListIndex = 1
               cboIOGbn.Text = "2. 출금"
               txtMoney.Text = Format(.ValueMatrix(.Row, 7), "#,0")
            End If
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
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "사업장 읽기 실패"
    Unload Me
    Exit Sub
End Sub

'+-----------+
'/// 추가 ///
'+-----------+
Private Sub cmdClear_Click()
Dim strSQL As String
    SubClearText
    vsfg1.Row = 0
    dtpCDate.SetFocus
End Sub
'+-----------+
'/// 조회 ///
'+-----------+
Private Sub cmdFind_Click()
    Screen.MousePointer = vbHourglass
    cmdFind.Enabled = False
    SubClearText
    Subvsfg1_FILL
    cmdFind.Enabled = True
    Screen.MousePointer = vbDefault
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
Dim intAddMode    As Integer '1.추가, Etc.저장
    '입력내역 검사
    blnOK = False
    FncCheckTextBox blnOK
    If blnOK = False Then
       Exit Sub
    End If
    If vsfg1.Row < vsfg1.FixedRows Then
       intAddMode = 1
       intRetVal = MsgBox("입력된 자료를 추가하시겠습니까 ?", vbQuestion + vbYesNo + vbDefaultButton1, "자료 추가")
    Else
       intRetVal = MsgBox("수정된 자료를 저장하시겠습니까 ?", vbQuestion + vbYesNo + vbDefaultButton1, "자료 저장")
    End If
    If intRetVal = vbNo Then
       vsfg1.SetFocus
       Exit Sub
    End If
    With vsfg1
         '서버시간 구하기
         cmdSave.Enabled = False
         Screen.MousePointer = vbHourglass
         P_adoRec.CursorLocation = adUseClient
         strSQL = "SELECT RIGHT(CONVERT(VARCHAR(23),GETDATE(), 121),12) AS 서버시간 "
         On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
         P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
         strServerTime = Mid(P_adoRec("서버시간"), 1, 2) + Mid(P_adoRec("서버시간"), 4, 2) + Mid(P_adoRec("서버시간"), 7, 2) _
                       + Mid(P_adoRec("서버시간"), 10)
         P_adoRec.Close
         PB_adoCnnSQL.BeginTrans
         If intAddMode = 1 Then '회계전표내역 추가
            strSQL = "INSERT INTO 회계전표내역(사업장코드, 작성일자, 작성시간, 계정코드, " _
                                            & "입출구분, 입금금액, 출금금액, 적요, 사용구분, 작성자코드, " _
                                            & "수정일자,사용자코드) Values( " _
                    & "'" & PB_regUserinfoU.UserBranchCode & "','" & DTOS(dtpCDate.Value) & "', " _
                    & "'" & strServerTime & "', '" & txtAccCode.Text & "', " _
                    & "" & cboIOGbn.ListIndex + 1 & ", " & IIf(cboIOGbn.ListIndex = 0, Vals(txtMoney.Text), 0) & ", " _
                    & "" & IIf(cboIOGbn.ListIndex = 1, Vals(txtMoney.Text), 0) & ", '" & txtJukyo.Text & "', " _
                    & "0, '" & PB_regUserinfoU.UserCode & "', " _
                    & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' )"
            On Error GoTo ERROR_TABLE_INSERT
            .AddItem .Rows
            .TextMatrix(.Rows - 1, 0) = PB_regUserinfoU.UserBranchCode & DTOS(dtpCDate.Value) & strServerTime & txtAccCode.Text
            .Cell(flexcpData, .Rows - 1, 0, .Rows - 1, 0) = Trim(.TextMatrix(.Rows - 1, 0))
            .TextMatrix(.Rows - 1, 1) = Format(DTOS(dtpCDate.Value), "0000-00-00")            '작성일자
            .TextMatrix(.Rows - 1, 2) = Format(Mid(strServerTime, 1, 6), "00:00:00")          '작성시간
            .TextMatrix(.Rows - 1, 3) = txtAccCode.Text                                       '계정코드
            .TextMatrix(.Rows - 1, 4) = txtAccName.Text                                       '계정명
            .TextMatrix(.Rows - 1, 5) = txtJukyo.Text                                         '적요
            .TextMatrix(.Rows - 1, 6) = IIf(cboIOGbn.ListIndex = 0, Format(Vals(txtMoney.Text), "#,0"), 0) '입금금액
            .TextMatrix(.Rows - 1, 7) = IIf(cboIOGbn.ListIndex = 1, Format(Vals(txtMoney.Text), "#,0"), 0) '출금금액
            .TextMatrix(.Rows - 1, 8) = PB_regUserinfoU.UserCode                              '작성자코드
            .TextMatrix(.Rows - 1, 9) = PB_regUserinfoU.UserName                              '작성자명
            .TextMatrix(.Rows - 1, 10) = "정    상"                                           '사용구분
            .TextMatrix(.Rows - 1, 11) = Format(PB_regUserinfoU.UserServerDate, "0000-00-00") '수정일자
            .TextMatrix(.Rows - 1, 12) = PB_regUserinfoU.UserCode                             '사용자코드
            .TextMatrix(.Rows - 1, 13) = PB_regUserinfoU.UserName                             '사용자명
            .TextMatrix(.Rows - 1, 14) = strServerTime                                        '작성시간
            If (dtpCDate.Value >= dtpFDate.Value And dtpCDate.Value <= dtpTDate.Value) Then
               If cboIOGbn.ListIndex = 0 Then
                  lblTotIn.Caption = Format(Vals(lblTotIn.Caption) + Vals(txtMoney.Text), "#,0")
               Else
                  lblTotOut.Caption = Format(Vals(lblTotOut.Caption) + Vals(txtMoney.Text), "#,0")
               End If
            End If
            If .Rows > PC_intRowCnt Then
               .ScrollBars = flexScrollBarBoth
               .TopRow = .Rows - 1
            End If
            .Row = .Rows - 1          '자동으로 vsfg1_EnterCell Event 발생
         Else                                         '회계전표내역 변경
            strSQL = "UPDATE 회계전표내역 SET " _
                          & "작성일자 = '" & DTOS(dtpCDate.Value) & "', " _
                          & "계정코드 = '" & txtAccCode.Text & "', " _
                          & "적요 = '" & txtJukyo.Text & "', " _
                          & "입금금액 = " & IIf(cboIOGbn.ListIndex = 0, Vals(txtMoney.Text), 0) & ", " _
                          & "출금금액 = " & IIf(cboIOGbn.ListIndex = 1, Vals(txtMoney.Text), 0) & ", " _
                          & "수정일자 = '" & PB_regUserinfoU.UserServerDate & "', " _
                          & "사용자코드 = '" & PB_regUserinfoU.UserCode & "' " _
                    & "WHERE 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                      & "AND 작성일자 = '" & DTOS(.TextMatrix(.Row, 1)) & "' " _
                      & "AND 작성시간 = '" & .TextMatrix(.Row, 14) & "' " _
                      & "AND 계정코드 = '" & .TextMatrix(.Row, 3) & "' "
            On Error GoTo ERROR_TABLE_UPDATE
            '해당기간에 포함되면 기존입출금액 합계에서 뺀다.
            If (DTOS(.TextMatrix(.Row, 1)) >= DTOS(dtpFDate.Value) And DTOS(.TextMatrix(.Row, 1)) <= DTOS(dtpTDate.Value)) Then
               lblTotIn.Caption = Format(Vals(lblTotIn.Caption) - .ValueMatrix(.Row, 6), "#,0")
               lblTotOut.Caption = Format(Vals(lblTotOut.Caption) - .ValueMatrix(.Row, 7), "#,0")
            End If
            .TextMatrix(.Row, 0) = PB_regUserinfoU.UserBranchCode & DTOS(dtpCDate.Value) & strServerTime & txtAccCode.Text
            .TextMatrix(.Row, 1) = Format(DTOS(dtpCDate.Value), "0000-00-00")            '작성일자
            .TextMatrix(.Row, 2) = Format(Mid(strServerTime, 1, 6), "00:00:00")          '작성시간
            .TextMatrix(.Row, 3) = txtAccCode.Text                                       '계정코드
            .TextMatrix(.Row, 4) = txtAccName.Text                                       '계정명
            .TextMatrix(.Row, 5) = txtJukyo.Text                                         '적요
            If (dtpCDate.Value >= dtpFDate.Value And dtpCDate.Value <= dtpTDate.Value) Then
               If cboIOGbn.ListIndex = 0 Then
                  lblTotIn.Caption = Format(Vals(lblTotIn.Caption) + Vals(txtMoney.Text), "#,0")
               Else
                  lblTotOut.Caption = Format(Vals(lblTotOut.Caption) + Vals(txtMoney.Text), "#,0")
               End If
            End If
            .TextMatrix(.Row, 6) = 0: .TextMatrix(.Row, 7) = 0
            .TextMatrix(.Row, 6) = IIf(cboIOGbn.ListIndex = 0, Format(Vals(txtMoney.Text), "#,0"), 0) '입금금액
            .TextMatrix(.Row, 7) = IIf(cboIOGbn.ListIndex = 1, Format(Vals(txtMoney.Text), "#,0"), 0) '출금금액
            .TextMatrix(.Row, 10) = "정    상"                                           '사용구분
            .TextMatrix(.Row, 11) = Format(PB_regUserinfoU.UserServerDate, "0000-00-00") '수정일자
            .TextMatrix(.Row, 12) = PB_regUserinfoU.UserCode                             '사용자코드
            .TextMatrix(.Row, 13) = PB_regUserinfoU.UserName                             '사용자명
         End If
         PB_adoCnnSQL.Execute strSQL
         PB_adoCnnSQL.CommitTrans
         Screen.MousePointer = vbDefault
    End With
    cmdSave.Enabled = True
    If intAddMode = 1 Then
       cmdClear_Click '추가모드로 바로가기
    Else
       vsfg1.SetFocus '그리드로 바로가기
    End If
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "(서버와의 연결 실패)"
    Unload Me
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "회계전표내역 읽기 실패"
    Unload Me
    Exit Sub
ERROR_TABLE_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "회계전표내역 추가 실패"
    Unload Me
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "회계전표내역 갱신 실패"
    Unload Me
    Exit Sub
End Sub

'+-----------+
'/// 삭제 ///
'+-----------+
Private Sub cmdDelete_Click()
Dim strSQL        As String
Dim intRetVal     As Integer
Dim lngCnt        As Long
Dim strServerTime As String
Dim strTime       As String
    With vsfg1
         If .Row >= .FixedRows Then
            intRetVal = MsgBox("등록된 자료를 삭제하시겠습니까 ?", vbCritical + vbYesNo + vbDefaultButton2, "자료 삭제")
            If intRetVal = vbYes Then
               Screen.MousePointer = vbHourglass
               cmdDelete.Enabled = False
               '삭제전 관련테이블 검사
               'P_adoRec.CursorLocation = adUseClient
               'strSQL = "SELECT Count(*) AS 해당건수 FROM TableName " _
               '        & "WHERE 사업장구분 = " & .TextMatrix(.Row, 0) & " "
               'On Error GoTo ERROR_TABLE_SELECT
               'P_adoRec.Open SQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
               'lngCnt = P_adoRec("해당건수")
               'P_adoRec.Close
               If lngCnt <> 0 Then
                  MsgBox "XXX(" & Format(lngCnt, "#,#") & "건)가 있으므로 삭제할 수 없습니다.", vbCritical, "회계전표내역 삭제 불가"
                  cmdDelete.Enabled = True
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
               '서버시간
               P_adoRec.CursorLocation = adUseClient
               strSQL = "SELECT RIGHT(CONVERT(VARCHAR(23),GETDATE(), 121),12) AS 서버시간 "
               On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
               P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
               strServerTime = Mid(P_adoRec("서버시간"), 1, 2) + Mid(P_adoRec("서버시간"), 4, 2) + Mid(P_adoRec("서버시간"), 7, 2) _
                             + Mid(P_adoRec("서버시간"), 10)
               P_adoRec.Close
               PB_adoCnnSQL.BeginTrans
               P_adoRec.CursorLocation = adUseClient
               strSQL = "UPDATE 회계전표내역 SET " _
                             & "사용구분 = 9, " _
                             & "수정일자 = '" & Mid(strServerTime, 1, 6) & "', " _
                             & "사용자코드 = '" & PB_regUserinfoU.UserCode & "' " _
                       & "WHERE 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                         & "AND 작성일자 = '" & DTOS(.TextMatrix(.Row, 1)) & "' " _
                         & "AND 작성시간 = '" & .TextMatrix(.Row, 14) & "' " _
                         & "AND 계정코드 = '" & .TextMatrix(.Row, 3) & "' "
               On Error GoTo ERROR_TABLE_DELETE
               PB_adoCnnSQL.Execute strSQL
               PB_adoCnnSQL.CommitTrans
               '총합계재계산
               If (DTOS(.TextMatrix(.Row, 1)) >= DTOS(dtpFDate.Value) And DTOS(.TextMatrix(.Row, 1)) <= DTOS(dtpTDate.Value)) Then
                  If .ValueMatrix(.Row, 6) > 0 Then lblTotIn.Caption = Format(Vals(lblTotIn.Caption) - .ValueMatrix(.Row, 6), "#,0")
                  If .ValueMatrix(.Row, 7) > 0 Then lblTotOut.Caption = Format(Vals(lblTotOut.Caption) - .ValueMatrix(.Row, 7), "#,0")
               End If
               .RemoveItem .Row
               If .Rows <= PC_intRowCnt Then
                  '.ScrollBars = flexScrollBarHorizontal
               End If
               cmdDelete.Enabled = True
               Screen.MousePointer = vbDefault
               If .Rows = 1 Then
                  SubClearText
                  .Row = 0
                  Exit Sub
               End If
               vsfg1_EnterCell
            End If
            .SetFocus
         End If
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "(서버와의 연결 실패)"
    Unload Me
    Exit Sub
ERROR_TABLE_DELETE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "회계전표내역 삭제 실패"
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
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
    Set frm금전출납관리 = Nothing
    frmMain.SBar.Panels(4).Text = ""
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
    'cboCode(0).Enabled = False               '계정코드 FLASE
    With vsfg1              'Rows 1, Cols 15, RowHeightMax(Min) 300
         .AllowBigSelection = False
         .AllowSelection = False
         .AllowUserResizing = flexResizeColumns
         .BackColorBkg = &H8000000F
         .BackColorSel = &H8000&
         .ExtendLastCol = True
         .ScrollBars = flexScrollBarVertical
         .ScrollTrack = True
         .SelectionMode = flexSelectionByRow
         .FixedRows = 1
         .FixedCols = 0
         .Rows = 1             'Subvsfg1_Fill수행시에 설정
         .Cols = 15
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1000   'KEY(사업장코드+작성일자+작성시간+계정코드)
         .ColWidth(1) = 1100   '작성일자(0000-00-00)
         .ColWidth(2) = 1000   '작성시간(00:00:00)
         .ColWidth(3) = 500    '코드
         .ColWidth(4) = 2800   '계정명
         .ColWidth(5) = 5500   '적요
         .ColWidth(6) = 1500   '입금금액
         .ColWidth(7) = 1500   '출금금액
         .ColWidth(8) = 1      '작성자코드
         .ColWidth(9) = 1000   '작성자명
         .ColWidth(10) = 1     '사용구분
         .ColWidth(11) = 1200  '수정일자
         .ColWidth(12) = 1     '사용자코드
         .ColWidth(13) = 1000  '사용자명
         .ColWidth(14) = 1000  '사간(밀리초 까지)
         
         .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "KEY"         'H
         .TextMatrix(0, 1) = "일자"
         .TextMatrix(0, 2) = "시간"
         .TextMatrix(0, 3) = "코드"
         .TextMatrix(0, 4) = "계정명"
         .TextMatrix(0, 5) = "적요"
         .TextMatrix(0, 6) = "입금금액"
         .TextMatrix(0, 7) = "출금금액"
         .TextMatrix(0, 8) = "작성자코드"  'H
         .TextMatrix(0, 9) = "작성자명"
         .TextMatrix(0, 10) = "사용구분"   'H
         .TextMatrix(0, 11) = "수정일자"   'H
         .TextMatrix(0, 12) = "사용자코드" 'H
         .TextMatrix(0, 13) = "사용자명"   'H
         .TextMatrix(0, 14) = "시간"       'H
         .ColHidden(0) = True: .ColHidden(8) = True: .ColHidden(10) = True
         .ColHidden(11) = True: .ColHidden(12) = True: .ColHidden(13) = True: .ColHidden(14) = True
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 0, 4, 5
                         .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 1, 2, 3, 8, 9, 10, 11, 12, 13, 14
                         .ColAlignment(lngC) = flexAlignCenterCenter
                    Case Else
                         .ColAlignment(lngC) = flexAlignRightCenter
             End Select
         Next lngC
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 6, 7
                         .ColFormat(lngC) = "#,#"
             End Select
         Next lngC
    End With
End Sub

'+---------------------------------+
'/// VsFlexGrid(vsfg1) 채우기///
'+---------------------------------+
Private Sub Subvsfg1_FILL()
Dim strSQL     As String
Dim strWhere   As String
Dim strOrderBy As String
Dim lngR       As Long
Dim lngC       As Long
Dim lngRR      As Long
    lblTotIn.Caption = "0": lblTotOut.Caption = "0"
    If dtpFDate.Value > dtpTDate.Value Then
       dtpFDate.SetFocus
       Exit Sub
    End If
    vsfg1.Rows = 1
    With vsfg1
         '검색조건 계정코드
         txtFindAccCode.Text = Trim(txtFindAccCode.Text)
         Select Case txtFindAccCode.Text
                Case ""         '계정코드 전체
                     strWhere = strWhere
                Case Else        '계정코드 전체 아니면
                     strWhere = IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                              & "T1.계정코드 = '" & txtFindAccCode.Text & "' "
         End Select
         strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                  & "T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' "
         strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                  & "(T1.작성일자 BETWEEN '" & DTOS(dtpFDate.Value) & "' AND '" & DTOS(dtpTDate.Value) & "') " _
                                     & "AND T1.사용구분 = 0 "
         strOrderBy = "ORDER BY T1.사업장코드, T1.작성일자, T1.작성시간 "
    End With
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.사업장코드 AS 사업장코드, T1.작성일자 AS 작성일자, " _
                  & "T1.작성시간 AS 작성시간, T1.계정코드 AS 계정코드, " _
                  & "ISNULL(T2.계정명, '') AS 계정명, ISNULL(T1.적요, '') AS 적요, " _
                  & "ISNULL(T1.입금금액, 0) AS 입금금액, ISNULL(T1.출금금액, 0) AS 출금금액, " _
                  & "T1.작성자코드 AS 작성자코드, " _
                  & "ISNULL(T3.사용자명, '') AS 작성자명, T1.사용구분 AS 사용구분, " _
                  & "T1.수정일자 AS 수정일자, T1.사용자코드 AS 사용자코드, " _
                  & "T4.사용자명 AS 사용자명 " _
             & "FROM 회계전표내역 T1 " _
            & "INNER JOIN 계정과목 T2 " _
                    & "ON T2.계정코드 = T1.계정코드 " _
             & "LEFT JOIN 사용자 T3 " _
                    & "ON T3.사용자코드 = T1.작성자코드 " _
             & "LEFT JOIN 사용자 T4 " _
                    & "ON T4.사용자코드 = T1.사용자코드 "
    strSQL = strSQL _
           & "" & strWhere & " " _
           & "" & strOrderBy & " "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    vsfg1.Rows = P_adoRec.RecordCount + 1
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       cmdExit.SetFocus
    Else
       With vsfg1
            .Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlack
            If .Rows <= PC_intRowCnt Then
               '.ScrollBars = flexScrollBarHorizontal
            Else
               '.ScrollBars = flexScrollBarBoth
            End If
            Do Until P_adoRec.EOF
               lngR = lngR + 1
               .TextMatrix(lngR, 0) = P_adoRec("사업장코드") & P_adoRec("작성일자") _
                                    & P_adoRec("작성시간") & P_adoRec("계정코드")
               .Cell(flexcpData, lngR, 0, lngR, 0) = Trim(.TextMatrix(lngR, 0)) 'FindRow 사용을 위해
               .TextMatrix(lngR, 1) = Format(P_adoRec("작성일자"), "0000-00-00")
               .TextMatrix(lngR, 2) = Format(Mid(P_adoRec("작성시간"), 1, 6), "00:00:00")
               .TextMatrix(lngR, 3) = P_adoRec("계정코드")
               .TextMatrix(lngR, 4) = P_adoRec("계정명")
               .TextMatrix(lngR, 5) = P_adoRec("적요")
               .TextMatrix(lngR, 6) = P_adoRec("입금금액")
               .TextMatrix(lngR, 7) = P_adoRec("출금금액")
               .TextMatrix(lngR, 8) = P_adoRec("작성자코드")
               .TextMatrix(lngR, 9) = P_adoRec("작성자명")
               If P_adoRec("사용구분") = 0 Then
                  .TextMatrix(lngR, 10) = "정    상"
               ElseIf _
                  P_adoRec("사용구분") = 9 Then
                  .TextMatrix(lngR, 10) = "삭    제"
               Else
                  .TextMatrix(lngR, 10) = "코드오류"
               End If
               .TextMatrix(lngR, 11) = Format(P_adoRec("수정일자"), "0000-00-00")
               .TextMatrix(lngR, 12) = P_adoRec("사용자코드")
               .TextMatrix(lngR, 13) = P_adoRec("사용자명")
               .TextMatrix(lngR, 14) = P_adoRec("작성시간")
               lblTotIn.Caption = Format(Vals(lblTotIn.Caption) + .ValueMatrix(lngR, 6), "#,0")
               lblTotOut.Caption = Format(Vals(lblTotOut.Caption) + .ValueMatrix(lngR, 7), "#,0")
               'If .TextMatrix(lngR, 0) = PB_regUserinfoU.UserBranchCode Then
               '   lngRR = lngR
               'End If
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
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "회계전표내역 읽기 실패"
    Unload Me
    Exit Sub
End Sub

Private Sub SubOther_FILL()
Dim strSQL        As String
Dim lngI          As Long
Dim lngJ          As Long
Dim intIndex      As Integer
    P_adoRec.CursorLocation = adUseClient
    With cboIOGbn
         .AddItem "1. 입금"
         .AddItem "2. 출금"
         .ListIndex = 0
    End With
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "계정코드 읽기 실패"
    Unload Me
    Exit Sub
End Sub

'+-------------------------+
'/// Clear text1(index) ///
'+-------------------------+
Private Sub SubClearText()
Dim lngC As Long
    txtAccCode.Text = "": txtAccName.Text = "": txtJukyo.Text = "": txtMoney.Text = ""
End Sub

'+-------------------------+
'/// Check text1(index) ///
'+-------------------------+
Private Function FncCheckTextBox(blnOK As Boolean)
    txtAccCode.Text = Trim(txtAccCode.Text) '계정과목
    If Not (Vals(txtAccCode.Text) > 0) Then
       txtAccCode.SetFocus
       Exit Function
    End If
    txtMoney.Text = Trim(txtMoney.Text) '적요
    If Not (LenH(txtJukyo.Text) <= 60) Then
       txtJukyo.SetFocus
       Exit Function
    End If
    txtMoney.Text = Trim(txtMoney.Text) '입출금액
    If Not (Vals(txtMoney.Text) <> 0) Then
       txtMoney.SetFocus
       Exit Function
    End If
    blnOK = True
End Function

'+---------------------------+
'/// 크리스탈 리포터 출력 ///
'+---------------------------+
Private Sub cmdPrint_Click()
    '
End Sub


