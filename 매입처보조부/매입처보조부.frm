VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm매입처보조부 
   BorderStyle     =   0  '없음
   Caption         =   "매입처보조부"
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
      TabIndex        =   9
      Top             =   0
      Width           =   15195
      Begin VB.OptionButton optPrtChk0 
         Caption         =   "코드순"
         Height          =   255
         Left            =   6840
         TabIndex        =   23
         Top             =   150
         Width           =   975
      End
      Begin VB.OptionButton optPrtChk1 
         Caption         =   "이름순"
         Height          =   255
         Left            =   6840
         TabIndex        =   22
         Top             =   390
         Value           =   -1  'True
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
         Picture         =   "매입처보조부.frx":0000
         Style           =   1  '그래픽
         TabIndex        =   16
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   9120
         Picture         =   "매입처보조부.frx":0963
         Style           =   1  '그래픽
         TabIndex        =   14
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   13635
         Picture         =   "매입처보조부.frx":1308
         Style           =   1  '그래픽
         TabIndex        =   13
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   390
         Left            =   12510
         Picture         =   "매입처보조부.frx":1C56
         Style           =   1  '그래픽
         TabIndex        =   12
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   11385
         Picture         =   "매입처보조부.frx":25DA
         Style           =   1  '그래픽
         TabIndex        =   11
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Height          =   390
         Left            =   10260
         Picture         =   "매입처보조부.frx":2E61
         Style           =   1  '그래픽
         TabIndex        =   4
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00008000&
         BorderStyle     =   1  '단일 고정
         Caption         =   "매 입 처 별 매 입 현 황"
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
         TabIndex        =   10
         Top             =   180
         Width           =   4650
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid vsfg1 
      Height          =   8415
      Left            =   60
      TabIndex        =   5
      Top             =   1649
      Width           =   15195
      _cx             =   26802
      _cy             =   14843
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
      Height          =   970
      Left            =   60
      TabIndex        =   6
      Top             =   630
      Width           =   15195
      Begin VB.CheckBox chkTotal 
         Caption         =   "전체 매입처"
         Height          =   255
         Left            =   5030
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   10  '한글 
         Index           =   1
         Left            =   2475
         MaxLength       =   50
         TabIndex        =   1
         Top             =   585
         Width           =   3855
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
         Top             =   225
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpF_Date 
         Height          =   270
         Left            =   10440
         TabIndex        =   2
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   56557569
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin MSComCtl2.DTPicker dtpT_Date 
         Height          =   270
         Left            =   12480
         TabIndex        =   3
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   56557569
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "까지"
         Height          =   240
         Index           =   12
         Left            =   13800
         TabIndex        =   20
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "부터"
         Height          =   240
         Index           =   11
         Left            =   11760
         TabIndex        =   19
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "기준일자"
         Height          =   240
         Index           =   10
         Left            =   9360
         TabIndex        =   18
         Top             =   285
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
         TabIndex        =   17
         Top             =   285
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "[Home]"
         Height          =   240
         Index           =   15
         Left            =   4080
         TabIndex        =   15
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "매입처명"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   1275
         TabIndex        =   8
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "매입처코드"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   1275
         TabIndex        =   7
         Top             =   285
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frm매입처보조부"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' 프로그램 제 목 : 매입처보조부
' 사용된 Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' 참조된 Table   :
' 업  무  설  명 :
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived       As Boolean
Private P_adoRec           As New ADODB.Recordset
Private P_intButton        As Integer
Private Const PC_intRowCnt As Integer = 25  '그리드 한 페이지 당 행수(FixedRows 포함)

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
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 50 '삭제, 추가, 저장, 조회, 인쇄
                   cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 99 '삭제, 추가, 저장, 조회, 인쇄
                   cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Else
                   cmdClear.Enabled = False: cmdFind.Enabled = False: cmdSave.Enabled = False
                   cmdDelete.Enabled = False: cmdExit.Enabled = True
       End Select
       dtpF_Date.Value = Format(Mid(PB_regUserinfoU.UserClientDate, 1, 6) & "01", "0000-00-00")
       dtpT_Date.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
       'dtpT_Date.Value = DateAdd("d", -1, DateAdd("m", 1, dtpF_Date.Value))
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
Private Sub dtpF_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub dtpT_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdFind.SetFocus
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
          dtpF_Date.SetFocus
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
                Case Else
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
End Sub
'+-----------+
'/// 저장 ///
'+-----------+
Private Sub cmdSave_Click()
    '
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
    Set frm매입처보조부 = Nothing
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
    With vsfg1              'Rows 1, Cols 17, RowHeightMax(Min) 300
         .AllowBigSelection = False
         .AllowSelection = False
         .AllowUserResizing = flexResizeColumns
         .BackColorBkg = &H8000000F
         .BackColorSel = &H8000&
         .ExtendLastCol = True
         .FocusRect = flexFocusHeavy
         .ScrollBars = flexScrollBarHorizontal
         .ScrollTrack = True
         .SelectionMode = flexSelectionByRow
         .FixedRows = 1
         .FixedCols = 5
         '.FrozenCols = 5
         .Rows = 1             'Subvsfg1_Fill수행시에 설정
         .Cols = 17
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1000   '사업장코드
         .ColWidth(1) = 1000   '매입처코드
         .ColWidth(2) = 2000   '매입처명
         .ColWidth(3) = 1100   '일자
         .ColWidth(4) = 2500   '적요
         .ColWidth(5) = 1900   '자재코드(분류+세부) 'H
         .ColWidth(6) = 2200   '자재명
         .ColWidth(7) = 2100   '규격
         .ColWidth(8) = 700    '단위
         .ColWidth(9) = 1000   '입출고구분(구분)
         .ColWidth(10) = 500   '입출고구분명(구분)
         .ColWidth(11) = 1000  '입고수량(수량)
         .ColWidth(12) = 1400  '입고단가(공급가)
         .ColWidth(13) = 1700  '입고금액(부가세미포함)
         .ColWidth(14) = 1600  '입고부가세
         .ColWidth(15) = 1700  '입고금액(합계)
         .ColWidth(16) = 2000  '비고
         
         .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "사업장코드"  'H
         .TextMatrix(0, 1) = "매입처코드"  'H
         .TextMatrix(0, 2) = "매입처명"    'H
         .TextMatrix(0, 3) = "날짜"
         .TextMatrix(0, 4) = "적요"
         .TextMatrix(0, 5) = "품목코드"    'H
         .TextMatrix(0, 6) = "품명"
         .TextMatrix(0, 7) = "규격"
         .TextMatrix(0, 8) = "단위"
         .TextMatrix(0, 9) = "구분"        'H
         .TextMatrix(0, 10) = "구분"
         .TextMatrix(0, 11) = "수량"
         .TextMatrix(0, 12) = "매입단가"   '품목단가
         .TextMatrix(0, 13) = "매입금액"   '수량 * 단가
         .TextMatrix(0, 14) = "매입부가"      '((수량 * 단가) * (PB_curVatRate + 1)) - (수량 * 단가)
         .TextMatrix(0, 15) = "매입금액(VAT)" '(수량 * 단가) * (PB_curVatRate + 1)
         .TextMatrix(0, 16) = "비고"
         .ColHidden(0) = True: .ColHidden(1) = True: .ColHidden(2) = True: .ColHidden(5) = True: .ColHidden(14) = True
         .ColHidden(9) = True
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 1, 2, 4, 5, 6, 7, 8
                         .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 0, 3, 9, 10, 17
                         .ColAlignment(lngC) = flexAlignCenterCenter
                    Case Else
                         .ColAlignment(lngC) = flexAlignRightCenter
             End Select
         Next lngC
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 11
                         .ColFormat(lngC) = "#,#"
                    Case 12 To 15
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

'+---------------------------------+
'/// VsFlexGrid(vsfg1) 채우기///
'+---------------------------------+
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

Dim curJanMny   As Currency  '업체별 잔액

Dim curMQAmt    As Currency  '업체별월수량
Dim curTQAmt    As Currency  '업체별수량누계
Dim curTTQAmt   As Currency  '전  체수량누계

Dim curMUMny    As Currency  '업체별월매출금액       '(단가 * 수량)
Dim curTUMny    As Currency  '업체별매출금액누계     '(단가 * 수량)
Dim curTTUMny   As Currency  '전  체매출금액누계     '(단가 * 수량)

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
    strOrderBy = "ORDER BY T1.사업장코드, " & IIf(optPrtChk0.Value = True, "T1.매입처코드, T3.매입처명 ", "T3.매입처명, T1.매입처코드 ") & ", 일자, 자재명, 구분 "
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.사업장코드 AS 사업장코드, T2.사업장명 AS 사업장명, " _
                  & "T1.매입처코드 AS 매입처코드, T3.매입처명 AS 매입처명, " _
                  & "(T1.마감년월 + '00') AS 일자, '(년 이 월)' AS 적요, " _
                  & "'' AS 분류코드, '' AS 분류명, " _
                  & "'' AS 세부코드, '' AS 자재명, " _
                  & "'' AS 규격, '' AS 단위, 0 AS 구분, " _
                  & "0 AS 입고수량, 0 AS 입고단가, " _
                  & "0 AS 입고부가, (T1.미지급금누계금액 - T1.미지급금지급누계금액) AS 입고금액 " _
             & "FROM 미지급금원장마감 T1 " _
             & "LEFT JOIN 사업장 T2 ON T2.사업장코드 = T1.사업장코드 " _
             & "LEFT JOIN 매입처 T3 ON T3.사업장코드 = T1.사업장코드 AND T3.매입처코드 = T1.매입처코드 " _
             & "" & strWhere & " AND SUBSTRING(T1.마감년월,5,2) = '00' " _
              & "AND (T1.미지급금누계금액 <> 0 OR T1.미지급금지급누계금액 <> 0) " _
              & "AND T1.마감년월 BETWEEN '" & (Mid(DTOS(dtpF_Date.Value), 1, 4) + "00") & "' " _
                                  & "AND '" & (Mid(DTOS(dtpT_Date.Value), 1, 4) + "00") & "' "
    strSQL = strSQL & "UNION ALL " _
           & "SELECT T1.사업장코드 AS 사업장코드, T2.사업장명 AS 사업장명, " _
                  & "T1.매입처코드 AS 매입처코드, T3.매입처명 AS 매입처명, " _
                  & "(T1.마감년월 + '00') AS 일자, '(월    계)' AS 적요, " _
                  & "'' AS 분류코드, '' AS 분류명, " _
                  & "'' AS 세부코드, '' AS 자재명, " _
                  & "'' AS 규격, '' AS 단위, 0 AS 구분, " _
                  & "0 AS 입고수량, 0 AS 입고단가, " _
                  & "0 AS 입고부가, (T1.미지급금누계금액 - T1.미지급금지급누계금액) AS 입고금액 " _
             & "FROM 미지급금원장마감 T1 " _
             & "LEFT JOIN 사업장 T2 ON T2.사업장코드 = T1.사업장코드 " _
             & "LEFT JOIN 매입처 T3 ON T3.사업장코드 = T1.사업장코드 AND T3.매입처코드 = T1.매입처코드 " _
             & "" & strWhere & " AND SUBSTRING(T1.마감년월,5,2) <> '00' " _
              & "AND (T1.미지급금누계금액 <> 0 OR T1.미지급금지급누계금액 <> 0) " _
              & "AND T1.마감년월 > '" & (Mid(DTOS(dtpF_Date.Value), 1, 4) + "00") & "' " _
              & "AND T1.마감년월 < '" & Mid(DTOS(dtpF_Date.Value), 1, 6) & "' "
    strSQL = strSQL & "UNION ALL " _
           & "SELECT T1.사업장코드 AS 사업장코드, T2.사업장명 AS 사업장명, " _
                  & "T1.매입처코드 AS 매입처코드, T3.매입처명 AS 매입처명, " _
                  & "(T1.입출고일자) AS 일자, (T1.원래입출고일자) AS 적요, " _
                  & "T1.분류코드 AS 분류코드, T4.분류명 AS 분류명, " _
                  & "T1.세부코드 AS 세부코드, ISNULL(T5.자재명,'ERROR!') AS 자재명, " _
                  & "ISNULL(T5.규격,'') AS 규격, ISNULL(T5.단위,'') AS 단위, T1.입출고구분 AS 구분, " _
                  & "SUM(T1.입고수량) AS 입고수량, T1.입고단가 AS 입고단가, " _
                  & "T1.입고부가 AS 입고부가, (SUM(T1.입고수량*T1.입고단가) * " & (PB_curVatRate + 1) & ") AS 입고금액 " _
             & "FROM 자재입출내역 T1 " _
             & "LEFT JOIN 사업장 T2 ON T2.사업장코드 = T1.사업장코드 " _
             & "LEFT JOIN 매입처 T3 ON T3.사업장코드 = T1.사업장코드 AND T3.매입처코드 = T1.매입처코드 " _
             & "LEFT JOIN 자재분류 T4 ON T4.분류코드 = T1.분류코드 " _
             & "LEFT JOIN 자재 T5 ON T5.분류코드 = T1.분류코드 AND T5.세부코드 = T1.세부코드 " _
             & "" & strWhere & " AND T1.입출고구분 = 1 AND T1.사용구분 = 0 " _
              & "AND T1.입출고일자 BETWEEN '" & (Mid(DTOS(dtpF_Date.Value), 1, 6) + "01") & "' AND '" & DTOS(dtpT_Date.Value) & "' "
    strSQL = strSQL _
           & "GROUP BY T1.사업장코드, T2.사업장명, T1.매입처코드, T3.매입처명, " _
                    & "T1.입출고일자, T1.원래입출고일자, " _
                    & "T1.분류코드, T4.분류명, T1.세부코드, T5.자재명, " _
                    & "T5.규격, T5.단위, T1.입출고구분, T1.입고단가, T1.입고부가 "
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
                         .TextMatrix(lngR, 11) = curMQAmt         '월수량
                         .TextMatrix(lngR, 13) = curMUMny         '월매입금액
                         .TextMatrix(lngR, 15) = (curMUMny * (PB_curVatRate + 1)) '월매입금액(VAT)
                         .TextMatrix(lngR, 14) = .ValueMatrix(lngR, 15) - .ValueMatrix(lngR, 13)  '월부가세금액
                         curMQAmt = 0: curMUMny = 0
                         lngR = lngR + 1
                     End If
                     '해당입(출)고누계
                     .TextMatrix(lngR, 4) = "(기간누계)"
                     .TextMatrix(lngR, 11) = curTQAmt             '업체수량누계
                     .TextMatrix(lngR, 13) = curTUMny             '업체매입금액누계
                     .TextMatrix(lngR, 15) = (curTUMny * (PB_curVatRate + 1))     '업체매입금액누계(VAT)
                     .TextMatrix(lngR, 14) = .ValueMatrix(lngR, 15) - .ValueMatrix(lngR, 13)  '업체부가세금액누계
                     curTQAmt = 0: curTUMny = 0
                     curJanMny = 0 '년(월)이월 잔액
                     .AddItem ""
                     lngR = lngR + 1
                     .AddItem ""
                     lngR = lngR + 1
                     '매입처명
                     .TextMatrix(lngR, 3) = P_adoRec("매입처코드"): .TextMatrix(lngR, 4) = P_adoRec("매입처명")
                     .Cell(flexcpBackColor, lngR, 0, lngR, .Cols - 1) = vbYellow
                     .AddItem ""
                     lngR = lngR + 1
                  Else                                                          '매입처코드 같으면
                     If .TextMatrix(lngR - 1, 3) <> "" And _
                         Mid(StrDate, 1, 6) <> Mid(P_adoRec("일자"), 1, 6) Then '일 집계면 And 월이 다르면
                         '해당월입(출)고누계
                         .AddItem ""
                         '.TextMatrix(lngR, 4) = "(" & Format(Mid(StrDate, 1, 6), "0000-00") & Space(1) & "월계)"
                         .TextMatrix(lngR, 4) = "(월    계)"
                         .TextMatrix(lngR, 11) = curMQAmt         '월수량
                         .TextMatrix(lngR, 13) = curMUMny         '월매입금액
                         .TextMatrix(lngR, 15) = (curMUMny * (PB_curVatRate + 1)) '월매입금액(VAT)
                         .TextMatrix(lngR, 14) = .ValueMatrix(lngR, 15) - .ValueMatrix(lngR, 13)  '월부가세금액
                         curMQAmt = 0: curMUMny = 0
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
                     .TextMatrix(lngR, 4) = "(" & Mid(P_adoRec("일자"), 1, 4) & "년 이월잔액)"
                  Else
                     .TextMatrix(lngR, 4) = "(" & Format(Mid(P_adoRec("일자"), 1, 6), "0000-00") & "월 잔액)"
                  End If
               End If
               If Len(P_adoRec("분류코드")) > 0 Then
                  .TextMatrix(lngR, 4) = P_adoRec("분류코드") & P_adoRec("세부코드")
               End If
               .TextMatrix(lngR, 5) = P_adoRec("분류코드") & P_adoRec("세부코드")
               .TextMatrix(lngR, 6) = P_adoRec("자재명")
               .TextMatrix(lngR, 7) = P_adoRec("규격")
               .TextMatrix(lngR, 8) = P_adoRec("단위")
               .TextMatrix(lngR, 9) = P_adoRec("구분")
               '10. 출고적요
               If P_adoRec("구분") = 0 Then
                  .TextMatrix(lngR, 10) = ""
               ElseIf _
                  P_adoRec("구분") = 1 Then
                  .TextMatrix(lngR, 10) = "매입"
               ElseIf _
                  P_adoRec("구분") = 7 Then
                  .TextMatrix(lngR, 10) = "임의"
               End If
               .TextMatrix(lngR, 11) = P_adoRec("입고수량")                          '수량
               .TextMatrix(lngR, 12) = P_adoRec("입고단가")                          '품목별 매출단가
               .TextMatrix(lngR, 13) = P_adoRec("입고수량") * P_adoRec("입고단가")   '해당행의 매입금액
               .TextMatrix(lngR, 14) = P_adoRec("입고금액") - .ValueMatrix(lngR, 13) '해당행의 매입부가세
                If Mid(P_adoRec("일자"), 7, 2) = "00" Then
                   curJanMny = curJanMny + P_adoRec("입고금액") '년(월)이월 누계
                   .TextMatrix(lngR, 15) = curJanMny
               Else
                  .TextMatrix(lngR, 15) = P_adoRec("입고금액")                       '해당행의 매입금액(부가세포함)
               End If
               If Mid(P_adoRec("일자"), 7, 2) <> "00" Then
                  curMQAmt = curMQAmt + P_adoRec("입고수량")
                  curMUMny = curMUMny + (P_adoRec("입고수량") * P_adoRec("입고단가"))
               End If
               '해당누계금액
               curTQAmt = curTQAmt + P_adoRec("입고수량")
               curTUMny = curTUMny + (P_adoRec("입고단가") * P_adoRec("입고수량"))
               curTTQAmt = curTTQAmt + P_adoRec("입고수량")
               curTTUMny = curTTUMny + (P_adoRec("입고단가") * P_adoRec("입고수량"))
               StrDate = P_adoRec("일자")
               'FindRow 사용을 위해
               '.TextMatrix(lngR, 4) = .TextMatrix(lngR, 0) & P_adoRec("세분류코드")
               '.Cell(flexcpData, lngR, 4, lngR, 4) = .TextMatrix(lngR, 4)
               P_adoRec.MoveNext
               If P_adoRec.EOF = True Then '마지막 레코드면
                  If .TextMatrix(lngR, 3) <> "" Then  '일 집계면
                     lngR = lngR + 1
                     '해당월누계
                     .AddItem ""
                     .TextMatrix(lngR, 4) = "(월    계)"
                     .TextMatrix(lngR, 11) = curMQAmt         '월수량
                     .TextMatrix(lngR, 13) = curMUMny         '월매입금액
                     .TextMatrix(lngR, 15) = (curMUMny * (PB_curVatRate + 1)) '월매입금액(VAT)
                     .TextMatrix(lngR, 14) = .ValueMatrix(lngR, 15) - .ValueMatrix(lngR, 13)  '월부가세금액
                     curMQAmt = 0: curMUMny = 0
                  End If
                  '해당누계계
                  lngR = lngR + 1
                  .AddItem ""
                  .TextMatrix(lngR, 4) = "(기간누계)"
                  .TextMatrix(lngR, 11) = curTQAmt            '업체수량누계
                  .TextMatrix(lngR, 13) = curTUMny            '업체매입금액누계
                  .TextMatrix(lngR, 15) = (curTUMny * (PB_curVatRate + 1))    '업체매입금액누계(VAT)
                  .TextMatrix(lngR, 14) = .ValueMatrix(lngR, 15) - .ValueMatrix(lngR, 13)  '업체부가세금액누계
                  curTQAmt = 0: curTUMny = 0
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
               .TextMatrix(lngR, 11) = curTTQAmt              '전체수량누계
               .TextMatrix(lngR, 13) = curTTUMny              '전체매입금액누계
               .TextMatrix(lngR, 15) = (curTTUMny * (PB_curVatRate + 1))      '전체매입금액누계(VAT)
               .TextMatrix(lngR, 14) = .ValueMatrix(lngR, 15) - .ValueMatrix(lngR, 13)  '전체부가세금액누계
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
            .TopRow = 1
            vsfg1_EnterCell       'vsfg1_EnterCell 자동실행(만약 한건 일때도 강제로 자동실행)
            .SetFocus
       End With
    End If
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "매입처보조부 읽기 실패"
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
    
    If Len(Trim(Text1(0).Text)) = 0 And (chkTotal.Value = 0) Then
       Exit Sub
    End If
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
            strExeFile = App.Path & ".\매입처보조부.rpt"
         Else
            strExeFile = App.Path & ".\매입처보조부T.rpt"
         End If
         varRetVal = Dir(strExeFile)
         If Len(varRetVal) = 0 Then
            intRetCHK = 0
         Else
            .ReportFileName = strExeFile
            On Error GoTo ERROR_CRYSTAL_REPORTS
            '--- Formula Fields ---
            .Formulas(0) = "ForAppPgDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'" '프로그램실행일자
            .Formulas(1) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                       '사업장명
            .Formulas(2) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                                  '출력일시
            '적용기준일자
            .Formulas(3) = "ForAppDate = '기준일자 : ' & '" & Format(DTOS(dtpF_Date.Value), "0000-00-00") & "' & ' 부터 ' & '" & Format(DTOS(dtpT_Date.Value), "0000-00-00") & "' & ' 까지' "
            '--- ParaMeter Fields ---
            '프로그램실행일자
            .StoredProcParam(0) = PB_regUserinfoU.UserClientDate
            '적용(기준)일자시작
            .StoredProcParam(1) = DTOS(dtpF_Date.Value)
            '적용(기준)일자종료
            .StoredProcParam(2) = DTOS(dtpT_Date.Value)
            '매입처코드
            If (Len(Text1(0).Text) = 0) Or (chkTotal.Value = 1) Then
               .StoredProcParam(3) = " "
            Else
               .StoredProcParam(3) = Trim(Text1(0).Text)
            End If
            .StoredProcParam(4) = PB_regUserinfoU.UserBranchCode '지점코드
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
            .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "매입처보조부"
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

