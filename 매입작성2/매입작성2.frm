VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm매입작성2 
   BorderStyle     =   0  '없음
   Caption         =   "매입작성2"
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
      TabIndex        =   6
      Top             =   0
      Width           =   15195
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   4320
         Top             =   195
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   375
         Left            =   7980
         Picture         =   "매입작성2.frx":0000
         Style           =   1  '그래픽
         TabIndex        =   13
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   9120
         Picture         =   "매입작성2.frx":0963
         Style           =   1  '그래픽
         TabIndex        =   12
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   13635
         Picture         =   "매입작성2.frx":1308
         Style           =   1  '그래픽
         TabIndex        =   11
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   390
         Left            =   12510
         Picture         =   "매입작성2.frx":1C56
         Style           =   1  '그래픽
         TabIndex        =   10
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   11385
         Picture         =   "매입작성2.frx":25DA
         Style           =   1  '그래픽
         TabIndex        =   9
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Height          =   390
         Left            =   10260
         Picture         =   "매입작성2.frx":2E61
         Style           =   1  '그래픽
         TabIndex        =   8
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00008000&
         BorderStyle     =   1  '단일 고정
         Caption         =   "매 입 전 표 입 력"
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
         TabIndex        =   7
         Top             =   180
         Width           =   4650
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid vsfg1 
      Height          =   8441
      Left            =   60
      TabIndex        =   2
      Top             =   1644
      Width           =   15195
      _cx             =   26802
      _cy             =   14889
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
      Rows            =   100
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
      Height          =   1035
      Left            =   60
      TabIndex        =   3
      Top             =   600
      Width           =   15195
      Begin VB.TextBox txtTelNo 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   10  '한글 
         Left            =   7920
         MaxLength       =   50
         TabIndex        =   21
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   10  '한글 
         Left            =   1275
         MaxLength       =   50
         TabIndex        =   20
         Top             =   600
         Width           =   5535
      End
      Begin MSComCtl2.DTPicker dtpJ_Date 
         Height          =   270
         Left            =   7920
         TabIndex        =   17
         ToolTipText     =   "현재 작업(매입)일자를 변경"
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         CalendarBackColor=   16777215
         Format          =   19857409
         CurrentDate     =   38301
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   10  '한글 
         Index           =   1
         Left            =   4035
         MaxLength       =   50
         TabIndex        =   1
         Top             =   225
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   0
         Left            =   1275
         MaxLength       =   8
         TabIndex        =   0
         Top             =   225
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "주소"
         Height          =   240
         Index           =   5
         Left            =   75
         TabIndex        =   23
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "전화번호"
         Height          =   240
         Index           =   4
         Left            =   7000
         TabIndex        =   22
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "전체금액"
         Height          =   240
         Index           =   2
         Left            =   9480
         TabIndex        =   19
         Top             =   285
         Width           =   735
      End
      Begin VB.Label lblJDate 
         Caption         =   "( 매입일자                )"
         Height          =   240
         Left            =   6960
         TabIndex        =   18
         ToolTipText     =   "현재 작업(매입)일자를 변경"
         Top             =   285
         Width           =   3135
      End
      Begin VB.Label lblTotMny 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "0.00"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   10560
         TabIndex        =   16
         Top             =   285
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "전체금액"
         Height          =   240
         Index           =   3
         Left            =   6960
         TabIndex        =   15
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "[Home]"
         Height          =   240
         Index           =   15
         Left            =   2520
         TabIndex        =   14
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "매입처명"
         Height          =   240
         Index           =   1
         Left            =   3075
         TabIndex        =   5
         Top             =   285
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "매입처코드"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   75
         TabIndex        =   4
         Top             =   285
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frm매입작성2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' 프로그램 제 목 : 매입작성2
' 사용된 Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' 참조된 Table   : 매입처, 매출처, 자재입출내역
' 업  무  설  명 : 발주없이 바로 매입
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived         As Boolean
Private P_adoRec             As New ADODB.Recordset
Private P_adoRecW            As New ADODB.Recordset
Private P_intButton          As Integer
Private P_strFindString2     As String
Private Const PC_intRowCnt   As Integer = 27  '그리드1의 한 페이지 당 행수(FixedRows 포함)

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
       Subvsfg1_INIT  '매입내역
       Select Case Val(PB_regUserinfoU.UserAuthority)
              Case Is <= 10 '조회
                   cmdPrint.Enabled = False: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 20 '인쇄, 조회
                   cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 40 '추가, 저장, 인쇄, 조회
                   cmdPrint.Enabled = True: cmdClear.Enabled = True: cmdFind.Enabled = True
                   cmdSave.Enabled = True: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 50 '삭제, 추가, 저장, 조회, 인쇄
                   cmdPrint.Enabled = True: cmdClear.Enabled = True: cmdFind.Enabled = True
                   cmdSave.Enabled = True: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 99 '삭제, 추가, 저장, 조회, 인쇄
                   cmdPrint.Enabled = True: cmdClear.Enabled = True: cmdFind.Enabled = True
                   cmdSave.Enabled = True: cmdDelete.Enabled = True: cmdExit.Enabled = True
              Case Else
                   cmdClear.Enabled = False: cmdFind.Enabled = False: cmdSave.Enabled = False
                   cmdDelete.Enabled = False: cmdExit.Enabled = True
       End Select
       '공통(제한)권한
       cmdFind.Enabled = False: cmdDelete.Enabled = False
       SubOther_FILL
       P_blnActived = True
       Screen.MousePointer = vbDefault
    End If
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "매입작성2(서버와의 연결 실패)"
    Unload Me
    Exit Sub
End Sub

'+--------------------+
'/// OtherControls ///
'+--------------------+
Private Sub dtpHopeDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    End If
End Sub
Private Sub dtpSactionDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    End If
End Sub
Private Sub cboSactionWay_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
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
       PB_strFMCCallFormName = "frm매입작성2"
       PB_strSupplierCode = UPPER(Trim(Text1(Index).Text))
       PB_strSupplierName = "" 'Trim(Text1(Index + 1).Text)
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

'+-------------+
'/// 매입처 ///
'+-------------+
Private Sub Text1_LostFocus(Index As Integer)
Dim strSQL As String
Dim lngR   As Long
    With Text1(Index)
         Select Case Index
                Case 0
                     .Text = UPPER(Trim(.Text))
                     If Len(.Text) < 1 Then
                        Text1(1).Text = "": txtAddress.Text = "": txtTelNo.Text = ""
                     End If
         End Select
    End With
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "TABLE 읽기 실패"
    Unload Me
    Exit Sub
End Sub

Private Sub chkCash_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
    Exit Sub
End Sub

'+-------------------+
'/// 작업일자선택 ///
'+-------------------+
Private Sub dtpJ_Date_Change()
    If PB_regUserinfoU.UserClientDate = DTOS(dtpJ_Date.Value) Then
       lblJDate.ForeColor = vbBlack
       With dtpJ_Date
            .CalendarBackColor = vbWhite
            .CalendarForeColor = vbBlack
       End With
    Else
       lblJDate.ForeColor = vbRed
       With dtpJ_Date
            .CalendarBackColor = vbRed
            .CalendarForeColor = vbWhite
       End With
    End If
End Sub
Private Sub dtpJ_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       vsfg1.SetFocus
    End If
End Sub

'+-----------+
'/// Grid ///
'+-----------+
Private Sub vsfg1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    With vsfg1
         P_intButton = Button
         If .MouseRow >= .FixedRows Then
            'If Button = vbLeftButton Then
             '  .Select .MouseRow, .MouseCol
             '  .EditCell
            'End If
         End If
    End With
End Sub
Private Sub vsfg1_Click()
Dim strData As String
Dim lngR1   As Long
Dim lngRH1  As Long
Dim lngR2   As Long
Dim lngRR2  As Long
    With vsfg1
         If (.MouseRow >= 0 And .MouseRow < .FixedRows) Then
            .Col = .MouseCol
            '.Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlack
            '.Cell(flexcpForeColor, 0, .MouseCol, 0, .MouseCol) = vbRed
            'strData = Trim(.Cell(flexcpData, .Row, 0))
            'Select Case .MouseCol
            '       Case 0
            '            .ColSel = 0
            '            .ColSort(0) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .Sort = flexSortUseColSort
            '       Case 1
            '            .ColSel = 1
            '            .ColSort(0) = flexSortNone
            '            .ColSort(1) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .Sort = flexSortUseColSort
            '       Case Else
            '            .Sort = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            'End Select
            'If .FindRow(strData, , 0) > 0 Then
            '   .Row = .FindRow(strData, , 0)
            'End If
            'If PC_intRowCnt < .Rows Then
            '   .TopRow = .Row
            'End If
         End If
    End With
End Sub
Private Sub vsfg1_DblClick()
    With vsfg1
         If .Row >= .FixedRows Then
             vsfg1_KeyDown vbKeyF1, 0  '자재시세검색 OR 매출처검색으로 바로 감.
         End If
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
Private Sub vsfg1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    With vsfg1
         If .Row >= .FixedRows Then
            If Len(.TextMatrix(.Row, 0)) <> 0 Then '0.자재코드, 3.발주량, 7.직송
               If (.Col = 3) Then   '3. 발주량
                  If Button = vbLeftButton Then
                     .Select .Row, .Col
                     .EditCell
                  End If
               ElseIf _
                  (.Col = 7) Then   '7. 직송
                  If Not (.ValueMatrix(.Row, 3) = 0 Or Len(.TextMatrix(.Row, 5)) = 0) Then
                     If Button = vbLeftButton Then
                        .Select .Row, .Col
                        .EditCell
                     End If
                  End If
               ElseIf _
                  (.Col = 8 Or .Col = 9) Then   '8.입고단가, 9.입고부가
                  If Button = vbLeftButton Then
                     .Select .Row, .Col
                     .EditCell
                  End If
               ElseIf _
                  (.Col = 16) Then   '16.적요
                  If Button = vbLeftButton Then
                     .Select .Row, .Col
                     .EditCell
                  End If
               End If
            End If
         End If
    End With
End Sub
Sub vsfg1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim blnModify As Boolean
Dim curTmpMny As Currency
    With vsfg1
         If Row >= .FixedRows Then
            If Len(.TextMatrix(Row, 0)) <> 0 Then
               If (Col = 3) Then         '수량
                  If .TextMatrix(Row, Col) <> .EditText Then
                     If (IsNumeric(.EditText) = False Or _
                        Fix(Vals(.EditText)) < Vals(.EditText) Or IsNumeric(Right(.EditText, 1)) = False) Then
                        Beep
                        Cancel = True
                     Else
                        blnModify = True
                        curTmpMny = .ValueMatrix(Row, 11)
                        .TextMatrix(Row, 11) = Vals(.EditText) * .ValueMatrix(Row, 8)   '합계금액 = 수량 * 단가
                     End If
                  End If
                  If Vals(.EditText) = 0 Or Len(.TextMatrix(Row, 5)) = 0 Then
                     .Cell(flexcpChecked, Row, 7, Row, 7) = flexUnchecked
                  End If
               ElseIf _
                  (Col = 8) Then '8.입고단가
                  If .TextMatrix(Row, Col) <> .EditText Then                            '변경인 경우 입력금액 검사
                     If (IsNumeric(.EditText) = False Or Vals(.EditText) < 0 Or _
                        IsNumeric(Right(.EditText, 1)) = False) Then
                        Beep
                        Cancel = True
                     Else
                        blnModify = True
                        .TextMatrix(Row, 9) = Fix(Vals(.EditText) * (PB_curVatRate))
                        curTmpMny = .ValueMatrix(Row, 11)
                        .TextMatrix(Row, 10) = Vals(.EditText) + .ValueMatrix(Row, 9)
                        .TextMatrix(Row, 11) = .ValueMatrix(Row, 3) * Vals(.EditText)
                     End If
                  End If
               ElseIf _
                  (Col = 9) Then '9.입고부가
                  If .TextMatrix(Row, Col) <> .EditText Then                            '변경인 경우 입력금액 검사
                     If (IsNumeric(.EditText) = False Or Vals(.EditText) < 0 Or _
                        Fix(Vals(.EditText)) < Vals(.EditText) Or IsNumeric(Right(.EditText, 1)) = False) Or _
                        (Vals(.EditText) > .ValueMatrix(Row, 8)) Then
                        Beep
                        Cancel = True
                     Else
                        blnModify = True
                        curTmpMny = .ValueMatrix(Row, 11)
                        .TextMatrix(Row, 10) = .ValueMatrix(Row, 8) + Vals(.EditText)
                        .TextMatrix(Row, 11) = .ValueMatrix(Row, 3) * .ValueMatrix(Row, 8)
                     End If
                  End If
               ElseIf _
                  (Col = 16) Then '적요 길이 검사
                  If .TextMatrix(Row, Col) <> .EditText Then
                     If Not (LenH(Trim(.EditText)) <= 50) Then
                        Beep
                        Cancel = True
                     Else
                        blnModify = True
                     End If
                  End If
               End If
            End If
            '변경표시 + 금액재계산
            If blnModify = True Then
               Select Case Col
                      Case 3, 8, 9
                           lblTotMny.Caption = Format(Vals(lblTotMny.Caption) - curTmpMny + .ValueMatrix(Row, 11), "#,#.00")
                      Case Else
                      
               End Select
            End If
         End If
    End With
End Sub
Private Sub vsfg1_EnterCell()
Dim lngR As Long
Dim lngC As Long
    With vsfg1
         .Editable = flexEDNone
         If .Row >= .FixedRows Then
             Select Case .Col
                    Case 3, 8, 16
                         .Editable = flexEDKbdMouse
                         vsfg1_MouseDown vbLeftButton, 0, 0, 0
             End Select
         End If
    End With
End Sub
Private Sub vsfg1_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    With vsfg1
         If KeyCode = vbKeyReturn Then
            If Col = 3 Then
               .Col = 8
            ElseIf _
               Col = 8 Then
               .Col = 16
            ElseIf _
               Col = 16 And Row < (.Rows - 1) Then
               .Col = 3: .Row = .Row + 1
               If .Row >= PC_intRowCnt Then
                  .TopRow = .TopRow + 1
               End If
            End If
         End If
    End With
End Sub
Private Sub vsfg1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lngR     As Long
Dim blnDupOK As Boolean
Dim intRetVal As Integer
    With vsfg1
         If .Row >= .FixedRows Then     '내역시세검색
            If KeyCode = vbKeyF2 And (Len(Text1(0).Text) > 0) And (Len(.TextMatrix(.Row, 0)) > 0) Then
               PB_strFMCCallFormName = "frm매입작성2"
               PB_strMaterialsCode = .TextMatrix(.Row, 0)
               PB_strMaterialsName = .TextMatrix(.Row, 1)
               If Len(Trim(Text1(0).Text)) = 0 Then
                  PB_strSupplierCode = ""
               Else
                  PB_strSupplierCode = Trim(Text1(0).Text)
               End If
               frm내역시세검색.Show vbModal
               If Len(PB_strMaterialsCode) <> 0 Then
                  PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
               End If
            End If
         End If
    End With
    With vsfg1
         If .Row >= .FixedRows Then
            If KeyCode = vbKeyF1 And Len(Trim(Text1(0).Text)) > 0 Then  '자재시세검색
               PB_strFMCCallFormName = "frm매입작성2"
               PB_strMaterialsCode = .TextMatrix(.Row, 0)
               PB_strMaterialsName = .TextMatrix(.Row, 1)
               If Len(Trim(Text1(0).Text)) = 0 Then
                  PB_strSupplierCode = ""
               Else
                  PB_strSupplierCode = Trim(Text1(0).Text)
               End If
               frm자재시세검색.Show vbModal
               If Len(PB_strMaterialsCode) <> 0 Then
                  PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
               End If
            ElseIf _
               KeyCode = vbKeyDelete And .Col = 6 Then
               .TextMatrix(.Row, 5) = "": .TextMatrix(.Row, 6) = ""
               .Cell(flexcpChecked, .Row, 7, .Row, 7) = flexUnchecked
            ElseIf _
               KeyCode = vbKeyDelete And (.Col <> 6) And (Len(.TextMatrix(.Row, 0)) > 0) Then 'And (.MouseRow > 0) Then
               intRetVal = MsgBox("입력한 매입내역을 삭제하시겠습니까 ?", vbQuestion + vbYesNo + vbDefaultButton2, "매입내역삭제")
               If intRetVal = vbYes Then
                  lblTotMny.Caption = Format(Vals(lblTotMny.Caption) - .ValueMatrix(.Row, 11), "#,#.00") '전체금액에서 제외
                  .Cell(flexcpText, .Row, 0, .Row, .Cols - 1) = ""
               End If
            End If
         End If
    End With
End Sub

'+-----------+
'/// 출력 ///
'+-----------+
Private Sub cmdPrint_Click()
    If (vsfg1.Rows) = 1 Then
       Exit Sub
    End If
End Sub
'+-----------+
'/// 조회 ///
'+-----------+
Private Sub cmdFind_Click()
    Screen.MousePointer = vbHourglass
    cmdFind.Enabled = False
    cmdFind.Enabled = True
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
Dim blnOK          As Boolean
Dim intRetVal      As Integer
Dim lngChkCnt      As Long
Dim lngDelCntS     As Long
Dim lngDelCntE     As Long
Dim strServerTime  As String
Dim strTime        As String
Dim strHH          As String
Dim strMM          As String
Dim strSS          As String
Dim strMS          As String
Dim lngLogCnt      As Long    '로그카운터
Dim intChkCash     As Integer '1.현금매입
Dim strJ_Date      As String  '작업일자
    If (vsfg1.Rows) = 1 Then
       Exit Sub
    End If
    If Len(Text1(0).Text) < 1 Then '매입처코드
       Text1(0).SetFocus
       Exit Sub
    End If
    With vsfg1
         For lngR = 1 To .Rows - 1
             If Len(.TextMatrix(lngR, 0)) > 0 Then 'And .ValueMatrix(lngR, 3) <> 0 Then
                lngChkCnt = lngChkCnt + 1
             End If
         Next lngR
         If lngChkCnt = 0 Then
            Exit Sub
         End If
    End With
    intRetVal = MsgBox("입력된 자료" & lngChkCnt & "(건)을 저장하시겠습니까 ?", vbQuestion + vbYesNo + vbDefaultButton1, "자료 저장")
    If intRetVal = vbNo Then
       vsfg1.SetFocus
       Exit Sub
    End If
    intRetVal = MsgBox("현금매입을 하시겠습니까 ?", vbQuestion + vbYesNoCancel + vbDefaultButton2, "현금매입")
    If intRetVal = vbYes Then
       intChkCash = 1
    ElseIf _
       intRetVal = vbCancel Then
       vsfg1.SetFocus
       Exit Sub
    End If
    '작업일자 구하기
    strJ_Date = DTOS(dtpJ_Date.Value)
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
    strTime = strServerTime
    '거래번호 구하기
    PB_adoCnnSQL.BeginTrans
    P_adoRec.CursorLocation = adUseClient
    strSQL = "spLogCounter '자재입출내역', '" & PB_regUserinfoU.UserBranchCode + strJ_Date + "1" & "', " _
                            & "0, 0, '" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' "
    On Error GoTo ERROR_STORED_PROCEDURE
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    lngLogCnt = P_adoRec(0)
    P_adoRec.Close
    
    lngChkCnt = 0
    With vsfg1
         For lngR = 1 To .Rows - 1
             If Len(.TextMatrix(lngR, 0)) > 0 Then 'And .ValueMatrix(lngR, 3) <> 0 Then
                '자재입출내역
                lngChkCnt = lngChkCnt + 1
                If lngChkCnt = 1 Then
                   strTime = strServerTime
                Else
                   strTime = Format((Val(strTime) + 10000), "000000000")
                   strHH = Mid(strTime, 1, 2): strMM = Mid(strTime, 3, 2): strSS = Mid(strTime, 5, 2): strMS = Mid(strTime, 7, 3)
                   If Val(strMS) > 999 Then
                      strMS = Format(0, "000")
                      strSS = Format(Val(strMM) + 1, "00")
                   End If
                   If Val(strSS) > 59 Then
                      strSS = Format(Val(strSS) - 60, "00")
                      strMM = Format(Val(strMM) + 1, "00")
                   End If
                   If Val(strMM) > 59 Then
                      strMM = Format(Val(strMM) - 60, "00")
                      strHH = Format(Val(strHH) + 1, "00")
                   End If
                   strTime = strHH & strMM & strSS & strMS
                End If
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
                             & "'" & PB_regUserinfoU.UserBranchCode & "', '" & Mid(.TextMatrix(lngR, 0), 1, 2) & "', " _
                             & "'" & Mid(.TextMatrix(lngR, 0), 3) & "', 1, " _
                             & "'" & strJ_Date & "', '" & strTime & "', " _
                             & "" & .ValueMatrix(lngR, 3) & ", " & .ValueMatrix(lngR, 8) & ", " _
                             & "" & .ValueMatrix(lngR, 9) & ", 0, " _
                             & "" & .ValueMatrix(lngR, 12) & ", " & .ValueMatrix(lngR, 13) & ", " _
                             & "'" & Trim(Text1(0).Text) & "' , '" & .TextMatrix(lngR, 5) & "', " _
                             & "'" & strJ_Date & "'," & IIf(.Cell(flexcpChecked, lngR, 7, lngR, 7) = flexUnchecked, 0, 1) & ", " _
                             & "'', 0, " _
                             & "'" & strJ_Date & "' , " & lngLogCnt & ", " _
                             & "0, " & intChkCash & ", 0, '" & .TextMatrix(lngR, 16) & "', '', 0, 0, " _
                             & "0, '" & PB_regUserinfoU.UserServerDate & "', " _
                             & "'" & PB_regUserinfoU.UserCode & "', '' ) "
                On Error GoTo ERROR_TABLE_INSERT
                PB_adoCnnSQL.Execute strSQL
                '최종입출고일자갱신(사업장코드, 분류코드, 세부코드, 입출고구분)
                strSQL = "sp자재최종입출고일자갱신 '" & PB_regUserinfoU.UserBranchCode & "', " _
                       & "'" & Mid(.TextMatrix(lngR, 0), 1, 2) & "', '" & Mid(.TextMatrix(lngR, 0), 3) & "', 1 "
                On Error GoTo ERROR_STORED_PROCEDURE
                PB_adoCnnSQL.Execute strSQL
                '자재최종단가갱신(사업장코드, 분류코드, 세부코드, 입출고구분, 업체코드, 단가, 거래일자)
                If .ValueMatrix(lngR, 8) > 0 And PB_intIAutoPriceGbn = 1 Then
                   strSQL = "sp자재최종단가갱신 '" & PB_regUserinfoU.UserBranchCode & "', " _
                          & "'" & Mid(.TextMatrix(lngR, 0), 1, 2) & "', '" & Mid(.TextMatrix(lngR, 0), 3) & "', 1, " _
                          & "'" & Trim(Text1(0).Text) & "', " _
                          & "" & .ValueMatrix(lngR, 8) & ", '" & strJ_Date & "' "
                   On Error GoTo ERROR_STORED_PROCEDURE
                   PB_adoCnnSQL.Execute strSQL
                End If
             End If
         Next lngR
    End With
    PB_adoCnnSQL.CommitTrans
    SubClearText
    Text1(0).SetFocus
    Screen.MousePointer = vbDefault
    cmdSave.Enabled = True
    'If chkPrint = 1 Then '저장후 출력
       '출력
    'End If
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
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "검색 실패"
    Unload Me
    Exit Sub
ERROR_TABLE_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "추가 실패"
    Unload Me
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "저장 실패"
    Unload Me
    Exit Sub
ERROR_TABLE_DELETE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "삭제 실패"
    Unload Me
    Exit Sub
ERROR_STORED_PROCEDURE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "서버 연결 실패"
    Unload Me
    Exit Sub
End Sub

'+-----------+
'/// 추가 ///
'+-----------+
Private Sub cmdClear_Click()
Dim strSQL As String
    vsfg1.Row = 0
    SubClearText
    Text1(Text1.LBound).Enabled = True
    Text1(Text1.LBound + 1).Enabled = False
    Text1(Text1.LBound).SetFocus
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
    Set frm매입작성2 = Nothing
End Sub

'+--------------------+
'/// Sub Procedure ///
'+--------------------+
Private Sub SubOther_FILL()
Dim strSQL        As String
Dim intIndex      As Integer
    dtpJ_Date.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
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
    txtAddress.Text = "": txtTelNo.Text = ""
    lblTotMny.Caption = "0.00"
    With vsfg1
         .Rows = 1: .Rows = 101
         .Row = 1: .Col = 3
         .TopRow = 1: .LeftCol = 3
         .Cell(flexcpChecked, 1, 7, .Rows - 1, 7) = flexUnchecked
         .Cell(flexcpText, 1, 7, .Rows - 1, 7) = "직 송"
    End With
End Sub
'+----------------------------------+
'/// VsFlexGrid(vsfg1) 초기화 ///
'+----------------------------------+
Private Sub Subvsfg1_INIT()
Dim lngR    As Long
Dim lngC    As Long
    With vsfg1              'Rows 101, Cols 17, RowHeightMax(Min) 300
         .AllowBigSelection = False
         .AllowSelection = False
         .AllowUserResizing = flexResizeColumns
         .BackColorBkg = &H8000000F
         .BackColorSel = &H8000&
         .ExtendLastCol = True
         .FocusRect = flexFocusHeavy
         .ScrollBars = flexScrollBarBoth
         .ScrollTrack = True
         .SelectionMode = flexSelectionByRow
         .FixedRows = 1
         .FixedCols = 3
         .Rows = 101
         .Cols = 17
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1900   '자재코드(분류코드+세부코드)
         .ColWidth(1) = 2500   '자재명
         .ColWidth(2) = 2200   '자재규격
         .ColWidth(3) = 1000   '수량
         .ColWidth(4) = 800    '자재단위
         .ColWidth(5) = 1200   '매출처코드   'H
         .ColWidth(6) = 2500   '매출처명     'H
         .ColWidth(7) = 800    '직송
         .ColWidth(8) = 1500   '입고단가
         .ColWidth(9) = 1300   '입고부가     'H
         .ColWidth(10) = 1500  '입고가격(단가 + 부가)
         .ColWidth(11) = 2000  '입고금액(발주량 * 입고가격)
         .ColWidth(12) = 1500  '출고단가     'H
         .ColWidth(13) = 1300  '출고부가     'H
         .ColWidth(14) = 1500  '출고가격(단가+부가) 'H
         .ColWidth(15) = 2000  '출고금액(발주량 * 출고가격) 'H
         .ColWidth(16) = 4500  '적요
         
         .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "코드"
         .TextMatrix(0, 1) = "품명"
         .TextMatrix(0, 2) = "규격"
         .TextMatrix(0, 3) = "수량"
         .TextMatrix(0, 4) = "단위"          '매입단위
         .TextMatrix(0, 5) = "매출처코드"    'H
         .TextMatrix(0, 6) = "매출처명"      'H
         .TextMatrix(0, 7) = "직송"          'H
         .TextMatrix(0, 8) = "매입단가"
         .TextMatrix(0, 9) = "매입부가"      'H
         .TextMatrix(0, 10) = "매입가격"     'H
         .TextMatrix(0, 11) = "매입금액"
         .TextMatrix(0, 12) = "매출단가"     'H
         .TextMatrix(0, 13) = "매출부가"     'H
         .TextMatrix(0, 14) = "매출단가"     'H
         .TextMatrix(0, 15) = "매출금액"     'H
         .TextMatrix(0, 16) = "적요"
         .ColHidden(5) = True: .ColHidden(6) = True: .ColHidden(7) = True
         .ColHidden(9) = True: .ColHidden(10) = True
         .ColHidden(12) = True: .ColHidden(13) = True: .ColHidden(14) = True: .ColHidden(15) = True
         .ColFormat(3) = "#,#"
         For lngC = 8 To 15
             .ColFormat(lngC) = "#,#.00"
         Next lngC
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 0, 1, 2, 4, 6, 7, 16
                         .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 5
                         .ColAlignment(lngC) = flexAlignCenterCenter
                    Case Else
                         .ColAlignment(lngC) = flexAlignRightCenter
             End Select
         Next lngC
         '.MergeCells = flexMergeRestrictRows 'flexMergeFixedOnly
         '.MergeRow(0) = True
         'For lngC = 0 To 2
         '    .MergeCol(lngC) = True
         'Next lngC
         .Cell(flexcpChecked, 1, 7, .Rows - 1, 7) = flexUnchecked
         .Cell(flexcpText, 1, 7, .Rows - 1, 7) = "직 송"
         .Cell(flexcpAlignment, 1, 7, .Rows - 1, 7) = flexAlignLeftCenter
         
         vsfg1_EnterCell
    End With
End Sub

