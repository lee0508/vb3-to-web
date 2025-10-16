VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm매입작성1 
   BorderStyle     =   0  '없음
   Caption         =   "발주서관리"
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
   Begin VSFlex7Ctl.VSFlexGrid vsfg2 
      Height          =   5727
      Left            =   60
      TabIndex        =   9
      Top             =   4338
      Width           =   15195
      _cx             =   26802
      _cy             =   10107
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
   Begin VB.Frame Frame1 
      Appearance      =   0  '평면
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   60
      TabIndex        =   15
      Top             =   0
      Width           =   15195
      Begin VB.OptionButton optPrtChk1 
         Caption         =   "발주서"
         Height          =   255
         Left            =   6840
         TabIndex        =   24
         Top             =   390
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optPrtChk0 
         Caption         =   "의뢰서"
         Height          =   255
         Left            =   6840
         TabIndex        =   23
         Top             =   150
         Width           =   975
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   375
         Left            =   7980
         Picture         =   "매입작성1.frx":0000
         Style           =   1  '그래픽
         TabIndex        =   19
         Top             =   195
         Width           =   1095
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   4320
         Top             =   195
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowBorderStyle=   0
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowGroupTree=   -1  'True
      End
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   9120
         Picture         =   "매입작성1.frx":0963
         Style           =   1  '그래픽
         TabIndex        =   10
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   13635
         Picture         =   "매입작성1.frx":1308
         Style           =   1  '그래픽
         TabIndex        =   12
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   390
         Left            =   12510
         Picture         =   "매입작성1.frx":1C56
         Style           =   1  '그래픽
         TabIndex        =   11
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   11385
         Picture         =   "매입작성1.frx":25DA
         Style           =   1  '그래픽
         TabIndex        =   13
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Height          =   390
         Left            =   10260
         Picture         =   "매입작성1.frx":2E61
         Style           =   1  '그래픽
         TabIndex        =   0
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00008000&
         BorderStyle     =   1  '단일 고정
         Caption         =   "발주서 매입 전표 처리"
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
         TabIndex        =   16
         Top             =   180
         Width           =   4650
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid vsfg1 
      Height          =   2715
      Left            =   60
      TabIndex        =   8
      Top             =   1695
      Width           =   15195
      _cx             =   26802
      _cy             =   4789
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
      Height          =   1035
      Left            =   60
      TabIndex        =   14
      Top             =   630
      Width           =   15195
      Begin VB.OptionButton optDate 
         Caption         =   "매입일자"
         Height          =   255
         Index           =   2
         Left            =   6120
         TabIndex        =   5
         Top             =   620
         Width           =   1215
      End
      Begin VB.OptionButton optDate 
         Caption         =   "매입예정일자"
         Height          =   255
         Index           =   1
         Left            =   4440
         TabIndex        =   4
         Top             =   620
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optDate 
         Caption         =   "발주일자"
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   3
         Top             =   620
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   1
         Left            =   4440
         TabIndex        =   2
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   0
         Left            =   3240
         MaxLength       =   8
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpF_Date 
         Height          =   270
         Left            =   7980
         TabIndex        =   6
         Top             =   600
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   19922945
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin MSComCtl2.DTPicker dtpT_Date 
         Height          =   270
         Left            =   10080
         TabIndex        =   7
         Top             =   600
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   19922945
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin VB.Label Label1 
         Caption         =   "]"
         Height          =   240
         Index           =   4
         Left            =   7520
         TabIndex        =   29
         Top             =   645
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "["
         Height          =   240
         Index           =   3
         Left            =   2520
         TabIndex        =   28
         Top             =   645
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "까지"
         Height          =   240
         Index           =   6
         Left            =   11400
         TabIndex        =   27
         Top             =   650
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "부터"
         Height          =   240
         Index           =   5
         Left            =   9360
         TabIndex        =   26
         Top             =   645
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "기준일자"
         Height          =   240
         Index           =   2
         Left            =   1200
         TabIndex        =   25
         Top             =   650
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "전체금액"
         Height          =   240
         Index           =   1
         Left            =   10800
         TabIndex        =   22
         Top             =   285
         Width           =   1095
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
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   12120
         TabIndex        =   21
         Top             =   285
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "[Home]"
         Height          =   240
         Index           =   15
         Left            =   2520
         TabIndex        =   20
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "매입처코드"
         Height          =   240
         Index           =   0
         Left            =   1200
         TabIndex        =   18
         Top             =   285
         Width           =   1095
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
   End
End
Attribute VB_Name = "frm매입작성1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' 프로그램 제 목 : 매입작성1
' 사용된 Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' 참조된 Table   : 발주, 발주내역, 자재입출내역, 사업장
' 업  무  설  명 : 발주내역을 이용하여 발주(내역)을 수정 변경후에 매입을 작성
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived         As Boolean
Private P_adoRec             As New ADODB.Recordset
Private P_adoRecW            As New ADODB.Recordset
Private P_intButton          As Integer
Private P_strFindString2     As String
Private Const PC_intRowCnt1  As Integer = 8   '그리드1의 한 페이지 당 행수(FixedRows 포함)
Private Const PC_intRowCnt2  As Integer = 18  '그리드2의 한 페이지 당 행수(FixedRows 포함)

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
       Subvsfg1_INIT  '발주
       Subvsfg2_INIT  '발주내역
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
                   cmdSave.Enabled = True: cmdDelete.Enabled = True: cmdExit.Enabled = True
              Case Is <= 99 '삭제, 추가, 저장, 조회, 인쇄
                   cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = True: cmdDelete.Enabled = True: cmdExit.Enabled = True
              Case Else
                   cmdPrint.Enabled = False: cmdClear.Enabled = False: cmdFind.Enabled = False
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
       End Select
       SubOther_FILL
       P_blnActived = True
       Screen.MousePointer = vbDefault
    End If
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "매입작성1(서버와의 연결 실패)"
    Unload Me
    Exit Sub
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
                        Text1(1).Text = ""
                        Exit Sub
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

'+-------------------+
'/// 기준일자선택 ///
'+-------------------+
Private Sub optDate_Click(Index As Integer)
    If (Index = 0 Or Index = 1) Then
       cmdSave.Enabled = True: cmdDelete.Enabled = True
    Else
       cmdSave.Enabled = False:: cmdDelete.Enabled = False
    End If
    If cmdFind.Enabled = True Then
       cmdFind_Click
    End If
End Sub
Private Sub dtpF_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub dtpT_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       If cmdFind.Enabled = True Then
          cmdFind_Click
       Else
          SendKeys "{tab}"
       End If
    End If
End Sub

'+------------+
'/// vsfg1 ///
'+------------+
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
    With vsfg1
         P_intButton = Button
         If .MouseRow >= .FixedRows Then
            If cmdSave.Enabled = False Or optDate(2).Value = True Then Exit Sub
            If (.MouseCol = 9) Then     '유효일수
               If Button = vbLeftButton Then
                  .Select .MouseRow, .MouseCol
                  .EditCell
                End If
            ElseIf _
               (.MouseCol = 10) Then    '결제방법
               If Button = vbLeftButton Then
                  .Select .MouseRow, .MouseCol
                  .EditCell
                End If
            ElseIf _
                (.MouseCol = 11) Then   '결제예정일자
                If Button = vbLeftButton Then
                  .Select .MouseRow, .MouseCol
                  .EditCell
                End If
            ElseIf _
                (.MouseCol = 12) Then   '제목
                If Button = vbLeftButton Then
                  .Select .MouseRow, .MouseCol
                  .EditCell
                End If
            ElseIf _
                (.MouseCol = 13) Then   '적요
                If Button = vbLeftButton Then
                  .Select .MouseRow, .MouseCol
                  .EditCell
                End If
            End If
         End If
    End With
End Sub
Sub vsfg1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim blnModify As Boolean
    With vsfg1
         If Row >= .FixedRows Then
            If (Col = 9) Then  '유효일수
               If .TextMatrix(Row, Col) <> .EditText Then
                  If (IsNumeric(.EditText) = False Or Vals(.EditText) < 0 Or _
                     Fix(Vals(.EditText)) < Vals(.EditText) Or IsNumeric(Right(.EditText, 1)) = False) Then
                     Beep
                     Cancel = True
                  Else
                     blnModify = True
                  End If
               End If
            ElseIf _
               (Col = 10) Then  '결제방법
               If .Cell(flexcpChecked, Row, Col) <> .EditText Then
                  blnModify = True
               End If
            ElseIf _
               (Col = 11) Then  '결제예정일자
               If .TextMatrix(Row, Col) <> .EditText Then
                  .EditText = Format(Replace(.EditText, "-", ""), "0000-00-00")
                  If Not ((Len(Trim(.EditText)) = 10) And IsDate(.EditText) And Val(.EditText) > 2000) Then
                     Beep
                     Cancel = True
                  Else
                     blnModify = True
                  End If
               End If
             ElseIf _
               (Col = 12) Then  '제목 길이 검사
               If .TextMatrix(Row, Col) <> .EditText Then
                  If Not (LenH(Trim(.EditText)) <= 30) Then
                     Beep
                     Cancel = True
                  Else
                     blnModify = True
                  End If
               End If
             ElseIf _
               (Col = 13) Then ''적요 길이 검사
               If .TextMatrix(Row, Col) <> .EditText Then
                  If Not (LenH(Trim(.EditText)) <= 50) Then
                     Beep
                     Cancel = True
                  Else
                     blnModify = True
                  End If
               End If
             End If
             If blnModify = True Then
               .Cell(flexcpBackColor, Row, Col, Row, Col) = vbRed
               .Cell(flexcpForeColor, Row, Col, Row, Col) = vbWhite
            End If
         End If
    End With
End Sub
'Private Sub vsfg1_MouseUP(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    P_intButton = Button
'End Sub
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
            'strData = Trim(.Cell(flexcpData, .Row, 1))
            'Select Case .MouseCol
            '       Case 0
            '            .ColSel = 5
            '            .ColSort(0) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .ColSort(1) = flexSortNone
            '            .ColSort(2) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .ColSort(3) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .ColSort(4) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .Sort = flexSortUseColSort
            '       Case 1
            '            .ColSel = 5
            '            .ColSort(0) = flexSortNone
            '            .ColSort(1) = flexSortNone
            '            .ColSort(2) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .ColSort(3) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .ColSort(4) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .Sort = flexSortUseColSort
            '       Case Else
            '            .Sort = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            'End Select
            'If .FindRow(strData, , 1) > 0 Then
            '   .Row = .FindRow(strData, , 1)
            'End If
            'If PC_intRowCnt1 < .Rows Then
            '   .TopRow = .Row
            'End If
         End If
    End With
End Sub
Private Sub vsfg1_EnterCell()
Dim lngR1  As Long
Dim lngR2  As Long
Dim lngC   As Long
Dim lngCnt As Long
    With vsfg1
         If .Row >= .FixedRows Then
            'For lngR2 = 1 To vsfg2.Rows - 1
            '    vsfg2.RowHidden(lngR2) = True
            'Next lngR2
            'For lngR1 = .ValueMatrix(.Row, 15) To .ValueMatrix(.Row, 16)
            '    vsfg2.RowHidden(lngR1) = False
            '    lngCnt = lngCnt + 1
            'Next lngR1
            'Not Used
            'vsfg2.Row = .ValueMatrix(.Row, 15)
            'vsfg2.Select .ValueMatrix(.Row, 15), vsfg2.FixedCols, .ValueMatrix(.Row, 16), vsfg2.Cols - 1
            'If PC_intRowCnt2 < vsfg2.Rows Then
            '   vsfg2.TopRow = vsfg2.Row
            'End If
            'If PC_intRowCnt2 < lngCnt Then
            '   vsfg2.TopRow = vsfg2.Row
            'End If
         End If
    End With
End Sub
Private Sub vsfg1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
Dim lngR1  As Long
Dim lngR2  As Long
Dim lngC   As Long
Dim lngCnt As Long
    With vsfg1
         If NewRow < 1 Then Exit Sub
         If NewRow <> OldRow Then
            For lngR2 = 1 To vsfg2.Rows - 1
                vsfg2.RowHidden(lngR2) = True
            Next lngR2
            For lngR1 = .ValueMatrix(.Row, 15) To .ValueMatrix(.Row, 16)
                If vsfg2.TextMatrix(lngR1, 28) = "D" Then
                   vsfg2.RowHidden(lngR1) = True
                Else
                   vsfg2.RowHidden(lngR1) = False
                   lngCnt = lngCnt + 1
                   vsfg2.TextMatrix(lngR1, 0) = lngCnt '순번
                End If
            Next lngR1
            'Not Used
            'vsfg2.Row = .ValueMatrix(.Row, 15)
            'vsfg2.Select .ValueMatrix(.Row, 15), vsfg2.FixedCols, .ValueMatrix(.Row, 16), vsfg2.Cols - 1
            'If PC_intRowCnt2 < vsfg2.Rows Then
            '   vsfg2.TopRow = vsfg2.Row
            'End If
            If PC_intRowCnt2 < lngCnt Then
               vsfg2.TopRow = vsfg2.Row
            End If
            vsfg2.Row = 0
         End If
    End With
End Sub
Private Sub vsfg1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL       As String
Dim intRetVal    As Integer
Dim lngCnt       As Long
    With vsfg1
    End With
    Exit Sub
End Sub

'+------------+
'/// vsfg2 ///
'+------------+
Private Sub vsfg2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    With vsfg2
         .ToolTipText = ""
         If .MouseRow < .FixedRows Or .MouseCol < 0 Then
            Exit Sub
         End If
         .ToolTipText = .TextMatrix(.MouseRow, .MouseCol)
    End With
End Sub
Private Sub vsfg2_DblClick()
    With vsfg1
         If .Row >= .FixedRows Then
             vsfg2_KeyDown vbKeyF1, 0  '자재시세검색
         End If
    End With
End Sub
Private Sub vsfg2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    With vsfg2
         P_intButton = Button
         If .Row >= .FixedRows Then
            If cmdSave.Enabled = False Or optDate(2).Value = True Then Exit Sub
            If (.Col = 16) Then      '수량
                If Button = vbLeftButton Then
                  .Select .Row, .Col
                  .EditCell
                End If
            ElseIf _
               (.Col = 18) Then      '직송구분
               If Button = vbLeftButton Then
                  .Select .Row, .Col
                  .EditCell
                End If
            ElseIf _
                (.Col = 19) Then     '입고단가
                If Button = vbLeftButton Then
                  .Select .Row, .Col
                  .EditCell
                End If
            ElseIf _
                (.Col = 20) Then     '입고부가
                If Button = vbLeftButton Then
                  .Select .Row, .Col
                  .EditCell
                End If
            ElseIf _
                (.Col = 27) Then     '적요
                If Button = vbLeftButton Then
                  .Select .Row, .Col
                  .EditCell
                End If
            End If
         End If
    End With
End Sub
Sub vsfg2_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim blnModify As Boolean
Dim curTmpMny As Currency
    With vsfg2
         If Row >= .FixedRows Then
            If (Col = 16) Then  '발주량
               If .TextMatrix(Row, Col) <> .EditText Then
                  If (IsNumeric(.EditText) = False Or _
                     Fix(Vals(.EditText)) < Vals(.EditText) Or IsNumeric(Right(.EditText, 1)) = False) Then
                     Beep
                     Cancel = True
                  Else
                     blnModify = True
                     curTmpMny = .ValueMatrix(Row, 22)
                     .TextMatrix(Row, 22) = Vals(.EditText) * .ValueMatrix(Row, 19) '힙계금액 = 수량 * 단가
                  End If
               End If
            ElseIf _
               (Col = 18) Then  '직송구분
               If (Len(.TextMatrix(Row, 9)) = 0) Then '매출처가 없음
                  .Cell(flexcpChecked, Row, 18, Row, 18) = flexUnchecked
                  Beep
                  Cancel = True
                  Exit Sub
               End If
               If .Cell(flexcpChecked, Row, Col) <> .EditText Then
                  blnModify = True
               End If
            ElseIf _
               (Col = 19) Then  '입고단가
               If .TextMatrix(Row, Col) <> .EditText Then
                  If (IsNumeric(.EditText) = False Or Vals(.EditText) < 0 Or _
                     IsNumeric(Right(.EditText, 1)) = False) Then                                            '소숫점이하 사용가
                     'fix(Vals(.EditText)) < Vals(.EditText) Or IsNumeric(Right(.EditText, 1)) = False) Then '소숫점이하 사용불가
                     Beep
                     Cancel = True
                  Else
                     blnModify = True
                     .TextMatrix(Row, 20) = Fix(Vals(.EditText) * (PB_curVatRate))  '부가세
                     curTmpMny = .ValueMatrix(Row, 22)
                     .TextMatrix(Row, 21) = Vals(.EditText) + .ValueMatrix(Row, 20)
                     .TextMatrix(Row, 22) = .ValueMatrix(Row, 16) * Vals(.EditText)
                  End If
               End If
            ElseIf _
               (Col = 20) Then  '입고부가
               If .TextMatrix(Row, Col) <> .EditText Then
                  If (IsNumeric(.EditText) = False Or Vals(.EditText) < 0 Or _
                     Fix(Vals(.EditText)) < Vals(.EditText) Or IsNumeric(Right(.EditText, 1)) = False) Then
                     Beep
                     Cancel = True
                  Else
                     blnModify = True
                     curTmpMny = .ValueMatrix(Row, 22)
                     .TextMatrix(Row, 21) = .ValueMatrix(Row, 19) + Vals(.EditText)
                     .TextMatrix(Row, 22) = .ValueMatrix(Row, 16) * .ValueMatrix(Row, 19)
                  End If
               End If
            ElseIf _
               (Col = 27) Then '적요 길이 검사
               If .TextMatrix(Row, Col) <> .EditText Then
                  If Not (LenH(Trim(.EditText)) <= 50) Then
                     Beep
                     Cancel = True
                  Else
                     blnModify = True
                  End If
               End If
            End If
            '변경표시 + 금액재계산
            If blnModify = True Then
               If .TextMatrix(Row, 28) = "" Then
                  .TextMatrix(Row, 28) = "U"
               End If
               .Cell(flexcpBackColor, Row, Col, Row, Col) = vbRed
               .Cell(flexcpForeColor, Row, Col, Row, Col) = vbWhite
               Select Case Col
                      Case 16, 19, 20
                           vsfg1.TextMatrix(vsfg1.Row, 7) = vsfg1.ValueMatrix(vsfg1.Row, 7) - curTmpMny + .ValueMatrix(Row, 22)
                           lblTotMny.Caption = Format(Vals(lblTotMny.Caption) - curTmpMny + .ValueMatrix(Row, 22), "#,#.00")
                      Case Else
               End Select
            End If
         End If
    End With
End Sub
Private Sub vsfg2_EnterCell()
Dim lngR As Long
Dim lngC As Long
    With vsfg2
         .Editable = flexEDNone
         If .Row >= .FixedRows Then
             Select Case .Col
                    Case 16, 19, 27
                         .Editable = flexEDKbdMouse
                         vsfg2_MouseUp vbLeftButton, 0, 0, 0
             End Select
         End If
    End With
End Sub
Private Sub vsfg2_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
Dim lngR    As Long
    With vsfg2
         If KeyCode = vbKeyReturn Then
            If Col = 16 Then
               .Col = 19
            ElseIf _
               Col = 19 Then
               .Col = 27
            ElseIf _
               Col = 27 Then
               For lngR = Row To vsfg1.ValueMatrix(vsfg1.Row, 16)
                   If .RowHidden(lngR) = False Then
                      Exit For
                   End If
               Next lngR
               If lngR <> vsfg1.ValueMatrix(vsfg1.Row, 16) Then
                  .Col = 16: .LeftCol = 15
                  .Row = .Row + 1
                  If .ValueMatrix(lngR, 0) > PC_intRowCnt2 Then
                     .TopRow = .TopRow + 1
                  End If
               End If
            End If
         End If
    End With
End Sub
Private Sub vsfg2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lngR     As Long
Dim lngPos   As Long
Dim blnDupOK As Boolean
Dim strTime  As String
Dim strHH    As String
Dim strMM    As String
Dim strSS    As String
Dim strMS    As String
Dim intRetVal As Integer
Dim CtrlDown  As Variant
    With vsfg2
         If (.Row >= .FixedRows) Then     '내역시세검색
            If cmdSave.Enabled = False Or optDate(2).Value = True Then
               Exit Sub
            End If
            If KeyCode = vbKeyF2 And (Len(vsfg1.TextMatrix(vsfg1.Row, 5)) > 0) And _
              (Len(.TextMatrix(.Row, 4)) > 0) Then
               PB_strFMCCallFormName = "frm매입작성1"
               PB_strMaterialsCode = .TextMatrix(.Row, 4)
               PB_strMaterialsName = .TextMatrix(.Row, 5)
               PB_strSupplierCode = vsfg1.TextMatrix(vsfg1.Row, 5)
               frm내역시세검색.Show vbModal
               If Len(PB_strMaterialsCode) <> 0 Then
                  PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
               End If
            End If
         End If
    End With
    With vsfg2
         '내역이 없는 경우 추가
         If .Row = 0 And KeyCode = vbKeyInsert Then
            For lngR = vsfg1.ValueMatrix(vsfg1.Row, 15) To vsfg1.ValueMatrix(vsfg1.Row, 16)
                If .RowHidden(lngR) = False Then
                   lngPos = lngPos + 1
                End If
            Next lngR
            If lngPos = 0 Then .Row = vsfg1.ValueMatrix(vsfg1.Row, 16)
         End If
         If (.Row >= .FixedRows) Then
            If cmdSave.Enabled = False Or optDate(2).Value = True Then Exit Sub
            If KeyCode = vbKeyF1 Then  '자재시세검색
               'If (.MouseCol = 5) Then
                  PB_strFMCCallFormName = "frm매입작성1"
                  PB_strMaterialsCode = .TextMatrix(.Row, 4)
                  PB_strMaterialsName = .TextMatrix(.Row, 5)
                  PB_strSupplierCode = vsfg1.TextMatrix(vsfg1.Row, 5)
                  frm자재시세검색.Show vbModal
                  If Len(PB_strMaterialsCode) <> 0 Then
                     PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                  End If
               'ElseIf _
               '   (.Col = 9) Then      '매출처검색
               '   PB_strSupplierCode = .TextMatrix(.Row, 8)
               '   PB_strSupplierName = .TextMatrix(.Row, 9)
               '   frm매출처검색.Show vbModal
               '   If Len(PB_strSupplierCode) <> 0 Then
               '      'For lngR = vsfg1.ValueMatrix(vsfg1.Row, 15) To vsfg1.ValueMatrix(vsfg1.Row, 16) '자재코드(변경후) + 매출처코드
               '      '    If .Row <> lngR And .TextMatrix(lngR, 0) = .TextMatrix(.Row, 0) And _
               '      '       .TextMatrix(lngR, 4) = PB_strSupplierCode Then
               '      '       blnDupOK = True
               '      '       Exit For
               '      '    End If
               '      'Next lngR
               '      If blnDupOK = False Then
               '         If PB_strSupplierCode <> .TextMatrix(.Row, 8) Then
               '            .Cell(flexcpBackColor, .Row, 9, .Row, 9) = vbRed
               '            .Cell(flexcpForeColor, .Row, 9, .Row, 9) = vbWhite
               '         End If
               '         .TextMatrix(.Row, 8) = PB_strSupplierCode
               '         .TextMatrix(.Row, 9) = PB_strSupplierName
               '         If .TextMatrix(.Row, 28) = "" Then
               '            .TextMatrix(.Row, 28) = "U"
               '         End If
               '      End If
               '   End If
               'End If
            ElseIf _
               KeyCode = vbKeyInsert And cmdSave.Enabled = True Then '발주내역 추가
               .AddItem "", .Row + 1
               .Row = .Row + 1
               .TopRow = .Row
               .TextMatrix(.Row, 0) = .ValueMatrix(.Row - 1, 0) + 1 '순번
               .TextMatrix(.Row, 1) = .TextMatrix(.Row - 1, 1)   '사업장코드
               .TextMatrix(.Row, 2) = .TextMatrix(.Row - 1, 2)   '발주일자
               .TextMatrix(.Row, 3) = .TextMatrix(.Row - 1, 3)   '발주번호
               .TextMatrix(.Row, 13) = .TextMatrix(.Row - 1, 13) '매입처코드
               .TextMatrix(.Row, 14) = .TextMatrix(.Row - 1, 14) '매입처명
               .Cell(flexcpChecked, .Row, 18) = flexUnchecked    '직송
               .Cell(flexcpText, .Row, 18) = "직 송"
               .Cell(flexcpAlignment, .Row, 18, .Row, 18) = flexAlignLeftCenter
               .TextMatrix(.Row, 28) = "I"                       'SQL구분
               '발주시간
               strTime = .TextMatrix(.Row - 1, 29)
               If .Row <= vsfg1.ValueMatrix(vsfg1.Row, 16) Then  '발주번호의 마지막 아니면
                  strTime = Format(Fix((.ValueMatrix(.Row - 1, 29) + .ValueMatrix(.Row + 1, 29)) / 2), "000000000")
                  '추가 가능한지 검사
                  If (strTime = .TextMatrix(.Row - 1, 29)) Or (strTime = .TextMatrix(.Row - 1, 29)) Then
                     MsgBox "이 행에는 더 이상 추가 할 수 없습니다. 다른 행에 추가하세요.", vbCritical + vbDefaultButton1, "추가"
                     .RemoveItem (.Row)
                     Exit Sub
                  End If
               Else                                              '발주번호의 마지막이면
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
               .TextMatrix(.Row, 29) = strTime                   '발주시간
               PB_strFMCCallFormName = "frm매입작성1"
               PB_strMaterialsCode = .TextMatrix(.Row, 4)
               PB_strMaterialsName = .TextMatrix(.Row, 5)
               PB_strSupplierCode = .TextMatrix(.Row, 13)
               frm자재시세검색.Show vbModal
               If Len(PB_strMaterialsCode) <> 0 Then '정상적인 선택이면
                  PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                  '순번
                  For lngR = (.Row + 1) To (vsfg1.ValueMatrix(vsfg1.Row, 16) + 1)
                      If .TextMatrix(lngR, 28) <> "D" Then
                         .TextMatrix(lngR, 0) = .ValueMatrix(lngR, 0) + 1
                      End If
                  Next lngR
                  '발주
                  vsfg1.TextMatrix(vsfg1.Row, 16) = vsfg1.ValueMatrix(vsfg1.Row, 16) + 1
                  For lngR = 1 To vsfg1.Rows - 1
                      If (lngR <> .Row) Then
                         If (vsfg1.ValueMatrix(lngR, 15) >= .Row) Then
                            vsfg1.TextMatrix(lngR, 15) = vsfg1.ValueMatrix(lngR, 15) + 1
                            vsfg1.TextMatrix(lngR, 16) = vsfg1.ValueMatrix(lngR, 16) + 1
                         End If
                      End If
                  Next lngR
               Else
                  .RemoveItem (.Row)
                  .Row = .Row - 1
               End If
            ElseIf _
               KeyCode = vbKeyDelete And .Col = 9 Then  '매출처 지움
               If (Len(.TextMatrix(.Row, 9)) <> 0) Then
                  .TextMatrix(.Row, 8) = "": .TextMatrix(.Row, 9) = ""
                  .Cell(flexcpBackColor, .Row, 9, .Row, 9) = vbRed
                  .Cell(flexcpForeColor, .Row, 9, .Row, 9) = vbWhite
                  If .TextMatrix(.Row, 28) = "" Then
                     .TextMatrix(.Row, 28) = "U"
                  End If
               End If
               If .Cell(flexcpChecked, .Row, 18, .Row, 18) = flexChecked Then
                  .Cell(flexcpChecked, .Row, 18, .Row, 18) = flexUnchecked
                  .Cell(flexcpBackColor, .Row, 18, .Row, 18) = vbRed
                  .Cell(flexcpForeColor, .Row, 18, .Row, 18) = vbWhite
               End If
            ElseIf _
               KeyCode = vbKeyDelete And (.Col <> 9) And (.Row > 0) And .RowHidden(.Row) = False Then
               intRetVal = MsgBox("입력한 발주내역을 삭제하시겠습니까 ?", vbQuestion + vbYesNo + vbDefaultButton2, "발주내역삭제")
               If intRetVal = vbYes Then
                  .TextMatrix(.Row, 28) = "D": .TextMatrix(.Row, 0) = "0"
                  vsfg1.TextMatrix(vsfg1.Row, 7) = vsfg1.ValueMatrix(vsfg1.Row, 7) - .ValueMatrix(.Row, 22)
                  lblTotMny.Caption = Format(Vals(lblTotMny.Caption) - .ValueMatrix(.Row, 22), "#,#.00")
                  .RowHidden(.Row) = True
                   For lngR = .Row + 1 To vsfg1.ValueMatrix(vsfg1.Row, 16)
                      If .TextMatrix(lngR, 28) <> "D" Then
                         .TextMatrix(lngR, 0) = .ValueMatrix(lngR, 0) - 1
                         If lngPos = 0 Then
                            lngPos = lngR
                         End If
                      End If
                  Next lngR
                  If lngPos = 0 Then
                     For lngR = vsfg1.ValueMatrix(vsfg1.Row, 15) To .Row
                         If .TextMatrix(lngR, 28) <> "D" And lngR < .Row Then
                            lngPos = lngR
                         End If
                     Next lngR
                  End If
                  .Row = lngPos
               End If
            End If
         End If
    End With
End Sub

'+-----------+
'/// 출력 ///
'+-----------+
Private Sub cmdPrint_Click()
    If ((vsfg1.Rows + vsfg2.Rows) = 2) Or (vsfg1.Row < 1) Then
       Exit Sub
    End If
    SubPrintCrystalReports
End Sub
'+-----------+
'/// 추가 ///
'+-----------+
Private Sub cmdClear_Click()
'
End Sub
'+-----------+
'/// 조회 ///
'+-----------+
Private Sub cmdFind_Click()
    Screen.MousePointer = vbHourglass
    cmdFind.Enabled = False
    lblTotMny.Caption = "0.00"
    Subvsfg1_FILL
    Subvsfg2_FILL
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
Dim lngDelCnt      As Long
Dim lngLogCnt      As Long
Dim curInputMoney  As Long
Dim curOutputMoney As Long
Dim strServerTime  As String
Dim strTime        As String
Dim strHH          As String
Dim strMM          As String
Dim strSS          As String
Dim strMS          As String
Dim intChkCash     As Integer '1.현금매입

    If vsfg1.Row >= vsfg1.FixedRows Then
       intRetVal = MsgBox("발주서에서 적성한 자료를 매입처리하시겠습니까 ?", vbQuestion + vbYesNo + vbDefaultButton1, "발주서매입저장")
       If intRetVal = vbNo Then
          vsfg1.SetFocus
          Exit Sub
       End If
       intRetVal = MsgBox("현금매입을 하시겠습니까 ?", vbQuestion + vbYesNoCancel + vbDefaultButton2, "현금매입")
       If intRetVal = vbYes Then
          intChkCash = 1
       ElseIf _
          intRetVal = vbCancel Then
          Exit Sub
       End If
       cmdSave.Enabled = False
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
       '거래번호 구하기
       PB_adoCnnSQL.BeginTrans
       P_adoRec.CursorLocation = adUseClient
       strSQL = "spLogCounter '자재입출내역', '" & PB_regUserinfoU.UserBranchCode + PB_regUserinfoU.UserClientDate + "1" & "', " _
                            & "0, 0, '" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' "
       On Error GoTo ERROR_STORED_PROCEDURE
       P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
       lngLogCnt = P_adoRec(0)
       P_adoRec.Close
       With vsfg2
            '발주내역
            For lngRR = vsfg1.ValueMatrix(vsfg1.Row, 15) To vsfg1.ValueMatrix(vsfg1.Row, 16)
                If (.TextMatrix(lngRR, 28) = "I") Then 'And .ValueMatrix(lngRR, 16) <> 0) Then '발주내역 추가
                   strSQL = "INSERT INTO 발주내역(사업장코드, 발주일자, " _
                                               & "발주번호, 발주시간, 자재코드, " _
                                               & "매입처코드, 발주량, " _
                                               & "직송구분, 매출처코드, " _
                                               & "입고단가, 입고부가, " _
                                               & "출고단가, 출고부가, " _
                                               & "상태코드, 입고일자, " _
                                               & "출고일자, 적요, " _
                                               & "사용구분, 수정일자, " _
                                               & "사용자코드) Values( " _
                             & "'" & .TextMatrix(lngRR, 1) & "', '" & DTOS(.TextMatrix(lngRR, 2)) & "', " _
                             & "" & .ValueMatrix(lngRR, 3) & ", '" & .TextMatrix(lngRR, 29) & "', '" & .TextMatrix(lngRR, 4) & "', " _
                             & "'" & .TextMatrix(lngRR, 13) & "', " & .ValueMatrix(lngRR, 16) & ", " _
                             & "" & IIf(.Cell(flexcpChecked, lngRR, 18, lngRR, 18) = flexUnchecked, 0, 1) & ", '" & .TextMatrix(lngRR, 8) & "', " _
                             & "" & .ValueMatrix(lngRR, 19) & ", " & .ValueMatrix(lngRR, 20) & ", " _
                             & "" & .ValueMatrix(lngRR, 23) & ", " & .ValueMatrix(lngRR, 24) & ", " _
                             & "1, '', " _
                             & "'', '" & .TextMatrix(lngRR, 27) & "', " _
                             & "0, '" & PB_regUserinfoU.UserServerDate & "', " _
                             & "'" & PB_regUserinfoU.UserCode & "' ) "
                   On Error GoTo ERROR_TABLE_INSERT
                   PB_adoCnnSQL.Execute strSQL
                ElseIf _
                   (.TextMatrix(lngRR, 28) = "D") Then '발주내역 취소
                   strSQL = "DELETE FROM 발주내역 " _
                           & "WHERE 사업장코드 = '" & .TextMatrix(lngRR, 1) & "' " _
                             & "AND 발주일자 = '" & DTOS(.TextMatrix(lngRR, 2)) & "' " _
                             & "AND 발주번호 = " & .ValueMatrix(lngRR, 3) & " " _
                             & "AND 발주시간 = '" & .TextMatrix(lngRR, 29) & "' " _
                             & "AND 자재코드 = '" & .TextMatrix(lngRR, 6) & "' " _
                             & "AND 매출처코드 = '" & .TextMatrix(lngRR, 10) & "' "
                   lngDelCnt = lngDelCnt + 1     '삭제할 Row수 계산
                   On Error GoTo ERROR_TABLE_DELETE
                   PB_adoCnnSQL.Execute strSQL
                ElseIf _
                   (.TextMatrix(lngRR, 28) = "U") Then 'And .ValueMatrix(lngRR, 16) <> 0) Then '발주내역 변경
                   strSQL = "UPDATE 발주내역 SET " _
                                 & "자재코드 = '" & .TextMatrix(lngRR, 4) & "', " _
                                 & "발주량 = " & .ValueMatrix(lngRR, 16) & ", " _
                                 & "직송구분 = " & IIf(.Cell(flexcpChecked, lngRR, 18, lngRR, 18) = flexUnchecked, 0, 1) & ", " _
                                 & "매출처코드 = '" & .TextMatrix(lngRR, 8) & "', " _
                                 & "입고단가 = " & .ValueMatrix(lngRR, 19) & ", " _
                                 & "입고부가 = " & .ValueMatrix(lngRR, 20) & ", " _
                                 & "출고단가 = " & .ValueMatrix(lngRR, 23) & "," _
                                 & "출고부가 = " & .ValueMatrix(lngRR, 24) & ", " _
                                 & "적요 = '" & .TextMatrix(lngRR, 27) & "', " _
                                 & "수정일자 = '" & PB_regUserinfoU.UserServerDate & "', " _
                                 & "사용자코드 = '" & PB_regUserinfoU.UserCode & "' " _
                           & "WHERE 사업장코드 = '" & .TextMatrix(lngRR, 1) & "' " _
                             & "AND 발주일자 = '" & DTOS(.TextMatrix(lngRR, 2)) & "' " _
                             & "AND 발주번호 = " & .ValueMatrix(lngRR, 3) & " " _
                             & "AND 발주시간 = '" & .TextMatrix(lngRR, 29) & "' " _
                             & "AND 자재코드 = '" & .TextMatrix(lngRR, 6) & "' " _
                             & "AND 매출처코드 = '" & .TextMatrix(lngRR, 10) & "' "
                   On Error GoTo ERROR_TABLE_INSERT
                   PB_adoCnnSQL.Execute strSQL
                End If
            Next lngRR
       End With
       With vsfg1
            '발주
            If (.ValueMatrix(.Row, 16) - .ValueMatrix(.Row, 15) + 1) = lngDelCnt Then '발주내역 모두 삭제
               strSQL = "DELETE FROM 발주 " _
                       & "WHERE 사업장코드 = '" & .TextMatrix(.Row, 2) & "' AND 발주일자 = '" & DTOS(.TextMatrix(.Row, 3)) & "' " _
                         & "AND 발주번호 = " & .ValueMatrix(.Row, 4) & " "
               On Error GoTo ERROR_TABLE_DELETE
               PB_adoCnnSQL.Execute strSQL
               lngDelCntS = .ValueMatrix(.Row, 15): lngDelCntE = .ValueMatrix(.Row, 16)
               For lngRR = .ValueMatrix(.Row, 16) To .ValueMatrix(.Row, 15) Step -1
                   vsfg2.RemoveItem lngRR
               Next lngRR
               .RemoveItem .Row
               For lngRRR = 1 To .Rows - 1
                   If lngDelCntS < .ValueMatrix(lngRRR, 15) Then
                      .TextMatrix(lngRRR, 15) = .ValueMatrix(lngRRR, 15) - (lngDelCntE - lngDelCntS + 1)
                      .TextMatrix(lngRRR, 16) = .ValueMatrix(lngRRR, 16) - (lngDelCntE - lngDelCntS + 1)
                   End If
               Next lngRRR
               .Row = 0 '현재선택된 발주 Row를 해제
            Else
               strSQL = "UPDATE 발주 SET " _
                             & "결제방법 = " & IIf(.Cell(flexcpChecked, .Row, 10) = flexChecked, 0, 1) & ", " _
                             & "결제예정일자 = '" & DTOS(Trim(.TextMatrix(.Row, 11))) & "', " _
                             & "유효일수 = " & .ValueMatrix(.Row, 9) & ", " _
                             & "제목 = '" & Trim(.TextMatrix(.Row, 12)) & "', " _
                             & "적요 = '" & Trim(.TextMatrix(.Row, 13)) & "', " _
                             & "수정일자 = '" & PB_regUserinfoU.UserServerDate & "' " _
                       & "WHERE 사업장코드 = '" & .TextMatrix(.Row, 2) & "' AND 발주일자 = '" & DTOS(.TextMatrix(.Row, 3)) & "' " _
                         & "AND 발주번호 = " & .ValueMatrix(.Row, 4) & " "
               On Error GoTo ERROR_TABLE_UPDATE
               PB_adoCnnSQL.Execute strSQL
               '발주(색상 원위치)
               .Cell(flexcpBackColor, .Row, .FixedCols, .Row, .Cols - 1) = vbWhite
               .Cell(flexcpForeColor, .Row, .FixedCols, .Row, .Cols - 1) = vbBlack
               lngDelCntS = .ValueMatrix(.Row, 15): lngDelCntE = .ValueMatrix(.Row, 16)
               For lngRR = .ValueMatrix(.Row, 16) To .ValueMatrix(.Row, 15) Step -1
                   If vsfg2.TextMatrix(lngRR, 28) = "D" Then
                      vsfg2.RemoveItem lngRR
                   End If
               Next lngRR
               vsfg1.TextMatrix(vsfg1.Row, 16) = vsfg1.ValueMatrix(vsfg1.Row, 16) - lngDelCnt
               For lngR = 1 To vsfg1.Rows - 1
                   If (lngR <> vsfg1.Row) Then
                      If (vsfg1.ValueMatrix(.Row, 16) < vsfg1.ValueMatrix(lngR, 15)) Then
                         vsfg1.TextMatrix(lngR, 15) = vsfg1.ValueMatrix(lngR, 15) - lngDelCnt
                         vsfg1.TextMatrix(lngR, 16) = vsfg1.ValueMatrix(lngR, 16) - lngDelCnt
                      End If
                   End If
               Next lngR
            End If
       End With
       With vsfg2
            '변경후(재정렬)
            If vsfg1.Row > 0 Then
               For lngRR = vsfg1.ValueMatrix(vsfg1.Row, 15) To vsfg1.ValueMatrix(vsfg1.Row, 16)
                   If (.TextMatrix(lngRR, 28) = "U" And .ValueMatrix(lngRR, 16) <> 0) Then   '발주내역 변경
                      .TextMatrix(lngRR, 6) = .TextMatrix(lngRR, 4)   '자재코드(변경후->변경전)
                      .TextMatrix(lngRR, 7) = .TextMatrix(lngRR, 5)   '자재명(변경후->변경전)
                      .TextMatrix(lngRR, 10) = .TextMatrix(lngRR, 8)  '매출처코드(변경후->변경전)
                      .TextMatrix(lngRR, 11) = .TextMatrix(lngRR, 9)  '매출처명(변경후->변경전)
                   End If
               Next lngRR
            End If
       End With
       With vsfg2
            '발주내역(색상 원위치)
            If vsfg1.Row > 0 Then '현재선택된 발주 Row를 해제
               .Cell(flexcpBackColor, vsfg1.ValueMatrix(vsfg1.Row, 15), 0, vsfg1.ValueMatrix(vsfg1.Row, 16), .FixedCols - 1) = _
               .Cell(flexcpBackColor, 0, 0, 0, 0)
               .Cell(flexcpForeColor, vsfg1.ValueMatrix(vsfg1.Row, 15), 0, vsfg1.ValueMatrix(vsfg1.Row, 16), .FixedCols - 1) = _
               .Cell(flexcpForeColor, 0, 0, 0, 0)
               .Cell(flexcpBackColor, vsfg1.ValueMatrix(vsfg1.Row, 15), .FixedCols, vsfg1.ValueMatrix(vsfg1.Row, 16), .Cols - 1) = vbWhite
               .Cell(flexcpForeColor, vsfg1.ValueMatrix(vsfg1.Row, 15), .FixedCols, vsfg1.ValueMatrix(vsfg1.Row, 16), .Cols - 1) = vbBlack
               .Cell(flexcpText, vsfg1.ValueMatrix(vsfg1.Row, 15), 28, vsfg1.ValueMatrix(vsfg1.Row, 16), 28) = "" 'SQL구분 지움
            End If
       End With
       '매입처리
       lngChkCnt = 0
       With vsfg2
            If vsfg1.Row > 0 Then
               '발주내역
               For lngRR = vsfg1.ValueMatrix(vsfg1.Row, 15) To vsfg1.ValueMatrix(vsfg1.Row, 16)
                   If vsfg2.RowHidden(lngRR) = False Then
                      lngChkCnt = lngChkCnt + 1
                      If lngChkCnt = 1 Then
                         strTime = strServerTime
                      Else
                         strTime = Format((Val(strTime) + 1000), "000000000")
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
                      strSQL = "UPDATE 발주내역 SET " _
                                    & "상태코드 = 2, " _
                                    & "입고일자 = '" & PB_regUserinfoU.UserClientDate & "', " _
                                    & "수정일자 = '" & PB_regUserinfoU.UserClientDate & "', " _
                                    & "사용자코드 = '" & PB_regUserinfoU.UserCode & "' " _
                              & "WHERE 사업장코드 = '" & .TextMatrix(lngRR, 1) & "' " _
                                & "AND 발주일자 = '" & DTOS(.TextMatrix(lngRR, 2)) & "' " _
                                & "AND 발주번호 = " & .ValueMatrix(lngRR, 3) & " " _
                                & "AND 발주시간 = '" & .TextMatrix(lngRR, 29) & "' " _
                                & "AND 자재코드 = '" & .TextMatrix(lngRR, 4) & "' " _
                                & "AND 매출처코드 = '" & .TextMatrix(lngRR, 8) & "' "
                      On Error GoTo ERROR_TABLE_UPDATE
                      PB_adoCnnSQL.Execute strSQL
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
                             & "'" & .TextMatrix(lngRR, 1) & "', '" & Mid(.TextMatrix(lngRR, 4), 1, 2) & "', " _
                             & "'" & Mid(.TextMatrix(lngRR, 4), 3) & "', 1, " _
                             & "'" & PB_regUserinfoU.UserClientDate & "', '" & strTime & "', " _
                             & "" & .ValueMatrix(lngRR, 16) & ", " & .ValueMatrix(lngRR, 19) & ", " _
                             & "" & .ValueMatrix(lngRR, 20) & ", 0, " _
                             & "" & .ValueMatrix(lngRR, 23) & ", " & .ValueMatrix(lngRR, 24) & ", " _
                             & "'" & .TextMatrix(lngRR, 13) & "' , '" & .TextMatrix(lngRR, 8) & "', " _
                             & "'" & PB_regUserinfoU.UserClientDate & "', " & IIf(.Cell(flexcpChecked, lngRR, 18, lngRR, 18) = flexUnchecked, 0, 1) & ", " _
                             & "'" & DTOS(.TextMatrix(lngRR, 2)) & "', " & .ValueMatrix(lngRR, 3) & ", " _
                             & "'" & PB_regUserinfoU.UserClientDate & "', " & lngLogCnt & ", " _
                             & "0, " & intChkCash & ", 0, " _
                             & "'" & Trim(.TextMatrix(lngRR, 27)) & "', '', 0, 0, " _
                             & "0, '" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "', '' ) "
                      On Error GoTo ERROR_TABLE_INSERT
                      PB_adoCnnSQL.Execute strSQL
                      '최종입출고일자갱신(사업장코드, 분류코드, 세부코드, 입출고구분)
                      strSQL = "sp자재최종입출고일자갱신 '" & PB_regUserinfoU.UserBranchCode & "', " _
                             & "'" & Mid(.TextMatrix(lngRR, 4), 1, 2) & "', '" & Mid(.TextMatrix(lngRR, 4), 3) & "', 1 "
                      On Error GoTo ERROR_STORED_PROCEDURE
                      PB_adoCnnSQL.Execute strSQL
                      '자재최종단가갱신(사업장코드, 분류코드, 세부코드, 입출고구분, 업체코드, 단가, 거래일자)
                      If .ValueMatrix(lngRR, 19) > 0 And PB_intIAutoPriceGbn = 1 Then
                         strSQL = "sp자재최종단가갱신 '" & PB_regUserinfoU.UserBranchCode & "', " _
                                & "'" & Mid(.TextMatrix(lngRR, 4), 1, 2) & "', '" & Mid(.TextMatrix(lngRR, 4), 3) & "', 1, " _
                                 & "'" & .TextMatrix(lngRR, 13) & "', " _
                                 & "" & .ValueMatrix(lngRR, 19) & ", '" & PB_regUserinfoU.UserClientDate & "' "
                         On Error GoTo ERROR_STORED_PROCEDURE
                         PB_adoCnnSQL.Execute strSQL
                      End If
                   End If
               Next lngRR
               '저장 내역 삭제
               lngDelCntS = vsfg1.ValueMatrix(vsfg1.Row, 15): lngDelCntE = vsfg1.ValueMatrix(vsfg1.Row, 16)
               For lngRR = vsfg1.ValueMatrix(vsfg1.Row, 16) To vsfg1.ValueMatrix(vsfg1.Row, 15) Step -1
                   vsfg2.RemoveItem lngRR
               Next lngRR
               '발주
               strSQL = "UPDATE 발주 SET " _
                             & "입고일자 = '" & PB_regUserinfoU.UserClientDate & "', " _
                             & "상태코드 = 2, " _
                             & "수정일자 = '" & PB_regUserinfoU.UserServerDate & "', " _
                             & "사용자코드 = '" & PB_regUserinfoU.UserCode & "' " _
                       & "WHERE 사업장코드 = '" & vsfg1.TextMatrix(vsfg1.Row, 2) & "' " _
                         & "AND 발주일자 = '" & DTOS(vsfg1.TextMatrix(vsfg1.Row, 3)) & "' " _
                         & "AND 발주번호 = " & vsfg1.ValueMatrix(vsfg1.Row, 4) & " "
               On Error GoTo ERROR_TABLE_UPDATE
               PB_adoCnnSQL.Execute strSQL
               vsfg1.RemoveItem vsfg1.Row
               For lngRRR = 1 To vsfg1.Rows - 1
                   If lngDelCntS < vsfg1.ValueMatrix(lngRRR, 15) Then
                      vsfg1.TextMatrix(lngRRR, 15) = vsfg1.ValueMatrix(lngRRR, 15) - (lngDelCntE - lngDelCntS + 1)
                      vsfg1.TextMatrix(lngRRR, 16) = vsfg1.ValueMatrix(lngRRR, 16) - (lngDelCntE - lngDelCntS + 1)
                   End If
               Next lngRRR
            End If
       End With
       PB_adoCnnSQL.CommitTrans
       vsfg1.Row = 0: vsfg2.Row = 0
       Screen.MousePointer = vbDefault
    End If
    cmdSave.Enabled = True
    vsfg1.SetFocus
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
'/// 삭제 ///
'+-----------+
Private Sub cmdDelete_Click()
Dim strSQL     As String
Dim lngR       As Long
Dim lngRR      As Long
Dim lngRRR     As Long
Dim lngC       As Long
Dim blnOK      As Boolean
Dim intRetVal  As Integer
Dim lngChkCnt  As Long
Dim lngDelCntS As Long
Dim lngDelCntE As Long
Dim lngLogCnt  As Long
    If vsfg1.Row >= vsfg1.FixedRows Then
       If cmdDelete.Enabled = False Or optDate(2).Value = True Then Exit Sub
       intRetVal = MsgBox("발주서를 삭제하시겠습니까 ?", vbCritical + vbYesNo + vbDefaultButton2, "발주서 삭제")
       If intRetVal = vbNo Then
          vsfg1.SetFocus
          Exit Sub
       End If
       cmdDelete.Enabled = False
       Screen.MousePointer = vbHourglass
       PB_adoCnnSQL.BeginTrans
       P_adoRec.CursorLocation = adUseClient
       With vsfg1
            lngDelCntS = .ValueMatrix(.Row, 15): lngDelCntE = .ValueMatrix(.Row, 16)
            '발주내역
            For lngRR = .ValueMatrix(.Row, 16) To .ValueMatrix(.Row, 15) Step -1
                strSQL = "UPDATE 발주내역 SET " _
                              & "사용구분 = 9, " _
                              & "수정일자 = '" & PB_regUserinfoU.UserServerDate & "', " _
                              & "사용자코드 = '" & PB_regUserinfoU.UserCode & "' " _
                        & "WHERE 사업장코드 = '" & vsfg2.TextMatrix(lngRR, 1) & "' " _
                          & "AND 발주일자 = '" & DTOS(vsfg2.TextMatrix(lngRR, 2)) & "' " _
                          & "AND 발주번호 = " & vsfg2.ValueMatrix(lngRR, 3) & " " _
                          & "AND 견적시간 = '" & vsfg2.TextMatrix(lngRR, 29) & "' " _
                          & "AND 자재코드 = '" & vsfg2.TextMatrix(lngRR, 6) & "' " _
                          & "AND 매출처코드 = '" & vsfg2.TextMatrix(lngRR, 10) & "' "
                On Error GoTo ERROR_TABLE_DELETE
                PB_adoCnnSQL.Execute strSQL
                vsfg2.RemoveItem lngRR
            Next lngRR
            '발주
            strSQL = "UPDATE 발주 SET " _
                          & "사용구분 = 9, " _
                          & "수정일자 = '" & PB_regUserinfoU.UserServerDate & "', " _
                          & "사용자코드 = '" & PB_regUserinfoU.UserCode & "' " _
                    & "WHERE 사업장코드 = '" & .TextMatrix(.Row, 2) & "' AND 발주일자 = '" & DTOS(.TextMatrix(.Row, 3)) & "' " _
                      & "AND 발주번호 = " & .ValueMatrix(.Row, 4) & " "
            On Error GoTo ERROR_TABLE_DELETE
            PB_adoCnnSQL.Execute strSQL
            lblTotMny.Caption = Format(Vals(lblTotMny.Caption) - .ValueMatrix(.Row, 7), "#,#.00") '전체금액에서 제외
            .RemoveItem .Row
            For lngRRR = 1 To .Rows - 1
                If lngDelCntS < .ValueMatrix(lngRRR, 15) Then
                   .TextMatrix(lngRRR, 15) = .ValueMatrix(lngRRR, 15) - (lngDelCntE - lngDelCntS + 1)
                   .TextMatrix(lngRRR, 16) = .ValueMatrix(lngRRR, 16) - (lngDelCntE - lngDelCntS + 1)
                End If
            Next lngRRR
            .Row = 0
       End With
       PB_adoCnnSQL.CommitTrans
       cmdFind.SetFocus
       Screen.MousePointer = vbDefault
    End If
    cmdDelete.Enabled = True
    Exit Sub
ERROR_TABLE_DELETE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "발주내역 삭제 실패"
    Unload Me
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
    Set frm매입작성1 = Nothing
    frmMain.SBar.Panels(4).Text = ""
End Sub

'+--------------------+
'/// Sub Procedure ///
'+--------------------+
Private Sub SubOther_FILL()
Dim strSQL        As String
Dim intIndex      As Integer
    'P_adoRec.CursorLocation = adUseClient
    'strSQL = "SELECT T1.사업장코드, T1.사업장명 " _
             & "FROM 사업장 T1 " _
            & "ORDER BY T1.사업장코드 "
    'On Error GoTo ERROR_TABLE_SELECT
    'P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    'If P_adoRec.RecordCount = 0 Then
    '   P_adoRec.Close
    '   cboBranch.Enabled = True
    '   Screen.MousePointer = vbDefault
    'Else
    '   cboBranch.AddItem "00. 전체사업장"
    '   Do Until P_adoRec.EOF
    '      cboBranch.AddItem Format(P_adoRec("사업장코드"), "00") & ". " & P_adoRec("사업장명")
    '      P_adoRec.MoveNext
    '   Loop
    '   P_adoRec.Close
    '   cboBranch.ListIndex = 0
    'End If
    Text1(0).Text = "": Text1(1).Text = ""
    dtpF_Date.Value = Format(Mid(PB_regUserinfoU.UserClientDate, 1, 6) & "01", "0000-00-00")
    dtpT_Date.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "사업장 읽기 실패"
    Unload Me
    Exit Sub
End Sub
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
         .FixedCols = 2
         .Rows = 1             'Subvsfg1_Fill수행시에 설정
         .Cols = 17
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1200   '입고희망일자(매입예정일자)
         .ColWidth(1) = 1730   '사업장코드+발주일자+발주번호(KEY)
         .ColWidth(2) = 1000   '사업장코드 'Hidden
         .ColWidth(3) = 1200   '발주일자   'Hidden
         .ColWidth(4) = 1000   '발주번호   'Hidden
         .ColWidth(5) = 1200   '매입처코드
         .ColWidth(6) = 2200   '매입처명
         .ColWidth(7) = 1800   '결제금액
         .ColWidth(8) = 1200   '매입일자
         .ColWidth(9) = 900    '유효일수
         .ColWidth(10) = 800   '결제방법
         .ColWidth(11) = 1200  '결제예정일자
         .ColWidth(12) = 2800  '제목
         .ColWidth(13) = 5000  '적요
         .ColWidth(14) = 1200  '발주자명
         .ColWidth(15) = 1000  'ROW(vsfg2.Row)   Not Used
         .ColWidth(16) = 1000  'COL(vsfg2.Row)   Not Used
         
         .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "매입예정일자"
         .TextMatrix(0, 1) = "발주번호"
         .TextMatrix(0, 2) = "사업장코드" 'H
         .TextMatrix(0, 3) = "발주일자"   'H
         .TextMatrix(0, 4) = "발주번호"   'H
         .TextMatrix(0, 5) = "매입처코드" 'H
         .TextMatrix(0, 6) = "매입처명"
         .TextMatrix(0, 7) = "금액"
         .TextMatrix(0, 8) = "매입일자"
         .TextMatrix(0, 9) = "유효일수"
         .TextMatrix(0, 10) = "결제"
         .TextMatrix(0, 11) = "결제예정일자"
         .TextMatrix(0, 12) = "제목"
         .TextMatrix(0, 13) = "적요"
         .TextMatrix(0, 14) = "발주자명"
         .TextMatrix(0, 15) = "Row"       'H
         .TextMatrix(0, 16) = "Col"       'H
         
         .ColHidden(2) = True: .ColHidden(3) = True: .ColHidden(4) = True: .ColHidden(5) = True
         .ColHidden(15) = True: .ColHidden(16) = True
         .ColFormat(7) = "#,#.00"
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 2, 3, 6, 12, 13
                         .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 0, 1, 5, 8, 10, 11, 14, 15, 16
                         .ColAlignment(lngC) = flexAlignCenterCenter
                    Case Else
                         .ColAlignment(lngC) = flexAlignRightCenter
             End Select
         Next lngC
         .MergeCells = flexMergeRestrictRows 'flexMergeFixedOnly
         .MergeRow(0) = True
         For lngC = 0 To 0
             .MergeCol(lngC) = True
         Next lngC
    End With
End Sub
Private Sub Subvsfg2_INIT()
Dim lngR    As Long
Dim lngC    As Long
    With vsfg2              'Rows 1, Cols 30, RowHeightMax(Min) 300
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
         .FixedCols = 6
         .Rows = 1             'Subvsfg2_Fill수행시에 설정
         .Cols = 30
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 400    'No
         .ColWidth(1) = 1000   '사업장코드
         .ColWidth(2) = 1200   '발주일자
         .ColWidth(3) = 1000   '발주번호
         .ColWidth(4) = 1900   '품목코드(변경후)
         .ColWidth(5) = 2600   '품명(변경후)
         .ColWidth(6) = 1900   '자재코드(변경전)   'H
         .ColWidth(7) = 2600   '자재명(변경전)     'H
         .ColWidth(8) = 1000   '매출처코드(변경후) 'H
         .ColWidth(9) = 2000   '매출처명(변경후)
         .ColWidth(10) = 1000  '매출처코드(변경전) 'H
         .ColWidth(11) = 2000  '매출처명(변경전) 'H
         .ColWidth(12) = 2000  '사업장코드+발주일자+발주번호+자재코드+매출처코드+(KEY) 'H
         .ColWidth(13) = 1000  '매입처코드 'H
         .ColWidth(14) = 2500  '매입처명   'H
         .ColWidth(15) = 2200  '자재규격
         .ColWidth(16) = 1000  '발주량
         .ColWidth(17) = 800   '발주단위
         .ColWidth(18) = 800   '직송       'H
         .ColWidth(19) = 1600  '입고단가
         .ColWidth(20) = 1200  '입고부가   'H
         .ColWidth(21) = 1600  '입고가격(단가+부가) 'H
         .ColWidth(22) = 1700  '입고금액
         .ColWidth(23) = 1600  '출고단가   'H
         .ColWidth(24) = 1200  '출고부가   'H
         .ColWidth(25) = 1600  '출고가격(단가+부가) 'H
         .ColWidth(26) = 1700  '출고금액   'H
         .ColWidth(27) = 5000  '적요
         .ColWidth(28) = 800   'SQL구분
         .ColWidth(29) = 1000  '발주시간   'H
         .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "No"
         .TextMatrix(0, 1) = "사업장코드"   'H
         .TextMatrix(0, 2) = "발주일자"     'H
         .TextMatrix(0, 3) = "발주번호"     'H
         .TextMatrix(0, 4) = "코드"         '변경후(Or 변경전)
         .TextMatrix(0, 5) = "품명"         '변경후(Or 변경전)
         .TextMatrix(0, 6) = "품목코드"     'H, 변경전
         .TextMatrix(0, 7) = "품명명"       'H, 변경전
         .TextMatrix(0, 8) = "매출처코드"   'H, 변경후
         .TextMatrix(0, 9) = "매출처명"
         .TextMatrix(0, 10) = "매출처코드"  'H, 변경전
         .TextMatrix(0, 11) = "매출처명"    'H, 변경전
         .TextMatrix(0, 12) = "KEY"         'H
         .TextMatrix(0, 13) = "매입처코드"  'H
         .TextMatrix(0, 14) = "매입처명"    'H
         .TextMatrix(0, 15) = "규격"
         .TextMatrix(0, 16) = "수량"
         .TextMatrix(0, 17) = "단위"
         .TextMatrix(0, 18) = "직송"        'H
         .TextMatrix(0, 19) = "매입단가"
         .TextMatrix(0, 20) = "매입부가"    'H
         .TextMatrix(0, 21) = "매입가격"    'H(단가 + 부가)
         .TextMatrix(0, 22) = "매입금액"
         .TextMatrix(0, 23) = "매출단가"    'H
         .TextMatrix(0, 24) = "매출부가"    'H
         .TextMatrix(0, 25) = "매출가격"    '(단가 + 부가) 'H
         .TextMatrix(0, 26) = "매출금액"    'H
         .TextMatrix(0, 27) = "적요"
         .TextMatrix(0, 28) = "구분"        'H(SQL구분:I.Insert, U.Update, D.Delete)
         .TextMatrix(0, 29) = "발주시간"
         .ColHidden(1) = True: .ColHidden(2) = True: .ColHidden(3) = True
         .ColHidden(6) = True: .ColHidden(7) = True: .ColHidden(8) = True:
         .ColHidden(9) = True: .ColHidden(10) = True: .ColHidden(11) = True: .ColHidden(12) = True
         .ColHidden(13) = True: .ColHidden(14) = True:: .ColHidden(18) = True
         .ColHidden(20) = True: .ColHidden(21) = True
         .ColHidden(23) = True: .ColHidden(24) = True: .ColHidden(25) = True: .ColHidden(26) = True
         .ColHidden(28) = True: .ColHidden(29) = True
         'For lngC = 0 To .Cols - 1
         '    .TextMatrix(0, lngC) = .TextMatrix(0, lngC) + CStr(lngC)
         'Next lngC
         .ColFormat(16) = "#,#"
         For lngC = 19 To 26
             .ColFormat(lngC) = "#,#.00"
         Next lngC
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 17, 27
                         .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 0, 1, 2, 3, 18, 28, 29
                         .ColAlignment(lngC) = flexAlignCenterCenter
                    Case Else
                         .ColAlignment(lngC) = flexAlignRightCenter
             End Select
         Next lngC
         .MergeCells = flexMergeRestrictRows 'flexMergeFixedOnly
         .MergeRow(0) = True
         For lngC = 0 To 5
             .MergeCol(lngC) = True
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
    vsfg1.Rows = 1
    If optDate(0).Value = True Then '발주일자 기준
       strWhere = "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
               & "AND T1.상태코드 = 1 AND T1.사용구분 = 0 " _
               & "AND (T1.발주일자 BETWEEN '" & DTOS(dtpF_Date.Value) & "' AND '" & DTOS(dtpT_Date.Value) & "' ) "
       strOrderBy = "ORDER BY T1.발주일자, T1.발주번호 "
    End If
    If optDate(1).Value = True Then   '매입예정일자 기준
       strWhere = "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
               & "AND T1.상태코드 = 1 AND T1.사용구분 = 0 " _
               & "AND (T1.입고희망일자 BETWEEN '" & DTOS(dtpF_Date.Value) & "' AND '" & DTOS(dtpT_Date.Value) & "' ) "
       strOrderBy = "ORDER BY T1.입고희망일자, T1.발주일자, 발주번호 "
    End If
    If optDate(2).Value = True Then   '매입일자 기준
       strWhere = "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
               & "AND T1.상태코드 = 2 AND T1.사용구분 = 0 " _
               & "AND (T1.입고일자 BETWEEN '" & DTOS(dtpF_Date.Value) & "' AND '" & DTOS(dtpT_Date.Value) & "' ) "
       strOrderBy = "ORDER BY T1.입고일자, T1.발주일자, T1.발주번호 "
    End If
    If Len(Trim(Text1(0).Text)) = 0 Then
       strWhere = strWhere
    Else
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
               & "T1.매입처코드 = '" & Trim(Text1(0).Text) & "' "
    End If
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.입고희망일자, T1.사업장코드, T1.발주일자, T1.발주번호, " _
                  & "T1.매입처코드, T2.매입처명, T1.입고일자 AS 입고일자, T1.유효일수 AS 유효일수, " _
                  & "T1.결제방법, T1.결제예정일자, T1.제목, T1.적요, T3.사용자명 " _
             & "FROM 발주 T1 " _
             & "LEFT JOIN 매입처 T2 ON T2.사업장코드 = T1.사업장코드 AND T2.매입처코드 = T1.매입처코드 " _
             & "LEFT JOIN 사용자 T3 ON T3.사용자코드 = T1.사용자코드 " _
            & "" & strWhere & " " _
            & "" & strOrderBy & " "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    vsfg1.Rows = P_adoRec.RecordCount + 1
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       cmdExit.SetFocus
       Exit Sub
    Else
       With vsfg1
            .Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlack
            If .Rows <= PC_intRowCnt1 Then
               .ScrollBars = flexScrollBarHorizontal
            Else
               .ScrollBars = flexScrollBarBoth
            End If
            Do Until P_adoRec.EOF
               lngR = lngR + 1
               .TextMatrix(lngR, 0) = Format(P_adoRec("입고희망일자"), "0000-00-00")
               .TextMatrix(lngR, 1) = P_adoRec("사업장코드") & "-" & Format(P_adoRec("발주일자"), "0000/00/00") _
                                    & "-" & CStr(P_adoRec("발주번호"))
               .Cell(flexcpData, lngR, 1, lngR, 1) = Trim(.TextMatrix(lngR, 1)) 'FindRow 사용을 위해
               .TextMatrix(lngR, 2) = IIf(IsNull(P_adoRec("사업장코드")), "", P_adoRec("사업장코드"))
               .TextMatrix(lngR, 3) = Format(P_adoRec("발주일자"), "0000-00-00")
               .TextMatrix(lngR, 4) = IIf(IsNull(P_adoRec("발주번호")), 0, P_adoRec("발주번호"))
               .TextMatrix(lngR, 5) = IIf(IsNull(P_adoRec("매입처코드")), "", P_adoRec("매입처코드"))
               .TextMatrix(lngR, 6) = IIf(IsNull(P_adoRec("매입처명")), "", P_adoRec("매입처명"))
               .TextMatrix(lngR, 8) = IIf(IsNull(P_adoRec("입고일자")), "", Format(P_adoRec("입고일자"), "0000-00-00"))
               .TextMatrix(lngR, 9) = IIf(IsNull(P_adoRec("유효일수")), 0, P_adoRec("유효일수"))
               .TextMatrix(lngR, 10) = IIf(IsNull(P_adoRec("결제방법")), 0, P_adoRec("결제방법"))
               If P_adoRec("결제방법") = 0 Then  '결제방법(0.현금, 1.어음)
                  .Cell(flexcpChecked, lngR, 10) = flexChecked    '1
               Else
                  .Cell(flexcpChecked, lngR, 10) = flexUnchecked  '2
               End If
               .Cell(flexcpText, lngR, 10) = "현 금"
               .TextMatrix(lngR, 11) = Format(P_adoRec("결제예정일자"), "0000-00-00")
               .TextMatrix(lngR, 12) = IIf(IsNull(P_adoRec("제목")), "", P_adoRec("제목"))
               .TextMatrix(lngR, 13) = IIf(IsNull(P_adoRec("적요")), "", P_adoRec("적요"))
               .TextMatrix(lngR, 14) = IIf(IsNull(P_adoRec("사용자명")), "", P_adoRec("사용자명"))
               'If .TextMatrix(lngR, 0) = PB_regUserinfoU.UserBranchCode Then
               '   lngRR = lngR
               'End If
               P_adoRec.MoveNext
            Loop
            P_adoRec.Close
            If lngRR = 0 Then
               .Row = lngRR       'vsfg1_EnterCell 자동실행(만약 한건 일때는 자동실행 안함)
               If .Rows > PC_intRowCnt1 Then
                  '.TopRow = .Rows - PC_intRowCnt1 + .FixedRows
                  .TopRow = 1
               End If
            Else
               .Row = lngRR       'vsfg1_EnterCell 자동실행(만약 한건 일때는 자동실행 안함)
               If .Rows > PC_intRowCnt1 Then
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
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "발주 읽기 실패"
    Unload Me
    Exit Sub
End Sub

Private Sub Subvsfg2_FILL()
Dim strSQL      As String
Dim strWhere    As String
Dim strOrderBy  As String
Dim lngR        As Long
Dim lngC        As Long
Dim lngRR       As Long
Dim lngRRR      As Long
Dim strCell     As String
Dim strSubTotal As String
    vsfg2.Rows = 1
    If optDate(0).Value = True Then '발주일자 기준
       strWhere = "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
               & "AND T1.상태코드 = 1 AND T1.사용구분 = 0 AND T2.상태코드 = 1 AND T2.사용구분 = 0 " _
               & "AND (T1.발주일자 BETWEEN '" & DTOS(dtpF_Date.Value) & "' AND '" & DTOS(dtpT_Date.Value) & "' ) "
       strOrderBy = "ORDER BY T1.발주일자, T1.발주번호, T2.발주시간 "
    End If
    If optDate(1).Value = True Then   '매입예정일자 기준
       strWhere = "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
               & "AND T1.상태코드 = 1 AND T1.사용구분 = 0 AND T2.상태코드 = 1 AND T2.사용구분 = 0 " _
               & "AND (T1.입고희망일자 BETWEEN '" & DTOS(dtpF_Date.Value) & "' AND '" & DTOS(dtpT_Date.Value) & "' ) "
       strOrderBy = "ORDER BY T1.입고희망일자, T1.발주일자, T1.발주번호, T2.발주시간 "
    End If
    If optDate(2).Value = True Then   '매입일자 기준
       strWhere = "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
               & "AND T1.상태코드 = 2 AND T1.사용구분 = 0 AND T2.상태코드 = 2 AND T2.사용구분 = 0 " _
               & "AND (T1.입고일자 BETWEEN '" & DTOS(dtpF_Date.Value) & "' AND '" & DTOS(dtpT_Date.Value) & "' ) "
       strOrderBy = "ORDER BY T1.입고일자, T1.발주일자, T1.발주번호, T2.발주시간 "
    End If
    If Len(Text1(0).Text) = 0 Then
       strWhere = strWhere
    Else
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                & "T1.매입처코드 = '" & Trim(Text1(0).Text) & "' "
    End If
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.입고희망일자, T1.사업장코드, T1.발주일자, T1.발주번호, T2.발주시간, " _
                  & "T2.자재코드, ISNULL(T5.자재명,'') AS 자재명, " _
                  & "ISNULL(T2.매출처코드,'') AS 매출처코드, ISNULL(T3.매출처명,'') AS 매출처명, " _
                  & "T1.매입처코드, T4.매입처명, T5.규격 AS 자재규격, T2.발주량, " _
                  & "T5.단위 AS 발주단위, T2.직송구분, T2.입고단가, T2.입고부가, " _
                  & "T2.출고단가 , T2.출고부가, T2.적요 AS 적요 " _
                  & "FROM 발주 T1 " _
           & "INNER JOIN 발주내역 T2 " _
                   & "ON T2.사업장코드 = T1.사업장코드 AND T2.발주일자 = T1.발주일자 AND T2.발주번호 = T1.발주번호 " _
            & "LEFT JOIN 매출처 T3 ON T3.사업장코드 = T2.사업장코드 AND T3.매출처코드 = T2.매출처코드 " _
            & "LEFT JOIN 매입처 T4 ON T4.사업장코드 = T2.사업장코드 AND T4.매입처코드 = T2.매입처코드 " _
            & "LEFT JOIN 자재 T5 ON (T5.분류코드 + T5.세부코드) = T2.자재코드 " _
           & "" & strWhere & " " _
           & "" & strOrderBy & " "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    vsfg2.Rows = P_adoRec.RecordCount + 1
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       cmdExit.SetFocus
       Exit Sub
    Else
       With vsfg2
            .Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlack
            If .Rows <= PC_intRowCnt2 Then
               '.ScrollBars = flexScrollBarHorizontal
            Else
               '.ScrollBars = flexScrollBarBoth
            End If
            Do Until P_adoRec.EOF
               lngR = lngR + 1
               '.Cell(flexcpData, lngR, 0, lngR, 0) = Trim(.TextMatrix(lngR, 0)) 'FindRow 사용을 위해
               '.TextMatrix(lngR, 0) = Format(P_adoRec("입고희망일자"), "0000-00-00")
               .TextMatrix(lngR, 1) = IIf(IsNull(P_adoRec("사업장코드")), "", P_adoRec("사업장코드"))
               .TextMatrix(lngR, 2) = Format(P_adoRec("발주일자"), "0000-00-00")
               .TextMatrix(lngR, 3) = IIf(IsNull(P_adoRec("발주번호")), 0, P_adoRec("발주번호"))
               .TextMatrix(lngR, 4) = IIf(IsNull(P_adoRec("자재코드")), "", P_adoRec("자재코드"))
               .TextMatrix(lngR, 5) = IIf(IsNull(P_adoRec("자재명")), "", P_adoRec("자재명"))
               .TextMatrix(lngR, 6) = IIf(IsNull(P_adoRec("자재코드")), "", P_adoRec("자재코드"))
               .TextMatrix(lngR, 7) = IIf(IsNull(P_adoRec("자재명")), "", P_adoRec("자재명"))
               .TextMatrix(lngR, 8) = IIf(IsNull(P_adoRec("매출처코드")), "", P_adoRec("매출처코드"))
               .TextMatrix(lngR, 9) = IIf(IsNull(P_adoRec("매출처명")), "", P_adoRec("매출처명"))
               .TextMatrix(lngR, 10) = IIf(IsNull(P_adoRec("매출처코드")), "", P_adoRec("매출처코드"))
               .TextMatrix(lngR, 11) = IIf(IsNull(P_adoRec("매출처명")), "", P_adoRec("매출처명"))
               .TextMatrix(lngR, 12) = P_adoRec("사업장코드") & "-" & Format(P_adoRec("발주일자"), "0000/00/00") _
                                     & "-" & CStr(P_adoRec("발주번호")) & "-" & P_adoRec("자재코드") & "-" & P_adoRec("매출처코드")
               .TextMatrix(lngR, 13) = IIf(IsNull(P_adoRec("매입처코드")), "", P_adoRec("매입처코드"))
               .TextMatrix(lngR, 14) = IIf(IsNull(P_adoRec("매입처명")), "", P_adoRec("매입처명"))
               .TextMatrix(lngR, 15) = IIf(IsNull(P_adoRec("자재규격")), "", P_adoRec("자재규격"))
               .TextMatrix(lngR, 16) = Format(P_adoRec("발주량"), "#,#")
               .TextMatrix(lngR, 17) = IIf(IsNull(P_adoRec("발주단위")), "", P_adoRec("발주단위"))
               If P_adoRec("직송구분") = 0 Then
                  .Cell(flexcpChecked, lngR, 18) = flexUnchecked
               Else
                  .Cell(flexcpChecked, lngR, 18) = flexChecked
               End If
               .Cell(flexcpText, lngR, 18) = "직 송"
               .Cell(flexcpAlignment, lngR, 18, lngR, 18) = flexAlignLeftCenter
               .TextMatrix(lngR, 19) = Format(P_adoRec("입고단가"), "#,#.00")
               .TextMatrix(lngR, 20) = Format(P_adoRec("입고부가"), "#,#.00")
               .TextMatrix(lngR, 21) = .ValueMatrix(lngR, 19) + .ValueMatrix(lngR, 20)
               .TextMatrix(lngR, 22) = .ValueMatrix(lngR, 16) * .ValueMatrix(lngR, 19)
               .TextMatrix(lngR, 23) = Format(P_adoRec("출고단가"), "#,#.00")
               .TextMatrix(lngR, 24) = Format(P_adoRec("출고부가"), "#,#.00")
               .TextMatrix(lngR, 25) = .ValueMatrix(lngR, 23) + .ValueMatrix(lngR, 24)
               .TextMatrix(lngR, 26) = .ValueMatrix(lngR, 16) * .ValueMatrix(lngR, 23)
               .TextMatrix(lngR, 27) = IIf(IsNull(P_adoRec("적요")), "", P_adoRec("적요"))
               .TextMatrix(lngR, 29) = IIf(IsNull(P_adoRec("발주시간")), "", P_adoRec("발주시간"))
               'If .TextMatrix(lngR, 0) = PB_regUserinfoU.UserBranchCode Then
               '   lngRR = lngR
               'End If
               .RowHidden(lngR) = True
               P_adoRec.MoveNext
            Loop
            P_adoRec.Close
            If lngRR = 0 Then
               .Row = lngRR       'vsfg2_EnterCell 자동실행(만약 한건 일때는 자동실행 안함)
               If .Rows > PC_intRowCnt2 Then
                  '.TopRow = .Rows - PC_intRowCnt2 + .FixedRows
                  .TopRow = 1
               End If
            Else
               .Row = lngRR       'vsfg2_EnterCell 자동실행(만약 한건 일때는 자동실행 안함)
               If .Rows > PC_intRowCnt2 Then
                  .TopRow = .Row
               End If
            End If
            '.MultiTotals = True '(default value : true)
            '.Subtotal flexSTClear
            '.SubtotalPosition = flexSTBelow
            '.Subtotal flexSTCount, 6, 8, "#", vbRed, vbWhite, , "%s", , False
            '.Subtotal flexSTSum, 6, 10, , vbRed, vbWhite, , "%s", , False
            For lngR = 1 To .Rows - 1
                strCell = .TextMatrix(lngR, 1) & "-" & Format(DTOS(.TextMatrix(lngR, 2)), "0000/00/00") & "-" & .TextMatrix(lngR, 3)
                For lngRRR = 1 To vsfg1.Rows - 1
                    If strCell = vsfg1.TextMatrix(lngRRR, 1) Then
                       If vsfg1.ValueMatrix(lngRRR, 15) = 0 Then
                          vsfg1.TextMatrix(lngRRR, 15) = lngR
                       End If
                       vsfg1.TextMatrix(lngRRR, 16) = lngR
                       '발주 합계금액 계산
                       vsfg1.TextMatrix(lngRRR, 7) = vsfg1.ValueMatrix(lngRRR, 7) + .ValueMatrix(lngR, 22)
                       lblTotMny.Caption = Format(Vals(lblTotMny.Caption) + .ValueMatrix(lngR, 22), "#,#.00")
                       Exit For
                    End If
                Next lngRRR
            Next lngR
            'vsfg2_EnterCell       'vsfg1_EnterCell 자동실행(만약 한건 일때도 강제로 자동실행) 'Not Used
            '.SetFocus                                                                         'Not Used
       End With
    End If
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "발주내역 읽기 실패"
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

Dim strEMail               As String

    Screen.MousePointer = vbHourglass
    '서버일시(출력일시)
    'strSQL = "SELECT CONVERT(VARCHAR(19),GETDATE(), 120) AS 서버시간 "
    'On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
    'P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    'strForPrtDateTime = Format(PB_regUserinfoU.UserServerDate, "0000-00-00") & Space(1) _
                      & Format(Right(P_adoRec("서버시간"), 8), "hh:mm:ss")
    'P_adoRec.Close
    
    intRetCHK = 99
    With CrystalReport1
         If PB_Test = 0 Then
            strExeFile = App.Path & ".\발주서.rpt"
         Else
            strExeFile = App.Path & ".\발주서T.rpt"
         End If
         varRetVal = Dir(strExeFile)
         If Len(varRetVal) = 0 Then
            intRetCHK = 0
         Else
            .ReportFileName = strExeFile
            On Error GoTo ERROR_CRYSTAL_REPORTS
            '--- Formula Fields ---
            .Formulas(0) = "ForPrtDate = '" & Mid(PB_regUserinfoU.UserServerDate, 1, 4) & "' + ' 년 ' " _
                                    & "+ '" & Mid(PB_regUserinfoU.UserServerDate, 5, 2) & "' + ' 월 ' " _
                                    & "+ '" & Mid(PB_regUserinfoU.UserServerDate, 7, 2) & "' + ' 일' "
            strSQL = "SELECT T1.사업자번호 AS 등록번호, T1.사업장명 AS 상호, " _
                          & "T1.대표자명 AS 대표, (T1.주소 + T1.번지) AS 주소, " _
                          & "T1.전화번호 AS 전화, T1.팩스번호 AS 팩스, T1.이메일주소 AS 이메일주소 " _
                     & "FROM 사업장 T1 " _
                    & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' "
            On Error GoTo ERROR_TABLE_SELECT
            P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
            If P_adoRec.RecordCount = 0 Then
               P_adoRec.Close
            Else
               .Formulas(1) = "ForEnterNo = '" & P_adoRec("등록번호") & "' "
               .Formulas(2) = "ForEnterName = '" & P_adoRec("상호") & "' "
               .Formulas(3) = "ForRepName = '" & P_adoRec("대표") & "' "
               .Formulas(4) = "ForAddress = '" & P_adoRec("주소") & "' "
               .Formulas(5) = "ForTelNo = '" & P_adoRec("전화") & "' "
               .Formulas(6) = "ForFaxNo = '" & P_adoRec("팩스") & "' "
               strEMail = P_adoRec("이메일주소")
               P_adoRec.Close
            End If
            '금액(한자, 숫자)
            If optPrtChk1.Value = True Then '발주서
               strSQL = "SELECT SUM(T1.입고단가 * T1.발주량) AS 금액 " _
                        & "FROM 발주내역 T1 " _
                       & "WHERE T1.사업장코드 = '" & vsfg1.TextMatrix(vsfg1.Row, 2) & "' " _
                         & "AND T1.발주일자 = '" & DTOS(vsfg1.TextMatrix(vsfg1.Row, 3)) & "' " _
                         & "AND T1.발주번호 = " & vsfg1.ValueMatrix(vsfg1.Row, 4) & " " _
                         & "AND (T1.상태코드 = 1 OR T1.상태코드 = 2) AND T1.사용구분 = 0 "
               On Error GoTo ERROR_TABLE_SELECT
               P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
               If P_adoRec.RecordCount = 0 Then
                  P_adoRec.Close
               Else
                  .Formulas(7) = "ForHanjaOrderMny = '" & hMValH(P_adoRec("금액")) & "' + space(1) + '元整' "
                  .Formulas(8) = "ForOrderMny = " & P_adoRec("금액") & " "
                  P_adoRec.Close
               End If
            Else                         '견적의뢰서
               .Formulas(7) = "ForHanjaOrderMny = '元整' "
               .Formulas(8) = "ForOrderMny = 0 "
            End If
            .Formulas(9) = "ForOrderGbn = " & IIf(optPrtChk1.Value = True, 1, 0) & " " '0.견적의뢰서, 1.발주서
            .Formulas(10) = "ForEMail = '" & strEMail & "' "
            '--- Parameter Fields ---
            .StoredProcParam(0) = vsfg1.TextMatrix(vsfg1.Row, 2)       '지점코드
            .StoredProcParam(1) = DTOS(vsfg1.TextMatrix(vsfg1.Row, 3)) '발주일자
            .StoredProcParam(2) = vsfg1.ValueMatrix(vsfg1.Row, 4)      '발주번호
            '0.견적의뢰서, 1.발주서
            .StoredProcParam(3) = IIf(optPrtChk1.Value = True, 1, 0)
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
            .WindowShowGroupTree = False
            .WindowShowPrintSetupBtn = True
            .WindowTop = 0: .WindowTop = 0: .WindowHeight = 11100: .WindowWidth = 15405
            .WindowState = crptMaximized
            .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & IIf(optPrtChk1.Value = True, "발주서", "견적의뢰서")
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
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "사업장 읽기 실패"
    Unload Me
    Exit Sub
ERROR_CRYSTAL_REPORTS:
    MsgBox Err.Number & Space(1) & Err.Description
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

