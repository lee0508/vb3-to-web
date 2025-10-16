VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frm계정코드검색 
   BackColor       =   &H00008000&
   BorderStyle     =   1  '단일 고정
   Caption         =   "계정코드 검색"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11100
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "계정코드검색.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   11100
   StartUpPosition =   2  '화면 가운데
   Begin VB.CheckBox chkTotal 
      BackColor       =   &H00008000&
      Caption         =   "전체 계정"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9600
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VSFlex7Ctl.VSFlexGrid vsfg1 
      Height          =   2150
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   10880
      _cx             =   19191
      _cy             =   3792
      _ConvInfo       =   1
      Appearance      =   0
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
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
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
   Begin VB.TextBox txtName 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   10  '한글 
      Left            =   3650
      TabIndex        =   2
      Top             =   105
      Width           =   5000
   End
   Begin VB.TextBox txtCode 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   8  '영문
      Left            =   1215
      MaxLength       =   4
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "계정명"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   2580
      TabIndex        =   4
      Top             =   165
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "계정코드"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   165
      Width           =   975
   End
End
Attribute VB_Name = "frm계정코드검색"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' 프로그램 제 목 : 계정코드검색
' 사용된 Control :
' 참조된 Table   : 계정과목
' 업  무  설  명 :
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived       As Boolean
Private P_intFindGbn       As Integer      '0.수동검색, 1.코드, 2.이름(명), 3.코드 + 이름(명)
Private P_intButton        As Integer
Private P_strFindString2   As String
Private P_adoRec           As New ADODB.Recordset
Private Const PC_intRowCnt As Integer = 6  '그리드 한 페이지 당 행수(FixedRows 포함)

'+--------------------------------+
'/// LOAD FORM ( 한번만 실행 ) ///
'+--------------------------------+
Private Sub Form_Load()
    P_blnActived = False
    txtCode.Text = PB_strAccCode: txtName.Text = PB_strAccName
    If Len(PB_strAccCode) = 0 And Len(PB_strAccName) = 0 Then
       P_intFindGbn = 0    '수동검색
    ElseIf _
       Len(PB_strAccCode) <> 0 And Len(PB_strAccName) = 0 Then
       P_intFindGbn = 1    '코드로만 자동검색
    ElseIf _
       Len(PB_strAccCode) = 0 And Len(PB_strAccName) <> 0 Then
       P_intFindGbn = 2    '이름(명)으로만 자동검색
    ElseIf _
       Len(PB_strAccCode) <> 0 And Len(PB_strAccName) <> 0 Then
       P_intFindGbn = 3    '코드와 이름(명)을 동시에 자동검색
    Else
       P_intFindGbn = 0    '수동검색
    End If
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
       If P_intFindGbn = 0 Then
          txtName.SetFocus
       Else
          Subvsfg1_FILL
       End If
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

'+-----------+
'/// 취소 ///
'+-----------+
Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
       SubEscape
    End If
End Sub
Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
       SubEscape
    End If
End Sub

Private Sub vsfg1_DblClick()
    vsfg1_KeyDown vbKeyReturn, 0
End Sub
Private Sub vsfg1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
       SubEscape
    End If
End Sub

Private Sub SubEscape()
    PB_strAccCode = ""
    PB_strAccName = ""
    Unload Me
End Sub

'+-----------+
'/// 코드 ///
'+-----------+
Private Sub txtCode_GotFocus()
    With txtCode
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       If Len(Trim(txtCode.Text) + Trim(txtName.Text)) = 0 Then
          PB_strAccCode = ""
          PB_strAccName = ""
          Unload Me
          Exit Sub
       End If
       If Len(Trim(txtCode.Text)) <> 0 Then
          P_intFindGbn = 1
          Subvsfg1_FILL
       Else
          txtCode.Text = ""
       End If
    End If
End Sub

'+---------------+
'/// 이름(명) ///
'+---------------+
Private Sub txtName_GotFocus()
    With txtName
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       If Len(Trim(txtCode.Text) + Trim(txtName.Text)) = 0 Then
          PB_strAccCode = ""
          PB_strAccName = ""
          Unload Me
          Exit Sub
       End If
       If Len(Trim(txtName.Text)) <> 0 Then
          P_intFindGbn = 2
          Subvsfg1_FILL
       Else
          txtName.Text = ""
       End If
    End If
End Sub

'+---------------+
'/// 이름(명) ///
'+---------------+
Private Sub chkTotal_Click()
    With chkTotal
         If .Value = 1 Then
            Subvsfg1_FILL
         End If
    End With
End Sub

'+-----------+
'/// Grid ///
'+-----------+
Private Sub vsfg1_BeforeSort(ByVal Col As Long, Order As Integer)
    With vsfg1
         'Not Used
         'P_strFindString2 = Trim(.Cell(flexcpData, .Row, 0)) 'Not Used
    End With
End Sub
Private Sub vsfg1_AfterSort(ByVal Col As Long, Order As Integer)
    With vsfg1
         'Not Used
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
Private Sub vsfg1_Click()
Dim strData As String
    With vsfg1
         If (.MouseRow >= 0 And .MouseRow < .FixedRows) Then
            .Col = .MouseCol
            .Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlack
            .Cell(flexcpForeColor, 0, .MouseCol, 0, .MouseCol) = vbRed
            strData = .TextMatrix(.Row, 0)
            Select Case .MouseCol
                   Case 0
                        .ColSort(0) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .Sort = flexSortUseColSort
                   Case 1
                        .ColSort(1) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
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
             txtCode.Text = .TextMatrix(.Row, 0)
             txtName.Text = .TextMatrix(.Row, 1)
         End If
    End With
End Sub
Private Sub vsfg1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim SQL       As String
Dim intRetVal As Integer
    With vsfg1
         If .Row >= .FixedRows Then
            If KeyCode = vbKeyReturn Then
               If .Row = 0 Then
                  PB_strAccCode = ""
                  PB_strAccName = ""
               Else
                  PB_strAccCode = .TextMatrix(.Row, 0)
                  PB_strAccName = .TextMatrix(.Row, 1)
               End If
               Unload Me
            End If
         End If
    End With
    Exit Sub
End Sub

'+-----------+
'/// 종료 ///
'+-----------+
Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
    If P_adoRec.State <> adStateClosed Then
       P_adoRec.Close
    End If
    Set P_adoRec = Nothing
    Set frm계정코드검색 = Nothing
End Sub

'+--------------------+
'/// Sub Procedure ///
'+--------------------+
'+-------------------------------+
'/// VsFlexGrid(vsfg1) 초기화 ///
'+-------------------------------+
Private Sub Subvsfg1_INIT()
Dim lngR    As Long
Dim lngC    As Long
    With vsfg1              'Rows 1, Cols 4, RowHeightMax(Min) 300
         .AllowBigSelection = False
         .AllowSelection = False
         .AllowUserResizing = flexResizeColumns
         '.BackColorBkg = &H8000000F
         .BackColorBkg = vbWhite
         .BackColorSel = &H8000&
         .Ellipsis = flexEllipsisEnd
         '.ExplorerBar = flexExSortShow
         .ExtendLastCol = True
         .FontSize = 9
         .ScrollBars = flexScrollBarVertical
         .ScrollTrack = True
         .SelectionMode = flexSelectionByRow
         .FixedRows = 1
         .FixedCols = 1
         .Rows = 1             'Subvsfg1_Fill수행시에 설정
         .Cols = 4
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1000   '계정코드
         .ColWidth(1) = 2500   '계정명
         .ColWidth(2) = 2000   '합계시산표연결여부
         .ColWidth(3) = 2000   '적요
         .Cell(flexcpFontBold, 0, 0, .FixedRows - 1, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "계정코드"
         .TextMatrix(0, 1) = "계정명"
         .TextMatrix(0, 2) = "합계시산표연결여부"
         .TextMatrix(0, 3) = "적요"
         '.ColFormat(X) = "#,#.00"
         '.ColHidden(X) = True
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 1, 3
                         .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 0, 2
                         .ColAlignment(lngC) = flexAlignCenterCenter
                    Case Else
                         .ColAlignment(lngC) = flexAlignRightCenter
             End Select
         Next lngC
         '.MergeCells = flexMergeRestrictColumns
         'For lngC = 0 To 1
         '    .MergeCol(lngC) = True
         'Next lngC
    End With
End Sub

'+------------------------------+
'/// VsFlexGrid(vsfg1) 채우기///
'+------------------------------+
Private Sub Subvsfg1_FILL()
Dim strSQL     As String
Dim strWhere   As String
Dim strOrderBy As String
Dim lngR       As Long
Dim lngC       As Long
Dim lngRR      As Long
Dim lngRRR     As Long
    Screen.MousePointer = vbHourglass
    txtCode.Text = Trim(txtCode.Text): txtName.Text = Trim(txtName.Text)
    If chkTotal.Value = 0 Then
       If P_intFindGbn = 1 Then
          strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                             & "T1.계정코드 Like '%" & Trim(txtCode.Text) & "%' " _
                         & "AND T1.사용구분 = 0 "
          strOrderBy = "ORDER BY T1.계정코드 "
       ElseIf _
          P_intFindGbn = 2 Then
          strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                             & "T1.계정명 Like '%" & Trim(txtName.Text) & "%' " _
                         & "AND T1.사용구분 = 0 "
          strOrderBy = "ORDER BY T1.계정명 "
       ElseIf _
          P_intFindGbn = 3 Then
          strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                             & "T1.계정코드 Like '%" & Trim(txtCode.Text) & "%' " _
                         & "AND T1.계정명 Like '%" & Trim(txtName.Text) & "%' " _
                         & "AND T1.사용구분 = 0 "
          strOrderBy = "ORDER BY T1.계정코드 "
       Else
          strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                             & "T1.계정명 Like '%" & Trim(txtName.Text) & "%' " _
                         & "AND T1.사용구분 = 0 "
          strOrderBy = "ORDER BY T1.계정명 "
       End If
    Else
       strOrderBy = "ORDER BY T1.계정코드 "
    End If
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.계정코드 AS 계정코드, T1.계정명 AS 계정명, " _
               & "합계시산표연결여부 AS 합계시산표연결여부, T1.적요 AS 적요 " _
             & "FROM 계정과목 T1 " _
             & "" & strWhere & " " _
             & "" & strOrderBy & " "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    vsfg1.Rows = P_adoRec.RecordCount + vsfg1.FixedRows
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       txtCode.Text = "": txtName.Text = ""
       txtCode.SetFocus
       Screen.MousePointer = vbDefault
       Exit Sub
    Else
       With vsfg1
            .Cell(flexcpForeColor, 0, 0, .FixedRows - 1, .Cols - 1) = vbBlack
            If .Rows <= PC_intRowCnt Then
               '.ScrollBars = flexScrollBarHorizontal
            Else
               '.ScrollBars = flexScrollBarBoth
            End If
            Do Until P_adoRec.EOF
               lngR = lngR + 1
               .TextMatrix(lngR, 0) = P_adoRec("계정코드")
               .Cell(flexcpData, lngR, 0, lngR, 0) = P_adoRec("계정코드")
               .TextMatrix(lngR, 1) = IIf(IsNull(P_adoRec("계정명")), "", P_adoRec("계정명"))
               .TextMatrix(lngR, 2) = IIf(IsNull(P_adoRec("합계시산표연결여부")), "", P_adoRec("합계시산표연결여부"))
               .TextMatrix(lngR, 3) = IIf(IsNull(P_adoRec("적요")), "", P_adoRec("적요"))
               If P_intFindGbn = 1 Then
                  If PB_strAccCode = Trim(.TextMatrix(lngR, 0)) Then
                     lngRR = lngR
                  End If
               ElseIf _
                  P_intFindGbn = 2 Then
                  If PB_strAccName = Trim(.TextMatrix(lngR, 1)) Then
                     lngRR = lngR
                  End If
               ElseIf _
                  P_intFindGbn = 3 Then
                  If PB_strAccCode = Trim(.TextMatrix(lngR, 0)) And _
                     PB_strAccName = Trim(.TextMatrix(lngR, 1)) Then
                     lngRR = lngR
                  End If
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
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_SELECT:
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "(서버와의 연결 실패)"
    Unload Me
    Exit Sub
End Sub

