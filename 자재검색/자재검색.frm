VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frm자재검색 
   BackColor       =   &H00008000&
   BorderStyle     =   1  '단일 고정
   Caption         =   "품목 검색"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13680
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "자재검색.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   13680
   StartUpPosition =   2  '화면 가운데
   Begin VSFlex7Ctl.VSFlexGrid vsfg1 
      Height          =   2145
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   13515
      _cx             =   23839
      _cy             =   3784
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
      BackColorBkg    =   -2147483643
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
      Left            =   3960
      MaxLength       =   50
      TabIndex        =   2
      Top             =   105
      Width           =   7155
   End
   Begin VB.TextBox txtCode 
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
      Left            =   975
      MaxLength       =   18
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "품명"
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
      Left            =   3000
      TabIndex        =   4
      Top             =   165
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "코드"
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
      Left            =   240
      TabIndex        =   3
      Top             =   165
      Width           =   615
   End
End
Attribute VB_Name = "frm자재검색"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' 프로그램 제 목 : 자재검색
' 사용된 Control :
' 참조된 Table   : 자재, 자재원장, 자재원장마감, 자재입출내역
' 업  무  설  명 :
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived       As Boolean
Private P_intFindGbn       As Integer      '0.수동검색, 1.코드, 2.이름(명), 3.코드 + 이름(명)
Private P_intButton        As Integer
Private P_strFindString2   As String
Private P_blnSelect        As Boolean
Private P_adoRec           As New ADODB.Recordset
Private Const PC_intRowCnt As Integer = 6  '그리드 한 페이지 당 행수(FixedRows 포함)

'+--------------------------------+
'/// LOAD FORM ( 한번만 실행 ) ///
'+--------------------------------+
Private Sub Form_Load()
    P_blnActived = False
    txtCode = PB_strMaterialsCode: txtName = PB_strMaterialsName
    If Len(PB_strMaterialsCode) = 0 And Len(PB_strMaterialsName) = 0 Then
       P_intFindGbn = 0    '수동검색
    ElseIf _
       Len(PB_strMaterialsCode) <> 0 And Len(PB_strMaterialsName) = 0 Then
       P_intFindGbn = 1    '코드로만 자동검색
    ElseIf _
       Len(PB_strMaterialsCode) = 0 And Len(PB_strMaterialsName) <> 0 Then
       P_intFindGbn = 2    '이름(명)으로만 자동검색
    ElseIf _
       Len(PB_strMaterialsCode) <> 0 And Len(PB_strMaterialsName) <> 0 Then
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
    vsfg1_KeyDown 13, 0
End Sub

Private Sub vsfg1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
       SubEscape
    End If
End Sub
Private Sub SubEscape()
    PB_strMaterialsCode = ""
    PB_strMaterialsName = ""
    Unload Me
End Sub

Private Sub txtCode_GotFocus()
    With txtCode
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       If Len(Trim(txtCode.Text) + Trim(txtName.Text)) = 0 Then
          PB_strMaterialsCode = ""
          PB_strMaterialsName = ""
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

Private Sub txtName_GotFocus()
    With txtName
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       If Len(Trim(txtCode.Text) + Trim(txtName.Text)) = 0 Then
          PB_strMaterialsCode = ""
          PB_strMaterialsName = ""
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
            strData = Trim(.Cell(flexcpData, .Row, 4))
            Select Case .MouseCol
                   Case 0, 2
                        .ColSel = 2
                        .ColSort(0) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .ColSort(1) = flexSortNone
                        .ColSort(2) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .Sort = flexSortUseColSort
                   Case 1
                        .ColSel = 2
                        .ColSort(0) = flexSortNone
                        .ColSort(1) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .ColSort(2) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .Sort = flexSortUseColSort
                   Case Else
                        .Sort = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            End Select
            If .FindRow(strData, , 3) > 0 Then
               .Row = .FindRow(strData, , 3)
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
             txtCode.Text = .TextMatrix(.Row, 4)
             txtName.Text = .TextMatrix(.Row, 3)
         End If
    End With
End Sub
Private Sub vsfg1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim SQL       As String
Dim intRetVal As Integer
    With vsfg1
         If .Row >= .FixedRows Then
            'If KeyCode = vbKeyF1 Then '자재시세검색
            '   PB_strMaterialsCode = .TextMatrix(.Row, 4)
            '   PB_strMaterialsName = .TextMatrix(.Row, 3)
            '   frm자재시세검색.Show vbModal
            'End If
            If KeyCode = vbKeyReturn Then
               If .Row = 0 Then
                  PB_strMaterialsCode = ""
                  PB_strMaterialsName = ""
               Else
                  PB_strMaterialsCode = .TextMatrix(.Row, 3)
                  PB_strMaterialsName = .TextMatrix(.Row, 4)
                  P_blnSelect = True
                  If PB_strCallFormName = "frm반입관리" Then
                     frm반입관리.Text1(4).Text = .TextMatrix(.Row, 5)  '규격
                     frm반입관리.Text1(6).Text = .TextMatrix(.Row, 6)  '단위
                  ElseIf _
                     PB_strCallFormName = "frm반품관리" Then
                     frm반품관리.Text1(4).Text = .TextMatrix(.Row, 5)  '규격
                     frm반품관리.Text1(6).Text = .TextMatrix(.Row, 6)  '단위
                  ElseIf _
                     PB_strCallFormName = "frm수불부" Then
                     frm수불부.Text1(2).Text = .TextMatrix(.Row, 5)    '규격
                     frm수불부.Text1(3).Text = .TextMatrix(.Row, 6)    '단위
                  ElseIf _
                     PB_strCallFormName = "frm자재원장" Then
                     frm자재원장.txtBarCode.Text = .TextMatrix(.Row, 8)  '바코드
                     frm자재원장.Text1(2).Text = .TextMatrix(.Row, 5)  '규격
                     frm자재원장.Text1(3).Text = .TextMatrix(.Row, 6)  '단위
                     frm자재원장.Text1(4).Text = Format(.ValueMatrix(.Row, 9), "#,0.00") '폐기율
                     If .ValueMatrix(.Row, 10) = 0 Then                '과세구분
                        frm자재원장.cboTaxGbn.ListIndex = 0            '비과세
                     Else
                        frm자재원장.cboTaxGbn.ListIndex = 1            '과  세
                     End If
                  'ElseIf _
                     'PB_strCallFormName = "frm자재시세" Then
                     'frm자재시세.Text1(2).Text = .TextMatrix(.Row, 5)  '규격
                     'frm자재시세.Text1(3).Text = .TextMatrix(.Row, 6)  '단위
                     'frm자재시세.Text1(6).Text = Format(.ValueMatrix(.Row, 9), "#,0.00") '폐기율
                     'If .ValueMatrix(.Row, 10) = 0 Then                 '과세구분
                     '   frm자재시세.cboTaxGbn.ListIndex = 0            '비과세
                     'Else
                     '   frm자재시세.cboTaxGbn.ListIndex = 1            '과  세
                     'End If
                  End If
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
    If P_blnSelect = False Then
       SubEscape
    End If
    Screen.MousePointer = vbDefault
    If P_adoRec.State <> adStateClosed Then
       P_adoRec.Close
    End If
    Set P_adoRec = Nothing
    Set frm자재검색 = Nothing
End Sub

'+--------------------+
'/// Sub Procedure ///
'+--------------------+
'+----------------------------------+
'/// VsFlexGrid(vsfgGrid) 초기화 ///
'+----------------------------------+
Private Sub Subvsfg1_INIT()
Dim lngR    As Long
Dim lngC    As Long
    With vsfg1           'Rows 1, Cols 15, RowHeightMax(Min) 300
         .AllowBigSelection = False
         .AllowSelection = False
         .AllowUserResizing = flexResizeColumns
         '.BackColorBkg = &H8000000F
         .BackColorBkg = vbWhite
         .BackColorSel = &H8000&
         .Ellipsis = flexEllipsisEnd
         '.ExplorerBar = flexExSortShow
         .ExtendLastCol = True
         .FocusRect = flexFocusHeavy
         .FontSize = 9
         .ScrollBars = flexScrollBarHorizontal
         .ScrollTrack = True
         .SelectionMode = flexSelectionByRow
         .FixedRows = 1
         .FixedCols = 4
         .Rows = 1             'Subvsfg1_Fill수행시에 설정
         .Cols = 15
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1000   '자재분류(분류코드) 'H
         .ColWidth(1) = 1000   '분류명(분류명)     'H
         .ColWidth(2) = 700    '세부코드           'H
         .ColWidth(3) = 2000   '분류코드 + 세부코드 = 자재코드
         .ColWidth(4) = 2700   '품명
         .ColWidth(5) = 2200   '규격
         .ColWidth(6) = 600    '단위
         .ColWidth(7) = 1200   '현재고량
         .ColWidth(8) = 1500   '바코드
         .ColWidth(9) = 700    '폐기율
         .ColWidth(10) = 1000  '과세구분 'H
         .ColWidth(11) = 900   '과세구분
         .ColWidth(12) = 1000  '사용구분 'H
         .ColWidth(13) = 900   '사용구분
         .ColWidth(14) = 500   '원장유무
         .Cell(flexcpFontBold, 0, 0, .FixedRows - 1, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "분류코드"   'H
         .TextMatrix(0, 1) = "분류명"
         .TextMatrix(0, 2) = "세부코드"   'H
         .TextMatrix(0, 3) = "품목코드"
         .TextMatrix(0, 4) = "품명"
         .TextMatrix(0, 5) = "규격"
         .TextMatrix(0, 6) = "단위"
         .TextMatrix(0, 7) = "현재고량"
         .TextMatrix(0, 8) = "바코드"
         .TextMatrix(0, 9) = "폐기율"
         .TextMatrix(0, 10) = "과세구분"
         .TextMatrix(0, 11) = "과세구분"
         .TextMatrix(0, 12) = "사용구분"
         .TextMatrix(0, 13) = "사용구분"
         .TextMatrix(0, 14) = "원장"
         .ColFormat(7) = "#,#.00": .ColFormat(9) = "#,#.00"
         .ColHidden(0) = True: .ColHidden(1) = True: .ColHidden(2) = True
         .ColHidden(10) = True: .ColHidden(12) = True:
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 1, 3, 4, 5, 6, 8
                         .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 0, 2, 10, 11, 13, 14
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

'+---------------------------------+
'/// VsFlexGrid(vsfgGrid) 채우기///
'+---------------------------------+
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
Dim lngRRR     As Long
    
    txtCode.Text = Trim(txtCode.Text): txtName.Text = Trim(txtName.Text)
    If P_intFindGbn = 1 And Len(txtCode.Text) = 0 Then
       txtCode.SetFocus
       Exit Sub
    ElseIf _
       P_intFindGbn = 2 And Len(txtName.Text) = 0 Then
       txtName.SetFocus
       Exit Sub
    Else
       If (Len(txtCode.Text) + Len(txtName.Text)) = 0 Then
          txtCode.SetFocus
          Exit Sub
       End If
    End If
    Screen.MousePointer = vbHourglass
    If P_intFindGbn = 1 Then
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                          & "(T1.분류코드 + T1.세부코드) Like '%" & Trim(txtCode.Text) & "%' " _
                       & "AND T1.사용구분 = 0 "
       strOrderBy = "ORDER BY T1.분류코드, T1.세부코드 "
    ElseIf _
       P_intFindGbn = 2 Then
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") + "T1.자재명 Like '%" & Trim(txtName.Text) & "%' " _
                                                                    & "AND T1.사용구분 = 0 "
       strOrderBy = "ORDER BY T1.자재명 "
    ElseIf _
       P_intFindGbn = 3 Then
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                          & "(T1.분류코드 + T1.세부코드) Like '%" & Trim(txtCode.Text) & "%' " _
                      & "AND T1.자재명 Like '%" & Trim(txtName.Text) & "%' " _
                      & "AND T1.사용구분 = 0 "
       strOrderBy = "ORDER BY T1.자재명 "
    Else
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") + "T1.자재명 Like '%" & Trim(txtName.Text) & "%' " _
                                                                    & "AND T1.사용구분 = 0 "
       strOrderBy = "ORDER BY T1.자재명 "
    End If
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT ISNULL(T1.분류코드,'') AS 분류코드, ISNULL(T2.분류명,'') AS 분류명, " _
                  & "ISNULL(T1.세부코드,'') AS 세부코드, T1.자재명 AS 자재명, " _
                  & "T1.규격 AS 규격, T1.단위 AS 단위, " _
                  & "T1.폐기율 AS 폐기율, T1.과세구분 AS 과세구분, " _
                  & "T1.사용구분 AS 사용구분, ISNULL(T5.사업장코드,'') AS 원장유무, ISNULL(T5.적정재고,0) AS 적정재고, " _
                  & "T1.바코드 AS 바코드, "
    strSQL = strSQL _
                  & "(SELECT ISNULL(SUM(입고누계수량-출고누계수량),0) " _
                     & "FROM 자재원장마감 " _
                    & "WHERE 분류코드 = T1.분류코드 And 세부코드 = T1.세부코드 " _
                      & "AND 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                      & "AND 마감년월 >= '" & Left(PB_regUserinfoU.UserClientDate, 4) + "00" & "' " _
                      & "AND 마감년월 <  '" & Left(PB_regUserinfoU.UserClientDate, 6) & "') AS 이월재고,"
    strSQL = strSQL _
                 & "(SELECT ISNULL(SUM(입고수량-출고수량),0) " _
                    & "FROM 자재입출내역 " _
                   & "WHERE 분류코드 = T1.분류코드 And 세부코드 = T1.세부코드 " _
                     & "AND 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                     & "AND 입출고일자 BETWEEN '" & Left(PB_regUserinfoU.UserClientDate, 6) + "01" & "' " _
                                             & "AND '" & PB_regUserinfoU.UserClientDate & "') AS 금월재고 "
    strSQL = strSQL _
             & "FROM 자재 T1 " _
             & "LEFT JOIN 자재분류 T2 " _
                    & "ON T2.분류코드 = T1.분류코드 " _
             & "LEFT JOIN 자재원장 T5 " _
                    & "ON T5.분류코드 = T1.분류코드 AND T5.세부코드 = T1.세부코드 " _
                   & "AND T5.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
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
               .ScrollBars = flexScrollBarHorizontal
            Else
               .ScrollBars = flexScrollBarBoth
            End If
            Do Until P_adoRec.EOF
               lngR = lngR + 1
               .TextMatrix(lngR, 0) = IIf(IsNull(P_adoRec("분류코드")), "", P_adoRec("분류코드"))
               .TextMatrix(lngR, 1) = IIf(IsNull(P_adoRec("분류명")), "", P_adoRec("분류명"))
               .TextMatrix(lngR, 2) = IIf(IsNull(P_adoRec("세부코드")), "", P_adoRec("세부코드"))
               'FindRow 사용을 위해
               .TextMatrix(lngR, 3) = .TextMatrix(lngR, 0) & P_adoRec("세부코드")
               .Cell(flexcpData, lngR, 3, lngR, 3) = .TextMatrix(lngR, 3)
               .TextMatrix(lngR, 4) = IIf(IsNull(P_adoRec("자재명")), "", P_adoRec("자재명"))
               .TextMatrix(lngR, 5) = IIf(IsNull(P_adoRec("규격")), "", P_adoRec("규격"))
               .TextMatrix(lngR, 6) = IIf(IsNull(P_adoRec("단위")), "", P_adoRec("단위"))
               .TextMatrix(lngR, 7) = P_adoRec("이월재고") + P_adoRec("금월재고")
               If P_adoRec("적정재고") <> 0 Then
                  If .ValueMatrix(lngR, 7) < P_adoRec("적정재고") Then
                     .Cell(flexcpForeColor, lngR, 7, lngR, 7) = vbRed
                  End If
               End If
               .TextMatrix(lngR, 8) = IIf(IsNull(P_adoRec("바코드")), "", P_adoRec("바코드"))
               .TextMatrix(lngR, 9) = IIf(IsNull(P_adoRec("폐기율")), "", P_adoRec("폐기율"))
               .TextMatrix(lngR, 10) = Vals(P_adoRec("과세구분"))
               If .ValueMatrix(lngR, 10) = 0 Then
                  .TextMatrix(lngR, 11) = "비과세"
               Else
                 .TextMatrix(lngR, 11) = "과  세"
               End If
               .TextMatrix(lngR, 12) = Vals(P_adoRec("사용구분"))
               If .ValueMatrix(lngR, 12) = 0 Then
                  .TextMatrix(lngR, 13) = "정   상"
               ElseIf _
                  .ValueMatrix(lngR, 12) = 9 Then
                  .TextMatrix(lngR, 13) = "사용불가"
               Else
                  .TextMatrix(lngR, 13) = "오    류"
               End If
               If Len(P_adoRec("원장유무")) = 0 Then
                  .Cell(flexcpBackColor, lngR, 14, lngR, 14) = vbYellow
                  .Cell(flexcpForeColor, lngR, 14, lngR, 14) = vbRed
                  .TextMatrix(lngR, 14) = "무"
               Else
                  .Cell(flexcpBackColor, lngR, 14, lngR, 14) = vbWhite
                  .Cell(flexcpForeColor, lngR, 14, lngR, 14) = vbBlack
                  .TextMatrix(lngR, 14) = "유"
               End If
               If P_intFindGbn = 1 Then
                  If PB_strMaterialsCode = Trim(.TextMatrix(lngR, 4)) Then
                     lngRR = lngR
                  End If
               ElseIf _
                  P_intFindGbn = 2 Then
                  If PB_strMaterialsName = Trim(.TextMatrix(lngR, 3)) Then
                     lngRR = lngR
                  End If
               ElseIf _
                  P_intFindGbn = 3 Then
                  If PB_strMaterialsCode = Trim(.TextMatrix(lngR, 4)) And _
                     PB_strMaterialsName = Trim(.TextMatrix(lngR, 3)) Then
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

