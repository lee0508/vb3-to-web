VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frm����˻� 
   BackColor       =   &H00008000&
   BorderStyle     =   1  '���� ����
   Caption         =   "ǰ�� �˻�"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13680
   BeginProperty Font 
      Name            =   "����ü"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "����˻�.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   13680
   StartUpPosition =   2  'ȭ�� ���
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
         Name            =   "����ü"
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
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "����ü"
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
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   8  '����
      Left            =   975
      MaxLength       =   18
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "ǰ��"
      BeginProperty Font 
         Name            =   "����ü"
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
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "�ڵ�"
      BeginProperty Font 
         Name            =   "����ü"
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
Attribute VB_Name = "frm����˻�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' ���α׷� �� �� : ����˻�
' ���� Control :
' ������ Table   : ����, �������, ������帶��, �������⳻��
' ��  ��  ��  �� :
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived       As Boolean
Private P_intFindGbn       As Integer      '0.�����˻�, 1.�ڵ�, 2.�̸�(��), 3.�ڵ� + �̸�(��)
Private P_intButton        As Integer
Private P_strFindString2   As String
Private P_blnSelect        As Boolean
Private P_adoRec           As New ADODB.Recordset
Private Const PC_intRowCnt As Integer = 6  '�׸��� �� ������ �� ���(FixedRows ����)

'+--------------------------------+
'/// LOAD FORM ( �ѹ��� ���� ) ///
'+--------------------------------+
Private Sub Form_Load()
    P_blnActived = False
    txtCode = PB_strMaterialsCode: txtName = PB_strMaterialsName
    If Len(PB_strMaterialsCode) = 0 And Len(PB_strMaterialsName) = 0 Then
       P_intFindGbn = 0    '�����˻�
    ElseIf _
       Len(PB_strMaterialsCode) <> 0 And Len(PB_strMaterialsName) = 0 Then
       P_intFindGbn = 1    '�ڵ�θ� �ڵ��˻�
    ElseIf _
       Len(PB_strMaterialsCode) = 0 And Len(PB_strMaterialsName) <> 0 Then
       P_intFindGbn = 2    '�̸�(��)���θ� �ڵ��˻�
    ElseIf _
       Len(PB_strMaterialsCode) <> 0 And Len(PB_strMaterialsName) <> 0 Then
       P_intFindGbn = 3    '�ڵ�� �̸�(��)�� ���ÿ� �ڵ��˻�
    Else
       P_intFindGbn = 0    '�����˻�
    End If
End Sub

'+-------------------------------------------+
'/// ACTIVATE FORM Ȱ��ȭ ( �ѹ��� ���� ) ///
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
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "(�������� ���� ����)"
    Unload Me
    Exit Sub
End Sub

'+-----------+
'/// ��� ///
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
            'If KeyCode = vbKeyF1 Then '����ü��˻�
            '   PB_strMaterialsCode = .TextMatrix(.Row, 4)
            '   PB_strMaterialsName = .TextMatrix(.Row, 3)
            '   frm����ü��˻�.Show vbModal
            'End If
            If KeyCode = vbKeyReturn Then
               If .Row = 0 Then
                  PB_strMaterialsCode = ""
                  PB_strMaterialsName = ""
               Else
                  PB_strMaterialsCode = .TextMatrix(.Row, 3)
                  PB_strMaterialsName = .TextMatrix(.Row, 4)
                  P_blnSelect = True
                  If PB_strCallFormName = "frm���԰���" Then
                     frm���԰���.Text1(4).Text = .TextMatrix(.Row, 5)  '�԰�
                     frm���԰���.Text1(6).Text = .TextMatrix(.Row, 6)  '����
                  ElseIf _
                     PB_strCallFormName = "frm��ǰ����" Then
                     frm��ǰ����.Text1(4).Text = .TextMatrix(.Row, 5)  '�԰�
                     frm��ǰ����.Text1(6).Text = .TextMatrix(.Row, 6)  '����
                  ElseIf _
                     PB_strCallFormName = "frm���Һ�" Then
                     frm���Һ�.Text1(2).Text = .TextMatrix(.Row, 5)    '�԰�
                     frm���Һ�.Text1(3).Text = .TextMatrix(.Row, 6)    '����
                  ElseIf _
                     PB_strCallFormName = "frm�������" Then
                     frm�������.txtBarCode.Text = .TextMatrix(.Row, 8)  '���ڵ�
                     frm�������.Text1(2).Text = .TextMatrix(.Row, 5)  '�԰�
                     frm�������.Text1(3).Text = .TextMatrix(.Row, 6)  '����
                     frm�������.Text1(4).Text = Format(.ValueMatrix(.Row, 9), "#,0.00") '�����
                     If .ValueMatrix(.Row, 10) = 0 Then                '��������
                        frm�������.cboTaxGbn.ListIndex = 0            '�����
                     Else
                        frm�������.cboTaxGbn.ListIndex = 1            '��  ��
                     End If
                  'ElseIf _
                     'PB_strCallFormName = "frm����ü�" Then
                     'frm����ü�.Text1(2).Text = .TextMatrix(.Row, 5)  '�԰�
                     'frm����ü�.Text1(3).Text = .TextMatrix(.Row, 6)  '����
                     'frm����ü�.Text1(6).Text = Format(.ValueMatrix(.Row, 9), "#,0.00") '�����
                     'If .ValueMatrix(.Row, 10) = 0 Then                 '��������
                     '   frm����ü�.cboTaxGbn.ListIndex = 0            '�����
                     'Else
                     '   frm����ü�.cboTaxGbn.ListIndex = 1            '��  ��
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
'/// ���� ///
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
    Set frm����˻� = Nothing
End Sub

'+--------------------+
'/// Sub Procedure ///
'+--------------------+
'+----------------------------------+
'/// VsFlexGrid(vsfgGrid) �ʱ�ȭ ///
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
         .Rows = 1             'Subvsfg1_Fill����ÿ� ����
         .Cols = 15
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1000   '����з�(�з��ڵ�) 'H
         .ColWidth(1) = 1000   '�з���(�з���)     'H
         .ColWidth(2) = 700    '�����ڵ�           'H
         .ColWidth(3) = 2000   '�з��ڵ� + �����ڵ� = �����ڵ�
         .ColWidth(4) = 2700   'ǰ��
         .ColWidth(5) = 2200   '�԰�
         .ColWidth(6) = 600    '����
         .ColWidth(7) = 1200   '�����
         .ColWidth(8) = 1500   '���ڵ�
         .ColWidth(9) = 700    '�����
         .ColWidth(10) = 1000  '�������� 'H
         .ColWidth(11) = 900   '��������
         .ColWidth(12) = 1000  '��뱸�� 'H
         .ColWidth(13) = 900   '��뱸��
         .ColWidth(14) = 500   '��������
         .Cell(flexcpFontBold, 0, 0, .FixedRows - 1, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "�з��ڵ�"   'H
         .TextMatrix(0, 1) = "�з���"
         .TextMatrix(0, 2) = "�����ڵ�"   'H
         .TextMatrix(0, 3) = "ǰ���ڵ�"
         .TextMatrix(0, 4) = "ǰ��"
         .TextMatrix(0, 5) = "�԰�"
         .TextMatrix(0, 6) = "����"
         .TextMatrix(0, 7) = "�����"
         .TextMatrix(0, 8) = "���ڵ�"
         .TextMatrix(0, 9) = "�����"
         .TextMatrix(0, 10) = "��������"
         .TextMatrix(0, 11) = "��������"
         .TextMatrix(0, 12) = "��뱸��"
         .TextMatrix(0, 13) = "��뱸��"
         .TextMatrix(0, 14) = "����"
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
'/// VsFlexGrid(vsfgGrid) ä���///
'+---------------------------------+
'+---------------------------------+
'/// VsFlexGrid(vsfg1) ä���///
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
                          & "(T1.�з��ڵ� + T1.�����ڵ�) Like '%" & Trim(txtCode.Text) & "%' " _
                       & "AND T1.��뱸�� = 0 "
       strOrderBy = "ORDER BY T1.�з��ڵ�, T1.�����ڵ� "
    ElseIf _
       P_intFindGbn = 2 Then
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") + "T1.����� Like '%" & Trim(txtName.Text) & "%' " _
                                                                    & "AND T1.��뱸�� = 0 "
       strOrderBy = "ORDER BY T1.����� "
    ElseIf _
       P_intFindGbn = 3 Then
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                          & "(T1.�з��ڵ� + T1.�����ڵ�) Like '%" & Trim(txtCode.Text) & "%' " _
                      & "AND T1.����� Like '%" & Trim(txtName.Text) & "%' " _
                      & "AND T1.��뱸�� = 0 "
       strOrderBy = "ORDER BY T1.����� "
    Else
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") + "T1.����� Like '%" & Trim(txtName.Text) & "%' " _
                                                                    & "AND T1.��뱸�� = 0 "
       strOrderBy = "ORDER BY T1.����� "
    End If
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT ISNULL(T1.�з��ڵ�,'') AS �з��ڵ�, ISNULL(T2.�з���,'') AS �з���, " _
                  & "ISNULL(T1.�����ڵ�,'') AS �����ڵ�, T1.����� AS �����, " _
                  & "T1.�԰� AS �԰�, T1.���� AS ����, " _
                  & "T1.����� AS �����, T1.�������� AS ��������, " _
                  & "T1.��뱸�� AS ��뱸��, ISNULL(T5.������ڵ�,'') AS ��������, ISNULL(T5.�������,0) AS �������, " _
                  & "T1.���ڵ� AS ���ڵ�, "
    strSQL = strSQL _
                  & "(SELECT ISNULL(SUM(�԰������-��������),0) " _
                     & "FROM ������帶�� " _
                    & "WHERE �з��ڵ� = T1.�з��ڵ� And �����ڵ� = T1.�����ڵ� " _
                      & "AND ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                      & "AND ������� >= '" & Left(PB_regUserinfoU.UserClientDate, 4) + "00" & "' " _
                      & "AND ������� <  '" & Left(PB_regUserinfoU.UserClientDate, 6) & "') AS �̿����,"
    strSQL = strSQL _
                 & "(SELECT ISNULL(SUM(�԰����-������),0) " _
                    & "FROM �������⳻�� " _
                   & "WHERE �з��ڵ� = T1.�з��ڵ� And �����ڵ� = T1.�����ڵ� " _
                     & "AND ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                     & "AND ��������� BETWEEN '" & Left(PB_regUserinfoU.UserClientDate, 6) + "01" & "' " _
                                             & "AND '" & PB_regUserinfoU.UserClientDate & "') AS �ݿ���� "
    strSQL = strSQL _
             & "FROM ���� T1 " _
             & "LEFT JOIN ����з� T2 " _
                    & "ON T2.�з��ڵ� = T1.�з��ڵ� " _
             & "LEFT JOIN ������� T5 " _
                    & "ON T5.�з��ڵ� = T1.�з��ڵ� AND T5.�����ڵ� = T1.�����ڵ� " _
                   & "AND T5.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
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
               .TextMatrix(lngR, 0) = IIf(IsNull(P_adoRec("�з��ڵ�")), "", P_adoRec("�з��ڵ�"))
               .TextMatrix(lngR, 1) = IIf(IsNull(P_adoRec("�з���")), "", P_adoRec("�з���"))
               .TextMatrix(lngR, 2) = IIf(IsNull(P_adoRec("�����ڵ�")), "", P_adoRec("�����ڵ�"))
               'FindRow ����� ����
               .TextMatrix(lngR, 3) = .TextMatrix(lngR, 0) & P_adoRec("�����ڵ�")
               .Cell(flexcpData, lngR, 3, lngR, 3) = .TextMatrix(lngR, 3)
               .TextMatrix(lngR, 4) = IIf(IsNull(P_adoRec("�����")), "", P_adoRec("�����"))
               .TextMatrix(lngR, 5) = IIf(IsNull(P_adoRec("�԰�")), "", P_adoRec("�԰�"))
               .TextMatrix(lngR, 6) = IIf(IsNull(P_adoRec("����")), "", P_adoRec("����"))
               .TextMatrix(lngR, 7) = P_adoRec("�̿����") + P_adoRec("�ݿ����")
               If P_adoRec("�������") <> 0 Then
                  If .ValueMatrix(lngR, 7) < P_adoRec("�������") Then
                     .Cell(flexcpForeColor, lngR, 7, lngR, 7) = vbRed
                  End If
               End If
               .TextMatrix(lngR, 8) = IIf(IsNull(P_adoRec("���ڵ�")), "", P_adoRec("���ڵ�"))
               .TextMatrix(lngR, 9) = IIf(IsNull(P_adoRec("�����")), "", P_adoRec("�����"))
               .TextMatrix(lngR, 10) = Vals(P_adoRec("��������"))
               If .ValueMatrix(lngR, 10) = 0 Then
                  .TextMatrix(lngR, 11) = "�����"
               Else
                 .TextMatrix(lngR, 11) = "��  ��"
               End If
               .TextMatrix(lngR, 12) = Vals(P_adoRec("��뱸��"))
               If .ValueMatrix(lngR, 12) = 0 Then
                  .TextMatrix(lngR, 13) = "��   ��"
               ElseIf _
                  .ValueMatrix(lngR, 12) = 9 Then
                  .TextMatrix(lngR, 13) = "���Ұ�"
               Else
                  .TextMatrix(lngR, 13) = "��    ��"
               End If
               If Len(P_adoRec("��������")) = 0 Then
                  .Cell(flexcpBackColor, lngR, 14, lngR, 14) = vbYellow
                  .Cell(flexcpForeColor, lngR, 14, lngR, 14) = vbRed
                  .TextMatrix(lngR, 14) = "��"
               Else
                  .Cell(flexcpBackColor, lngR, 14, lngR, 14) = vbWhite
                  .Cell(flexcpForeColor, lngR, 14, lngR, 14) = vbBlack
                  .TextMatrix(lngR, 14) = "��"
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
               .Row = lngRR       'vsfg1_EnterCell �ڵ�����(���� �Ѱ� �϶��� �ڵ����� ����)
               If .Rows > PC_intRowCnt Then
                  '.TopRow = .Rows - PC_intRowCnt + .FixedRows
                  .TopRow = 1
               End If
            Else
               .Row = lngRR       'vsfg1_EnterCell �ڵ�����(���� �Ѱ� �϶��� �ڵ����� ����)
               If .Rows > PC_intRowCnt Then
                  .TopRow = .Row
               End If
            End If
            vsfg1_EnterCell       'vsfg1_EnterCell �ڵ�����(���� �Ѱ� �϶��� ������ �ڵ�����)
            .SetFocus
       End With
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_SELECT:
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "(�������� ���� ����)"
    Unload Me
    Exit Sub
End Sub

