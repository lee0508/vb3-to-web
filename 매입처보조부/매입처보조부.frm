VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm����ó������ 
   BorderStyle     =   0  '����
   Caption         =   "����ó������"
   ClientHeight    =   10095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15405
   BeginProperty Font 
      Name            =   "����ü"
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
   ScaleMode       =   0  '�����
   ScaleWidth      =   15405
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  '���
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   60
      TabIndex        =   9
      Top             =   0
      Width           =   15195
      Begin VB.OptionButton optPrtChk0 
         Caption         =   "�ڵ��"
         Height          =   255
         Left            =   6840
         TabIndex        =   23
         Top             =   150
         Width           =   975
      End
      Begin VB.OptionButton optPrtChk1 
         Caption         =   "�̸���"
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
         Picture         =   "����ó������.frx":0000
         Style           =   1  '�׷���
         TabIndex        =   16
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   9120
         Picture         =   "����ó������.frx":0963
         Style           =   1  '�׷���
         TabIndex        =   14
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   13635
         Picture         =   "����ó������.frx":1308
         Style           =   1  '�׷���
         TabIndex        =   13
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   390
         Left            =   12510
         Picture         =   "����ó������.frx":1C56
         Style           =   1  '�׷���
         TabIndex        =   12
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   11385
         Picture         =   "����ó������.frx":25DA
         Style           =   1  '�׷���
         TabIndex        =   11
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Height          =   390
         Left            =   10260
         Picture         =   "����ó������.frx":2E61
         Style           =   1  '�׷���
         TabIndex        =   4
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00008000&
         BorderStyle     =   1  '���� ����
         Caption         =   "�� �� ó �� �� �� �� Ȳ"
         BeginProperty Font 
            Name            =   "����ü"
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
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "����ü"
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
         Caption         =   "��ü ����ó"
         Height          =   255
         Left            =   5030
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Index           =   1
         Left            =   2475
         MaxLength       =   50
         TabIndex        =   1
         Top             =   585
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
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
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   12
         Left            =   13800
         TabIndex        =   20
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   11
         Left            =   11760
         TabIndex        =   19
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��������"
         Height          =   240
         Index           =   10
         Left            =   9360
         TabIndex        =   18
         Top             =   285
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�˻�����)"
         BeginProperty Font 
            Name            =   "����ü"
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
         Alignment       =   1  '������ ����
         Caption         =   "����ó��"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   1275
         TabIndex        =   8
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����ó�ڵ�"
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
Attribute VB_Name = "frm����ó������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' ���α׷� �� �� : ����ó������
' ���� Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' ������ Table   :
' ��  ��  ��  �� :
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived       As Boolean
Private P_adoRec           As New ADODB.Recordset
Private P_intButton        As Integer
Private Const PC_intRowCnt As Integer = 25  '�׸��� �� ������ �� ���(FixedRows ����)

'+--------------------------------+
'/// LOAD FORM ( �ѹ��� ���� ) ///
'+--------------------------------+
Private Sub Form_Load()
    P_blnActived = False
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
       Select Case Val(PB_regUserinfoU.UserAuthority)
              Case Is <= 10 '��ȸ
                   cmdPrint.Enabled = False: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 20 '�μ�, ��ȸ
                   cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 40 '�߰�, ����, �μ�, ��ȸ
                   cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 50 '����, �߰�, ����, ��ȸ, �μ�
                   cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 99 '����, �߰�, ����, ��ȸ, �μ�
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
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����� ����(�������� ���� ����)"
    Unload Me
    Exit Sub
End Sub

'+---------------+
'/// �˻����� ///
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
    If (Index = 0 And (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn)) Then  '����ó�˻�
       PB_strSupplierCode = UPPER(Trim(Text1(Index).Text))
       PB_strSupplierName = ""  'Trim(Text1(Index + 1).Text)
       frm����ó�˻�.Show vbModal
       If (Len(PB_strSupplierCode) + Len(PB_strSupplierName)) = 0 Then '�˻����� ���(ESC)
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
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����ó �б� ����"
    Unload Me
    Exit Sub
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim strSQL As String
    With Text1(Index)
         .Text = Trim(.Text)
         Select Case Index
                Case 0 '����ó
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
'/// �߰� ///
'+-----------+
Private Sub cmdClear_Click()
Dim strSQL As String
    '
End Sub
'+-----------+
'/// ��ȸ ///
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
'/// ���� ///
'+-----------+
Private Sub cmdSave_Click()
    '
End Sub

'+-----------+
'/// ���� ///
'+-----------+
Private Sub cmdDelete_Click()
    '
End Sub

'+-----------+
'/// ���� ///
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
    Set frm����ó������ = Nothing
End Sub

'+--------------------+
'/// Sub Procedure ///
'+--------------------+
'+----------------------------------+
'/// VsFlexGrid(vsfg1) �ʱ�ȭ ///
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
         .Rows = 1             'Subvsfg1_Fill����ÿ� ����
         .Cols = 17
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1000   '������ڵ�
         .ColWidth(1) = 1000   '����ó�ڵ�
         .ColWidth(2) = 2000   '����ó��
         .ColWidth(3) = 1100   '����
         .ColWidth(4) = 2500   '����
         .ColWidth(5) = 1900   '�����ڵ�(�з�+����) 'H
         .ColWidth(6) = 2200   '�����
         .ColWidth(7) = 2100   '�԰�
         .ColWidth(8) = 700    '����
         .ColWidth(9) = 1000   '�������(����)
         .ColWidth(10) = 500   '������и�(����)
         .ColWidth(11) = 1000  '�԰����(����)
         .ColWidth(12) = 1400  '�԰�ܰ�(���ް�)
         .ColWidth(13) = 1700  '�԰�ݾ�(�ΰ���������)
         .ColWidth(14) = 1600  '�԰�ΰ���
         .ColWidth(15) = 1700  '�԰�ݾ�(�հ�)
         .ColWidth(16) = 2000  '���
         
         .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "������ڵ�"  'H
         .TextMatrix(0, 1) = "����ó�ڵ�"  'H
         .TextMatrix(0, 2) = "����ó��"    'H
         .TextMatrix(0, 3) = "��¥"
         .TextMatrix(0, 4) = "����"
         .TextMatrix(0, 5) = "ǰ���ڵ�"    'H
         .TextMatrix(0, 6) = "ǰ��"
         .TextMatrix(0, 7) = "�԰�"
         .TextMatrix(0, 8) = "����"
         .TextMatrix(0, 9) = "����"        'H
         .TextMatrix(0, 10) = "����"
         .TextMatrix(0, 11) = "����"
         .TextMatrix(0, 12) = "���Դܰ�"   'ǰ��ܰ�
         .TextMatrix(0, 13) = "���Աݾ�"   '���� * �ܰ�
         .TextMatrix(0, 14) = "���Ժΰ�"      '((���� * �ܰ�) * (PB_curVatRate + 1)) - (���� * �ܰ�)
         .TextMatrix(0, 15) = "���Աݾ�(VAT)" '(���� * �ܰ�) * (PB_curVatRate + 1)
         .TextMatrix(0, 16) = "���"
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
'/// VsFlexGrid(vsfg1) ä���///
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
Dim StrDate     As String    '�ش�����

Dim curJanMny   As Currency  '��ü�� �ܾ�

Dim curMQAmt    As Currency  '��ü��������
Dim curTQAmt    As Currency  '��ü����������
Dim curTTQAmt   As Currency  '��  ü��������

Dim curMUMny    As Currency  '��ü��������ݾ�       '(�ܰ� * ����)
Dim curTUMny    As Currency  '��ü������ݾ״���     '(�ܰ� * ����)
Dim curTTUMny   As Currency  '��  ü����ݾ״���     '(�ܰ� * ����)

    vsfg1.Rows = 1
    With vsfg1
         '�˻����� ����ó
         If chkTotal.Value = 0 Then '�Ǻ� ��ȸ
            If Len(Text1(0).Text) > 0 Then
               strWhere = "WHERE T1.����ó�ڵ� = '" & Trim(Text1(0).Text) & "' "
            Else
               Text1(0).SetFocus
               Exit Sub
            End If
         End If
         strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                  & "T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' "
    End With
    strOrderBy = "ORDER BY T1.������ڵ�, " & IIf(optPrtChk0.Value = True, "T1.����ó�ڵ�, T3.����ó�� ", "T3.����ó��, T1.����ó�ڵ� ") & ", ����, �����, ���� "
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.������ڵ� AS ������ڵ�, T2.������ AS ������, " _
                  & "T1.����ó�ڵ� AS ����ó�ڵ�, T3.����ó�� AS ����ó��, " _
                  & "(T1.������� + '00') AS ����, '(�� �� ��)' AS ����, " _
                  & "'' AS �з��ڵ�, '' AS �з���, " _
                  & "'' AS �����ڵ�, '' AS �����, " _
                  & "'' AS �԰�, '' AS ����, 0 AS ����, " _
                  & "0 AS �԰����, 0 AS �԰�ܰ�, " _
                  & "0 AS �԰�ΰ�, (T1.�����ޱݴ���ݾ� - T1.�����ޱ����޴���ݾ�) AS �԰�ݾ� " _
             & "FROM �����ޱݿ��帶�� T1 " _
             & "LEFT JOIN ����� T2 ON T2.������ڵ� = T1.������ڵ� " _
             & "LEFT JOIN ����ó T3 ON T3.������ڵ� = T1.������ڵ� AND T3.����ó�ڵ� = T1.����ó�ڵ� " _
             & "" & strWhere & " AND SUBSTRING(T1.�������,5,2) = '00' " _
              & "AND (T1.�����ޱݴ���ݾ� <> 0 OR T1.�����ޱ����޴���ݾ� <> 0) " _
              & "AND T1.������� BETWEEN '" & (Mid(DTOS(dtpF_Date.Value), 1, 4) + "00") & "' " _
                                  & "AND '" & (Mid(DTOS(dtpT_Date.Value), 1, 4) + "00") & "' "
    strSQL = strSQL & "UNION ALL " _
           & "SELECT T1.������ڵ� AS ������ڵ�, T2.������ AS ������, " _
                  & "T1.����ó�ڵ� AS ����ó�ڵ�, T3.����ó�� AS ����ó��, " _
                  & "(T1.������� + '00') AS ����, '(��    ��)' AS ����, " _
                  & "'' AS �з��ڵ�, '' AS �з���, " _
                  & "'' AS �����ڵ�, '' AS �����, " _
                  & "'' AS �԰�, '' AS ����, 0 AS ����, " _
                  & "0 AS �԰����, 0 AS �԰�ܰ�, " _
                  & "0 AS �԰�ΰ�, (T1.�����ޱݴ���ݾ� - T1.�����ޱ����޴���ݾ�) AS �԰�ݾ� " _
             & "FROM �����ޱݿ��帶�� T1 " _
             & "LEFT JOIN ����� T2 ON T2.������ڵ� = T1.������ڵ� " _
             & "LEFT JOIN ����ó T3 ON T3.������ڵ� = T1.������ڵ� AND T3.����ó�ڵ� = T1.����ó�ڵ� " _
             & "" & strWhere & " AND SUBSTRING(T1.�������,5,2) <> '00' " _
              & "AND (T1.�����ޱݴ���ݾ� <> 0 OR T1.�����ޱ����޴���ݾ� <> 0) " _
              & "AND T1.������� > '" & (Mid(DTOS(dtpF_Date.Value), 1, 4) + "00") & "' " _
              & "AND T1.������� < '" & Mid(DTOS(dtpF_Date.Value), 1, 6) & "' "
    strSQL = strSQL & "UNION ALL " _
           & "SELECT T1.������ڵ� AS ������ڵ�, T2.������ AS ������, " _
                  & "T1.����ó�ڵ� AS ����ó�ڵ�, T3.����ó�� AS ����ó��, " _
                  & "(T1.���������) AS ����, (T1.�������������) AS ����, " _
                  & "T1.�з��ڵ� AS �з��ڵ�, T4.�з��� AS �з���, " _
                  & "T1.�����ڵ� AS �����ڵ�, ISNULL(T5.�����,'ERROR!') AS �����, " _
                  & "ISNULL(T5.�԰�,'') AS �԰�, ISNULL(T5.����,'') AS ����, T1.������� AS ����, " _
                  & "SUM(T1.�԰����) AS �԰����, T1.�԰�ܰ� AS �԰�ܰ�, " _
                  & "T1.�԰�ΰ� AS �԰�ΰ�, (SUM(T1.�԰����*T1.�԰�ܰ�) * " & (PB_curVatRate + 1) & ") AS �԰�ݾ� " _
             & "FROM �������⳻�� T1 " _
             & "LEFT JOIN ����� T2 ON T2.������ڵ� = T1.������ڵ� " _
             & "LEFT JOIN ����ó T3 ON T3.������ڵ� = T1.������ڵ� AND T3.����ó�ڵ� = T1.����ó�ڵ� " _
             & "LEFT JOIN ����з� T4 ON T4.�з��ڵ� = T1.�з��ڵ� " _
             & "LEFT JOIN ���� T5 ON T5.�з��ڵ� = T1.�з��ڵ� AND T5.�����ڵ� = T1.�����ڵ� " _
             & "" & strWhere & " AND T1.������� = 1 AND T1.��뱸�� = 0 " _
              & "AND T1.��������� BETWEEN '" & (Mid(DTOS(dtpF_Date.Value), 1, 6) + "01") & "' AND '" & DTOS(dtpT_Date.Value) & "' "
    strSQL = strSQL _
           & "GROUP BY T1.������ڵ�, T2.������, T1.����ó�ڵ�, T3.����ó��, " _
                    & "T1.���������, T1.�������������, " _
                    & "T1.�з��ڵ�, T4.�з���, T1.�����ڵ�, T5.�����, " _
                    & "T5.�԰�, T5.����, T1.�������, T1.�԰�ܰ�, T1.�԰�ΰ� "
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
            .TextMatrix(lngR, 3) = P_adoRec("����ó�ڵ�"): .TextMatrix(lngR, 4) = P_adoRec("����ó��")
            .Cell(flexcpBackColor, lngR, 0, lngR, .Cols - 1) = vbYellow
            Do Until P_adoRec.EOF
               lngR = lngR + 1
               If lngR <> 2 Then    'ó�� ���ڵ� �ƴϸ�
                  If .TextMatrix(lngR - 1, 1) <> P_adoRec("����ó�ڵ�") Then '����ó�ڵ尡 �ٸ���
                     If .TextMatrix(lngR - 1, 3) <> "" Then '�� �����
                        '�ش���������
                         .AddItem ""
                         '.TextMatrix(lngR, 4) = "(" & Format(Mid(StrDate, 1, 6), "0000-00") & Space(1) & "����)"
                         .TextMatrix(lngR, 4) = "(��    ��)"
                         .TextMatrix(lngR, 11) = curMQAmt         '������
                         .TextMatrix(lngR, 13) = curMUMny         '�����Աݾ�
                         .TextMatrix(lngR, 15) = (curMUMny * (PB_curVatRate + 1)) '�����Աݾ�(VAT)
                         .TextMatrix(lngR, 14) = .ValueMatrix(lngR, 15) - .ValueMatrix(lngR, 13)  '���ΰ����ݾ�
                         curMQAmt = 0: curMUMny = 0
                         lngR = lngR + 1
                     End If
                     '�ش���(��)����
                     .TextMatrix(lngR, 4) = "(�Ⱓ����)"
                     .TextMatrix(lngR, 11) = curTQAmt             '��ü��������
                     .TextMatrix(lngR, 13) = curTUMny             '��ü���Աݾ״���
                     .TextMatrix(lngR, 15) = (curTUMny * (PB_curVatRate + 1))     '��ü���Աݾ״���(VAT)
                     .TextMatrix(lngR, 14) = .ValueMatrix(lngR, 15) - .ValueMatrix(lngR, 13)  '��ü�ΰ����ݾ״���
                     curTQAmt = 0: curTUMny = 0
                     curJanMny = 0 '��(��)�̿� �ܾ�
                     .AddItem ""
                     lngR = lngR + 1
                     .AddItem ""
                     lngR = lngR + 1
                     '����ó��
                     .TextMatrix(lngR, 3) = P_adoRec("����ó�ڵ�"): .TextMatrix(lngR, 4) = P_adoRec("����ó��")
                     .Cell(flexcpBackColor, lngR, 0, lngR, .Cols - 1) = vbYellow
                     .AddItem ""
                     lngR = lngR + 1
                  Else                                                          '����ó�ڵ� ������
                     If .TextMatrix(lngR - 1, 3) <> "" And _
                         Mid(StrDate, 1, 6) <> Mid(P_adoRec("����"), 1, 6) Then '�� ����� And ���� �ٸ���
                         '�ش����(��)����
                         .AddItem ""
                         '.TextMatrix(lngR, 4) = "(" & Format(Mid(StrDate, 1, 6), "0000-00") & Space(1) & "����)"
                         .TextMatrix(lngR, 4) = "(��    ��)"
                         .TextMatrix(lngR, 11) = curMQAmt         '������
                         .TextMatrix(lngR, 13) = curMUMny         '�����Աݾ�
                         .TextMatrix(lngR, 15) = (curMUMny * (PB_curVatRate + 1)) '�����Աݾ�(VAT)
                         .TextMatrix(lngR, 14) = .ValueMatrix(lngR, 15) - .ValueMatrix(lngR, 13)  '���ΰ����ݾ�
                         curMQAmt = 0: curMUMny = 0
                         lngR = lngR + 1
                     End If
                  End If
               End If
               .TextMatrix(lngR, 0) = P_adoRec("������ڵ�")
               .TextMatrix(lngR, 1) = P_adoRec("����ó�ڵ�")
               .TextMatrix(lngR, 2) = P_adoRec("����ó��")
               '3. ����
               If Mid(P_adoRec("����"), 7, 2) = "00" Then
                  .TextMatrix(lngR, 3) = ""
               Else
                  .TextMatrix(lngR, 3) = Format(P_adoRec("����"), "0000-00-00")
               End If
               '4. ����
               If Mid(P_adoRec("����"), 7, 2) = "00" Then
                  If Mid(P_adoRec("����"), 5, 2) = "00" Then
                     .TextMatrix(lngR, 4) = "(" & Mid(P_adoRec("����"), 1, 4) & "�� �̿��ܾ�)"
                  Else
                     .TextMatrix(lngR, 4) = "(" & Format(Mid(P_adoRec("����"), 1, 6), "0000-00") & "�� �ܾ�)"
                  End If
               End If
               If Len(P_adoRec("�з��ڵ�")) > 0 Then
                  .TextMatrix(lngR, 4) = P_adoRec("�з��ڵ�") & P_adoRec("�����ڵ�")
               End If
               .TextMatrix(lngR, 5) = P_adoRec("�з��ڵ�") & P_adoRec("�����ڵ�")
               .TextMatrix(lngR, 6) = P_adoRec("�����")
               .TextMatrix(lngR, 7) = P_adoRec("�԰�")
               .TextMatrix(lngR, 8) = P_adoRec("����")
               .TextMatrix(lngR, 9) = P_adoRec("����")
               '10. �������
               If P_adoRec("����") = 0 Then
                  .TextMatrix(lngR, 10) = ""
               ElseIf _
                  P_adoRec("����") = 1 Then
                  .TextMatrix(lngR, 10) = "����"
               ElseIf _
                  P_adoRec("����") = 7 Then
                  .TextMatrix(lngR, 10) = "����"
               End If
               .TextMatrix(lngR, 11) = P_adoRec("�԰����")                          '����
               .TextMatrix(lngR, 12) = P_adoRec("�԰�ܰ�")                          'ǰ�� ����ܰ�
               .TextMatrix(lngR, 13) = P_adoRec("�԰����") * P_adoRec("�԰�ܰ�")   '�ش����� ���Աݾ�
               .TextMatrix(lngR, 14) = P_adoRec("�԰�ݾ�") - .ValueMatrix(lngR, 13) '�ش����� ���Ժΰ���
                If Mid(P_adoRec("����"), 7, 2) = "00" Then
                   curJanMny = curJanMny + P_adoRec("�԰�ݾ�") '��(��)�̿� ����
                   .TextMatrix(lngR, 15) = curJanMny
               Else
                  .TextMatrix(lngR, 15) = P_adoRec("�԰�ݾ�")                       '�ش����� ���Աݾ�(�ΰ�������)
               End If
               If Mid(P_adoRec("����"), 7, 2) <> "00" Then
                  curMQAmt = curMQAmt + P_adoRec("�԰����")
                  curMUMny = curMUMny + (P_adoRec("�԰����") * P_adoRec("�԰�ܰ�"))
               End If
               '�ش紩��ݾ�
               curTQAmt = curTQAmt + P_adoRec("�԰����")
               curTUMny = curTUMny + (P_adoRec("�԰�ܰ�") * P_adoRec("�԰����"))
               curTTQAmt = curTTQAmt + P_adoRec("�԰����")
               curTTUMny = curTTUMny + (P_adoRec("�԰�ܰ�") * P_adoRec("�԰����"))
               StrDate = P_adoRec("����")
               'FindRow ����� ����
               '.TextMatrix(lngR, 4) = .TextMatrix(lngR, 0) & P_adoRec("���з��ڵ�")
               '.Cell(flexcpData, lngR, 4, lngR, 4) = .TextMatrix(lngR, 4)
               P_adoRec.MoveNext
               If P_adoRec.EOF = True Then '������ ���ڵ��
                  If .TextMatrix(lngR, 3) <> "" Then  '�� �����
                     lngR = lngR + 1
                     '�ش������
                     .AddItem ""
                     .TextMatrix(lngR, 4) = "(��    ��)"
                     .TextMatrix(lngR, 11) = curMQAmt         '������
                     .TextMatrix(lngR, 13) = curMUMny         '�����Աݾ�
                     .TextMatrix(lngR, 15) = (curMUMny * (PB_curVatRate + 1)) '�����Աݾ�(VAT)
                     .TextMatrix(lngR, 14) = .ValueMatrix(lngR, 15) - .ValueMatrix(lngR, 13)  '���ΰ����ݾ�
                     curMQAmt = 0: curMUMny = 0
                  End If
                  '�ش紩���
                  lngR = lngR + 1
                  .AddItem ""
                  .TextMatrix(lngR, 4) = "(�Ⱓ����)"
                  .TextMatrix(lngR, 11) = curTQAmt            '��ü��������
                  .TextMatrix(lngR, 13) = curTUMny            '��ü���Աݾ״���
                  .TextMatrix(lngR, 15) = (curTUMny * (PB_curVatRate + 1))    '��ü���Աݾ״���(VAT)
                  .TextMatrix(lngR, 14) = .ValueMatrix(lngR, 15) - .ValueMatrix(lngR, 13)  '��ü�ΰ����ݾ״���
                  curTQAmt = 0: curTUMny = 0
               End If
            Loop
            P_adoRec.Close
            '��ü �հ�
            If chkTotal.Value = 1 Then
               lngR = lngR + 1
               .AddItem ""
               lngR = lngR + 1
               .AddItem ""
               .TextMatrix(lngR, 4) = "(��ü����)"
               .TextMatrix(lngR, 11) = curTTQAmt              '��ü��������
               .TextMatrix(lngR, 13) = curTTUMny              '��ü���Աݾ״���
               .TextMatrix(lngR, 15) = (curTTUMny * (PB_curVatRate + 1))      '��ü���Աݾ״���(VAT)
               .TextMatrix(lngR, 14) = .ValueMatrix(lngR, 15) - .ValueMatrix(lngR, 13)  '��ü�ΰ����ݾ״���
               .Cell(flexcpBackColor, lngR, 0, lngR, .Cols - 1) = vbYellow
            End If
            If .Rows <= PC_intRowCnt Then
               .ScrollBars = flexScrollBarHorizontal
            Else
               .ScrollBars = flexScrollBarBoth
            End If
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
            .TopRow = 1
            vsfg1_EnterCell       'vsfg1_EnterCell �ڵ�����(���� �Ѱ� �϶��� ������ �ڵ�����)
            .SetFocus
       End With
    End If
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����ó������ �б� ����"
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
               Case 0  '����ó�ڵ�
                    If Not (Len(Text1(lngC).Text) > 0) Then
                       Exit Function
                    End If
               Case 1  '����ó��
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
'/// ũ����Ż ������ ��� ///
'+---------------------------+
Private Sub cmdPrint_Click()
Dim strSQL                 As String
Dim strWhere               As String
Dim strOrderBy             As String

Dim varRetVal              As Variant '������ ����
Dim strExeFile             As String
Dim strExeMode             As String
Dim intRetCHK              As Integer '���࿩��

Dim lngR                   As Long
Dim lngC                   As Long
Dim strForPrtDateTime      As String  '����Ͻ�           (Formula)
    
    If Len(Trim(Text1(0).Text)) = 0 And (chkTotal.Value = 0) Then
       Exit Sub
    End If
    If DTOS(dtpF_Date.Value) > DTOS(dtpT_Date.Value) Then
       Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    '�����Ͻ�(����Ͻ�)
    strSQL = "SELECT CONVERT(VARCHAR(19),GETDATE(), 120) AS �����ð� "
    On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    strForPrtDateTime = Format(PB_regUserinfoU.UserServerDate, "0000-00-00") & Space(1) _
                      & Format(Right(P_adoRec("�����ð�"), 8), "hh:mm:ss")
    P_adoRec.Close
    
    intRetCHK = 99
    With CrystalReport1
         If PB_Test = 0 Then
            strExeFile = App.Path & ".\����ó������.rpt"
         Else
            strExeFile = App.Path & ".\����ó������T.rpt"
         End If
         varRetVal = Dir(strExeFile)
         If Len(varRetVal) = 0 Then
            intRetCHK = 0
         Else
            .ReportFileName = strExeFile
            On Error GoTo ERROR_CRYSTAL_REPORTS
            '--- Formula Fields ---
            .Formulas(0) = "ForAppPgDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'" '���α׷���������
            .Formulas(1) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                       '������
            .Formulas(2) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                                  '����Ͻ�
            '�����������
            .Formulas(3) = "ForAppDate = '�������� : ' & '" & Format(DTOS(dtpF_Date.Value), "0000-00-00") & "' & ' ���� ' & '" & Format(DTOS(dtpT_Date.Value), "0000-00-00") & "' & ' ����' "
            '--- ParaMeter Fields ---
            '���α׷���������
            .StoredProcParam(0) = PB_regUserinfoU.UserClientDate
            '����(����)���ڽ���
            .StoredProcParam(1) = DTOS(dtpF_Date.Value)
            '����(����)��������
            .StoredProcParam(2) = DTOS(dtpT_Date.Value)
            '����ó�ڵ�
            If (Len(Text1(0).Text) = 0) Or (chkTotal.Value = 1) Then
               .StoredProcParam(3) = " "
            Else
               .StoredProcParam(3) = Trim(Text1(0).Text)
            End If
            .StoredProcParam(4) = PB_regUserinfoU.UserBranchCode '�����ڵ�
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
            .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "����ó������"
            .Action = 1
            .Reset
         End If
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "(�������� ���� ����)"
    Unload Me
    Exit Sub
ERROR_CRYSTAL_REPORTS:
    MsgBox Err.Number & Space(1) & Err.Description
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

