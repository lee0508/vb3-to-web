VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm������� 
   BorderStyle     =   0  '����
   Caption         =   "�������"
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
      TabIndex        =   8
      Top             =   0
      Width           =   15195
      Begin VB.CommandButton cmdPrint 
         Height          =   375
         Left            =   7980
         Picture         =   "�������.frx":0000
         Style           =   1  '�׷���
         TabIndex        =   17
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   9120
         Picture         =   "�������.frx":0963
         Style           =   1  '�׷���
         TabIndex        =   13
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   13635
         Picture         =   "�������.frx":1308
         Style           =   1  '�׷���
         TabIndex        =   12
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   390
         Left            =   12510
         Picture         =   "�������.frx":1C56
         Style           =   1  '�׷���
         TabIndex        =   11
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   11385
         Picture         =   "�������.frx":25DA
         Style           =   1  '�׷���
         TabIndex        =   10
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Height          =   390
         Left            =   10260
         Picture         =   "�������.frx":2E61
         Style           =   1  '�׷���
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
         Alignment       =   2  '��� ����
         Caption         =   "-"
         Height          =   240
         Index           =   11
         Left            =   6240
         TabIndex        =   20
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00008000&
         BorderStyle     =   1  '���� ����
         Caption         =   "�� �� �� ��"
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
      Height          =   1005
      Left            =   60
      TabIndex        =   7
      Top             =   630
      Width           =   15195
      Begin VB.TextBox txtFindNM 
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '����
         Left            =   7800
         MaxLength       =   50
         TabIndex        =   4
         Top             =   600
         Width           =   3525
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Left            =   7800
         MaxLength       =   18
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox cboTaxGbn 
         Height          =   300
         Left            =   2400
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   1
         Top             =   600
         Width           =   1110
      End
      Begin VB.ComboBox cboMt 
         Height          =   300
         Left            =   2400
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   0
         Top             =   195
         Width           =   3735
      End
      Begin VB.ComboBox cboState 
         Height          =   300
         Left            =   5050
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   2
         Top             =   600
         Width           =   1080
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "ǰ��"
         Height          =   240
         Index           =   13
         Left            =   6960
         TabIndex        =   24
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "ǰ���ڵ�"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   6600
         TabIndex        =   23
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         Caption         =   "[Home]"
         Height          =   240
         Index           =   1
         Left            =   10245
         TabIndex        =   22
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��������"
         Height          =   240
         Index           =   0
         Left            =   1245
         TabIndex        =   18
         Top             =   660
         Width           =   975
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
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�з�"
         Height          =   240
         Index           =   34
         Left            =   1245
         TabIndex        =   15
         Top             =   250
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��뱸��"
         Height          =   240
         Index           =   25
         Left            =   4000
         TabIndex        =   14
         Top             =   660
         Width           =   855
      End
   End
End
Attribute VB_Name = "frm�������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' ���α׷� �� �� : �������
' ���� Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' ������ Table   : �������, ����, ����з�
'                  ������帶��, �����������
' ��  ��  ��  �� :
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived       As Boolean
Private P_adoRec           As New ADODB.Recordset
Private P_intButton        As Integer
Private P_strFindString1   As String
Private P_strFindString2   As String
Private P_strSortM(1000)   As String
Private P_strSortS(1000)   As String
Private Const PC_intRowCnt As Integer = 22  '�׸��� �� ������ �� ���(FixedRows ����)

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
                   cmdSave.Enabled = True: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 50 '����, �߰�, ����, ��ȸ, �μ�
                   cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = True: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 99 '����, �߰�, ����, ��ȸ, �μ�
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
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "(�������� ���� ����)"
    Unload Me
    Exit Sub
End Sub

'+---------------+
'/// �μ����� ///
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
    '�ڵ� ��ħ
    'SendKeys "{F4}"
    '�ڵ� ��ħ
    'nRet = SendMessage(cboFdMtGp(Index).hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
    'ListIndex�� ���� �ٲپ Click �̺�Ʈ�� �߻����� �ʵ��� ��.
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
    If (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn) Then '����˻�
       PB_strCallFormName = "frm�������"
       PB_strMaterialsCode = Trim(txtCode.Text)
       PB_strMaterialsName = txtFindNM.Text
       frm����˻�.Show vbModal
       If (Len(PB_strMaterialsCode) + Len(PB_strMaterialsName)) = 0 Then '�˻����� ���(ESC)
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
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "������� �б� ����"
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
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "������� ����"
    Unload Me
    Exit Sub
End Sub

'+----------------------------+
'/// txtFindNM(�����˻�) ///
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
            If Len(.TextMatrix(.Row, 0)) <> 0 Then '15.���������, 18.����
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
               If (Col = 15) Then         '���������
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
                  (Col = 18) Then '���� ���� �˻�
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
    If Len(txtCode.Text) = 0 Then
       txtCode.SetFocus
       Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    cmdFind.Enabled = False
    P_strFindString1 = Trim(txtCode.Text)     '��ȸ�� ��� �˻��� �����ڵ�
    'P_strFindString2 = Trim(txtFindNM.Text)  '��ȸ�� ��� �˻��� ����� ����
    Subvsfg1_FILL
    cmdFind.Enabled = True
    vsfg1.SetFocus
    Screen.MousePointer = vbDefault
End Sub
'+-----------+
'/// ���� ///
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
Dim CurInputMny    As Currency '�԰�ܰ�
Dim CurInputVat    As Currency '�԰�ΰ�
Dim CurOutPutMny   As Currency '���ܰ�
Dim CurOutPutVat   As Currency '���ΰ�

    If (vsfg1.Rows) = 1 Then
       Exit Sub
    End If
    With vsfg1
         If .ValueMatrix(.Row, 15) = 0 Then
            Exit Sub
         End If
    End With
    intRetVal = MsgBox("������ ��������� ���� ���� �����Ͻðڽ��ϱ� ?", vbQuestion + vbYesNo + vbDefaultButton1, "�ڷ� ����")
    If intRetVal = vbNo Then
       vsfg1.SetFocus
       Exit Sub
    End If
    '�����ð� ���ϱ�
    Screen.MousePointer = vbHourglass
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT RIGHT(CONVERT(VARCHAR(23),GETDATE(), 121),12) AS �����ð� "
    On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    strServerTime = Mid(P_adoRec("�����ð�"), 1, 2) + Mid(P_adoRec("�����ð�"), 4, 2) + Mid(P_adoRec("�����ð�"), 7, 2) _
                  + Mid(P_adoRec("�����ð�"), 10)
    P_adoRec.Close
    strTime = strServerTime
    
    If vsfg1.ValueMatrix(vsfg1.Row, 15) > 0 Then '�԰�(+)
       strSQL = "SELECT TOP 1 " _
                  & "ISNULL(T1.�԰�ܰ�1, 0) AS �԰�ܰ�, ISNULL(ROUND(T1.�԰�ܰ�1 * (" & PB_curVatRate & "), 0, 1), 0) AS �԰�ΰ�, " _
                  & "ISNULL(T1.���ܰ�1, 0) AS ���ܰ�, ISNULL(ROUND(T1.���ܰ�1 * (" & PB_curVatRate & "), 0, 1), 0) AS ���ΰ� " _
                & "FROM ������� T1 " _
               & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                 & "AND T1.�з��ڵ� = '" & Mid(vsfg1.TextMatrix(vsfg1.Row, 4), 1, 2) & "' " _
                 & "AND T1.�����ڵ� = '" & Mid(vsfg1.TextMatrix(vsfg1.Row, 4), 3) & "' "
    Else
       strSQL = "SELECT TOP 1 " _
                  & "ISNULL(T1.�԰�ܰ�1, 0) AS �԰�ܰ�, ISNULL(ROUND(T1.�԰�ܰ�1 * (" & PB_curVatRate & "), 0, 1), 0) AS �԰�ΰ�, " _
                  & "ISNULL(T1.���ܰ�1, 0) AS ���ܰ�, ISNULL(ROUND(T1.���ܰ�1 * (" & PB_curVatRate & "), 0, 1), 0) AS ���ΰ� " _
                & "FROM ������� T1 " _
               & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                 & "AND T1.�з��ڵ� = '" & Mid(vsfg1.TextMatrix(vsfg1.Row, 4), 1, 2) & "' " _
                 & "AND T1.�����ڵ� = '" & Mid(vsfg1.TextMatrix(vsfg1.Row, 4), 3) & "' "
    End If
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount <> 0 Then
       CurInputMny = P_adoRec("�԰�ܰ�"): CurInputVat = P_adoRec("�԰�ΰ�")
       CurOutPutMny = P_adoRec("���ܰ�"): CurOutPutVat = P_adoRec("���ΰ�")
    End If
    P_adoRec.Close
    PB_adoCnnSQL.BeginTrans
    P_adoRec.CursorLocation = adUseClient
    With vsfg1
         If .ValueMatrix(.Row, 15) > 0 Then '�԰�(+)
            '�ŷ���ȣ ���ϱ�
            P_adoRec.CursorLocation = adUseClient
            strSQL = "spLogCounter '�������⳻��', '" & PB_regUserinfoU.UserBranchCode + PB_regUserinfoU.UserClientDate + "5" & "', " _
                                 & "0, 0, '" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' "
            On Error GoTo ERROR_STORED_PROCEDURE
            P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
            lngLogCnt = P_adoRec(0)
            P_adoRec.Close
            '�������⳻��
            strSQL = "INSERT INTO �������⳻��(������ڵ�, �з��ڵ�, " _
                                            & "�����ڵ�, �������, " _
                                            & "���������, �����ð�, " _
                                            & "�԰����, �԰�ܰ�, " _
                                            & "�԰�ΰ�, ������, " _
                                            & "���ܰ�, ���ΰ�, " _
                                            & "����ó�ڵ�, ����ó�ڵ�, " _
                                            & "�������������, ���۱���, " _
                                            & "�߰�����, �߰߹�ȣ, �ŷ�����, �ŷ���ȣ, " _
                                            & "��꼭���࿩��, ���ݱ���, ��������, ����, �ۼ��⵵, å��ȣ, �Ϸù�ȣ, " _
                                            & "��뱸��, ��������, " _
                                            & "������ڵ�, ����̵�������ڵ�) VALUES( " _
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
         Else                            '���(-)
            '�ŷ���ȣ ���ϱ�
            P_adoRec.CursorLocation = adUseClient
            strSQL = "spLogCounter '�������⳻��', '" & PB_regUserinfoU.UserBranchCode + PB_regUserinfoU.UserClientDate + "6" & "', " _
                                 & "0, 0, '" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' "
            On Error GoTo ERROR_STORED_PROCEDURE
            P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
            lngLogCnt = P_adoRec(0)
            P_adoRec.Close
            strSQL = "INSERT INTO �������⳻��(������ڵ�, �з��ڵ�, " _
                                            & "�����ڵ�, �������, " _
                                            & "���������, �����ð�, " _
                                            & "�԰����, �԰�ܰ�, " _
                                            & "�԰�ΰ�, ������, " _
                                            & "���ܰ�, ���ΰ�, " _
                                            & "����ó�ڵ�, ����ó�ڵ�, " _
                                            & "�������������, ���۱���, " _
                                            & "�߰�����, �߰߹�ȣ, �ŷ�����, �ŷ���ȣ, " _
                                            & "��꼭���࿩��, ���ݱ���, ��������, ����, �ۼ��⵵, å��ȣ, �Ϸù�ȣ, " _
                                            & "��뱸��, ��������, " _
                                            & "������ڵ�, ����̵�������ڵ�) VALUES( " _
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
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "(�������� ���� ����)"
    Unload Me
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "������� �߰� ����"
    Unload Me
    Exit Sub
ERROR_TABLE_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "������� �߰� ����"
    Unload Me
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "������� ���� ����"
    Unload Me
    Exit Sub
ERROR_STORED_PROCEDURE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�α� ���� ����"
    Unload Me
    Exit Sub
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
    Set frm������� = Nothing
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
         .Rows = 1             'Subvsfg1_Fill����ÿ� ����
         .Cols = 19
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1000   '�з��ڵ�       'H
         .ColWidth(1) = 1200   '�з���
         .ColWidth(2) = 1900   'ǰ���ڵ�(�з��ڵ�+�����ڵ�)
         .ColWidth(3) = 2500   'ǰ���
         .ColWidth(4) = 1900   '�з��ڵ� + �����ڵ� 'H
         .ColWidth(5) = 2500   '�԰�
         .ColWidth(6) = 1000   '����
         .ColWidth(7) = 1000   '�����         'H
         .ColWidth(8) = 1000   '��������       'H
         .ColWidth(9) = 1000   '��뱸��       'H
         .ColWidth(10) = 1200  '�������
         .ColWidth(11) = 1200  '�̿����       'H
         .ColWidth(12) = 1200  '�԰����       'H
         .ColWidth(13) = 1200  '������       'H
         .ColWidth(14) = 1200  '�������
         .ColWidth(15) = 1200  '��������
         .ColWidth(16) = 1200  '�����԰�����   'H
         .ColWidth(17) = 1200  '�����������   'H
         .ColWidth(18) = 5000  '����
         .Cell(flexcpFontBold, 0, 0, .FixedRows - 1, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "�з��ڵ�"         'H
         .TextMatrix(0, 1) = "�з���"
         .TextMatrix(0, 2) = "ǰ���ڵ�"
         .TextMatrix(0, 3) = "ǰ��"
         .TextMatrix(0, 4) = "(�з�+����)�ڵ�"  'H
         .TextMatrix(0, 5) = "�԰�"
         .TextMatrix(0, 6) = "����"
         .TextMatrix(0, 7) = "�����"           'H
         .TextMatrix(0, 8) = "��������"         'H
         .TextMatrix(0, 9) = "��뱸��"         'H
         .TextMatrix(0, 10) = "�������"
         .TextMatrix(0, 11) = "�̿����"        'H
         .TextMatrix(0, 12) = "���Լ���"        'H
         .TextMatrix(0, 13) = "�������"        'H
         .TextMatrix(0, 14) = "�������"
         .TextMatrix(0, 15) = "������(+/-)"
         .TextMatrix(0, 16) = "�����԰�����"    'H
         .TextMatrix(0, 17) = "�����������"    'H
         .TextMatrix(0, 18) = "����"
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
    strSQL = "SELECT T1.�з��ڵ� AS �з��ڵ�, " _
                  & "ISNULL(T1.�з���,'') AS �з��� " _
             & "FROM ����з� T1 " _
            & "ORDER BY T1.�з��ڵ� "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       cboMt.ListIndex = -1
       Exit Sub
    Else
       cboMt.AddItem "00. " & "��ü"
       Do Until P_adoRec.EOF
          cboMt.AddItem P_adoRec("�з��ڵ�") & ". " & P_adoRec("�з���")
          P_adoRec.MoveNext
       Loop
       P_adoRec.Close
       cboMt.ListIndex = 0
    End If
    With cboState
         .AddItem "��    ü"
         .AddItem "��    ��"
         .AddItem "���Ұ�"
         .AddItem "��    Ÿ"
         .ListIndex = 1
    End With
    With cboTaxGbn
         .AddItem "��    ü"
         .AddItem "�� �� ��"
         .AddItem "��    ��"
         .ListIndex = 0
    End With
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����з� �б� ����"
    Unload Me
    Exit Sub
End Sub

'+---------------------------------+
'/// VsFlexGrid(vsfg1) ä���///
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
         '�˻����� ����з�
         Select Case Mid(Trim(cboMt.Text), 1, 2)
                Case "00"      '�з� ��ü
                     strWhere = ""
                Case Else      '�з� ��ü �ƴϸ�
                     strWhere = "WHERE T1.�з��ڵ� = '" & Mid(Trim(cboMt.Text), 1, 2) & "' "
         End Select
         strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                  & "T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' "
    End With
    '�˻����� ��������
    Select Case cboTaxGbn.ListIndex
           Case 0 '��ü
                strWhere = strWhere
           Case 1 '�����
                strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") + "T3.�������� = 0 "
           Case 2 '����
                strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") + "T3.�������� = 1 "
    End Select
    '�˻����� ��뱸��
    Select Case cboState.ListIndex
           Case 0 '��ü
                strWhere = strWhere
           Case 1 '����
                strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") + "T1.��뱸�� = 0 "
           Case 2 '���Ұ�
                strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") + "T1.��뱸�� = 9 "
           Case Else
                strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") + "NOT(T1.��뱸�� = 0 OR T1.��뱸�� = 9) "
    End Select
    If Len(P_strFindString1) = 0 Then            '�������� ��ȸ
       strOrderBy = "ORDER BY T1.������ڵ�, T1.�з��ڵ�, T1.�����ڵ� "
    Else
       'strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                           & "T1.�����ڵ� LIKE '%" & P_strFindString1 & "%' " _
                           & "AND T3.����� LIKE '%" & P_strFindString2 & "%' "
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                           & "T1.�з��ڵ� = '" & Mid(P_strFindString1, 1, 2) & "' AND T1.�����ڵ� = '" & Mid(P_strFindString1, 3) & "' "
                           
       strOrderBy = "ORDER BY T3.����� "
    End If
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.������ڵ� AS ������ڵ�, T2.������ AS ������, " _
                  & "ISNULL(T1.�з��ڵ�,'') AS �з��ڵ�, ISNULL(T4.�з���,'') AS �з���, " _
                  & "ISNULL(T1.�����ڵ�,'') AS �����ڵ�, T3.����� AS �����, " _
                  & "T3.�԰� AS �԰�, T3.���� AS ����, T3.����� AS �����, T3.�������� AS ��������, " _
                  & "T1.��뱸�� AS ��뱸��, T1.������� AS �������, " _
                  & "ISNULL(T1.�����԰�����,'') AS �����԰�����, ISNULL(T1.�����������,'') AS �����������, " _
                  & "(SELECT ISNULL(SUM(�԰������-��������),0) " _
                     & "FROM ������帶�� " _
                    & "WHERE �з��ڵ� = T1.�з��ڵ� AND �����ڵ� = T1.�����ڵ� " _
                      & "AND ������ڵ� = T1.������ڵ� " _
                      & "AND ������� >= '" & Left(PB_regUserinfoU.UserClientDate, 4) + "00" & "' " _
                      & "AND ������� < '" & Left(PB_regUserinfoU.UserClientDate, 6) & "') AS �̿����, "
    strSQL = strSQL _
                  & "(SELECT ISNULL(SUM(�԰����),0) " _
                     & "FROM �������⳻�� " _
                    & "WHERE �з��ڵ� = T1.�з��ڵ� AND �����ڵ� = T1.�����ڵ� " _
                      & "AND ������ڵ� = T1.������ڵ� AND ��뱸�� = 0 " _
                      & "AND ��������� BETWEEN '" & Left(PB_regUserinfoU.UserClientDate, 6) + "01" & "' " _
                                         & "AND '" & PB_regUserinfoU.UserClientDate & "') AS �԰����, "
    strSQL = strSQL _
                  & "(SELECT ISNULL(SUM(������),0) " _
                     & "FROM �������⳻�� " _
                    & "WHERE �з��ڵ� = T1.�з��ڵ� AND �����ڵ� = T1.�����ڵ� " _
                      & "AND ������ڵ� = T1.������ڵ� AND ��뱸�� = 0 " _
                      & "AND ��������� BETWEEN '" & Left(PB_regUserinfoU.UserClientDate, 6) + "01" & "' " _
                                         & "AND '" & PB_regUserinfoU.UserClientDate & "') AS ������ "
    strSQL = strSQL _
             & "FROM ������� T1 " _
             & "LEFT JOIN ����� T2 ON T2.������ڵ� = T1.������ڵ� " _
             & "LEFT JOIN ���� T3 " _
                    & "ON T3.�з��ڵ� = T1.�з��ڵ� AND T3.�����ڵ� = T1.�����ڵ� " _
             & "LEFT JOIN ����з� T4 ON T4.�з��ڵ� = T1.�з��ڵ� "
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
               .TextMatrix(lngR, 0) = P_adoRec("�з��ڵ�")
               .TextMatrix(lngR, 1) = IIf(IsNull(P_adoRec("�з���")), "", P_adoRec("�з���"))
               .TextMatrix(lngR, 2) = P_adoRec("�з��ڵ�") & P_adoRec("�����ڵ�")
               .TextMatrix(lngR, 3) = IIf(IsNull(P_adoRec("�����")), "", P_adoRec("�����"))
               'FindRow ����� ����
               .TextMatrix(lngR, 4) = .TextMatrix(lngR, 0) & P_adoRec("�����ڵ�")
               .Cell(flexcpData, lngR, 4, lngR, 4) = .TextMatrix(lngR, 4)
               .TextMatrix(lngR, 5) = IIf(IsNull(P_adoRec("�԰�")), "", P_adoRec("�԰�"))
               .TextMatrix(lngR, 6) = IIf(IsNull(P_adoRec("����")), "", P_adoRec("����"))
               .TextMatrix(lngR, 7) = IIf(IsNull(P_adoRec("�����")), "", P_adoRec("�����"))
               If P_adoRec("��������") = 0 Then
                  .TextMatrix(lngR, 8) = "�� �� ��"
               Else
                  .TextMatrix(lngR, 8) = "��    ��"
               End If
               If P_adoRec("��뱸��") = 0 Then
                  .TextMatrix(lngR, 9) = "��    ��"
               ElseIf _
                  P_adoRec("��뱸��") = 9 Then
                  .TextMatrix(lngR, 9) = "���Ұ�"
               Else
                  .TextMatrix(lngR, 9) = "�ڵ����"
               End If
               .TextMatrix(lngR, 10) = IIf(IsNull(P_adoRec("�������")), "", P_adoRec("�������"))
               .TextMatrix(lngR, 11) = IIf(IsNull(P_adoRec("�̿����")), "", P_adoRec("�̿����"))
               .TextMatrix(lngR, 12) = IIf(IsNull(P_adoRec("�԰����")), "", P_adoRec("�԰����"))
               .TextMatrix(lngR, 13) = IIf(IsNull(P_adoRec("������")), "", P_adoRec("������"))
               .TextMatrix(lngR, 14) = .ValueMatrix(lngR, 11) + .ValueMatrix(lngR, 12) - .ValueMatrix(lngR, 13) '�������
               .TextMatrix(lngR, 15) = ""
               'If .ValueMatrix(lngR, 14) < .ValueMatrix(lngR, 10) Then
               '   .Cell(flexcpForeColor, lngR, 15, lngR, 15) = vbRed
               'End If
               If Len(P_adoRec("�����԰�����")) = 8 Then
                  .TextMatrix(lngR, 16) = Format(P_adoRec("�����԰�����"), "0000-00-00")
               End If
               If Len(P_adoRec("�����������")) = 8 Then
                  .TextMatrix(lngR, 17) = Format(P_adoRec("�����������"), "0000-00-00")
               End If
               If .TextMatrix(lngR, 2) = P_strFindString1 Then
                  lngRR = lngR
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
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "������� �б� ����"
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
            strExeFile = App.Path & ".\�����������.rpt"
         Else
            strExeFile = App.Path & ".\�����������T.rpt"
         End If
         varRetVal = Dir(strExeFile)
         If Len(varRetVal) = 0 Then
            intRetCHK = 0
         Else
            .ReportFileName = strExeFile
            On Error GoTo ERROR_CRYSTAL_REPORTS
            '--- Formula Fields ---
            .Formulas(0) = "ForPgDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'" '���α׷���������
            .Formulas(1) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                    '������
            .Formulas(2) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                               '����Ͻ�
            .Formulas(3) = "ForAppDate = '�������� : ' & '" & Format(DTOS(dtpFDate.Value), "0000-00-00") & "' & ' ���� ' & '" & Format(DTOS(dtpTDate.Value), "0000-00-00") & "' & ' ����' "
            '--- Formula Fields(Select Record) ---
            .Formulas(4) = "ForSelKindCode = '" & Mid(cboMt.Text, 1, 2) & "'"                           '�з��ڵ�
            If cboTaxGbn.ListIndex = 1 Then       '�����
               .Formulas(5) = "ForSelTaxGbn = 0"
            ElseIf _
               cboTaxGbn.ListIndex = 2 Then       '��  ��
               .Formulas(5) = "ForSelTaxGbn = 1"
            Else
               .Formulas(5) = "ForSelTaxGbn = 2"  '��  ü
            End If
            If cboState.ListIndex = 1 Then         '��    ��
               .Formulas(6) = "ForSelUsageGbn = 0"
            ElseIf _
               cboState.ListIndex = 2 Then         '���Ұ�
               .Formulas(6) = "ForSelUsageGbn = 9"
            Else
               .Formulas(6) = "ForSelUsageGbn = 2"  '��   ü
            End If
            '--- Parameter Fields ---
            .StoredProcParam(0) = PB_regUserinfoU.UserBranchCode  '�����ڵ�
            .StoredProcParam(1) = DTOS(dtpFDate.Value)           '��������(��������)
            .StoredProcParam(2) = DTOS(dtpTDate.Value)           '��������(��������)
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
            .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "�����������"
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

