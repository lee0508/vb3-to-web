VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm�����ޱݿ��� 
   BorderStyle     =   0  '����
   Caption         =   "�����ޱݿ���"
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
      TabIndex        =   16
      Top             =   0
      Width           =   15195
      Begin VB.OptionButton optPrtChk1 
         Caption         =   "�̸���"
         Height          =   255
         Left            =   6840
         TabIndex        =   33
         Top             =   390
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optPrtChk0 
         Caption         =   "�ڵ��"
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
         Picture         =   "�����ޱݿ���.frx":0000
         Style           =   1  '�׷���
         TabIndex        =   22
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   9120
         Picture         =   "�����ޱݿ���.frx":0963
         Style           =   1  '�׷���
         TabIndex        =   20
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   13635
         Picture         =   "�����ޱݿ���.frx":1308
         Style           =   1  '�׷���
         TabIndex        =   19
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   390
         Left            =   12510
         Picture         =   "�����ޱݿ���.frx":1C56
         Style           =   1  '�׷���
         TabIndex        =   18
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   11385
         Picture         =   "�����ޱݿ���.frx":25DA
         Style           =   1  '�׷���
         TabIndex        =   7
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Height          =   390
         Left            =   10260
         Picture         =   "�����ޱݿ���.frx":2E61
         Style           =   1  '�׷���
         TabIndex        =   12
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00008000&
         BorderStyle     =   1  '���� ����
         Caption         =   "�� �� ó �� �� ó ��"
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
      Height          =   1299
      Left            =   60
      TabIndex        =   14
      Top             =   630
      Width           =   15195
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   10  '�ѱ� 
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
            Name            =   "����ü"
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
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   2
         Top             =   560
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   3
         Left            =   2475
         MaxLength       =   14
         TabIndex        =   3
         Top             =   915
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Index           =   4
         Left            =   9240
         MaxLength       =   50
         TabIndex        =   6
         Top             =   915
         Width           =   5535
      End
      Begin VB.CheckBox chkTotal 
         Caption         =   "��ü ����ó"
         Height          =   255
         Left            =   7785
         TabIndex        =   9
         Top             =   200
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Index           =   1
         Left            =   4515
         MaxLength       =   50
         TabIndex        =   1
         Top             =   185
         Width           =   2895
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
         Alignment       =   1  '������ ����
         Caption         =   "������ȣ"
         Height          =   240
         Index           =   5
         Left            =   7920
         TabIndex        =   31
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����)"
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
         Index           =   4
         Left            =   120
         TabIndex        =   30
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��������"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   3
         Left            =   4155
         TabIndex        =   29
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�ݾ�"
         Height          =   240
         Index           =   2
         Left            =   1275
         TabIndex        =   28
         Top             =   945
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   8
         Left            =   7920
         TabIndex        =   27
         Top             =   945
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   12
         Left            =   13800
         TabIndex        =   26
         Top             =   245
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   11
         Left            =   11760
         TabIndex        =   25
         Top             =   245
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��������"
         Height          =   240
         Index           =   10
         Left            =   9360
         TabIndex        =   24
         Top             =   245
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
         Alignment       =   1  '������ ����
         Caption         =   "�������"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   1275
         TabIndex        =   15
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����ó�ڵ�"
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
Attribute VB_Name = "frm�����ޱݿ���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' ���α׷� �� �� : �����ޱݿ���
' ���� Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' ������ Table   :
' ��  ��  ��  �� : ��������̷�(�ŷ�(��ǥ)����)��ȸ + ����ó����ó��
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived       As Boolean
Private P_adoRec           As New ADODB.Recordset
Private P_intButton        As Integer
Private Const PC_intRowCnt As Integer = 26  '�׸��� �� ������ �� ���(FixedRows ����)

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
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����� ����(�������� ���� ����)"
    Unload Me
    Exit Sub
End Sub

'+---------------+
'/// �˻����� ///
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
'/// ������� ///
'+---------------+
Private Sub cboSactionWay_Click()
    With cboSactionWay
        If .ListIndex = 0 Or .ListIndex = 1 Then '���� �Ǵ� ��ǥ
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
'/// �������� ///
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
    'If chkTotal.Value = 0 And cmdSave.Enabled = True And vsfg1.Rows > 1 Then
    '   cboSactionWay.Enabled = True: Text1(3).Enabled = True: Text1(4).Enabled = True
    '   cboSactionWay.SetFocus
    '   Exit Sub
    'End If
End Sub

'+-----------+
'/// ���� ///
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
    '�Է³��� �˻�
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
       intRetVal = MsgBox("�Էµ� �ڷḦ �����Ͻðڽ��ϱ� ?", vbQuestion + vbYesNo + vbDefaultButton1, "�ڷ� ����")
       If intRetVal = vbNo Then
          vsfg1.SetFocus
          Exit Sub
       End If
       cmdSave.Enabled = False
    End If
    Screen.MousePointer = vbHourglass
    '�����ð� ���ϱ�
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT RIGHT(CONVERT(VARCHAR(23),GETDATE(), 121),12) AS �����ð� "
    On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    strServerTime = Mid(P_adoRec("�����ð�"), 1, 2) + Mid(P_adoRec("�����ð�"), 4, 2) + Mid(P_adoRec("�����ð�"), 7, 2) _
                  + Mid(P_adoRec("�����ð�"), 10)
    P_adoRec.Close
    strTime = strServerTime
    '�����ޱݳ���
    PB_adoCnnSQL.BeginTrans
    strSQL = "INSERT INTO �����ޱݳ���(������ڵ�, ����ó�ڵ�, " _
                                    & "�����ޱ���������, �����ޱ����޽ð�," _
                                    & "�����ޱ����ޱݾ�, �������, " _
                                    & "��������, ������ȣ, " _
                                    & "����, ��������, " _
                                    & "������ڵ�) VALUES(" _
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
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "(�������� ���� ����)"
    Unload Me
    Exit Sub
ERROR_TABLE_SELECT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�����ޱ� �߰� ����"
    Unload Me
    Exit Sub
ERROR_TABLE_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�����ޱ� �߰� ����"
    Unload Me
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�����ޱ� ���� ����"
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
    Set frm�����ޱݿ��� = Nothing
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
         .Rows = 1             'Subvsfg1_Fill����ÿ� ����
         .Cols = 13
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1000   '������ڵ�
         .ColWidth(1) = 1000   '����ó�ڵ�
         .ColWidth(2) = 2000   '����ó��
         .ColWidth(3) = 1200   '����
         .ColWidth(4) = 3000   '����
         .ColWidth(5) = 1000   '�������(����)
         .ColWidth(6) = 1000   '������и�(����)
         .ColWidth(7) = 1000   '��������(0000-00-00)
         .ColWidth(8) = 2000   '������ȣ
         .ColWidth(9) = 2000   '�����ޱݾ�
         .ColWidth(10) = 2000  '���޾�
         .ColWidth(11) = 2000  '�ܾ�
         .ColWidth(12) = 4000  '���
         
         .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "������ڵ�"  'H
         .TextMatrix(0, 1) = "����ó�ڵ�"  'H
         .TextMatrix(0, 2) = "����ó��"    'H
         .TextMatrix(0, 3) = "��¥"
         .TextMatrix(0, 4) = "����"
         .TextMatrix(0, 5) = "����"        'H
         .TextMatrix(0, 6) = "����"
         .TextMatrix(0, 7) = "��������"
         .TextMatrix(0, 8) = "������ȣ"
         .TextMatrix(0, 9) = "�ݾ�"
         .TextMatrix(0, 10) = "���޾�"
         .TextMatrix(0, 11) = "�ܾ�"
         .TextMatrix(0, 12) = "���"
         
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
'/// VsFlexGrid(vsfg1) ä���(�����ޱݹ߻����� : 1.��ǥ, 2.��꼭) ///
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
Dim StrDate     As String    '�ش�����
Dim curMonIMny  As Currency  '�ش������ݾ�(�԰�)
Dim curMonOMny  As Currency  '�ش������ݾ�(����)
Dim curMonTMny  As Currency  '�ش������ݾ�(�ܾ�)
Dim curTotIMny  As Currency  '�ش紩��ݾ�(�԰�)
Dim curTotOMny  As Currency  '�ش紩��ݾ�(����)
Dim curTotTMny  As Currency  '�ش紩��ݾ�(�ܾ�)
Dim curTotTIMny As Currency  '��ü����ݾ�(�԰�)
Dim curTotTOMny As Currency  '��ü����ݾ�(����)
Dim curTotTTMny As Currency  '��ü����ݾ�(�ܾ�)
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
    strOrderBy = "ORDER BY T1.������ڵ�, " & IIf(optPrtChk0.Value = True, "T1.����ó�ڵ�, T3.����ó�� ", "T3.����ó��, T1.����ó�ڵ� ") & ", ����, �ð�, ���� "
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.������ڵ� AS ������ڵ�, T2.������ AS ������, " _
                  & "T1.����ó�ڵ� AS ����ó�ڵ�, T3.����ó�� AS ����ó��, " _
                  & "(T1.������� + '00') AS ����, '(�� �� ��)' AS ����, " _
                  & "0 AS ����, (T1.�����ޱݴ���ݾ�) AS �԰�ݾ�, " _
                  & "(T1.�����ޱ����޴���ݾ�) AS ���ޱݾ�, " _
                  & "'' AS �������, '' AS ��������, '' AS ������ȣ, '' AS �ð� " _
             & "FROM �����ޱݿ��帶�� T1 " _
             & "LEFT JOIN ����� T2 ON T2.������ڵ� = T1.������ڵ� " _
             & "LEFT JOIN ����ó T3 ON T3.������ڵ� = T1.������ڵ� AND T3.����ó�ڵ� = T1.����ó�ڵ� " _
             & "" & strWhere & " AND SUBSTRING(T1.�������,5,2) = '00' AND (T1.�����ޱݴ���ݾ� <> 0 OR T1.�����ޱ����޴���ݾ� <> 0) " _
              & "AND T1.������� BETWEEN '" & (Mid(DTOS(dtpF_Date.Value), 1, 4) + "00") & "' " _
                                  & "AND '" & (Mid(DTOS(dtpT_Date.Value), 1, 4) + "00") & "' "
    strSQL = strSQL & "UNION ALL " _
           & "SELECT T1.������ڵ� AS ������ڵ�, T2.������ AS ������, " _
                  & "T1.����ó�ڵ� AS ����ó�ڵ�, T3.����ó�� AS ����ó��, " _
                  & "(T1.������� + '00') AS ����, '(��    ��)' AS ����, " _
                  & "0 AS ����, (T1.�����ޱݴ���ݾ�) AS �԰�ݾ�, " _
                  & "(T1.�����ޱ����޴���ݾ�) AS ���ޱݾ�, " _
                  & "'' AS �������, '' AS ��������, '' AS ������ȣ, '' AS �ð� " _
             & "FROM �����ޱݿ��帶�� T1 " _
             & "LEFT JOIN ����� T2 ON T2.������ڵ� = T1.������ڵ� " _
             & "LEFT JOIN ����ó T3 ON T3.������ڵ� = T1.������ڵ� AND T3.����ó�ڵ� = T1.����ó�ڵ� " _
             & "" & strWhere & " AND SUBSTRING(T1.�������,5,2) <> '00' AND (T1.�����ޱݴ���ݾ� <> 0 OR T1.�����ޱ����޴���ݾ� <> 0) " _
              & "AND T1.������� > '" & (Mid(DTOS(dtpF_Date.Value), 1, 4) + "00") & "' " _
              & "AND T1.������� < '" & Mid(DTOS(dtpF_Date.Value), 1, 6) & "' "
    If PB_regUserinfoU.UserMJGbn = "1" Then
       strSQL = strSQL & "UNION ALL " _
              & "SELECT T1.������ڵ� AS ������ڵ�, T2.������ AS ������, " _
                  & "T1.����ó�ڵ� AS ����ó�ڵ�, T3.����ó�� AS ����ó��, " _
                  & "(T1.���������) AS ����, (T1.����) AS ����, " _
                  & "T1.������� AS ����, (SUM(T1.�԰���� * T1.�԰�ܰ�) * " & (PB_curVatRate + 1) & ") AS �԰�ݾ�, " _
                  & "0 AS ���ޱݾ�, " _
                  & "������� = CASE WHEN T1.���ݱ��� = 1 THEN '����' ELSE '�ܻ�' END,  '' AS ��������, '' AS ������ȣ, " _
                  & "T1.�����ð� AS �ð� " _
             & "FROM �������⳻�� T1 " _
             & "LEFT JOIN ����� T2 ON T2.������ڵ� = T1.������ڵ� " _
             & "LEFT JOIN ����ó T3 ON T3.������ڵ� = T1.������ڵ� AND T3.����ó�ڵ� = T1.����ó�ڵ� " _
             & "" & strWhere & " AND T1.������� = 1 " _
              & "AND (T1.��뱸�� = 0) " _
              & "AND T1.��������� BETWEEN '" & (Mid(DTOS(dtpF_Date.Value), 1, 6) + "01") & "' AND '" & DTOS(dtpT_Date.Value) & "' " _
            & "GROUP BY T1.������ڵ�, T2.������, T1.����ó�ڵ�, T3.����ó��, " _
                     & "T1.���������, T1.�����ð�, T1.���ݱ���, " _
                     & "T1.������� "
    Else
       strSQL = strSQL & "UNION ALL " _
           & "SELECT T1.������ڵ� AS ������ڵ�, T2.������ AS ������, " _
                  & "T1.����ó�ڵ� AS ����ó�ڵ�, T3.����ó�� AS ����ó��, " _
                  & "(T1.�ۼ�����) AS ����, T1.���� AS ����, " _
                  & "1 AS ����, (SUM(T1.���ް��� + T1.����)) AS �԰�ݾ�, " _
                  & "0 AS ���ޱݾ�, " _
                  & "������� = CASE WHEN T1.�ݾױ��� = 0 THEN '����' WHEN T1.�ݾױ��� = 1 THEN '��ǥ' " _
                                  & "WHEN T1.�ݾױ��� = 2 THEN '����' ELSE '�ܻ�' END,  '' AS ��������, '' AS ������ȣ, " _
                  & "T1.�ۼ��ð� AS �ð� " _
             & "FROM ���Լ��ݰ�꼭��� T1 " _
             & "LEFT JOIN ����� T2 ON T2.������ڵ� = T1.������ڵ� " _
             & "LEFT JOIN ����ó T3 ON T3.������ڵ� = T1.������ڵ� AND T3.����ó�ڵ� = T1.����ó�ڵ� " _
             & "" & strWhere & " " _
              & "AND (T1.��뱸�� = 0) AND T1.�����ޱ��� = 1 " _
              & "AND T1.�ۼ����� BETWEEN '" & (Mid(DTOS(dtpF_Date.Value), 1, 6) + "01") & "' AND '" & DTOS(dtpT_Date.Value) & "' " _
            & "GROUP BY T1.������ڵ�, T2.������, T1.����ó�ڵ�, T3.����ó��, " _
                     & "T1.�ۼ�����, T1.�ۼ��ð�, T1.����, T1.�ݾױ��� "
    End If
    strSQL = strSQL & "UNION ALL " _
           & "SELECT T1.������ڵ� AS ������ڵ�, T2.������ AS ������, " _
                  & "T1.����ó�ڵ� AS ����ó�ڵ�, T3.����ó�� AS ����ó��, " _
                  & "(T1.�����ޱ���������) AS ����, '' AS ����, " _
                  & "0 AS ����, 0 AS �԰�ݾ�, " _
                  & "ISNULL(SUM(T1.�����ޱ����ޱݾ�), 0) As ���ޱݾ�, " _
                  & "������� = CASE WHEN T1.������� = 0 THEN '����' WHEN T1.������� = 1 THEN '��ǥ' " _
                                  & "WHEN T1.������� = 2 THEN '����' ELSE '��Ÿ' END, " _
                  & "T1.��������, T1.������ȣ, T1.�����ޱ����޽ð� AS �ð� " _
             & "FROM �����ޱݳ��� T1 " _
             & "LEFT JOIN ����� T2 ON T2.������ڵ� = T1.������ڵ� " _
             & "LEFT JOIN ����ó T3 ON T3.������ڵ� = T1.������ڵ� AND T3.����ó�ڵ� = T1.����ó�ڵ� " _
            & "" & strWhere & " AND T1.�����ޱ��������� " _
                         & "BETWEEN '" & (Mid(DTOS(dtpF_Date.Value), 1, 6) + "01") & "' AND '" & DTOS(dtpT_Date.Value) & "' " _
            & "GROUP BY T1.������ڵ�, T2.������, T1.����ó�ڵ�, T3.����ó��, " _
                     & "T1.�����ޱ���������, T1.�����ޱ����޽ð�, T1.����, T1.�������, T1.��������, T1.������ȣ "
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
                         .TextMatrix(lngR, 9) = curMonIMny   '������ݾ�(�԰�)
                         .TextMatrix(lngR, 10) = curMonOMny  '������ݾ�(����)
                         curMonIMny = 0: curMonOMny = 0: curMonTMny = 0
                         lngR = lngR + 1
                     End If
                     '�ش���(��)����
                     .TextMatrix(lngR, 4) = "(��    ��)"
                     .TextMatrix(lngR, 9) = curTotIMny       '����ó����ݾ�(�԰�)
                     .TextMatrix(lngR, 10) = curTotOMny      '����ó����ݾ�(����)
                     curTotIMny = 0: curTotOMny = 0: curTotTMny = 0
                     .AddItem ""
                     lngR = lngR + 1
                     .AddItem ""
                     lngR = lngR + 1
                     '����ó��
                     .TextMatrix(lngR, 3) = P_adoRec("����ó�ڵ�"): .TextMatrix(lngR, 4) = P_adoRec("����ó��")
                     .Cell(flexcpBackColor, lngR, 0, lngR, .Cols - 1) = vbYellow
                     .AddItem ""
                     lngR = lngR + 1
                  Else
                     If .TextMatrix(lngR - 1, 3) <> "" And _
                         Mid(StrDate, 1, 6) <> Mid(P_adoRec("����"), 1, 6) Then '�� ����� And ���� �ٸ���
                         '�ش����(��)����
                         .AddItem ""
                         '.TextMatrix(lngR, 4) = "(" & Format(Mid(StrDate, 1, 6), "0000-00") & Space(1) & "����)"
                         .TextMatrix(lngR, 4) = "(��    ��)"
                         .TextMatrix(lngR, 9) = curMonIMny   '������ݾ�(�԰�)
                         .TextMatrix(lngR, 10) = curMonOMny  '������ݾ�(����)
                         curMonIMny = 0: curMonOMny = 0: curMonTMny = 0
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
                     .TextMatrix(lngR, 4) = "(" & Mid(P_adoRec("����"), 1, 4) & " ���̿�)"
                  Else
                     .TextMatrix(lngR, 4) = "(" & Format(Mid(P_adoRec("����"), 1, 6), "0000-00") & " ����)"
                  End If
               End If
               '5. �����ڵ�
               .TextMatrix(lngR, 5) = P_adoRec("����")
               '6. ����
               If PB_regUserinfoU.UserMJGbn = "1" Then
                  If P_adoRec("����") = 0 Then
                     .TextMatrix(lngR, 6) = IIf(Mid(P_adoRec("����"), 7, 2) <> "00", "����", "") + IIf(P_adoRec("�������") = "", "", "(" + P_adoRec("�������") + ")")
                  ElseIf _
                     P_adoRec("����") = 1 Then
                     .TextMatrix(lngR, 6) = "����" + IIf(P_adoRec("�������") = "", "", "(" + P_adoRec("�������") + ")")
                  End If
               Else
                  If P_adoRec("����") = 0 Then
                     .TextMatrix(lngR, 6) = IIf(Mid(P_adoRec("����"), 7, 2) <> "00", "����", "") + IIf(P_adoRec("�������") = "", "", "(" + P_adoRec("�������") + ")")
                  ElseIf _
                     P_adoRec("����") = 2 Then
                     .TextMatrix(lngR, 6) = "����" + IIf(P_adoRec("�������") = "", "", "(" + P_adoRec("�������") + ")")
                  End If
               End If
               '7.��������, 8.������ȣ
               If Mid(P_adoRec("����"), 7, 2) = "00" Then
               Else
                  .TextMatrix(lngR, 7) = IIf(Len(P_adoRec("��������")) > 0, Format(P_adoRec("��������"), "0000-00-00"), "")
                  .TextMatrix(lngR, 8) = IIf(Len(P_adoRec("������ȣ")) > 0, P_adoRec("������ȣ"), "")
               End If
               '9. �԰�ݾ�
               .TextMatrix(lngR, 9) = P_adoRec("�԰�ݾ�")
               '10. ���޾�
               .TextMatrix(lngR, 10) = P_adoRec("���ޱݾ�")
               '11. �ܾ�
               .TextMatrix(lngR, 11) = curTotTMny + (P_adoRec("�԰�ݾ�") - P_adoRec("���ޱݾ�"))
               '12. ����
               If Mid(P_adoRec("����"), 7, 2) = "00" Then
               Else
                  .TextMatrix(lngR, 12) = P_adoRec("����")
               End If
               If Mid(P_adoRec("����"), 7, 2) <> 0 Then
                  curMonIMny = curMonIMny + P_adoRec("�԰�ݾ�")
                  curMonOMny = curMonOMny + P_adoRec("���ޱݾ�")
                  curMonTMny = curMonTMny + (P_adoRec("�԰�ݾ�") - P_adoRec("���ޱݾ�"))
               End If
               '�ش紩��ݾ�
               curTotIMny = curTotIMny + P_adoRec("�԰�ݾ�")
               curTotOMny = curTotOMny + P_adoRec("���ޱݾ�")
               curTotTMny = curTotTMny + (P_adoRec("�԰�ݾ�") - P_adoRec("���ޱݾ�"))
               '��ü����ݾ�
               curTotTIMny = curTotTIMny + P_adoRec("�԰�ݾ�")
               curTotTOMny = curTotTOMny + P_adoRec("���ޱݾ�")
               curTotTTMny = curTotTTMny + (P_adoRec("�԰�ݾ�") - P_adoRec("���ޱݾ�"))
               StrDate = P_adoRec("����")
               
               'FindRow ����� ����
               '.TextMatrix(lngR, 4) = .TextMatrix(lngR, 0) & P_adoRec("���з��ڵ�")
               '.Cell(flexcpData, lngR, 4, lngR, 4) = .TextMatrix(lngR, 4)
               P_adoRec.MoveNext
               If P_adoRec.EOF = True Then '������ ���ڵ��
                  If .TextMatrix(lngR, 3) <> "" Then     '�� �����
                     lngR = lngR + 1
                     '�ش������
                     .AddItem ""
                     .TextMatrix(lngR, 4) = "(��    ��)"
                     .TextMatrix(lngR, 9) = curMonIMny    '������ݾ�(�԰�)
                     .TextMatrix(lngR, 10) = curMonOMny   '������ݾ�(����)
                     curMonIMny = 0: curMonOMny = 0: curMonTMny = 0
                  End If
                  '�ش紩��
                  lngR = lngR + 1
                  .AddItem ""
                  .TextMatrix(lngR, 4) = "(��    ��)"
                  .TextMatrix(lngR, 9) = curTotIMny       '����ó����ݾ�(�԰�)
                  .TextMatrix(lngR, 10) = curTotOMny      '����ó����ݾ�(����)
                  curTotIMny = 0: curTotOMny = 0: curTotTMny = 0
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
               .TextMatrix(lngR, 9) = curTotTIMny        '��ü����ݾ�(�԰�)
               .TextMatrix(lngR, 10) = curTotTOMny       '��ü����ݾ�(����)
               .TextMatrix(lngR, 11) = curTotTTMny       '��ü����ݾ�(�ܾ�)
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
            If chkTotal.Value = 1 Then
               .TopRow = 1
            End If
            vsfg1_EnterCell       'vsfg1_EnterCell �ڵ�����(���� �Ѱ� �϶��� ������ �ڵ�����)
            .SetFocus
       End With
    End If
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�����ޱ� �б� ����"
    Unload Me
    Exit Sub
End Sub

Private Sub SubOther_FILL()
Dim strSQL        As String
Dim intIndex      As Integer
    cboSactionWay.AddItem "0. �� ��"
    cboSactionWay.AddItem "1. �� ǥ"
    cboSactionWay.AddItem "2. �� ��"
    cboSactionWay.ListIndex = 0
    dtpExpired_Date.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "TABLE �б� ����"
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
Dim curJanMny As Currency '���ް��ɱݾ�
    'If Not (chkTotal.Value = 0) Then '��ü�δ� ���� �Ұ�
    '   lngC = -1
    '   Exit Function
    'End If
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
               Case 2  '������ȣ
                    If Not (LenH(Text1(lngC).Text) <= 20) Then
                       Exit Function
                    End If
               Case 3  '���ޱݾ�
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
               Case 4  '����
                    If Not (LenH(Text1(lngC).Text) <= 50) Then
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
            strExeFile = App.Path & ".\�����ޱݿ���.rpt"
         Else
            strExeFile = App.Path & ".\�����ޱݿ���T.rpt"
         End If
         varRetVal = Dir(strExeFile)
         If Len(varRetVal) = 0 Then
            intRetCHK = 0
         Else
            .ReportFileName = strExeFile
            On Error GoTo ERROR_CRYSTAL_REPORTS
            '--- Grid Size = 0.101 ---
            '--- Formula Fields ---
            .Formulas(0) = "ForAppPgDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'" '���α׷���������
            .Formulas(1) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                       '������
            .Formulas(2) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                                  '����Ͻ�
            '�����������
            .Formulas(3) = "ForAppDate = '�������� : ' & '" & Format(DTOS(dtpF_Date.Value), "0000-00-00") & "' & ' ���� ' & '" & Format(DTOS(dtpT_Date.Value), "0000-00-00") & "' & ' ����' "
            'DECLARE @ParAppPgDate VarChar(8), @ParAppFDate VarChar(8),  @ParAppTDate VarChar(8), @ParSupplierCode VarChar(10)
            .Formulas(4) = "ForMJGbn = '" & PB_regUserinfoU.UserMJGbn & "'"                                '�����ޱݹ߻�����
            '���α׷���������
            '--- Parameter Fields ---
            .StoredProcParam(0) = PB_regUserinfoU.UserBranchCode
            .StoredProcParam(1) = PB_regUserinfoU.UserClientDate
            '����(����)���ڽ���
            .StoredProcParam(2) = DTOS(dtpF_Date.Value)
            '����(����)��������
            .StoredProcParam(3) = DTOS(dtpT_Date.Value)
            '����ó�ڵ�
            If chkTotal.Value = 0 Then
               If Len(Text1(0).Text) = 0 Then
                  .StoredProcParam(4) = " "
               Else
                  .StoredProcParam(4) = Trim(Text1(0).Text)
               End If
            Else
               .StoredProcParam(4) = " "
            End If
            .StoredProcParam(5) = CInt(PB_regUserinfoU.UserMJGbn)                 '�����ޱݹ߻�����(1.��ǥ, 2.��꼭)
            '.StoredProcParam(6) = IIf(optPrtChk0.Value = True, 0, 1)             '���ļ���(0.����ó�ڵ�, 1.����ó��)
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
            .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "�����ޱݿ���"
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

