VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm����ó���� 
   BorderStyle     =   0  '����
   Caption         =   "����ó����"
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
      TabIndex        =   37
      Top             =   0
      Width           =   15195
      Begin VB.OptionButton optPrtChk1 
         Caption         =   "�̸���"
         Height          =   255
         Left            =   6840
         TabIndex        =   54
         Top             =   390
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optPrtChk0 
         Caption         =   "�ڵ��"
         Height          =   255
         Left            =   6840
         TabIndex        =   53
         Top             =   150
         Width           =   975
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   375
         Left            =   7980
         Picture         =   "����ó����.frx":0000
         Style           =   1  '�׷���
         TabIndex        =   49
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   9120
         Picture         =   "����ó����.frx":0963
         Style           =   1  '�׷���
         TabIndex        =   44
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   13635
         Picture         =   "����ó����.frx":1308
         Style           =   1  '�׷���
         TabIndex        =   42
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   390
         Left            =   12510
         Picture         =   "����ó����.frx":1C56
         Style           =   1  '�׷���
         TabIndex        =   40
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   11385
         Picture         =   "����ó����.frx":25DA
         Style           =   1  '�׷���
         TabIndex        =   22
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Height          =   390
         Left            =   10260
         Picture         =   "����ó����.frx":2E61
         Style           =   1  '�׷���
         TabIndex        =   39
         Top             =   195
         Width           =   1095
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   4920
         Top             =   200
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Label2 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00008000&
         BorderStyle     =   1  '���� ����
         Caption         =   "�� �� �� �� ó �� ��"
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
         TabIndex        =   38
         Top             =   180
         Width           =   4650
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid vsfg1 
      Height          =   6917
      Left            =   60
      TabIndex        =   41
      Top             =   3198
      Width           =   15195
      _cx             =   26802
      _cy             =   12201
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
      Height          =   2549
      Left            =   60
      TabIndex        =   23
      Top             =   630
      Width           =   15195
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   14
         Left            =   14280
         MaxLength       =   1
         TabIndex        =   17
         Top             =   1320
         Width           =   430
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Index           =   18
         Left            =   1275
         TabIndex        =   21
         Top             =   2160
         Width           =   9345
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   13
         Left            =   12420
         TabIndex        =   16
         Top             =   1305
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Index           =   11
         Left            =   8940
         MaxLength       =   30
         TabIndex        =   12
         Top             =   1305
         Width           =   2175
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "����"
         Height          =   255
         Left            =   3480
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox cboBank 
         Height          =   300
         Left            =   12420
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   14
         Top             =   585
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   12
         Left            =   12420
         MaxLength       =   20
         TabIndex        =   15
         Top             =   945
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker dtpOpenDate 
         Height          =   270
         Left            =   12420
         TabIndex        =   13
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   57540609
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Index           =   4
         Left            =   5430
         MaxLength       =   20
         TabIndex        =   5
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Index           =   17
         Left            =   9660
         TabIndex        =   20
         Top             =   1785
         Width           =   5055
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   15
         Left            =   1275
         MaxLength       =   7
         TabIndex        =   18
         Top             =   1785
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Index           =   16
         Left            =   3435
         TabIndex        =   19
         Top             =   1785
         Width           =   5505
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   10
         Left            =   8940
         MaxLength       =   1
         TabIndex        =   11
         Top             =   945
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   9
         Left            =   8940
         MaxLength       =   14
         TabIndex        =   10
         Top             =   585
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   8
         Left            =   8940
         MaxLength       =   20
         TabIndex        =   9
         Top             =   233
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Index           =   7
         Left            =   5430
         TabIndex        =   8
         Top             =   1305
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Index           =   6
         Left            =   5430
         TabIndex        =   7
         Top             =   945
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   5
         Left            =   5430
         MaxLength       =   14
         TabIndex        =   6
         Top             =   585
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   3
         Left            =   1275
         MaxLength       =   14
         TabIndex        =   4
         Top             =   1305
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   2
         Left            =   1275
         MaxLength       =   14
         TabIndex        =   3
         Top             =   945
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Index           =   1
         Left            =   1275
         MaxLength       =   50
         TabIndex        =   2
         Top             =   585
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '����
         Index           =   0
         Left            =   1275
         MaxLength       =   8
         TabIndex        =   0
         Top             =   225
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�ܰ�����"
         Height          =   240
         Index           =   21
         Left            =   13320
         TabIndex        =   57
         ToolTipText     =   "1, 2, 3"
         Top             =   1365
         Width           =   855
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   15015
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   23
         Left            =   75
         TabIndex        =   56
         Top             =   2220
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         Caption         =   "[Home]"
         Height          =   240
         Index           =   22
         Left            =   2160
         TabIndex        =   55
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��꼭������"
         Height          =   240
         Index           =   20
         Left            =   11235
         TabIndex        =   52
         ToolTipText     =   "����(�Ƿ�)�� ������"
         Top             =   1365
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��꼭"
         Height          =   240
         Index           =   19
         Left            =   2750
         TabIndex        =   51
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "(0.����)"
         Height          =   240
         Index           =   18
         Left            =   9960
         TabIndex        =   50
         Top             =   1005
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "���¹�ȣ"
         Height          =   240
         Index           =   16
         Left            =   11235
         TabIndex        =   48
         Top             =   1005
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��������"
         Height          =   240
         Index           =   17
         Left            =   11235
         TabIndex        =   47
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "[Home]  �ּ�"
         Height          =   240
         Index           =   15
         Left            =   2280
         TabIndex        =   46
         Top             =   1845
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��������"
         Height          =   240
         Index           =   14
         Left            =   11235
         TabIndex        =   45
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��뱸��"
         Height          =   240
         Index           =   9
         Left            =   7740
         TabIndex        =   43
         ToolTipText     =   "0.����, ��Ÿ.���Ұ�"
         Top             =   1005
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   13
         Left            =   8460
         TabIndex        =   36
         Top             =   1845
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�ѽ���ȣ"
         Height          =   240
         Index           =   12
         Left            =   7740
         TabIndex        =   35
         ToolTipText     =   "14�ڸ� ����"
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����ڸ�"
         Height          =   240
         Index           =   11
         Left            =   7740
         TabIndex        =   34
         Top             =   1365
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�����ȣ"
         Height          =   240
         Index           =   10
         Left            =   75
         TabIndex        =   33
         Top             =   1845
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��ȭ��ȣ"
         Height          =   240
         Index           =   8
         Left            =   7740
         TabIndex        =   32
         ToolTipText     =   "20�ڸ� ����"
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   7
         Left            =   4230
         TabIndex        =   31
         Top             =   1365
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   6
         Left            =   4230
         TabIndex        =   30
         Top             =   1005
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�ֹι�ȣ"
         Height          =   240
         Index           =   5
         Left            =   4230
         TabIndex        =   29
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��ǥ�ڸ�"
         Height          =   240
         Index           =   4
         Left            =   4230
         TabIndex        =   28
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "���ι�ȣ"
         Height          =   240
         Index           =   3
         Left            =   75
         TabIndex        =   27
         Top             =   1365
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����ڹ�ȣ"
         Height          =   240
         Index           =   2
         Left            =   75
         TabIndex        =   26
         Top             =   1005
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����ó��"
         Height          =   240
         Index           =   1
         Left            =   75
         TabIndex        =   25
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����ó�ڵ�"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   75
         TabIndex        =   24
         Top             =   285
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frm����ó����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' ���α׷� �� �� : ����ó����
' ���� Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' ������ Table   : ����ó
' ��  ��  ��  �� :
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived       As Boolean
Private P_adoRec           As New ADODB.Recordset
Private P_intButton        As Integer
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
       'dtpOpenDate.Value = Format(Left(PB_regUserinfoU.UserClientDate, 4) + "0101", "0000-00-00")
       dtpOpenDate.Value = Format("19000101", "0000-00-00")
       Subvsfg1_INIT
       Select Case Val(PB_regUserinfoU.UserAuthority)
              Case Is <= 10 '��ȸ
                   cmdPrint.Enabled = False: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 20 '�μ�, ��ȸ
                   cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 40 '�߰�, ����, �μ�, ��ȸ
                   cmdPrint.Enabled = True: cmdClear.Enabled = True: cmdFind.Enabled = True
                   cmdSave.Enabled = True: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 50 '����, �߰�, ����, ��ȸ, �μ�
                   cmdPrint.Enabled = True: cmdClear.Enabled = True: cmdFind.Enabled = True
                   cmdSave.Enabled = True: cmdDelete.Enabled = True: cmdExit.Enabled = True
              Case Is <= 99 '����, �߰�, ����, ��ȸ, �μ�
                   cmdPrint.Enabled = True: cmdClear.Enabled = True: cmdFind.Enabled = True
                   cmdSave.Enabled = True: cmdDelete.Enabled = True: cmdExit.Enabled = True
              Case Else
                   cmdClear.Enabled = False: cmdFind.Enabled = False: cmdSave.Enabled = False
                   cmdDelete.Enabled = False: cmdExit.Enabled = True
       End Select
       SubOther_FILL
       '���� �����ϸ� �߰����� �ٷν���
       If cmdSave.Enabled = True Then
          cmdClear_Click
       End If
       P_blnActived = True
       Screen.MousePointer = vbDefault
    End If
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����ó ����(�������� ���� ����)"
    Unload Me
    Exit Sub
End Sub

'+---------------------+
'/// ��꼭���࿩�� ///
'+---------------------+
Private Sub chkPrint_KeyDown(KeyCode As Integer, Shift As Integer)
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
Dim strExeFile As String
Dim varRetVal  As Variant
    If (Index = 0 And (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn)) Then  '����ó�˻�
       PB_strSupplierCode = UPPER(Trim(Text1(Index).Text))
       PB_strSupplierName = "" 'Trim(Text1(Index + 1).Text)
       frm����ó�˻�.Show vbModal
       If (Len(PB_strSupplierCode) + Len(PB_strSupplierName)) = 0 Then '�˻����� ���(ESC)
       Else
          Text1(Index).Text = PB_strSupplierCode
          Text1(Index + 1).Text = PB_strSupplierName
          Subvsfg1_FILL
          Exit Sub
       End If
       If PB_strSupplierCode <> "" Then
          SendKeys "{TAB}"
       End If
       PB_strSupplierCode = "": PB_strSupplierName = ""
    ElseIf _
       (Index = 15 And (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn)) Then
       If Len(Trim(Text1(Index).Text)) = 6 Then
          Text1(Index).Text = Format(Trim(Text1(Index).Text), "###-###")
       End If
       PB_strPostCode = Trim(Text1(Index).Text)
       PB_strPostName = Trim(Text1(Index + 1).Text)
       frm�����ȣ�˻�.Show vbModal
       If (Len(PB_strPostCode) + Len(PB_strPostName)) = 0 Then '�˻����� ���(ESC)
       Else
          Text1(Index).Text = PB_strPostCode
          Text1(Index + 1).Text = PB_strPostName
       End If
       If PB_strPostCode <> "" Then
          Text1(Index + 2).SetFocus
       Else
          Text1(Index + 1).SetFocus
       End If
    ElseIf _
       KeyCode = vbKeyReturn Then
       Select Case Index
       End Select
       SendKeys "{tab}"
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim strSQL As String
    With Text1(Index)
         Select Case Index
                Case 0
                     .Text = UPPER(Trim(.Text))
                     If Len(Trim(Text1(Index).Text)) < 1 Then
                        Text1(Index).Text = ""
                     End If
                     If Text1(Index).Enabled = True Then
                        'P_adoRec.CursorLocation = adUseClient
                        'strSQL = "SELECT * FROM ����ó " _
                        '        & "WHERE ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                        '          & "AND ����ó�ڵ� = '" & Trim(.Text) & "' "
                        'On Error GoTo ERROR_TABLE_SELECT
                        'P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
                        'If P_adoRec.RecordCount <> 0 Then
                        '   P_adoRec.Close
                        '   .Text = ""
                        '   .SetFocus
                        '   Exit Sub
                        'End If
                        'P_adoRec.Close
                     End If
                Case 2 '����ڵ�Ϲ�ȣ
                     If Len(Trim(.Text)) = 10 Then
                        .Text = Format(Trim(.Text), "###-##-#####")
                     End If
                Case 5 '�ֹε�Ϲ�ȣ
                     If Len(Trim(.Text)) = 13 Then
                        .Text = Format(Trim(.Text), "######-#######")
                     End If
                Case 11, 12
                     .Text = Trim(.Text)
                Case 13 '��꼭������
                     .Text = Format(Vals(Trim(.Text)), "#0.00")
                Case 14 '�ܰ�����
                     .Text = Format(Vals(Trim(.Text)), "0")
                Case 15 '�����ȣ
                     If Len(Trim(.Text)) = 6 Then
                        .Text = Format(Trim(.Text), "###-###")
                     End If
         End Select
    End With
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����ó���� �б� ����"
    Unload Me
    Exit Sub
End Sub

Private Sub dtpOpenDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub cboBank_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
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
            .Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlack '.ForeColorSel
            .Cell(flexcpForeColor, 0, .MouseCol, 0, .MouseCol) = vbRed
            strData = Trim(.Cell(flexcpData, .Row, 0))
            .Sort = flexSortGenericAscending
            If P_intButton = 1 Then
               .Sort = flexSortGenericAscending
            Else
               .Sort = flexSortGenericDescending
            End If
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
         If .Row < .FixedRows Then
            Text1(Text1.LBound).Enabled = True
         Else
            Text1(Text1.LBound).Enabled = False
            For lngC = 0 To .Cols - 1
                Select Case lngC
                       Case Is <= 10
                            Text1(lngC).Text = .TextMatrix(.Row, lngC)
                       Case 12
                            If Len(.TextMatrix(.Row, lngC)) = 0 Then
                               dtpOpenDate.Value = Format("19000101", "0000-00-00")
                            Else
                               dtpOpenDate.Value = Format(DTOS(.TextMatrix(.Row, lngC)), "0000-00-00")
                            End If
                       Case 13 To 15
                            Text1(lngC + 2).Text = .TextMatrix(.Row, lngC)
                       Case 17 '�����ڵ�
                            If Len(.TextMatrix(.Row, lngC)) = 0 Then
                               cboBank.ListIndex = 0
                            Else
                               For lngR = 0 To cboBank.ListCount - 1
                                   If .TextMatrix(.Row, lngC) = Left(cboBank.List(lngR), 2) Then
                                      cboBank.ListIndex = lngR
                                      Exit For
                                   End If
                               Next lngR
                            End If
                       Case 19 '���¹�ȣ
                            Text1(12).Text = .TextMatrix(.Row, lngC)
                       Case 21 '��꼭���࿩��(0, 1)
                            chkPrint.Value = .ValueMatrix(.Row, lngC)
                       Case 22 '��꼭������
                            Text1(13).Text = Format(.ValueMatrix(.Row, lngC), "#0.00")
                       Case 23 '�ܰ�����
                            Text1(14).Text = .ValueMatrix(.Row, lngC)
                       Case 24 '����ڸ�
                            Text1(11).Text = .TextMatrix(.Row, lngC)
                       Case 25 '����
                            Text1(18).Text = .TextMatrix(.Row, lngC)
                       Case Else
                End Select
            Next lngC
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
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����ó �б� ����"
    Unload Me
    Exit Sub

End Sub

'+-----------+
'/// �߰� ///
'+-----------+
Private Sub cmdClear_Click()
    SubClearText
    vsfg1.Row = 0
    Text1(Text1.LBound).Enabled = True
    Text1(Text1.LBound).SetFocus
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
Dim strSQL    As String
Dim lngR      As Long
Dim lngC      As Long
Dim blnOK     As Boolean
Dim intRetVal As Integer
    '�Է³��� �˻�
    FncCheckTextBox lngC, blnOK
    If blnOK = False Then
       If Text1(lngC).Enabled = False Then
          Text1(lngC).Enabled = True
       End If
       Text1(lngC).SetFocus
       Exit Sub
    End If
    If Text1(Text1.LBound).Enabled = True Then
       intRetVal = MsgBox("�Էµ� �ڷḦ �߰��Ͻðڽ��ϱ� ?", vbQuestion + vbYesNo + vbDefaultButton1, "�ڷ� �߰�")
    Else
       intRetVal = MsgBox("������ �ڷḦ �����Ͻðڽ��ϱ� ?", vbQuestion + vbYesNo + vbDefaultButton1, "�ڷ� ����")
    End If
    If intRetVal = vbNo Then
       vsfg1.SetFocus
       Exit Sub
    End If
    cmdSave.Enabled = False
    With vsfg1
         Screen.MousePointer = vbHourglass
         P_adoRec.CursorLocation = adUseClient
         If Text1(Text1.LBound).Enabled = True Then '����ó �߰��� �˻�
            strSQL = "SELECT * FROM ����ó " _
                    & "WHERE ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                      & "AND ����ó�ڵ� = '" & Trim(Text1(Text1.LBound).Text) & "' "
            On Error GoTo ERROR_TABLE_SELECT
            P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
            If P_adoRec.RecordCount <> 0 Then
               P_adoRec.Close
               Text1(Text1.LBound).SetFocus
               Screen.MousePointer = vbDefault
               MsgBox "�̹� ��ϵ� ����ó�Դϴ�. ����ó�ڵ带 �ٽ� Ȯ���Ͽ��ּ���!", vbCritical + vbOKOnly, "����ó ���"
               cmdSave.Enabled = True
               Exit Sub
            End If
            P_adoRec.Close
         End If
         PB_adoCnnSQL.BeginTrans
         If Text1(Text1.LBound).Enabled = True Then '����ó �߰�
            strSQL = "INSERT INTO ����ó(������ڵ�, ����ó�ڵ�, ����ó��, ����ڹ�ȣ," _
                                      & "���ι�ȣ, ��ǥ�ڸ�, ��ǥ���ֹι�ȣ," _
                                      & "��������," _
                                      & "�����ȣ, �ּ�, ����," _
                                      & "����, ����, ��ȭ��ȣ," _
                                      & "�ѽ���ȣ, " _
                                      & "�����ڵ�, ���¹�ȣ, " _
                                      & "��꼭���࿩��, ��꼭������, ����ڸ�, " _
                                      & "��뱸��, " _
                                      & "�������� , ������ڵ�, " _
                                      & "����, �ܰ�����) VALUES( " _
                                      & "'" & PB_regUserinfoU.UserBranchCode & "', '" & Trim(Text1(0).Text) & "', " _
                    & "'" & Trim(Text1(1).Text) & "','" & Trim(Text1(2).Text) & "', " _
                    & "'" & Trim(Text1(3).Text) & "','" & Trim(Text1(4).Text) & "','" & Trim(Text1(5).Text) & "', " _
                    & "'" & IIf(DTOS(dtpOpenDate.Value) = "19000101", "", DTOS(dtpOpenDate.Value)) & "', " _
                    & "'" & Trim(Text1(15).Text) & "','" & Trim(Text1(16).Text) & "','" & Trim(Text1(17).Text) & "', " _
                    & "'" & Trim(Text1(6).Text) & "','" & Trim(Text1(7).Text) & "','" & Trim(Text1(8).Text) & "', " _
                    & "'" & Trim(Text1(9).Text) & "', " _
                    & "'" & IIf(Left(cboBank.Text, 2) = "00", "", Left(cboBank.Text, 2)) & "', '" & Trim(Text1(12).Text) & "', " _
                    & "" & chkPrint.Value & ", " & Vals(Text1(13).Text) & ", '" & Trim(Text1(11).Text) & "', " _
                    & "" & Val(Text1(10).Text) & ", " _
                    & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "', " _
                    & "'" & Trim(Text1(18).Text) & "', " & Vals(Text1(14).Text) & " )"
            On Error GoTo ERROR_TABLE_INSERT
            PB_adoCnnSQL.Execute strSQL
            .AddItem .Rows
            For lngC = Text1.LBound To Text1.UBound
                Select Case lngC
                       Case Is <= 9
                            .TextMatrix(.Rows - 1, lngC) = Text1(lngC).Text
                            If lngC = 0 Then .Cell(flexcpData, .Rows - 1, lngC, .Rows - 1, lngC) = Text1(lngC).Text
                       Case 10
                            .TextMatrix(.Rows - 1, lngC) = Val(Trim(Text1(lngC).Text))
                            Select Case Val(Text1(lngC).Text)
                                   Case 0
                                        .TextMatrix(.Rows - 1, lngC + 1) = "����"
                                   Case 9
                                        .TextMatrix(.Rows - 1, lngC + 1) = "���Ұ�"
                                   Case Else
                                        .TextMatrix(.Rows - 1, lngC + 1) = "���п���"
                            End Select
                       Case 11
                            .TextMatrix(.Rows - 1, 24) = Trim(Text1(lngC).Text)
                       Case 12
                            .TextMatrix(.Rows - 1, 19) = Trim(Text1(lngC).Text)
                       Case 13
                            .TextMatrix(.Rows - 1, 22) = Vals(Text1(lngC).Text)
                       Case 14
                            .TextMatrix(.Rows - 1, 23) = Vals(Text1(lngC).Text)
                       Case 15 To 17
                            .TextMatrix(.Rows - 1, lngC - 2) = Text1(lngC).Text
                       Case 18
                            .TextMatrix(.Rows - 1, 25) = Text1(lngC).Text
                       Case Else
                End Select
            Next lngC
            .TextMatrix(.Rows - 1, 12) = Format(DTOS(dtpOpenDate.Value), "0000-00-00")
            .TextMatrix(.Rows - 1, 16) = Trim(.TextMatrix(.Rows - 1, 14)) & Space(1) & Trim(.TextMatrix(.Rows - 1, 15))
            .TextMatrix(.Rows - 1, 17) = IIf(Left(cboBank.Text, 2) = "00", "", Left(cboBank.Text, 2))
            .TextMatrix(.Rows - 1, 18) = IIf(Left(cboBank.Text, 2) = "00", "", Mid(cboBank.Text, 5))
            .TextMatrix(.Rows - 1, 20) = chkPrint.Value
            Select Case chkPrint.Value
                   Case 0
                        .TextMatrix(.Rows - 1, 21) = "0. �̹���"
                   Case 1
                        .TextMatrix(.Rows - 1, 21) = "1. ��  ��"
                   Case Else
                        .TextMatrix(.Rows - 1, 21) = CStr(chkPrint.Value) + ". �� ��"
            End Select
            If .Rows > PC_intRowCnt Then
               .ScrollBars = flexScrollBarBoth
               .TopRow = .Rows - 1
            End If
            Text1(Text1.LBound).Enabled = False
            .Row = .Rows - 1          '�ڵ����� vsfg1_EnterCell Event �߻�
         Else                                          '����ó���� ����
            strSQL = "UPDATE ����ó SET " _
                          & "����ó�� = '" & Trim(Text1(1).Text) & "', " _
                          & "����ڹ�ȣ = '" & Trim(Text1(2).Text) & "', " _
                          & "���ι�ȣ = '" & Trim(Text1(3).Text) & "', " _
                          & "��ǥ�ڸ� = '" & Trim(Text1(4).Text) & "', " _
                          & "��ǥ���ֹι�ȣ = '" & Trim(Text1(5).Text) & "', " _
                          & "�������� = '" & IIf(DTOS(dtpOpenDate.Value) = "19000101", "", DTOS(dtpOpenDate.Value)) & "', " _
                          & "�����ȣ = '" & Trim(Text1(15).Text) & "', " _
                          & "�ּ� = '" & Trim(Text1(16).Text) & "', " _
                          & "���� = '" & Trim(Text1(17).Text) & "', " _
                          & "���� = '" & Trim(Text1(6).Text) & "', " _
                          & "���� = '" & Trim(Text1(7).Text) & "', " _
                          & "��ȭ��ȣ = '" & Trim(Text1(8).Text) & "', " _
                          & "�ѽ���ȣ = '" & Trim(Text1(9).Text) & "', " _
                          & "�����ڵ� = '" & IIf(Left(cboBank.Text, 2) = "00", "", Left(cboBank.Text, 2)) & "', " _
                          & "���¹�ȣ = '" & Trim(Text1(12).Text) & "', " _
                          & "��꼭���࿩�� = " & chkPrint.Value & ", " _
                          & "��꼭������ = " & Vals(Text1(13).Text) & ", " _
                          & "����ڸ� = '" & Trim(Text1(11).Text) & "', " _
                          & "��뱸�� = " & Val(Text1(10).Text) & ", " _
                          & "�������� = '" & PB_regUserinfoU.UserServerDate & "', " _
                          & "������ڵ� = '" & PB_regUserinfoU.UserCode & "', " _
                          & "���� = '" & Trim(Text1(18).Text) & "', �ܰ����� = " & Vals(Text1(14).Text) & " " _
                    & "WHERE ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                      & "AND ����ó�ڵ� = '" & Trim(Text1(Text1.LBound).Text) & "' "
            On Error GoTo ERROR_TABLE_UPDATE
            PB_adoCnnSQL.Execute strSQL
            For lngC = Text1.LBound To Text1.UBound
                Select Case lngC
                       Case Is <= 9
                            .TextMatrix(.Row, lngC) = Text1(lngC).Text
                       Case 10
                            .TextMatrix(.Row, lngC) = Val(Trim(Text1(lngC).Text))
                            Select Case Val(Text1(lngC).Text)
                                   Case 0
                                        .TextMatrix(.Row, lngC + 1) = "����"
                                   Case 9
                                        .TextMatrix(.Row, lngC + 1) = "���Ұ�"
                                   Case Else
                                        .TextMatrix(.Row, lngC + 1) = "���п���"
                            End Select
                       Case 11
                            .TextMatrix(.Row, 24) = Trim(Text1(lngC).Text)
                       Case 12
                            .TextMatrix(.Row, 19) = Trim(Text1(lngC).Text)
                       Case 13
                            .TextMatrix(.Row, 22) = Vals(Text1(lngC).Text)
                       Case 14
                            .TextMatrix(.Row, 23) = Vals(Text1(lngC).Text)
                       Case 15 To 17
                            .TextMatrix(.Row, lngC - 2) = Text1(lngC).Text
                       Case 18
                            .TextMatrix(.Row, 25) = Text1(lngC).Text
                       Case Else
                End Select
            Next lngC
            .TextMatrix(.Row, 12) = Format(DTOS(dtpOpenDate.Value), "0000-00-00")
            .TextMatrix(.Row, 16) = Trim(.TextMatrix(.Row, 14)) & Space(1) & Trim(.TextMatrix(.Row, 15))
            .TextMatrix(.Row, 17) = IIf(Left(cboBank.Text, 2) = "00", "", Left(cboBank.Text, 2))
            .TextMatrix(.Row, 18) = IIf(Left(cboBank.Text, 2) = "00", "", Mid(cboBank.Text, 5))
            .TextMatrix(.Row, 20) = chkPrint.Value
            Select Case chkPrint.Value
                   Case 0
                        .TextMatrix(.Row, 21) = "0. �̹���"
                   Case 1
                        .TextMatrix(.Row, 21) = "1. ��  ��"
                   Case Else
                        .TextMatrix(.Row, 21) = .TextMatrix(.Row, 21) + ". �� ��"
            End Select
         End If
         PB_adoCnnSQL.CommitTrans
         .SetFocus
         Screen.MousePointer = vbDefault
    End With
    cmdSave.Enabled = True
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����ó �б� ����"
    Unload Me
    Exit Sub
ERROR_TABLE_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����ó �߰� ����"
    Unload Me
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����ó ���� ����"
    Unload Me
    Exit Sub
End Sub

'+-----------+
'/// ���� ///
'+-----------+
Private Sub cmdDelete_Click()
Dim strSQL       As String
Dim intRetVal    As Integer
Dim lngCnt       As Long
    With vsfg1
         If .Row >= .FixedRows Then
            intRetVal = MsgBox("��ϵ� �ڷḦ �����Ͻðڽ��ϱ� ?", vbCritical + vbYesNo + vbDefaultButton2, "�ڷ� ����")
            If intRetVal = vbYes Then
               Screen.MousePointer = vbHourglass
               cmdDelete.Enabled = False
               '������ �������̺� �˻�
               'P_adoRec.CursorLocation = adUseClient
               'strSQL = "SELECT Count(*) AS �ش�Ǽ� FROM TableName " _
               '        & "WHERE ����ó���� = " & .TextMatrix(.Row, 0) & " "
               'On Error GoTo ERROR_TABLE_SELECT
               'P_adoRec.Open SQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
               'lngCnt = P_adoRec("�ش�Ǽ�")
               'P_adoRec.Close
               If lngCnt <> 0 Then
                  MsgBox "XXX(" & Format(lngCnt, "#,#") & "��)�� �����Ƿ� ������ �� �����ϴ�.", vbCritical, "����ó ���� �Ұ�"
                  cmdDelete.Enabled = True
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
               PB_adoCnnSQL.BeginTrans
               P_adoRec.CursorLocation = adUseClient
               strSQL = "DELETE FROM ����ó " _
                       & "WHERE ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                         & "AND ����ó�ڵ� = '" & .TextMatrix(.Row, 0) & "' "
               On Error GoTo ERROR_TABLE_DELETE
               PB_adoCnnSQL.Execute strSQL
               PB_adoCnnSQL.CommitTrans
               .RemoveItem .Row
               Text1(Text1.LBound).Enabled = False
               If .Rows <= PC_intRowCnt Then
                  .ScrollBars = flexScrollBarHorizontal
               End If
               cmdDelete.Enabled = True
               Screen.MousePointer = vbDefault
               If .Rows = 1 Then
                  SubClearText
                  .Row = 0
                  Text1(Text1.LBound).Enabled = True
                  Text1(Text1.LBound).SetFocus
                  Exit Sub
               End If
               vsfg1_EnterCell
            End If
            .SetFocus
         End If
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_DELETE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����ó ���� ����"
    cmdDelete.Enabled = True
    vsfg1.SetFocus
    'Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
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
    Set frm����ó���� = Nothing
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
    Text1(Text1.LBound).Enabled = False                '����ó�ڵ� FLASE
    With vsfg1              'Rows 1, Cols 26, RowHeightMax(Min) 300
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
         .FixedCols = 2
         .Rows = 1             'Subvsfg1_Fill����ÿ� ����
         .Cols = 26
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 850    '����ó�ڵ�
         .ColWidth(1) = 2550   '����ó��
         .ColWidth(2) = 1400   '����ڹ�ȣ
         .ColWidth(3) = 1500   '���ι�ȣ
         .ColWidth(4) = 1300   '��ǥ�ڸ�
         .ColWidth(5) = 1400   '��ǥ���ֹι�ȣ
         .ColWidth(6) = 2500   '����
         .ColWidth(7) = 2500   '����
         .ColWidth(8) = 2000   '��ȭ��ȣ
         .ColWidth(9) = 1600   '�ѽ���ȣ
         .ColWidth(10) = 1     '��뱸��
         .ColWidth(11) = 1000  '��뱸��
         
         .ColWidth(12) = 1000  '��������
         .ColWidth(13) = 1000  '�����ȣ
         .ColWidth(14) = 1     '����ó�ּ�
         .ColWidth(15) = 1     '����ó����
         .ColWidth(16) = 7000  '����ó�ּ�(�ּ�+����)
         .ColWidth(17) = 1000  '�����ڵ�
         .ColWidth(18) = 1400  '�����
         .ColWidth(19) = 1700  '���¹�ȣ
         .ColWidth(20) = 1200  '��꼭���࿩��(0, 1)
         .ColWidth(21) = 900   '��꼭(0.�̹���, 1. �� ��,?. ��  ��)
         .ColWidth(22) = 900   '��꼭������
         .ColWidth(23) = 1000  '�ܰ�����
         .ColWidth(24) = 2410  '����ڸ�
         .ColWidth(25) = 9000  '����
         
         .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "�ڵ�"
         .TextMatrix(0, 1) = "����ó��"
         .TextMatrix(0, 2) = "����ڹ�ȣ"
         .TextMatrix(0, 3) = "���ι�ȣ"
         .TextMatrix(0, 4) = "��ǥ�ڸ�"
         .TextMatrix(0, 5) = "�ֹι�ȣ"
         .TextMatrix(0, 6) = "����"
         .TextMatrix(0, 7) = "����"
         .TextMatrix(0, 8) = "��ȭ��ȣ"
         .TextMatrix(0, 9) = "�ѽ���ȣ"
         .TextMatrix(0, 10) = "��뱸��"
         .TextMatrix(0, 11) = "��뱸��"
         .TextMatrix(0, 12) = "��������"
         .TextMatrix(0, 13) = "�����ȣ"
         .TextMatrix(0, 14) = "����ó�ּ�"
         .TextMatrix(0, 15) = "����ó����"
         .TextMatrix(0, 16) = "����ó�ּ�" '�ּ�+����
         .TextMatrix(0, 17) = "�����ڵ�"
         .TextMatrix(0, 18) = "�����"
         .TextMatrix(0, 19) = "���¹�ȣ"
         .TextMatrix(0, 20) = "��꼭���࿩��"
         .TextMatrix(0, 21) = "��꼭"
         .TextMatrix(0, 22) = "������"
         .TextMatrix(0, 23) = "�ܰ�����"
         .TextMatrix(0, 24) = "����ڸ�"
         .TextMatrix(0, 25) = "����"
         
         .ColHidden(10) = True: .ColHidden(14) = True: .ColHidden(15) = True
         .ColHidden(17) = True: .ColHidden(20) = True
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 1, 6 To 9, 14 To 16, 18, 19, 24, 25
                         .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 0, 2 To 5, 10 To 13, 17, 21, 22, 23
                         .ColAlignment(lngC) = flexAlignCenterCenter
                    Case Else
                         .ColAlignment(lngC) = flexAlignLeftCenter
             End Select
         Next lngC
         .ColFormat(22) = "#0.00"
    End With
End Sub

Private Sub SubOther_FILL()
Dim strSQL        As String
Dim intIndex      As Integer
    Text1(Text1.LBound).Enabled = False
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.�����ڵ� AS �����ڵ�, " _
                  & "T1.����� AS ����� " _
             & "FROM ���� T1 " _
            & "ORDER BY T1.�����ڵ� "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       cboBank.AddItem "00. ����"
       cboBank.ListIndex = 0
       cboBank.Enabled = False
       cmdPrint.Enabled = False: cmdClear.Enabled = False: cmdFind.Enabled = False
       cmdSave.Enabled = False: cmdDelete.Enabled = False
       Exit Sub
    Else
       cboBank.AddItem "00. ����"
       Do Until P_adoRec.EOF
          cboBank.AddItem P_adoRec("�����ڵ�") & ". " & P_adoRec("�����")
          P_adoRec.MoveNext
       Loop
       P_adoRec.Close
       cboBank.ListIndex = 0
    End If
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "���� �б� ����"
    Unload Me
    Exit Sub
End Sub

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
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.������ڵ� AS ������ڵ�, T1.����ó�ڵ� AS ����ó�ڵ�, " _
                  & "T1.����ó�� AS ����ó��, T1.����ڹ�ȣ AS ����ڹ�ȣ, " _
                  & "T1.���ι�ȣ AS ���ι�ȣ, T1.��ǥ�ڸ� AS ��ǥ�ڸ�, " _
                  & "T1.��ǥ���ֹι�ȣ AS ��ǥ���ֹι�ȣ, T1.���� AS ����, " _
                  & "T1.���� AS ����, T1.��ȭ��ȣ AS ��ȭ��ȣ, " _
                  & "T1.�ѽ���ȣ AS �ѽ���ȣ, T1.��뱸�� AS ��뱸��, " _
                  & "T1.�������� AS ��������, T1.�����ȣ AS �����ȣ, " _
                  & "T1.�ּ� AS �ּ�, T1.���� AS ����, " _
                  & "ISNULL(T1.�����ڵ�,'') AS �����ڵ�, ISNULL(T3.�����,'') AS �����, " _
                  & "ISNULL(T1.���¹�ȣ,'') AS ���¹�ȣ, T1.��꼭���࿩�� AS ��꼭���࿩��, T1.��꼭������ AS ��꼭������, " _
                  & "ISNULL(T1.����ڸ�,'') AS ����ڸ�, ISNULL(T1.����, '') AS ����, T1.�ܰ����� AS �ܰ����� " _
          & "FROM ����ó T1 " _
          & "LEFT JOIN ���� T3 " _
                 & "ON T3.�����ڵ� = T1.�����ڵ� " _
         & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
           & "AND T1.����ó�ڵ� LIKE '%" & Trim(Text1(0).Text) & "%' " _
         & "ORDER BY " & IIf(optPrtChk0.Value = True, "T1.����ó�ڵ�", "T1.����ó��") & " "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    vsfg1.Rows = P_adoRec.RecordCount + 1
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       Text1(Text1.LBound).Enabled = True
       Text1(Text1.LBound).SetFocus
       Exit Sub
    Else
       With vsfg1
            .Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlack
            If .Rows <= PC_intRowCnt Then
               .ScrollBars = flexScrollBarHorizontal
            Else
               .ScrollBars = flexScrollBarBoth
            End If
            Do Until P_adoRec.EOF
               lngR = lngR + 1
               .TextMatrix(lngR, 0) = P_adoRec("����ó�ڵ�")
               .Cell(flexcpData, lngR, 0, lngR, 0) = Trim(.TextMatrix(lngR, 0)) 'FindRow ����� ����
               .TextMatrix(lngR, 1) = IIf(IsNull(P_adoRec("����ó��")), "", P_adoRec("����ó��"))
               .TextMatrix(lngR, 2) = IIf(IsNull(P_adoRec("����ڹ�ȣ")), "", P_adoRec("����ڹ�ȣ"))
               .TextMatrix(lngR, 3) = IIf(IsNull(P_adoRec("���ι�ȣ")), "", P_adoRec("���ι�ȣ"))
               .TextMatrix(lngR, 4) = IIf(IsNull(P_adoRec("��ǥ�ڸ�")), "", P_adoRec("��ǥ�ڸ�"))
               .TextMatrix(lngR, 5) = IIf(IsNull(P_adoRec("��ǥ���ֹι�ȣ")), "", P_adoRec("��ǥ���ֹι�ȣ"))
               .TextMatrix(lngR, 6) = IIf(IsNull(P_adoRec("����")), "", P_adoRec("����"))
               .TextMatrix(lngR, 7) = IIf(IsNull(P_adoRec("����")), "", P_adoRec("����"))
               .TextMatrix(lngR, 8) = IIf(IsNull(P_adoRec("��ȭ��ȣ")), "", P_adoRec("��ȭ��ȣ"))
               .TextMatrix(lngR, 9) = IIf(IsNull(P_adoRec("�ѽ���ȣ")), "", P_adoRec("�ѽ���ȣ"))
               .TextMatrix(lngR, 10) = IIf(IsNull(P_adoRec("��뱸��")), "", P_adoRec("��뱸��"))
               Select Case .ValueMatrix(lngR, 10)
                      Case 0
                           .TextMatrix(lngR, 11) = "����"
                      Case 9
                           .TextMatrix(lngR, 11) = "���Ұ�"
                      Case Else
                           .TextMatrix(lngR, 11) = "���п���"
               End Select
               .TextMatrix(lngR, 12) = IIf(IsNull(P_adoRec("��������")), "", Format(P_adoRec("��������"), "0000-00-00"))
               .TextMatrix(lngR, 13) = IIf(IsNull(P_adoRec("�����ȣ")), "", P_adoRec("�����ȣ"))
               .TextMatrix(lngR, 14) = IIf(IsNull(P_adoRec("�ּ�")), "", P_adoRec("�ּ�"))
               .TextMatrix(lngR, 15) = IIf(IsNull(P_adoRec("����")), "", P_adoRec("����"))
               .TextMatrix(lngR, 16) = Trim(.TextMatrix(lngR, 14)) & Space(1) & Trim(.TextMatrix(lngR, 13))
               .TextMatrix(lngR, 17) = IIf(IsNull(P_adoRec("�����ڵ�")), "", P_adoRec("�����ڵ�"))
               .TextMatrix(lngR, 18) = IIf(IsNull(P_adoRec("�����")), "", P_adoRec("�����"))
               .TextMatrix(lngR, 19) = IIf(IsNull(P_adoRec("���¹�ȣ")), "", P_adoRec("���¹�ȣ"))
               .TextMatrix(lngR, 20) = IIf(IsNull(P_adoRec("��꼭���࿩��")), "", P_adoRec("��꼭���࿩��"))
               Select Case .ValueMatrix(lngR, 20)
                      Case 0
                           .TextMatrix(lngR, 21) = "0. �̹���"
                      Case 1
                           .TextMatrix(lngR, 21) = "1. ��  ��"
                      Case Else
                           .TextMatrix(lngR, 21) = .TextMatrix(lngR, 21) + ". �� ��"
               End Select
               .TextMatrix(lngR, 22) = P_adoRec("��꼭������")
               .TextMatrix(lngR, 23) = P_adoRec("�ܰ�����")
               .TextMatrix(lngR, 24) = IIf(IsNull(P_adoRec("����ڸ�")), "", P_adoRec("����ڸ�"))
               .TextMatrix(lngR, 25) = IIf(IsNull(P_adoRec("����")), "", P_adoRec("����"))
               If .TextMatrix(lngR, 0) = Trim(Text1(0).Text) Then
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
               Text1(Text1.LBound).Enabled = True
               Text1(Text1.LBound).SetFocus
               Exit Sub
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
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����ó���� �б� ����"
    Unload Me
    Exit Sub
End Sub

'+-------------------------+
'/// Clear text1(index) ///
'+-------------------------+
Private Sub SubClearText()
Dim lngC As Long
    For lngC = Text1.LBound To Text1.UBound
        If lngC = 13 Then
           Text1(lngC).Text = "100.00"
        ElseIf _
           lngC = 14 Then
           Text1(lngC).Text = "1"
        Else
           Text1(lngC).Text = ""
        End If
    Next lngC
    'dtpOpenDate.Value = Format(Left(PB_regUserinfoU.UserClientDate, 4) + "0101", "0000-00-00")
    dtpOpenDate.Value = Format("19000101", "0000-00-00")
    cboBank.ListIndex = -1
End Sub

'+-------------------------+
'/// Check text1(index) ///
'+-------------------------+
Private Function FncCheckTextBox(lngC As Long, blnOK As Boolean)
    For lngC = Text1.LBound To Text1.UBound
        Select Case lngC
               Case 0  '����ó�ڵ�
                    If Len(Trim(Text1(lngC).Text)) < 1 Then
                       Exit Function
                    End If
               Case 1  '����ó��
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (LenH(Text1(lngC).Text) > 0 And LenH(Text1(lngC).Text) <= 30) Then
                       Exit Function
                    End If
               Case 2  '����ڹ�ȣ
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (LenH(Text1(lngC).Text) <= 14) Then
                       Exit Function
                    End If
               Case 3  '���ι�ȣ
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (LenH(Text1(lngC).Text) <= 14) Then
                       Exit Function
                    End If
               Case 4  '��ǥ�ڸ�
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (LenH(Text1(lngC).Text) <= 30) Then
                       Exit Function
                    End If
               Case 5  '�ֹι�ȣ
                    If Len(Trim(Text1(lngC).Text)) > 14 Then
                       Exit Function
                    End If
               Case 6  '����
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (LenH(Text1(lngC).Text) <= 30) Then
                       Exit Function
                    End If
               Case 7  '����
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (LenH(Text1(lngC).Text) <= 30) Then
                       Exit Function
                    End If
               Case 8  '��ȭ��ȣ
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (LenH(Text1(lngC).Text) <= 20) Then
                       Exit Function
                    End If
               Case 9  '�ѽ���ȣ
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (LenH(Text1(lngC).Text) <= 14) Then
                       Exit Function
                    End If
               Case 10  '��뱸��
                    Text1(lngC).Text = Format(Val(Trim(Text1(lngC).Text)), "0")
                    If Vals(Text1(lngC).Text) <> 0 Then
                       Text1(lngC).Text = "9"
                    End If
                    If Not (Text1(lngC).Text = "0" Or Text1(lngC).Text = "9") Then
                       Exit Function
                    End If
               Case 11  '����ڸ�
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (LenH(Text1(lngC).Text) <= 30) Then
                       Exit Function
                    End If
               Case 12  '���¹�ȣ
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (LenH(Text1(lngC).Text) <= 20) Then
                       Exit Function
                    End If
               Case 13  '��꼭������
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (Vals(Text1(lngC).Text) >= 0) Then
                       Exit Function
                    End If
               Case 14  '�ܰ�����
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (Text1(lngC).Text >= "1" And Text1(lngC) <= "3") Then
                       Exit Function
                    End If
               Case 15  '�����ȣ
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (LenH(Text1(lngC).Text) <= 7) Then
                       Exit Function
                    End If
               Case 16  '�ּ�
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (LenH(Text1(lngC).Text) <= 60) Then
                       Exit Function
                    End If
               Case 17  '����
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (LenH(Text1(lngC).Text) <= 60) Then
                       Exit Function
                    End If
               Case 18  '����
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (LenH(Text1(lngC).Text) <= 100) Then
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
            strExeFile = App.Path & ".\����ó����.rpt"
         Else
            strExeFile = App.Path & ".\����ó����.rpt"
         End If
         varRetVal = Dir(strExeFile)
         If Len(varRetVal) = 0 Then
            intRetCHK = 0
         Else
            .ReportFileName = strExeFile
            On Error GoTo ERROR_CRYSTAL_REPORTS
            '--- Formula Fields ---
            .Formulas(0) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                       '������
            .Formulas(1) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                                  '����Ͻ�
            .Formulas(2) = "ForSortGbn = 1"                           '���ļ���(0.����ó�ڵ�, 1.����ó��)  'Mot Used
            '--- Parameter Fields ---
            '���α׷���������
            .StoredProcParam(0) = PB_regUserinfoU.UserBranchCode '�����ڵ�
            .StoredProcParam(1) = 0                              '��뱸��                                 'Mot Used
            .StoredProcParam(2) = IIf(optPrtChk0.Value = True, 0, 1)             '���ļ���(0.����ó�ڵ�, 1.����ó��)
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
            .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "����ó����"
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
