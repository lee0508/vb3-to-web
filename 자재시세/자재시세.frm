VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm����ü� 
   BorderStyle     =   0  '����
   Caption         =   "����ü�"
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
      TabIndex        =   22
      Top             =   0
      Width           =   15195
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
         Picture         =   "����ü�.frx":0000
         Style           =   1  '�׷���
         TabIndex        =   18
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   9120
         Picture         =   "����ü�.frx":0963
         Style           =   1  '�׷���
         TabIndex        =   19
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   13635
         Picture         =   "����ü�.frx":1308
         Style           =   1  '�׷���
         TabIndex        =   21
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   390
         Left            =   12510
         Picture         =   "����ü�.frx":1C56
         Style           =   1  '�׷���
         TabIndex        =   16
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   11385
         Picture         =   "����ü�.frx":25DA
         Style           =   1  '�׷���
         TabIndex        =   15
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Height          =   390
         Left            =   10260
         Picture         =   "����ü�.frx":2E61
         Style           =   1  '�׷���
         TabIndex        =   34
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00008000&
         BorderStyle     =   1  '���� ����
         Caption         =   "�� �� �� �� �� ��"
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
         TabIndex        =   23
         Top             =   180
         Width           =   4650
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid vsfg1 
      Height          =   6630
      Left            =   60
      TabIndex        =   17
      Top             =   3255
      Width           =   15195
      _cx             =   26802
      _cy             =   11695
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
      Height          =   2565
      Left            =   60
      TabIndex        =   20
      Top             =   630
      Width           =   15195
      Begin VB.TextBox txtFindCD 
         Appearance      =   0  '���
         Height          =   285
         Left            =   9840
         MaxLength       =   18
         TabIndex        =   29
         Top             =   960
         Width           =   1800
      End
      Begin VB.TextBox txtFindSZ 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Left            =   13200
         MaxLength       =   50
         TabIndex        =   32
         Top             =   1320
         Width           =   1800
      End
      Begin VB.CheckBox chkGbn 
         Caption         =   "�����ü���"
         Height          =   255
         Left            =   13200
         TabIndex        =   55
         Top             =   240
         Value           =   1  'Ȯ��
         Width           =   1215
      End
      Begin VB.TextBox txtFindNM 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Left            =   9840
         MaxLength       =   50
         TabIndex        =   31
         Top             =   1320
         Width           =   2280
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Index           =   5
         Left            =   1515
         MaxLength       =   20
         TabIndex        =   6
         Top             =   2160
         Width           =   1995
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '����
         Index           =   4
         Left            =   1515
         MaxLength       =   8
         TabIndex        =   5
         Top             =   1800
         Width           =   1320
      End
      Begin VB.ComboBox cboSupplier 
         Height          =   300
         Left            =   9840
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   28
         Top             =   555
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '����
         Index           =   11
         Left            =   13185
         MaxLength       =   14
         TabIndex        =   13
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   10
         Left            =   9840
         MaxLength       =   20
         TabIndex        =   12
         Top             =   2160
         Width           =   1590
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   9
         Left            =   9840
         MaxLength       =   20
         TabIndex        =   11
         Top             =   1800
         Width           =   1590
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '����
         Index           =   6
         Left            =   4560
         MaxLength       =   20
         TabIndex        =   7
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   7
         Left            =   7080
         MaxLength       =   20
         TabIndex        =   9
         Top             =   1800
         Width           =   1590
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   8
         Left            =   7080
         MaxLength       =   20
         TabIndex        =   10
         Top             =   2160
         Width           =   1590
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   1  '�Է� ���� ����
         Index           =   3
         Left            =   1515
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1320
         Width           =   1320
      End
      Begin VB.ComboBox cboState 
         Height          =   300
         Index           =   1
         Left            =   13185
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   14
         Top             =   2160
         Width           =   1350
      End
      Begin VB.ComboBox cboTaxGbn 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4560
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   8
         Top             =   2160
         Width           =   1350
      End
      Begin VB.ComboBox cboMtGp 
         Height          =   300
         Left            =   9840
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   27
         Top             =   200
         Width           =   3135
      End
      Begin VB.ComboBox cboState 
         Height          =   300
         Index           =   0
         Left            =   13185
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   35
         Top             =   600
         Width           =   1470
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   1  '�Է� ���� ����
         Index           =   2
         Left            =   1515
         MaxLength       =   20
         TabIndex        =   2
         Top             =   960
         Width           =   2040
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Index           =   1
         Left            =   1515
         MaxLength       =   50
         TabIndex        =   1
         Top             =   585
         Width           =   4155
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '����
         Index           =   0
         Left            =   1515
         MaxLength       =   18
         TabIndex        =   0
         Top             =   225
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker dtpStandardDate 
         Height          =   270
         Left            =   13200
         TabIndex        =   30
         Top             =   960
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   19791873
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin MSComCtl2.DTPicker dtpAppDate 
         Height          =   270
         Left            =   4560
         TabIndex        =   4
         Top             =   1320
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
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
         Format          =   19791873
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�����ڵ�"
         Height          =   240
         Index           =   16
         Left            =   8760
         TabIndex        =   57
         Top             =   1005
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�԰�"
         Height          =   240
         Index           =   14
         Left            =   12045
         TabIndex        =   56
         Top             =   1365
         Width           =   975
      End
      Begin VB.Line Line2 
         X1              =   7080
         X2              =   7080
         Y1              =   480
         Y2              =   1680
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "ǰ��"
         Height          =   240
         Index           =   6
         Left            =   8760
         TabIndex        =   54
         Top             =   1365
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��������"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   12
         Left            =   3420
         TabIndex        =   53
         Top             =   1365
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����ó��"
         Height          =   240
         Index           =   11
         Left            =   315
         TabIndex        =   52
         Top             =   2205
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         Caption         =   "[Home]"
         Height          =   240
         Index           =   10
         Left            =   3000
         TabIndex        =   51
         Top             =   1845
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��뱸��"
         Height          =   240
         Index           =   8
         Left            =   11805
         TabIndex        =   50
         Top             =   2220
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "����"
         Height          =   240
         Index           =   19
         Left            =   14640
         TabIndex        =   49
         Top             =   1005
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����ó"
         Height          =   240
         Index           =   18
         Left            =   8760
         TabIndex        =   48
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��������"
         Height          =   240
         Index           =   13
         Left            =   12045
         TabIndex        =   47
         Top             =   1005
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         Caption         =   "[Home]"
         Height          =   240
         Index           =   1
         Left            =   4320
         TabIndex        =   46
         Top             =   255
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "ǰ���ڵ�"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   315
         TabIndex        =   45
         Top             =   260
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����ΰ�"
         Height          =   240
         Index           =   15
         Left            =   8760
         TabIndex        =   44
         Top             =   2220
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����ܰ�"
         Height          =   240
         Index           =   9
         Left            =   8760
         TabIndex        =   43
         Top             =   1845
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "������"
         Height          =   240
         Index           =   7
         Left            =   11805
         TabIndex        =   42
         Top             =   1845
         Width           =   1215
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   15000
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "���Դܰ�"
         Height          =   240
         Index           =   5
         Left            =   5940
         TabIndex        =   41
         Top             =   1845
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "���Ժΰ�"
         Height          =   240
         Index           =   4
         Left            =   5940
         TabIndex        =   40
         Top             =   2220
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��������"
         Height          =   240
         Index           =   3
         Left            =   3405
         TabIndex        =   39
         Top             =   2220
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   2
         Left            =   315
         TabIndex        =   38
         Top             =   1380
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "ǰ��"
         Height          =   240
         Index           =   36
         Left            =   315
         TabIndex        =   37
         Top             =   645
         Width           =   1095
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
         Left            =   7110
         TabIndex        =   36
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�з�"
         Height          =   240
         Index           =   34
         Left            =   8760
         TabIndex        =   33
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����ó�ڵ�"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   31
         Left            =   315
         TabIndex        =   26
         Top             =   1845
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�԰�"
         Height          =   240
         Index           =   26
         Left            =   315
         TabIndex        =   25
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�����"
         Height          =   240
         Index           =   24
         Left            =   3405
         TabIndex        =   24
         Top             =   1845
         Width           =   975
      End
   End
End
Attribute VB_Name = "frm����ü�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' ���α׷� �� �� : ����ü�(�μ�/�߰�/��ȸ/����/����)
' ���� Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' ������ Table   : �������, ����ó, ����, ����з�
' ��  ��  ��  �� : ���� ���ÿ� ����ü� �ڵ�����
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived            As Boolean
Private P_intButton             As Integer
Private P_strFindString1        As String
Private P_strFindString2        As String
Private P_strFindString3        As String
Private P_adoRec                As New ADODB.Recordset
Private Const PC_intRowCnt      As Integer = 21  '�׸��� �� ������ �� ���(FixedRows ����)

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
       P_blnActived = True
       Screen.MousePointer = vbDefault
    End If
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����ü� (�������� ���� ����)"
    Unload Me
    Exit Sub
End Sub

'+----------------+
'/// cboMtGp ///
'+----------------+
Private Sub cboMtGp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       cboSupplier.SetFocus
    End If
End Sub
'+------------------+
'/// cboSuppiler ///
'+------------------+
Private Sub cboSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       cboState(0).SetFocus
    End If
End Sub

'+-----------------------+
'/// cboState(index) ///
'+-----------------------+
Private Sub cboState_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       Select Case Index
              Case 0
                   dtpStandardDate.SetFocus
                   Exit Sub
       End Select
       SendKeys "{tab}"
    End If
End Sub

'+-----------------+
'/// dtpAppDate ///
'+-----------------+
Private Sub dtpAppDate_KeyDown(KeyCode As Integer, Shift As Integer)
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
    If (Index = 0 And (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn)) Then '����ü� �߰��ÿ���
       PB_strCallFormName = "frm����ü�"
       PB_strMaterialsCode = Trim(Text1(0).Text)
       PB_strMaterialsName = "" 'Trim(Text1(1).Text)
       frm����˻�.Show vbModal
       If (Len(PB_strMaterialsCode) + Len(PB_strMaterialsName)) = 0 Then '�˻����� ���(ESC)
       Else
          Text1(0).Text = PB_strMaterialsCode
          Text1(1).Text = PB_strMaterialsName
       End If
       If PB_strMaterialsCode <> "" Then
          SendKeys "{tab}"
       End If
       PB_strCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
    ElseIf _
       (Index = 4 And (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn)) Then '����ó �˻�
       PB_strSupplierCode = UPPER(Trim(Text1(Index).Text))
       PB_strSupplierName = "" 'Trim(Text1(Index + 1).Text)
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
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����ü� �б� ����"
    Unload Me
    Exit Sub
End Sub
Private Sub Text1_LostFocus(Index As Integer)
Dim strSQL As String
Dim lngR   As Long
    With Text1(Index)
         Select Case Index
                Case 0
                     .Text = Trim(.Text)
                     If Len(Text1(Index).Text) < 1 Then
                        Text1(Index).Text = ""
                        Exit Sub
                     End If
                Case 4
                     .Text = UPPER(Trim(.Text))
                Case 7 To 10
                     .Text = Format(Vals(Trim(.Text)), "#,0.00")
                     If Index = 7 Or Index = 9 Then
                        If Vals(Left(cboTaxGbn.Text, 1)) = 0 Then '�����
                           Text1(Index + 1).Text = 0
                        Else                                      '��  ��
                           Text1(Index + 1).Text = Format(Fix(Vals(Trim(.Text)) * (PB_curVatRate)), "#,0.00")
                        End If
                     End If
                     '11.������ = (((���ܰ�+���ΰ�)/(�԰�ܰ�+�԰�ΰ�))-1)*100, 7.�԰�ܰ�, 8.�԰�ΰ�, 9.����ܰ�, 10.���ΰ�
                     If (Vals(Text1(9).Text) + Vals(Text1(10).Text)) > 0 Then
                        If (Vals(Text1(7).Text) + Vals(Text1(8).Text)) > 0 Then
                           Text1(11).Text = Fix((((Vals(Text1(9).Text) + Vals(Text1(10).Text)) _
                                          / (Vals(Text1(7).Text) + Vals(Text1(8).Text)) - 1) * 100) * 100) / 100
                        Else
                           Text1(11).Text = ""
                        End If
                     Else
                        Text1(11).Text = ""
                     End If
         End Select
    End With
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "TABLE �б� ����"
    Unload Me
    Exit Sub
End Sub

'+--------------+
'/// txtFind ///
'+--------------+
Private Sub txtFindCD_GotFocus()
    With txtFindCD
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtFindCD_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL     As String
Dim strExeFile As String
Dim varRetVal  As Variant
    If KeyCode = vbKeyReturn Then
       dtpStandardDate.SetFocus
    End If
End Sub

'+----------------------+
'/// dtpStandardDate ///
'+----------------------+
Private Sub dtpStandardDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       txtFindNM.SetFocus
    End If
End Sub

'+--------------+
'/// txtFind ///
'+--------------+
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
       txtFindSZ.SetFocus
    End If
End Sub
Private Sub txtFindSZ_GotFocus()
    With txtFindSZ
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtFindSZ_KeyDown(KeyCode As Integer, Shift As Integer)
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
Private Sub vsfg1_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsfg1
         If .Row >= .FixedRows Then
            If KeyCode = vbKeyHome Then
               PB_strFMCCallFormName = "frm����ü�"
               PB_strMaterialsCode = .TextMatrix(.Row, 4)
               PB_strMaterialsName = .TextMatrix(.Row, 3)
               PB_strSupplierCode = .TextMatrix(.Row, 6)
               frm����ü��˻�.Show vbModal
            End If
         End If
    End With
End Sub
Private Sub vsfg1_Click()
Dim strData As String
    With vsfg1
         If (.MouseRow >= 0 And .MouseRow < .FixedRows) Then
            .Col = .MouseCol
            .Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlack
            .Cell(flexcpForeColor, 0, .MouseCol, 0, .MouseCol) = vbRed
            strData = Trim(.Cell(flexcpData, .Row, 8))
            Select Case .MouseCol
                   Case 0, 2
                        .ColSel = 2
                        .ColSort(0) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .ColSort(1) = flexSortNone
                        .ColSort(2) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .Sort = flexSortUseColSort
                   Case 1 '����з�
                        .ColSel = 2
                        .ColSort(0) = flexSortNone
                        .ColSort(1) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .ColSort(2) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .Sort = flexSortUseColSort
                   Case 3 '�����
                        .ColSel = 3
                        .ColSort(0) = flexSortNone
                        .ColSort(1) = flexSortNone
                        .ColSort(2) = flexSortNone
                        .ColSort(3) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .Sort = flexSortUseColSort
                   Case Else
                        .Sort = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            End Select
            If .FindRow(strData, , 8) > 0 Then
               .Row = .FindRow(strData, , 8)
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
         Text1(Text1.LBound).Enabled = False: Text1(Text1.LBound + 1).Enabled = True
         Text1(Text1.LBound + 4).Enabled = False: dtpAppDate.Enabled = False
         If .Row >= .FixedRows Then
            For lngC = 0 To .Cols - 1
                Select Case lngC
                       Case 3    '3.�����
                            Text1(1).Text = .TextMatrix(.Row, lngC)
                       Case 4    '4.�����ڵ�
                            Text1(0).Text = .TextMatrix(.Row, lngC)
                       Case 5    '5.��������
                            dtpAppDate.Value = .TextMatrix(.Row, 5)
                       Case 6    '6.����ó�ڵ�
                            Text1(4).Text = .TextMatrix(.Row, lngC)
                       Case 7    '7.����ó��
                            Text1(5).Text = .TextMatrix(.Row, lngC)
                       Case 9 To 10 '9.�԰�, 10.����
                            Text1(lngC - 7).Text = .TextMatrix(.Row, lngC)
                       Case 11   '11.�����
                            Text1(lngC - 5).Text = Format(.ValueMatrix(.Row, lngC), "#,0.00")
                       Case 12   '12.��������
                            cboTaxGbn.ListIndex = .ValueMatrix(.Row, lngC)
                       Case 14 To 18     '14.�԰�ܰ�, 18.������
                            Text1(lngC - 7).Text = Format(.ValueMatrix(.Row, lngC), "#,0.00")
                       Case 20   '��뱸�� ListIndex
                            cboState(1).ListIndex = .ValueMatrix(.Row, lngC)
                End Select
            Next lngC
         End If
    End With
End Sub

'+-----------+
'/// �߰� ///
'+-----------+
Private Sub cmdClear_Click()
Dim strSQL As String
    SubClearText
    vsfg1.Row = 0
    Text1(Text1.LBound).Enabled = True
    Text1(Text1.LBound + 1).Enabled = False
    Text1(Text1.LBound + 4).Enabled = True
    dtpAppDate.Enabled = True
    Text1(Text1.LBound).SetFocus
End Sub
'+-----------+
'/// ��ȸ ///
'+-----------+
Private Sub cmdFind_Click()
    Screen.MousePointer = vbHourglass
    cmdFind.Enabled = False
    P_strFindString1 = Trim(txtFindCD.Text)  '��ȸ�� ��� �˻��� �����ڵ� ����
    P_strFindString2 = Trim(txtFindNM.Text)  '��ȸ�� ��� �˻��� ����� ����
    P_strFindString3 = Trim(txtFindSZ.Text)  '��ȸ�� ��� �˻��� �԰� ����
    SubClearText
    Subvsfg1_FILL
    cmdFind.Enabled = True
    Screen.MousePointer = vbDefault
End Sub
'+-----------+
'/// ���� ///
'+-----------+
Private Sub cmdSave_Click()
Dim strSQL       As String
Dim lngR         As Long
Dim lngC         As Long
Dim blnOK        As Boolean
Dim intRetVal    As Integer
    '�Է³��� �˻�
    blnOK = False
    FncCheckTextBox lngC, blnOK
    If blnOK = False Then
       If lngC = 0 Then
          Text1(lngC).Enabled = True
       End If
       Text1(lngC).SetFocus
       Exit Sub
    End If
    If Text1(Text1.LBound).Enabled = True Then
       intRetVal = MsgBox("�Էµ� �ڷḦ �߰��Ͻðڽ��ϱ� ?", vbQuestion + vbYesNo + vbDefaultButton2, "�ڷ� �߰�")
    Else
       intRetVal = MsgBox("������ �ڷḦ �����Ͻðڽ��ϱ� ?", vbQuestion + vbYesNo + vbDefaultButton2, "�ڷ� ����")
    End If
    If intRetVal = vbNo Then
       vsfg1.SetFocus
       Exit Sub
    End If
    cmdSave.Enabled = False
    With vsfg1
         Screen.MousePointer = vbHourglass
         P_adoRec.CursorLocation = adUseClient
         If Text1(Text1.LBound).Enabled = True Then '����ü� �߰��� �˻�
            strSQL = "SELECT * FROM ����ü� T1 " _
                    & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                      & "AND T1.�з��ڵ� = '" & Mid(Text1(0).Text, 1, 2) & "' AND T1.�����ڵ� = '" & Mid(Text1(0).Text, 3) & "' " _
                      & "AND T1.����ó�ڵ� = '" & Trim(Text1(4).Text) & "' " _
                      & "AND T1.�������� = '" & DTOS(dtpAppDate.Value) & "' "
            On Error GoTo ERROR_TABLE_SELECT
            P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
            If P_adoRec.RecordCount <> 0 Then
               P_adoRec.Close
               Text1(Text1.LBound).SetFocus
               Screen.MousePointer = vbDefault
               cmdSave.Enabled = True
               Exit Sub
            End If
            P_adoRec.Close
            strSQL = "SELECT * FROM ����ó T1 " _
                    & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                      & "AND T1.����ó�ڵ� = '" & Text1(Text1.LBound + 4).Text & "' "
            On Error GoTo ERROR_TABLE_SELECT
            P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
            If P_adoRec.RecordCount = 0 Then
               P_adoRec.Close
               Text1(Text1.LBound + 4).SetFocus
               Screen.MousePointer = vbDefault
               cmdSave.Enabled = True
               Exit Sub
            End If
            P_adoRec.Close
         End If
         PB_adoCnnSQL.BeginTrans
         If Text1(Text1.LBound).Enabled = True Then '����ü� �߰�
            strSQL = "INSERT INTO ����ü�(������ڵ�, �з��ڵ�, �����ڵ�, " _
                                        & "����ó�ڵ�, ��������," _
                                        & "�԰�ܰ�, �԰�ΰ�, " _
                                        & "���ܰ�, ���ΰ�, " _
                                        & "������, ��뱸��," _
                                        & "��������, ������ڵ�) Values('" & PB_regUserinfoU.UserBranchCode & "', " _
                    & "'" & Mid(Text1(Text1.LBound).Text, 1, 2) & "','" & Mid(Text1(Text1.LBound).Text, 3) & "', " _
                    & "'" & Trim(Text1(4).Text) & "','" & DTOS(dtpAppDate.Value) & "', " _
                    & "" & Vals(Trim(Text1(7).Text)) & "," & Vals(Trim(Text1(8).Text)) & ", " _
                    & "" & Vals(Trim(Text1(9).Text)) & "," & Vals(Trim(Text1(10).Text)) & ", " _
                    & "" & Vals(Trim(Text1(11).Text)) & "," & Vals(Left(cboState(1).Text, 1)) & "," _
                    & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' ) "
            On Error GoTo ERROR_TABLE_INSERT
            PB_adoCnnSQL.Execute strSQL
            strSQL = "SELECT * FROM ����ü� T1 " _
                    & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                      & "AND T1.�з��ڵ� = '" & Mid(Text1(0).Text, 1, 2) & "' AND T1.�����ڵ� = '" & Mid(Text1(0).Text, 3) & "' " _
                      & "AND T1.����ó�ڵ� = '" & Trim(Text1(4).Text) & "' " _
                      & "AND T1.�������� = '" & DTOS(dtpAppDate.Value) & "' "
            On Error GoTo ERROR_TABLE_SELECT
            P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
            If P_adoRec.RecordCount = 0 Then '������ ����ü� �߰�
               P_adoRec.Close
               strSQL = "INSERT INTO ����ü�(������ڵ�, �з��ڵ�, �����ڵ�, " _
                                           & "����ó�ڵ�, ��������," _
                                           & "�԰�ܰ�, �԰�ΰ�, " _
                                           & "���ܰ�, ���ΰ�, " _
                                           & "������, ��뱸��," _
                                           & "��������, ������ڵ�) Values('" & PB_regUserinfoU.UserBranchCode & "', " _
                    & "'" & Mid(Text1(Text1.LBound).Text, 1, 2) & "','" & Mid(Text1(Text1.LBound).Text, 3) & "', " _
                    & "'" & Trim(Text1(4).Text) & "','" & DTOS(dtpAppDate.Value) & "', " _
                    & "" & Vals(Trim(Text1(7).Text)) & "," & Vals(Trim(Text1(8).Text)) & ", " _
                    & "" & Vals(Trim(Text1(9).Text)) & "," & Vals(Trim(Text1(10).Text)) & ", " _
                    & "" & Vals(Trim(Text1(11).Text)) & "," & Vals(Left(cboState(1).Text, 1)) & "," _
                    & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' ) "
               On Error GoTo ERROR_TABLE_INSERT
            Else                             '������ ����
               P_adoRec.Close
               strSQL = "UPDATE ����ü� SET " _
                             & "�԰�ܰ� = " & Vals(Trim(Text1(7).Text)) & ", �԰�ΰ� = " & Vals(Trim(Text1(8).Text)) & ", " _
                             & "���ܰ� = " & Vals(Trim(Text1(9).Text)) & ", ���ΰ� = " & Vals(Trim(Text1(10).Text)) & ", " _
                             & "������ = " & Vals(Trim(Text1(11).Text)) & ", ��뱸�� = " & Vals(Left(cboState(1).Text, 1)) & ", " _
                             & "�������� = '" & PB_regUserinfoU.UserServerDate & "', " _
                             & "������ڵ� = '" & PB_regUserinfoU.UserCode & "' " _
                       & "WHERE ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                         & "AND �з��ڵ� = '" & Mid(Text1(Text1.LBound).Text, 1, 2) & "' " _
                         & "AND �����ڵ� = '" & Mid(Text1(Text1.LBound).Text, 3) & "' " _
                         & "AND ����ó�ڵ� = '" & Trim(Text1(4).Text) & "' " _
                         & "AND �������� = '" & DTOS(dtpAppDate.Value) & "' "
               On Error GoTo ERROR_TABLE_UPDATE
            End If
            .AddItem .Rows
            For lngC = 0 To .Cols - 1
                Select Case lngC
                       Case 0  '0.�з��ڵ�
                            .TextMatrix(.Rows - 1, 0) = Left(Text1(0).Text, 2)
                       Case 1  '1.�з���
                            For lngR = 0 To cboMtGp.ListCount - 1
                                If Left(Text1(0).Text, 2) = Left(cboMtGp.List(lngR), 2) Then
                                   .TextMatrix(.Rows - 1, 1) = Trim(Mid(cboMtGp.List(lngR), 5))
                                   Exit For
                                End If
                            Next lngR
                       Case 2  '2.�����ڵ�
                            .TextMatrix(.Rows - 1, 2) = Mid(Text1(0).Text, 3)
                       Case 3  '3.�����
                            .TextMatrix(.Rows - 1, 3) = Trim(Text1(1).Text)
                       Case 4  '4.�����ڵ�
                            .TextMatrix(.Rows - 1, 4) = .TextMatrix(.Rows - 1, 0) + .TextMatrix(.Rows - 1, 2)
                       Case 5  '5.��������
                            .TextMatrix(.Rows - 1, 5) = Format(DTOS(dtpAppDate.Value), "0000-00-00")
                       Case 6  '6.����ó�ڵ�
                            .TextMatrix(.Rows - 1, 6) = Trim(Text1(4).Text)
                       Case 7  '7.����ó��
                            .TextMatrix(.Rows - 1, 7) = Trim(Text1(5).Text)
                       Case 8  '8.����ó��
                            .TextMatrix(.Rows - 1, 8) = Trim(Text1(0).Text) & Trim(Text1(6).Text) _
                                                      & Format(DTOS(dtpAppDate.Value), "0000-00-00")
                            .Cell(flexcpData, .Rows - 1, 8, .Rows - 1, 8) = .TextMatrix(.Rows - 1, 8)
                       Case 9 To 10 '9.�԰�, 10.����
                            .TextMatrix(.Rows - 1, lngC) = Trim(Text1(lngC - 7).Text)
                       Case 11 '11.�����
                            .TextMatrix(.Rows - 1, 11) = Vals(Trim(Text1(6).Text))
                       Case 12 '12.��������
                            .TextMatrix(.Rows - 1, 12) = Vals(Left(cboTaxGbn.Text, 1))
                       Case 13 '13.��������
                            .TextMatrix(.Rows - 1, 13) = Mid(cboTaxGbn.Text, 4)
                       Case 14 To 18 '14.�԰�ܰ�, 18.������
                            .TextMatrix(.Rows - 1, lngC) = Vals(Trim(Text1(lngC - 7).Text))
                       Case 19 '19.��뱸��
                            .TextMatrix(.Rows - 1, 19) = Vals(Left(cboState(1).Text, 1))
                       Case 20 '20.��뱸�� ListIndex
                            .TextMatrix(.Rows - 1, 20) = cboState(1).ListIndex
                       Case 21 '21.��뱸��
                            .TextMatrix(.Rows - 1, 21) = Mid(cboState(1).Text, 4)
                       Case Else
                End Select
            Next lngC
            If .Rows > PC_intRowCnt Then
               .ScrollBars = flexScrollBarBoth
               .TopRow = .Rows - 1
            End If
            Text1(Text1.LBound + 0).Enabled = False: Text1(Text1.LBound + 1).Enabled = True
            Text1(Text1.LBound + 4).Enabled = False: dtpAppDate.Enabled = False
            .Row = .Rows - 1          '�ڵ����� vsfg1_EnterCell Event �߻�
         Else
            strSQL = "UPDATE ����ü� SET " _
                          & "�԰�ܰ� = " & Vals(Trim(Text1(7).Text)) & ", �԰�ΰ� = " & Vals(Trim(Text1(8).Text)) & ", " _
                          & "���ܰ� = " & Vals(Trim(Text1(9).Text)) & ", ���ΰ� = " & Vals(Trim(Text1(10).Text)) & ", " _
                          & "������ = " & Vals(Trim(Text1(11).Text)) & ", ��뱸�� = " & Vals(Left(cboState(1).Text, 1)) & ", " _
                          & "�������� = '" & PB_regUserinfoU.UserServerDate & "', " _
                          & "������ڵ� = '" & PB_regUserinfoU.UserCode & "' " _
                    & "WHERE ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                      & "AND �з��ڵ� = '" & Mid(Text1(Text1.LBound).Text, 1, 2) & "' " _
                      & "AND �����ڵ� = '" & Mid(Text1(Text1.LBound).Text, 3) & "' " _
                      & "AND ����ó�ڵ� = '" & Trim(Text1(4).Text) & "' " _
                      & "AND �������� = '" & DTOS(dtpAppDate.Value) & "' "
            On Error GoTo ERROR_TABLE_UPDATE
            PB_adoCnnSQL.Execute strSQL
            strSQL = "UPDATE ����ü� SET " _
                          & "�԰�ܰ� = " & Vals(Trim(Text1(7).Text)) & ", �԰�ΰ� = " & Vals(Trim(Text1(8).Text)) & ", " _
                          & "���ܰ� = " & Vals(Trim(Text1(9).Text)) & ", ���ΰ� = " & Vals(Trim(Text1(10).Text)) & ", " _
                          & "������ = " & Vals(Trim(Text1(11).Text)) & ", ��뱸�� = " & Vals(Left(cboState(1).Text, 1)) & ", " _
                          & "�������� = '" & PB_regUserinfoU.UserServerDate & "', " _
                          & "������ڵ� = '" & PB_regUserinfoU.UserCode & "' " _
                    & "WHERE ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                      & "AND �з��ڵ� = '" & Mid(Text1(Text1.LBound).Text, 1, 2) & "' " _
                      & "AND �����ڵ� = '" & Mid(Text1(Text1.LBound).Text, 3) & "' " _
                      & "AND ����ó�ڵ� = '" & Trim(Text1(4).Text) & "' " _
                      & "AND �������� = '" & DTOS(dtpAppDate.Value) & "' "
            On Error GoTo ERROR_TABLE_UPDATE
            For lngC = 0 To .Cols - 1
                Select Case lngC
                       Case 0  '0.�з��ڵ�
                            .TextMatrix(.Row, 0) = Left(Text1(0).Text, 2)
                       Case 1  '1.�з���
                            For lngR = 0 To cboMtGp.ListCount - 1
                                If Left(Text1(0).Text, 2) = Left(cboMtGp.List(lngR), 2) Then
                                   .TextMatrix(.Row, 1) = Trim(Mid(cboMtGp.List(lngR), 5))
                                   Exit For
                                End If
                            Next lngR
                       Case 2  '2.�����ڵ�
                            .TextMatrix(.Row, 2) = Right(Text1(0).Text, 4)
                       Case 3  '3.�����
                            .TextMatrix(.Row, 3) = Trim(Text1(1).Text)
                       Case 4  '4.�����ڵ�
                            .TextMatrix(.Row, 4) = .TextMatrix(.Row, 0) + .TextMatrix(.Row, 2)
                       Case 5  '5.��������
                            .TextMatrix(.Row, 5) = Format(DTOS(dtpAppDate.Value), "0000-00-00")
                       Case 6  '6.����ó�ڵ�
                            .TextMatrix(.Row, 6) = Trim(Text1(4).Text)
                       Case 7  '7.����ó��
                            .TextMatrix(.Row, 7) = Trim(Text1(5).Text)
                       Case 8  '8.����ó��
                            .TextMatrix(.Row, 8) = Trim(Text1(0).Text) & Trim(Text1(6).Text) _
                                                 & Format(DTOS(dtpAppDate.Value), "0000-00-00")
                            .Cell(flexcpData, .Row, 8, .Row, 8) = .TextMatrix(.Row, 8)
                       Case 9 To 10 '9.�԰�, 10.����
                            .TextMatrix(.Row, lngC) = Trim(Text1(lngC - 7).Text)
                       Case 11 '11.�����
                            .TextMatrix(.Row, 11) = Vals(Trim(Text1(6).Text))
                       Case 12 '12.��������
                            .TextMatrix(.Row, 12) = Vals(Left(cboTaxGbn.Text, 1))
                       Case 13 '13.��������
                            .TextMatrix(.Row, 13) = Mid(cboTaxGbn.Text, 4)
                       Case 14 To 18 '14.�԰�ܰ�, 18.������
                            .TextMatrix(.Row, lngC) = Vals(Trim(Text1(lngC - 7).Text))
                       Case 19 '19.��뱸��
                            .TextMatrix(.Row, 19) = Vals(Left(cboState(1).Text, 1))
                       Case 20 '20.��뱸�� ListIndex
                            .TextMatrix(.Row, 20) = cboState(1).ListIndex
                       Case 21 '21.��뱸��
                            .TextMatrix(.Row, 21) = Mid(cboState(1).Text, 4)
                       Case Else
                End Select
            Next lngC
         End If
         PB_adoCnnSQL.Execute strSQL
         PB_adoCnnSQL.CommitTrans
         .SetFocus
         Screen.MousePointer = vbDefault
    End With
    cmdSave.Enabled = True
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����ü� �б� ����"
    Unload Me
    Exit Sub
ERROR_TABLE_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����ü� �߰� ����"
    Unload Me
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����ü� ���� ����"
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
               cmdDelete.Enabled = False
               Screen.MousePointer = vbHourglass
               '������ �������̺� �˻�
               'P_adoRec.CursorLocation = adUseClient
               'strSQL = "SELECT Count(*) AS �ش�Ǽ� FROM ����ü� " _
                       & "WHERE �з��ڵ� = '" & Left(Trim(.TextMatrix(.Row, 0)), 2) & "' " _
                         & "AND �����ڵ� = '" & Trim(.TextMatrix(.Row, 2)) & "' "
               'On Error GoTo ERROR_TABLE_SELECT
               'P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
               'lngCnt = P_adoRec("�ش�Ǽ�")
               'P_adoRec.Close
               If lngCnt <> 0 Then
                  MsgBox "����ü�(" & Format(lngCnt, "#,#") & "��)�� �����Ƿ� ������ �� �����ϴ�.", vbCritical, "����ü� ���� �Ұ�"
                  cmdDelete.Enabled = True
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
               PB_adoCnnSQL.BeginTrans
               P_adoRec.CursorLocation = adUseClient
               strSQL = "DELETE FROM ����ü� " _
                       & "WHERE ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                         & "AND �з��ڵ� = '" & Mid(Trim(.TextMatrix(.Row, 4)), 1, 2) & "' " _
                         & "AND �����ڵ� = '" & Mid(Trim(.TextMatrix(.Row, 4)), 3, 16) & "' " _
                         & "AND ����ó�ڵ� = '" & Trim(.TextMatrix(.Row, 6)) & "' " _
                         & "AND �������� = '" & DTOS(.TextMatrix(.Row, 5)) & "' "
               On Error GoTo ERROR_TABLE_DELETE
               PB_adoCnnSQL.Execute strSQL
               strSQL = "DELETE FROM ����ü� " _
                       & "WHERE �з��ڵ� = '" & Mid(Trim(.TextMatrix(.Row, 4)), 1, 2) & "' " _
                         & "AND �����ڵ� = '" & Mid(Trim(.TextMatrix(.Row, 4)), 3, 16) & "' " _
                         & "AND ����ó�ڵ� = '0001' " _
                         & "AND �������� = '" & DTOS(.TextMatrix(.Row, 5)) & "' "
               On Error GoTo ERROR_TABLE_DELETE
               PB_adoCnnSQL.Execute strSQL
               PB_adoCnnSQL.CommitTrans
               .RemoveItem .Row
               Text1(Text1.LBound).Enabled = False: Text1(Text1.LBound + 1).Enabled = True
               Text1(Text1.LBound + 4).Enabled = False: dtpAppDate.Enabled = False
               If .Rows <= PC_intRowCnt Then
                  .ScrollBars = flexScrollBarHorizontal
               End If
               Screen.MousePointer = vbDefault
               SubClearText
               .Row = 0
               vsfg1_EnterCell
               cmdDelete.Enabled = True
            End If
            .SetFocus
         End If
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_SELECT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����ü� �б� ����"
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_DELETE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����ü� ���� ����"
    Unload Me
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
    Set frm����ü� = Nothing
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
    Text1(Text1.LBound).Enabled = False      '�����ڵ� FLASE
    Text1(Text1.LBound + 1).Enabled = True   '����� FLASE
    Text1(Text1.LBound + 4).Enabled = False  '����ó�ڵ� FLASE
    With vsfg1                 'Rows 1, Cols 22, RowHeightMax(Min) 300
         .AllowBigSelection = False
         .AllowSelection = False
         .AllowUserResizing = flexResizeColumns
         .BackColorBkg = &H8000000F
         .BackColorSel = &H8000&
         .Ellipsis = flexEllipsisEnd
         '.ExplorerBar = flexExSortShow
         .ExtendLastCol = True
         .ScrollBars = flexScrollBarHorizontal
         .ScrollTrack = True
         .SelectionMode = flexSelectionByRow
         .FixedRows = 1
         .FixedCols = 9
         .Rows = 1             'Subvsfg1_Fill����ÿ� ����
         .Cols = 22
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1000   '����з�(�з��ڵ�)
         .ColWidth(1) = 1000   '�з���
         .ColWidth(2) = 2000   '�����ڵ�
         .ColWidth(3) = 2500   '�����
         .ColWidth(4) = 1000   '�з��ڵ� + �����ڵ�
         .ColWidth(5) = 1200   '��������
         .ColWidth(6) = 1000   '����ó�ڵ�
         .ColWidth(7) = 2000   '����ó��
         .ColWidth(8) = 1000   '�����ڵ�+����ó�ڵ�+��������
         .ColWidth(9) = 2000   '�԰�
         .ColWidth(10) = 1100  '����
         .ColWidth(11) = 800   '�����
         .ColWidth(12) = 1000  '��������
         .ColWidth(13) = 800   '��������
         .ColWidth(14) = 1300  '�԰�ܰ�
         .ColWidth(15) = 1200  '�԰�ΰ�
         .ColWidth(16) = 1300  '���ܰ�
         .ColWidth(17) = 1200  '���ΰ�
         .ColWidth(18) = 1000  '������
         .ColWidth(19) = 1000  '��뱸��
         .ColWidth(20) = 1000  '��뱸�� ListIndex
         .ColWidth(21) = 1000  '��뱸��
         .Cell(flexcpFontBold, 0, 0, .FixedRows - 1, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "�з��ڵ�"         'H(�з��ڵ�)
         .TextMatrix(0, 1) = "�з���"
         .TextMatrix(0, 2) = "�ڵ�"             '(�����ڵ�)
         .TextMatrix(0, 3) = "ǰ��"
         .TextMatrix(0, 4) = "�����ڵ�"         'H(�з��ڵ� + �����ڵ�)
         .TextMatrix(0, 5) = "��������"
         .TextMatrix(0, 6) = "����ó�ڵ�"       'H
         .TextMatrix(0, 7) = "����ó��"
         .TextMatrix(0, 8) = "KEY"              'H
         .TextMatrix(0, 9) = "�԰�"
         .TextMatrix(0, 10) = "����"
         .TextMatrix(0, 11) = "�����"
         .TextMatrix(0, 12) = "����"            'H
         .TextMatrix(0, 13) = "����"
         .TextMatrix(0, 14) = "���Դܰ�"
         .TextMatrix(0, 15) = "���Ժΰ�"
         .TextMatrix(0, 16) = "����ܰ�"
         .TextMatrix(0, 17) = "����ΰ�"
         .TextMatrix(0, 18) = "������"
         .TextMatrix(0, 19) = "��뱸��"        'H
         .TextMatrix(0, 20) = "��뱸��"        'H
         .TextMatrix(0, 21) = "��뱸��"
         For lngC = 11 To 18
             .ColFormat(lngC) = "#,#.00"
         Next lngC
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 0, 4, 6, 8, 12, 19, 20
                        .ColHidden(lngC) = True
             End Select
         Next lngC
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 1, 2, 3, 7, 9, 10
                         .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 0, 4, 5, 6, 8, 13, 19, 20, 21
                         .ColAlignment(lngC) = flexAlignCenterCenter
                    Case Else
                         .ColAlignment(lngC) = flexAlignRightCenter
             End Select
         Next lngC
         .MergeCells = flexMergeRestrictColumns
         For lngC = 0 To 4
             .MergeCol(lngC) = True
         Next lngC
    End With
End Sub

Private Sub SubOther_FILL()
Dim strSQL        As String
Dim intIndex      As Integer
    Text1(Text1.LBound).Enabled = False: Text1(Text1.LBound + 1).Enabled = True
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.�з��ڵ� AS �з��ڵ�, ISNULL(T1.�з���,'') AS �з��� " _
             & "FROM ����з� T1 " _
            & "ORDER BY T1.�з��ڵ� "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       cboMtGp.ListIndex = -1
       cboMtGp.Enabled = False
       cmdClear.Enabled = False: cmdFind.Enabled = False
       cmdSave.Enabled = False: cmdDelete.Enabled = False
       Exit Sub
    Else
       cboMtGp.AddItem "00. " & "��ü"
       Do Until P_adoRec.EOF
          cboMtGp.AddItem P_adoRec("�з��ڵ�") & ". " & P_adoRec("�з���")
          P_adoRec.MoveNext
       Loop
       P_adoRec.Close
       cboMtGp.ListIndex = 0
    End If
    strSQL = "SELECT T1.����ó�ڵ� AS ����ó�ڵ�, " _
                  & "T1.����ó�� AS ����ó�� " _
             & "FROM ����ó T1 " _
            & "ORDER BY T1.����ó�ڵ� "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       cboSupplier.ListIndex = -1
       cboSupplier.Enabled = False
       cmdClear.Enabled = False: cmdFind.Enabled = False
       cmdSave.Enabled = False: cmdDelete.Enabled = False
       Exit Sub
    Else
       cboSupplier.AddItem "0000. " & "��ü"
       Do Until P_adoRec.EOF
          cboSupplier.AddItem P_adoRec("����ó�ڵ�") & ". " & P_adoRec("����ó��")
          P_adoRec.MoveNext
       Loop
       P_adoRec.Close
       cboSupplier.ListIndex = 0
    End If
    With cboState(0)
         .AddItem "0. ��    ü"
         .AddItem "1. ��    ��"
         .AddItem "2. ���Ұ�"
         .AddItem "3. ��    Ÿ"
         .ListIndex = 1
    End With
    With cboTaxGbn
         .AddItem "0. �����"
         .AddItem "1. ��  ��"
         .ListIndex = 1
    End With
    With cboState(1)
         .AddItem "0. ��    ��"
         .AddItem "9. ���Ұ�"
         .ListIndex = 0
    End With
    dtpStandardDate.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
    dtpAppDate.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
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
Dim strGroupBy As String
Dim strHaving  As String
Dim strWhere   As String
Dim strOrderBy As String
Dim lngR       As Long
Dim lngC       As Long
Dim lngRR      As Long
Dim lngRRR     As Long
    
    If (Len(P_strFindString1) + Len(P_strFindString2) + Len(P_strFindString3)) = 0 Then
       txtFindCD.SetFocus
       Exit Sub
    End If
    '������ڵ�, ���س�¥ �˻�����
    strWhere = "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' "
               '& "AND T1.�������� >= '" & DTOS(dtpStandardDate.Value) & "' "
    '�����ڵ� �˻�����
    If Len(Text1(0).Text) <> 0 Then
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                & "(T1.�з��ڵ� + T1.�����ڵ�) LIKE '%" & Trim(Text1(0).Text) & "%' "
    End If
    '�˻�����(����з�)
    Select Case Left(cboMtGp.Text, 2)
           Case "00" '��ü
                strWhere = strWhere
           Case Else
                strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                                          & "T1.�з��ڵ� = '" & Mid(Trim(cboMtGp.Text), 1, 2) & "' "
    End Select
    '�˻�����(����ó)
    Select Case Left(cboSupplier.Text, 4)
           Case "0000" '��ü
                strWhere = strWhere
           Case Else
                strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                                          & "T1.����ó�ڵ� = '" & Mid(Trim(cboSupplier.Text), 1, 4) & "' "
    End Select
    '�˻�����(��뱸��)
    Select Case cboState(0).ListIndex
           Case 0 '��ü
                strWhere = strWhere
           Case 1 '����
                strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") + "T1.��뱸�� = 0 "
           Case 2 '���Ұ�
                strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") + "T1.��뱸�� = 9 "
           Case Else
                strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") + "NOT(T1.��뱸�� = 0 OR T1.��뱸�� = 9) "
    End Select
    If chkGbn.Value = 1 Then
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                & "T1.�������� = (SELECT TOP 1 �������� FROM ����ü� " _
                                & "WHERE �з��ڵ� = T1.�з��ڵ� AND �����ڵ� = T1.�����ڵ� " _
                                  & "AND ����ó�ڵ� = T1.����ó�ڵ� " _
                                  & "AND �������� <= '" & DTOS(dtpStandardDate.Value) & "' " _
                                  & "AND ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                                & "ORDER BY �������� DESC) "
    Else
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                & "T1.�������� BETWEEN " _
                    & "(SELECT TOP 1 �������� " _
                       & "FROM ����ü� " _
                      & "WHERE �з��ڵ� = T1.�з��ڵ� AND �����ڵ� = T1.�����ڵ� " _
                        & "AND �������� <= '" & DTOS(dtpStandardDate.Value) & "' " _
                        & "AND ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                      & "ORDER BY T1.�������� DESC) " _
                 & "AND '" & DTOS(dtpStandardDate.Value) & "' "
    End If
    If (Len(P_strFindString1) + Len(P_strFindString2) + Len(P_strFindString3)) = 0 Then          '�������� ��ȸ
       strOrderBy = "ORDER BY T1.������ڵ�, T1.�����ڵ�, T2.�����, T1.�������� DESC "
    Else
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") + "T1.�����ڵ� LIKE '%" & P_strFindString1 & "%' " _
                & "AND T2.����� LIKE '%" & P_strFindString2 & "%' AND T2.�԰� LIKE '%" & P_strFindString3 & "%' "
       strOrderBy = "ORDER BY T1.������ڵ�, T2.�����, T1.�������� DESC "
    End If
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT ISNULL(T1.�з��ڵ�,'') AS �з��ڵ�, ISNULL(T1.�����ڵ�,'') AS �����ڵ�, " _
                  & "ISNULL(T1.����ó�ڵ�,'') AS ����ó�ڵ�, ISNULL(T1.��������,'') AS ��������, " _
                  & "ISNULL(T4.�з���,'') AS �з���, " _
                  & "ISNULL(T2.�����,'') AS �����, ISNULL(T3.����ó��,'') AS ����ó��, " _
                  & "ISNULL(T2.�԰�,'') AS �԰�, ISNULL(T2.����,'') ����, " _
                  & "ISNULL(T2.�����,0) AS �����, ISNULL(T2.��������,0) AS ��������, " _
                  & "ISNULl(T1.�԰�ܰ�,0) AS �԰�ܰ�, ISNULL(T1.�԰�ΰ�,0) AS �԰�ΰ�, " _
                  & "ISNULl(T1.���ܰ�,0) AS ���ܰ�, ISNULL(T1.���ΰ�,0) AS ���ΰ�, " _
                  & "ISNULL(T1.������,0) AS ������ , ISNULL(T1.��뱸��,0) AS ��뱸�� " _
             & "FROM ����ü� T1 " _
             & "LEFT JOIN ���� T2 " _
                    & "ON T2.�з��ڵ� = T1.�з��ڵ� AND T2.�����ڵ� = T1.�����ڵ� " _
             & "LEFT JOIN ����ó T3 " _
                    & "ON T3.������ڵ� = T1.������ڵ� AND T3.����ó�ڵ� = T1.����ó�ڵ� " _
             & "LEFT JOIN ����з� T4 ON T4.�з��ڵ� = T1.�з��ڵ� " _
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
               .TextMatrix(lngR, 2) = IIf(IsNull(P_adoRec("�����ڵ�")), "", P_adoRec("�����ڵ�"))
               .TextMatrix(lngR, 3) = IIf(IsNull(P_adoRec("�����")), "", P_adoRec("�����"))
               .TextMatrix(lngR, 4) = .TextMatrix(lngR, 0) & P_adoRec("�����ڵ�")
               If Len(P_adoRec("��������")) = 8 Then
                  .TextMatrix(lngR, 5) = Format(P_adoRec("��������"), "0000-00-00")
               End If
               .TextMatrix(lngR, 6) = IIf(IsNull(P_adoRec("����ó�ڵ�")), "", P_adoRec("����ó�ڵ�"))
               .TextMatrix(lngR, 7) = IIf(IsNull(P_adoRec("����ó��")), "", P_adoRec("����ó��"))
               'FindRow ����� ����
               .TextMatrix(lngR, 8) = .TextMatrix(lngR, 4) & .TextMatrix(lngR, 5) & .TextMatrix(lngR, 6)
               .Cell(flexcpData, lngR, 8, lngR, 8) = .TextMatrix(lngR, 8)
               .TextMatrix(lngR, 9) = P_adoRec("�԰�")
               .TextMatrix(lngR, 10) = P_adoRec("����")
               .TextMatrix(lngR, 11) = P_adoRec("�����")
               .TextMatrix(lngR, 12) = P_adoRec("��������")
               If P_adoRec("��������") = 0 Then
                  .TextMatrix(lngR, 13) = "�����"
               Else
                  .TextMatrix(lngR, 13) = "��  ��"
               End If
               .TextMatrix(lngR, 14) = P_adoRec("�԰�ܰ�")
               .TextMatrix(lngR, 15) = P_adoRec("�԰�ΰ�")
               .TextMatrix(lngR, 16) = P_adoRec("���ܰ�")
               .TextMatrix(lngR, 17) = P_adoRec("���ΰ�")
               .TextMatrix(lngR, 18) = P_adoRec("������")
               .TextMatrix(lngR, 19) = IIf(IsNull(P_adoRec("��뱸��")), "", P_adoRec("��뱸��"))
               'ListIndex
               For lngRRR = 0 To cboState(1).ListCount - 1
                   If .ValueMatrix(lngR, 19) = Vals(Left(cboState(1).List(lngRRR), 1)) Then
                      .TextMatrix(lngR, 20) = lngRRR
                      .TextMatrix(lngR, 21) = Right(Trim(cboState(1).List(lngRRR)), Len(Trim(cboState(1).List(lngRRR))) - 3)
                      Exit For
                   End If
               Next lngRRR
               If .TextMatrix(lngR, 3) = P_strFindString2 Then
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
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "���� �б� ����"
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
Dim strSQL As String
    For lngC = Text1.LBound To Text1.UBound
        Select Case lngC
               Case 0  '�����ڵ�
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Len(Text1(lngC).Text) < 1 Then
                       Text1(lngC).Text = ""
                       Exit Function
                    Else
                       '�����ڵ� �˻�
                       strSQL = "SELECT T1.����� AS ����� FROM ���� T1 " _
                               & "WHERE T1.�з��ڵ� = '" & Mid(Text1(lngC).Text, 1, 2) & "' " _
                                 & "AND T1.�����ڵ� = '" & Mid(Text1(lngC).Text, 3) & "' "
                       On Error GoTo ERROR_TABLE_SELECT
                       P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
                       If P_adoRec.RecordCount = 0 Then
                          P_adoRec.Close
                          Exit Function
                       Else
                          Text1(lngC + 1).Text = P_adoRec("�����")
                          P_adoRec.Close
                       End If
                    End If
               Case 1  '�����
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not Len(Text1(lngC).Text) > 0 Then
                       Text1(lngC).Text = ""
                       lngC = 0
                       Exit Function
                    End If
              Case 4  '����ó�ڵ�
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Len(Text1(lngC).Text) < 1 Then
                       Text1(lngC).Text = ""
                       Exit Function
                    Else
                       '����ó�ڵ� �˻�
                       strSQL = "SELECT T1.����ó�ڵ� AS ����ó�ڵ�, T1.����ó�� AS ����ó�� FROM ����ó T1 " _
                               & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                                 & "AND T1.����ó�ڵ� = '" & Text1(lngC).Text & "' "
                       On Error GoTo ERROR_TABLE_SELECT
                       P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
                       If P_adoRec.RecordCount = 0 Then
                          P_adoRec.Close
                          Exit Function
                       Else
                          Text1(lngC + 1).Text = P_adoRec("����ó��")
                          P_adoRec.Close
                       End If
                    End If
              Case 5  '����ó��
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (Len(Text1(lngC).Text) <> 0) Then
                       Text1(lngC).Text = ""
                       lngC = 2
                       Exit Function
                    End If
              Case 7, 9  '7.�԰�ܰ�, 9.���ܰ�
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If (Vals(Text1(7).Text) < 1) And (Vals(Text1(9).Text) < 1) Then
                       Exit Function
                    End If
        End Select
    Next lngC
    blnOK = True
    Exit Function
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "TABLE �б� ����"
    Unload Me
    Exit Function
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

Dim strForAppDate          As String  '��������       (Formula)
Dim strForBranchName       As String  '������       (Formula)
Dim strForPrtDateTime      As String  '����Ͻ�       (Formula)
Dim strParGroupCode        As Integer '�з��ڵ�       (Parameter)
Dim intParStateCode        As Integer '��뱸��       (Parameter)

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
            strExeFile = App.Path & ".\����ü�.rpt"
         Else
            strExeFile = App.Path & ".\����ü�T.rpt"
         End If
         varRetVal = Dir(strExeFile)
         If Len(varRetVal) = 0 Then
            intRetCHK = 0
         Else
            .ReportFileName = strExeFile
            On Error GoTo ERROR_CRYSTAL_REPORTS
            '--- Formula Fields ---
            .Formulas(0) = "ForAppDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'" '���α׷�����
            .Formulas(1) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                     '������
            .Formulas(2) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                                '����Ͻ�
            .Formulas(3) = "ForStdDate = '" & Format(DTOS(dtpStandardDate.Value), "0000-00-00") & "'"    '��������
            '--- Parameter Fields ---
            '����з�(�з��ڵ�)
            If Mid(cboMtGp.Text, 1, 2) = "00" Then
               .StoredProcParam(0) = " "
            Else
               .StoredProcParam(0) = Mid(cboMtGp.Text, 1, 2)
            End If
            '����ó�ڵ�
            If Mid(cboSupplier.Text, 1, 4) = "0000" Then
               .StoredProcParam(1) = " "
            Else
               .StoredProcParam(1) = Mid(cboSupplier.Text, 1, 4)
            End If
            .StoredProcParam(2) = DTOS(dtpStandardDate.Value)   '��������
            '�����
            If Len(txtFindNM.Text) = 0 Then
               .StoredProcParam(3) = " "
            Else
               .StoredProcParam(3) = Trim(txtFindNM.Text)
            End If
            .StoredProcParam(4) = chkGbn.Value                  '(0.��ü�ü�, 1.�����ü�)
            If cboState(0).ListIndex < 2 Then                   '��뱸��(0.��ü, 1.����, 2.����, 3.�� ��)
               .StoredProcParam(5) = 0
            Else
               .StoredProcParam(5) = 9
            End If
            .StoredProcParam(6) = PB_regUserinfoU.UserBranchCode
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
            .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "����ü�"
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
 
