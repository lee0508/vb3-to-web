VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm�������� 
   BorderStyle     =   0  '����
   Caption         =   "��������"
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
      TabIndex        =   23
      Top             =   0
      Width           =   15195
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   4920
         Top             =   200
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowLeft      =   0
         WindowTop       =   0
         WindowWidth     =   15405
         WindowHeight    =   11100
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   375
         Left            =   7980
         Picture         =   "��������.frx":0000
         Style           =   1  '�׷���
         TabIndex        =   45
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   9120
         Picture         =   "��������.frx":0963
         Style           =   1  '�׷���
         TabIndex        =   27
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   13635
         Picture         =   "��������.frx":1308
         Style           =   1  '�׷���
         TabIndex        =   26
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   390
         Left            =   12510
         Picture         =   "��������.frx":1C56
         Style           =   1  '�׷���
         TabIndex        =   25
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   11385
         Picture         =   "��������.frx":25DA
         Style           =   1  '�׷���
         TabIndex        =   18
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Height          =   390
         Left            =   10260
         Picture         =   "��������.frx":2E61
         Style           =   1  '�׷���
         TabIndex        =   38
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
         TabIndex        =   24
         Top             =   180
         Width           =   4650
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid vsfg1 
      Height          =   6450
      Left            =   60
      TabIndex        =   19
      Top             =   3585
      Width           =   15195
      _cx             =   26802
      _cy             =   11377
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
      Height          =   2925
      Left            =   60
      TabIndex        =   20
      Top             =   630
      Width           =   15195
      Begin VB.CheckBox chkCodeException 
         Caption         =   "CODE ǰ�� ����"
         Height          =   180
         Left            =   10860
         TabIndex        =   67
         Top             =   240
         Value           =   1  'Ȯ��
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Index           =   8
         Left            =   1155
         TabIndex        =   11
         Top             =   2535
         Width           =   9465
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   9
         Left            =   11985
         MaxLength       =   20
         TabIndex        =   12
         Top             =   1755
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   10
         Left            =   11985
         MaxLength       =   20
         TabIndex        =   13
         Top             =   2115
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   11
         Left            =   11985
         MaxLength       =   20
         TabIndex        =   14
         Top             =   2475
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   12
         Left            =   13410
         MaxLength       =   20
         TabIndex        =   15
         Top             =   1755
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   13
         Left            =   13410
         MaxLength       =   20
         TabIndex        =   16
         Top             =   2115
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   14
         Left            =   13410
         MaxLength       =   20
         TabIndex        =   17
         Top             =   2475
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   6
         Left            =   7920
         MaxLength       =   8
         TabIndex        =   9
         Top             =   1440
         Width           =   825
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '����
         Index           =   7
         Left            =   7920
         MaxLength       =   20
         TabIndex        =   10
         Top             =   1800
         Width           =   2685
      End
      Begin VB.TextBox txtSebuCodeRe 
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '����
         Left            =   3795
         MaxLength       =   16
         TabIndex        =   55
         Top             =   570
         Width           =   2055
      End
      Begin VB.ComboBox cboMtGpRe 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1155
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   53
         Top             =   570
         Width           =   1575
      End
      Begin VB.TextBox txtFindBarCode 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Left            =   13400
         MaxLength       =   13
         TabIndex        =   36
         Top             =   945
         Width           =   1350
      End
      Begin VB.TextBox txtFindCD 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Left            =   10860
         MaxLength       =   18
         TabIndex        =   33
         Top             =   585
         Width           =   1605
      End
      Begin VB.TextBox txtFindSZ 
         Appearance      =   0  '���
         Height          =   285
         Left            =   10860
         MaxLength       =   30
         TabIndex        =   35
         Top             =   945
         Width           =   1605
      End
      Begin VB.TextBox txtFindNM 
         Appearance      =   0  '���
         Height          =   285
         Left            =   7920
         MaxLength       =   30
         TabIndex        =   34
         Top             =   945
         Width           =   1845
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   2
         Left            =   4500
         MaxLength       =   13
         TabIndex        =   3
         Top             =   945
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   5
         Left            =   1155
         MaxLength       =   8
         TabIndex        =   6
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   1  '�Է� ���� ����
         Index           =   4
         Left            =   1155
         MaxLength       =   20
         TabIndex        =   5
         Top             =   1800
         Width           =   1335
      End
      Begin VB.ComboBox cboState 
         Height          =   300
         Index           =   1
         Left            =   4500
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   7
         Top             =   1440
         Width           =   1350
      End
      Begin VB.ComboBox cboTaxGbn 
         Height          =   300
         Left            =   4515
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   8
         Top             =   1800
         Width           =   1350
      End
      Begin VB.ComboBox cboMtGp 
         Height          =   300
         Index           =   1
         Left            =   7920
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   32
         Top             =   555
         Width           =   1850
      End
      Begin VB.ComboBox cboState 
         Height          =   300
         Index           =   0
         Left            =   13650
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   31
         Top             =   555
         Width           =   1100
      End
      Begin VB.ComboBox cboMtGp 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   1155
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   0
         Top             =   200
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   3
         Left            =   1155
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1440
         Width           =   2400
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         Index           =   1
         Left            =   1140
         MaxLength       =   30
         TabIndex        =   2
         Top             =   945
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '����
         Index           =   0
         Left            =   3795
         MaxLength       =   16
         TabIndex        =   1
         Top             =   225
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����ó��"
         Height          =   240
         Index           =   27
         Left            =   6520
         TabIndex        =   66
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         Caption         =   " ���Դܰ� "
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   18
         Left            =   12105
         TabIndex        =   65
         ToolTipText     =   "�������"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         Caption         =   " ����ܰ� "
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   19
         Left            =   13545
         TabIndex        =   64
         ToolTipText     =   "�������"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "1."
         Height          =   240
         Index           =   20
         Left            =   11640
         TabIndex        =   63
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "2."
         Height          =   240
         Index           =   21
         Left            =   11640
         TabIndex        =   62
         Top             =   2175
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "3."
         Height          =   240
         Index           =   22
         Left            =   11640
         TabIndex        =   61
         Top             =   2535
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����ó�ڵ�"
         Height          =   240
         Index           =   17
         Left            =   6520
         TabIndex        =   60
         ToolTipText     =   "�������"
         Top             =   1485
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         Caption         =   "[Home]"
         Height          =   240
         Index           =   16
         Left            =   9225
         TabIndex        =   59
         Top             =   1500
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   23
         Left            =   195
         TabIndex        =   58
         ToolTipText     =   "�������"
         Top             =   2595
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         Caption         =   ")"
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   15
         Left            =   5880
         TabIndex        =   57
         ToolTipText     =   "�ڵ庯��"
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         Caption         =   "[Home]"
         Height          =   240
         Index           =   14
         Left            =   6120
         TabIndex        =   56
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�����ڵ�"
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   13
         Left            =   2835
         TabIndex        =   54
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�з�"
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   12
         Left            =   550
         TabIndex        =   52
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         Caption         =   "=>("
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   4
         Left            =   240
         TabIndex        =   51
         ToolTipText     =   "�ڵ庯��"
         Top             =   600
         Width           =   375
      End
      Begin VB.Line Line2 
         X1              =   7080
         X2              =   7080
         Y1              =   240
         Y2              =   1320
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "���ڵ�"
         Height          =   240
         Index           =   9
         Left            =   12480
         TabIndex        =   50
         Top             =   1005
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         Caption         =   "[Home]"
         Height          =   240
         Index           =   8
         Left            =   6120
         TabIndex        =   49
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�����ڵ�"
         Height          =   240
         Index           =   7
         Left            =   9975
         TabIndex        =   48
         Top             =   645
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�԰�"
         Height          =   240
         Index           =   5
         Left            =   10200
         TabIndex        =   47
         Top             =   1005
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "ǰ��"
         Height          =   240
         Index           =   11
         Left            =   7140
         TabIndex        =   46
         Top             =   1005
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "���ڵ�"
         Height          =   240
         Index           =   10
         Left            =   3660
         TabIndex        =   44
         Top             =   1005
         Width           =   735
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   15000
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�����"
         Height          =   240
         Index           =   6
         Left            =   195
         TabIndex        =   43
         Top             =   2205
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��������"
         Height          =   240
         Index           =   3
         Left            =   3660
         TabIndex        =   42
         Top             =   1845
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   2
         Left            =   195
         TabIndex        =   41
         Top             =   1845
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "ǰ��"
         Height          =   240
         Index           =   36
         Left            =   300
         TabIndex        =   40
         Top             =   1005
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         Caption         =   "(�˻�����)"
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
         Left            =   9600
         TabIndex        =   39
         ToolTipText     =   "300"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�з�"
         Height          =   240
         Index           =   34
         Left            =   7005
         TabIndex        =   37
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�԰�"
         Height          =   240
         Index           =   26
         Left            =   195
         TabIndex        =   30
         Top             =   1500
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��뱸��"
         Height          =   240
         Index           =   25
         Left            =   12800
         TabIndex        =   29
         Top             =   615
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��뱸��"
         Height          =   240
         Index           =   24
         Left            =   3660
         TabIndex        =   28
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�����ڵ�"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   1
         Left            =   2835
         TabIndex        =   22
         Top             =   285
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�з�"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   195
         TabIndex        =   21
         Top             =   255
         Width           =   855
      End
   End
End
Attribute VB_Name = "frm��������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' ���α׷� �� �� : ��������
' ���� Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' ������ Table   : ����, (����з�)
' ��  ��  ��  �� :
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived       As Boolean
Private P_adoRec           As New ADODB.Recordset
Private P_intButton        As Integer
Private P_strFindString1   As String
Private P_strFindString2   As String
Private P_strFindString3   As String
Private P_strFindString4   As String
Private Const PC_intRowCnt As Integer = 20  '�׸��� �� ������ �� ���(FixedRows ����)

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
       frmMain.SBar.Panels(4).Text = ""
       P_blnActived = True
       Screen.MousePointer = vbDefault
    End If
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "���� (�������� ���� ����)"
    Unload Me
    Exit Sub
End Sub

'+-----------------------+
'/// cboMtGp(index) ///
'+-----------------------+
Private Sub cboMtGp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       Select Case Index
       End Select
       SendKeys "{tab}"
    End If
End Sub
'+------------------------------+
'/// cboMtGpRe(�з��ڵ庯��) ///
'+------------------------------+
Private Sub cboMtGpRe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       txtSebuCodeRe.SetFocus
    End If
End Sub
Private Sub txtSebuCodeRe_GotFocus()
    With txtSebuCodeRe
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtSebuCodeRe_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strExeFile As String
Dim varRetVal  As Variant
Dim inti       As Integer
    If (Len(Trim(txtSebuCodeRe.Text)) > 0 And (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn)) Then 'ǰ���ڵ� ����ÿ���
       PB_strCallFormName = "frm��������"
       PB_strMaterialsCode = Trim(txtSebuCodeRe.Text) 'Mid(cboMtGp(0).Text, 1, 2)
       PB_strMaterialsName = "" 'Trim(Text1(1).Text)
       frm����˻�.Show vbModal
       If (Len(PB_strMaterialsCode) + Len(PB_strMaterialsName)) = 0 Then '�˻����� ���(ESC)
       Else
          For inti = 0 To cboMtGpRe.ListCount - 1
              cboMtGpRe.ListIndex = inti
              If Mid(cboMtGpRe.Text, 1, 2) = Mid(PB_strMaterialsCode, 1, 2) Then
                 Exit For
              End If
          Next inti
          txtSebuCodeRe.Text = Mid(PB_strMaterialsCode, 3) '�����ڵ�
       End If
       'If PB_strMaterialsCode = "" Then
       '   PB_strCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
       'Else
          PB_strCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
       'End If
       Text1(1).SetFocus
    Else
       If KeyCode = vbKeyReturn Then
          Text1(1).SetFocus
       End If
    End If
End Sub

Private Sub txtSebuCodeRe_LostFocus()
    txtSebuCodeRe.Text = UPPER(Trim(txtSebuCodeRe.Text))
End Sub
'+-----------------------+
'/// cboState(index) ///
'+-----------------------+
Private Sub cboState_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       Select Case Index
       End Select
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
'/// dtpInputDate ///
'+-------------------+
Private Sub dtpInputDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    End If
End Sub
'+--------------------+
'/// dtpOutputDate ///
'+--------------------+
Private Sub dtpOutputDate_KeyDown(KeyCode As Integer, Shift As Integer)
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
Dim inti       As Integer
    If (Index = 0 And (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn)) Then '�������� �߰��ÿ���
       PB_strCallFormName = "frm��������"
       PB_strMaterialsCode = (Text1(0).Text) 'Mid(cboMtGp(0).Text, 1, 2)
       PB_strMaterialsName = "" 'Trim(Text1(1).Text)
       frm����˻�.Show vbModal
       If (Len(PB_strMaterialsCode) + Len(PB_strMaterialsName)) = 0 Then '�˻����� ���(ESC)
       Else
          For inti = 0 To cboMtGp(0).ListCount - 1
              cboMtGp(0).ListIndex = inti
              If Mid(cboMtGp(0).Text, 1, 2) = Mid(PB_strMaterialsCode, 1, 2) Then
                 Exit For
              End If
          Next inti
          Text1(0).Text = Mid(PB_strMaterialsCode, 3) '�����ڵ�
          Text1(1).Text = PB_strMaterialsName         'ǰ��
       End If
       If PB_strMaterialsCode <> "" Then
          SendKeys "{tab}"
       End If
       PB_strCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
    ElseIf _
       (Index = 6 And (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn)) Then '����ó�˻�
       PB_strFMCCallFormName = "frm��������"
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
End Sub
Private Sub Text1_LostFocus(Index As Integer)
Dim strSQL As String
Dim lngR   As Long
    With Text1(Index)
         Select Case Index
                Case 0
                     .Text = UPPER(Trim(.Text))
                     If Len(Text1(Index).Text) < 1 Then
                        Text1(Index).Text = ""
                        Exit Sub
                     End If
                     'If Text1(Index).Enabled = True Then
                     '   P_adoRec.CursorLocation = adUseClient
                     '   strSQL = "SELECT * FROM ���� " _
                     '           & "WHERE �з��ڵ� = '" & Left(cboMtGp(0).List(lngR), 2) & "' " _
                     '             & "AND �����ڵ� = '" & Trim(.Text) & "' "
                     '   On Error GoTo ERROR_TABLE_SELECT
                     '   P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
                     '   If P_adoRec.RecordCount <> 0 Then
                     '      P_adoRec.Close
                     '      .Text = ""
                     '      .SetFocus
                     '      Exit Sub
                     '   End If
                     '   P_adoRec.Close
                     'End If
                Case 6
                     .Text = UPPER(Trim(.Text))
                Case 5, 9 To 14
                     .Text = Format(Vals(Trim(.Text)), "#,0.00")
         End Select
    End With
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "���� �б� ����"
    Unload Me
    Exit Sub
End Sub

'+--------------+
'/// txtFind ///
'+--------------+
'+-------------------------------------+
'/// chkCodeException(�ڵ�ǰ������) ///
'+-------------------------------------+
Private Sub chkCodeException_Click()
    cboMtGp(1).SetFocus
End Sub
Private Sub txtFindCD_GotFocus()
    With txtFindCD
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtFindCD_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strExeFile As String
Dim varRetVal  As Variant
    If KeyCode = vbKeyReturn Then
       txtFindNM.SetFocus
    End If
End Sub
Private Sub txtFindCD_LostFocus()
    With txtFindCD
         .Text = UPPER(Trim(.Text))
    End With
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
Dim strExeFile As String
Dim varRetVal  As Variant
    If KeyCode = vbKeyReturn Then
       txtFindSZ.SetFocus
    End If
End Sub
'+--------------+
'/// txtFind ///
'+--------------+
Private Sub txtFindSZ_GotFocus()
    With txtFindSZ
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtFindSZ_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strExeFile As String
Dim varRetVal  As Variant
    If KeyCode = vbKeyReturn Then
       txtFindBarCode.SetFocus
    End If
End Sub
'+--------------+
'/// txtFind ///
'+--------------+
Private Sub txtFindBarCode_GotFocus()
    With txtFindBarCode
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtFindBarCode_KeyDown(KeyCode As Integer, Shift As Integer)
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
            If KeyCode = vbKeyF1 Then '����ü��˻�
               'PB_strFMCCallFormName = "frm��������"
               'PB_strMaterialsCode = .TextMatrix(.Row, 4)
               'PB_strMaterialsName = .TextMatrix(.Row, 3)
               'PB_strSupplierCode = ""
               'frm����ü��˻�.Show vbModal
            ElseIf _
               KeyCode = vbKeyF2 Then '�������
               'PB_strFMWCallFormName = "frm��������"
               'PB_strMaterialsCode = .TextMatrix(.Row, 4)
               'PB_strMaterialsName = .TextMatrix(.Row, 3)
               'frm����ü�.Show vbModal
               'MsgBox "�������"
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
            If .FindRow(strData, , 4) > 0 Then
               .Row = .FindRow(strData, , 4)
            End If
            If PC_intRowCnt < .Rows Then
               .TopRow = .Row
            End If
         End If
    End With
End Sub
Private Sub vsfg1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
Dim lngR As Long
Dim lngC As Long
    With vsfg1
         cboMtGp(cboMtGp.LBound).Enabled = False
         Text1(Text1.LBound).Enabled = False
         cboMtGpRe.Enabled = True
         txtSebuCodeRe.Enabled = True
         txtSebuCodeRe.Text = ""
         If .Row >= .FixedRows And OldRow <> NewRow Then
            For lngC = 0 To .Cols - 1
                Select Case lngC
                       Case 0
                            cboMtGp(0).Text = .TextMatrix(.Row, 0) + ". " + .TextMatrix(.Row, 1)
                       Case 2 To 3
                            Text1(lngC - 2).Text = .TextMatrix(.Row, lngC) '�����ڵ�
                       Case 16 '���ڵ�
                            Text1(2).Text = .TextMatrix(.Row, lngC)
                       Case 5 To 6 '�԰�, ����
                            Text1(lngC - 2).Text = .TextMatrix(.Row, lngC)
                       Case 17 '�����
                            Text1(5).Text = Format(.ValueMatrix(.Row, lngC), "#,0.00")
                       Case 22 '��뱸�� listindex
                            cboState(1).ListIndex = .ValueMatrix(.Row, lngC)
                       Case 19 '�������� listindex
                            cboTaxGbn.ListIndex = .ValueMatrix(.Row, lngC)
                       Case 7 To 8 '����ó
                            Text1(lngC - 1).Text = .TextMatrix(.Row, lngC)
                       Case 15 '����
                            Text1(8).Text = .TextMatrix(.Row, lngC)
                       Case 9 To 14
                            Text1(lngC).Text = Format(.ValueMatrix(.Row, lngC), "#,0.00")
                       Case Else
                End Select
            Next lngC
         End If
    End With
End Sub
Private Sub vsfg1_EnterCell()
Dim lngR As Long
Dim lngC As Long
    With vsfg1
         'cboMtGp(cboMtGp.LBound).Enabled = False
         'Text1(Text1.LBound).Enabled = False
         'cboMtGpRe.Enabled = True
         'txtSebuCodeRe.Enabled = True
         'txtSebuCodeRe.Text = ""
    End With
End Sub

'+-----------+
'/// �߰� ///
'+-----------+
Private Sub cmdClear_Click()
Dim strSQL As String
    SubClearText
    vsfg1.Row = 0
    cboMtGp(cboMtGp.LBound).Enabled = True
    Text1(Text1.LBound).Enabled = True  'Log Counter���ÿ��� False
    cboMtGpRe.Enabled = False
    txtSebuCodeRe.Enabled = False
    cboMtGp(cboMtGp.LBound).SetFocus
End Sub
'+-----------+
'/// ��ȸ ///
'+-----------+
Private Sub cmdFind_Click()
    Screen.MousePointer = vbHourglass
    cmdFind.Enabled = False
    P_strFindString1 = Trim(txtFindCD.Text)       '��ȸ�� ��� �˻��� �����ڵ� ����
    P_strFindString2 = Trim(txtFindNM.Text)       '��ȸ�� ��� �˻��� ����� ����
    P_strFindString3 = Trim(txtFindSZ.Text)       '��ȸ�� ��� �˻��� �԰� ����
    P_strFindString4 = Trim(txtFindBarCode.Text)  '��ȸ�� ��� �˻��� ���ڵ� ����
    SubClearText
    Subvsfg1_FILL
    cmdFind.Enabled = True
    Screen.MousePointer = vbDefault
End Sub
'+-----------+
'/// ���� ///
'+-----------+
Private Sub cmdSave_Click()
Dim strSQL        As String
Dim lngR          As Long
Dim lngC          As Long
Dim blnOK         As Boolean
Dim intRetVal     As Integer
Dim strServerDate As String
Dim strServerTime As String
    '�߰��̸� �̹��ִ� ǰ������ �˻�
    If Text1(0).Enabled = True Then
       P_adoRec.CursorLocation = adUseClient
       strSQL = "SELECT �з��ڵ�, �����ڵ� FROM ���� " _
               & "WHERE �з��ڵ� = '" & Left(cboMtGp(0).Text, 2) & "' " _
                 & "AND �����ڵ� = '" & Trim(Text1(0).Text) & "' "
       On Error GoTo ERROR_TABLE_SELECT
       P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
       If P_adoRec.RecordCount <> 0 Then
          P_adoRec.Close
          MsgBox "�̹� ����� �Ϸ�� ǰ���Դϴ�. Ȯ���� �ٽ� �Է��Ͽ� �ּ���.", vbCritical, "ǰ�� �ߺ� ���"
          Text1(0).SetFocus
          Exit Sub
       End If
       P_adoRec.Close
    End If
    If cboMtGp(cboMtGp.LBound).Enabled = False Then  '�����̸�
       If Len(Trim(Text1(Text1.LBound))) = 0 Then
          Text1(1).SetFocus
          Exit Sub
       End If
       If Len(Trim(txtSebuCodeRe.Text)) > 0 Then     '�ڵ庯���̸�
          P_adoRec.CursorLocation = adUseClient
          strSQL = "SELECT �з��ڵ�, �����ڵ� FROM ���� " _
                  & "WHERE �з��ڵ� = '" & Left(cboMtGpRe.Text, 2) & "' " _
                    & "AND �����ڵ� = '" & Trim(txtSebuCodeRe.Text) & "' "
          On Error GoTo ERROR_TABLE_SELECT
          P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
          If P_adoRec.RecordCount <> 0 Then
             P_adoRec.Close
             MsgBox "�̹� ����� �Ϸ�� ǰ���Դϴ�. Ȯ���� �ٽ� �Է��Ͽ� �ּ���.", vbCritical, "ǰ���ڵ� ���� �Ұ�"
             txtSebuCodeRe.SetFocus
             Exit Sub
          End If
          P_adoRec.Close
       End If
    End If
    '���ڵ� �ߺ�üũ
    If Len(Trim(Text1(2).Text)) > 0 Then
       P_adoRec.CursorLocation = adUseClient
       strSQL = "SELECT �з��ڵ�, �����ڵ�, ���ڵ� FROM ���� " _
               & "WHERE NOT (�з��ڵ� = '" & Left(cboMtGp(0).Text, 2) & "' " _
                 & "AND �����ڵ� = '" & Trim(Text1(0).Text) & "') " _
                 & "AND ���ڵ� = '" & Trim(Text1(2).Text) & "' "
       On Error GoTo ERROR_TABLE_SELECT
       P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
       If P_adoRec.RecordCount <> 0 Then
          P_adoRec.Close
          MsgBox "�̹� ����� �Ϸ�� ���ڵ��Դϴ�. Ȯ���� �ٽ� �Է��Ͽ� �ּ���.", vbCritical, "ǰ�� ���ڵ� �ߺ� ���"
          Text1(2).SetFocus
          Exit Sub
       End If
       P_adoRec.Close
    End If
    '�Է³��� �˻�
    blnOK = False
    FncCheckTextBox lngC, blnOK
    If blnOK = False Then
       'If lngC = 0 Then
       '   Text1(lngC).Enabled = True
       'End If
       Text1(lngC).SetFocus
       Exit Sub
    End If
    If cboMtGp(cboMtGp.LBound).Enabled = True Then
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
         PB_adoCnnSQL.BeginTrans
         If cboMtGp(cboMtGp.LBound).Enabled = True Then '���� �߰��� �˻�  '�α�
            'strSQL = "SELECT * FROM ���� " _
            '        & "WHERE �����ڵ� = '" & Trim(Text1(Text1.LBound).Text) & "' "
            'On Error GoTo ERROR_TABLE_SELECT
            'P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
            'If P_adoRec.RecordCount <> 0 Then
            '   P_adoRec.Close
            '   Text1(Text1.LBound).SetFocus
            '   Screen.MousePointer = vbDefault
            '   cmdSave.Enabled = True
            '   Exit Sub
            'End If
            'P_adoRec.Close
            '// Log Counter
            'P_adoRec.CursorLocation = adUseClient
            'strSQL = "spLogCounter '����', '" & Left(cboMtGp(0).Text, 2) & "', 0, 0, '" & PB_regUserinfoU.UserCode & "','' "
            'On Error GoTo ERROR_STORED_PROCEDURE
            'P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
            'Text1(0).Text = Format(P_adoRec(0), "0000")
            'P_adoRec.Close
         End If
         If cboMtGp(cboMtGp.LBound).Enabled = True Then '���� �߰�
            strSQL = "INSERT INTO ����(�з��ڵ�, �����ڵ�, " _
                                    & "�����, ���ڵ�, �԰�," _
                                    & "����, �����, " _
                                    & "��������, ����, ��뱸��, " _
                                    & "��������, ������ڵ�) VALUES( " _
                    & "'" & Left(Trim(cboMtGp(0).Text), 2) & "', '" & Trim(Text1(0).Text) & "', " _
                    & "'" & Trim(Text1(1).Text) & "', '" & Trim(Text1(2).Text) & "', '" & Trim(Text1(3).Text) & "', " _
                    & "'" & Trim(Text1(4).Text) & "'," & Vals(Trim(Text1(5).Text)) & "," _
                    & "" & Vals(Left(cboTaxGbn.Text, 1)) & ", '', " & Vals(Left(cboState(1).Text, 1)) & ", " _
                    & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' ) "
            On Error GoTo ERROR_TABLE_INSERT
            PB_adoCnnSQL.Execute strSQL
            '��������߰�
            strSQL = "INSERT INTO ������� " _
                   & "SELECT '" & PB_regUserinfoU.UserBranchCode & "' AS ������ڵ�, " _
                         & "ISNULL(T1.�з��ڵ� , '') AS �з��ڵ�, ISNULL(T1.�����ڵ� , '') AS �����ڵ�, " _
                         & "0 AS �������, 0 AS �������, '' AS �����԰�����, '' AS �����������, " _
                         & "T1.��뱸�� AS ��뱸��, T1.�������� AS ��������, T1.������ڵ� AS ������ڵ�, " _
                         & "'" & Trim(Text1(8).Text) & "' AS ����, '" & Trim(Text1(6).Text) & "' AS �ָ���ó�ڵ�, " _
                         & "" & Vals(Trim(Text1(9).Text)) & " AS �԰�ܰ�1, " & Vals(Trim(Text1(10).Text)) & " AS �԰�ܰ�2, " _
                         & "" & Vals(Trim(Text1(11).Text)) & " AS �԰�ܰ�3, " & Vals(Trim(Text1(12).Text)) & " AS ���ܰ�1, " _
                         & "" & Vals(Trim(Text1(13).Text)) & " AS ���ܰ�2, " & Vals(Trim(Text1(14).Text)) & " AS ���ܰ�3 " _
                     & "FROM ���� T1 " _
                    & "WHERE T1.�з��ڵ� = '" & Left(Trim(cboMtGp(0).Text), 2) & "' AND T1.�����ڵ� = '" & Trim(Text1(0).Text) & "' "
            On Error GoTo ERROR_TABLE_INSERT
            PB_adoCnnSQL.Execute strSQL
            P_adoRec.CursorLocation = adUseClient
            strSQL = "SELECT ������ڵ� FROM ����� " _
                    & "WHERE ������ڵ� <> '" & PB_regUserinfoU.UserBranchCode & "' "
            On Error GoTo ERROR_TABLE_SELECT
            P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
            If P_adoRec.RecordCount = 0 Then
               P_adoRec.Close
            Else
               Do Until P_adoRec.EOF
                  strSQL = "INSERT INTO ������� " _
                         & "SELECT '" & P_adoRec("������ڵ�") & "' AS ������ڵ�, " _
                                & "ISNULL(T1.�з��ڵ� , '') AS �з��ڵ�, ISNULL(T1.�����ڵ� , '') AS �����ڵ�, " _
                                & "0 AS �������, 0 AS �������, '' AS �����԰�����, '' AS �����������, " _
                                & "T1.��뱸�� AS ��뱸��, T1.�������� AS ��������, T1.������ڵ� AS ������ڵ�, " _
                                & "'' AS ����, '' AS �ָ���ó�ڵ�, " _
                                & "T1.�԰�ܰ�1 AS �԰�ܰ�1, T1.�԰�ܰ�2 AS �԰�ܰ�2, T1.�԰�ܰ�3 AS �԰�ܰ�3, " _
                                & "T1.���ܰ�1 AS ���ܰ�1, T1.���ܰ�2 AS ���ܰ�2, T1.���ܰ�3 AS ���ܰ�3 " _
                           & "FROM ���� T1 " _
                          & "WHERE T1.�з��ڵ� = '" & Left(Trim(cboMtGp(0).Text), 2) & "' AND T1.�����ڵ� = '" & Trim(Text1(0).Text) & "' "
                  On Error GoTo ERROR_TABLE_INSERT
                  PB_adoCnnSQL.Execute strSQL
                  P_adoRec.MoveNext
               Loop
               P_adoRec.Close
            End If
            .AddItem .Rows
            For lngC = Text1.LBound To Text1.UBound
                Select Case lngC
                       Case 0
                            .TextMatrix(.Rows - 1, 0) = Left(cboMtGp(0).Text, 2)
                            .TextMatrix(.Rows - 1, 1) = Right(Trim(cboMtGp(0).Text), Len(Trim(cboMtGp(0).Text)) - 4)
                            .TextMatrix(.Rows - 1, 2) = Trim(Text1(0).Text)
                            .TextMatrix(.Rows - 1, 4) = .TextMatrix(.Rows - 1, 0) + .TextMatrix(.Rows - 1, 2)
                            .Cell(flexcpData, .Rows - 1, 4, .Rows - 1, 4) = .TextMatrix(.Rows - 1, 4)
                            .TextMatrix(.Rows - 1, 18) = Vals(Left(cboTaxGbn.Text, 1)) '��������
                            .TextMatrix(.Rows - 1, 19) = cboTaxGbn.ListIndex
                            .TextMatrix(.Rows - 1, 20) = Right(Trim(cboTaxGbn.Text), Len(Trim(cboTaxGbn.Text)) - 3)
                            .TextMatrix(.Rows - 1, 21) = Vals(Left(cboState(1).Text, 1)) '��뱸��
                            .TextMatrix(.Rows - 1, 22) = cboState(1).ListIndex
                            .TextMatrix(.Rows - 1, 23) = Right(Trim(cboState(1).Text), Len(Trim(cboState(1).Text)) - 3)
                       Case 1 '1.ǰ��
                            .TextMatrix(.Rows - 1, 3) = Trim(Text1(1).Text)
                       Case 2 '2.���ڵ�
                            .TextMatrix(.Rows - 1, 16) = Trim(Text1(lngC).Text)
                       Case 3 To 4 '3.�԰�, 4.����
                            .TextMatrix(.Rows - 1, lngC + 2) = Trim(Text1(lngC).Text)
                       Case 5 '5.�����
                            .TextMatrix(.Rows - 1, 17) = Vals(Trim(Text1(lngC).Text))
                       Case 6 To 7 '6.�ָ���ó�ڵ�, 7.�ָ���ó��
                            .TextMatrix(.Rows - 1, lngC + 1) = Trim(Text1(lngC).Text)
                       Case 8 '8.����
                            .TextMatrix(.Rows - 1, 15) = Trim(Text1(lngC).Text)
                       Case 9 To 14
                            .TextMatrix(.Rows - 1, lngC) = Vals(Trim(Text1(lngC).Text))
                       Case Else
                End Select
            Next lngC
            If .Rows > PC_intRowCnt Then
               .ScrollBars = flexScrollBarBoth
               .TopRow = .Rows - 1
            End If
            cboMtGp(0).Enabled = False
            .Row = .Rows - 1          '�ڵ����� vsfg1_EnterCell Event �߻�
         Else                                          '���� ����
            strSQL = "UPDATE ���� SET " _
                          & "����� = '" & Trim(Text1(1).Text) & "', ���ڵ� = '" & Trim(Text1(2).Text) & "', " _
                          & "�԰� = '" & Trim(Text1(3).Text) & "', " _
                          & "���� = '" & Trim(Text1(4).Text) & "', ����� = " & Vals(Trim(Text1(5).Text)) & ", " _
                          & "�������� = " & Vals(Left(cboTaxGbn.Text, 1)) & ", " _
                          & "��뱸�� = " & Vals(Left(cboState(1).Text, 1)) & ", " _
                          & "�������� = '" & PB_regUserinfoU.UserServerDate & "', " _
                          & "������ڵ� = '" & PB_regUserinfoU.UserCode & "' " _
                    & "WHERE �з��ڵ� = '" & Left(Trim(cboMtGp(0).Text), 2) & "' " _
                      & "AND �����ڵ� = '" & Trim(Text1(0).Text) & "' "
            On Error GoTo ERROR_TABLE_UPDATE
            PB_adoCnnSQL.Execute strSQL
            strSQL = "UPDATE ������� SET " _
                          & "�ָ���ó�ڵ� = '" & Trim(Text1(6).Text) & "', �԰�ܰ�1 = " & Vals(Trim(Text1(9).Text)) & ", " _
                          & "�԰�ܰ�2 = " & Vals(Trim(Text1(10).Text)) & ", �԰�ܰ�3 = " & Vals(Trim(Text1(11).Text)) & ", " _
                          & "���ܰ�1 = " & Vals(Trim(Text1(12).Text)) & ", ���ܰ�2 = " & Vals(Trim(Text1(13).Text)) & ", " _
                          & "���ܰ�3 = " & Vals(Trim(Text1(14).Text)) & ", ���� = '" & Trim(Text1(8).Text) & "', " _
                          & "��뱸�� = " & Vals(Left(cboState(1).Text, 1)) & ", " _
                          & "�������� = '" & PB_regUserinfoU.UserServerDate & "', " _
                          & "������ڵ� = '" & PB_regUserinfoU.UserCode & "' " _
                    & "WHERE ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                      & "AND �з��ڵ� = '" & Left(Trim(cboMtGp(0).Text), 2) & "' " _
                      & "AND �����ڵ� = '" & Trim(Text1(0).Text) & "' "
            On Error GoTo ERROR_TABLE_UPDATE
            PB_adoCnnSQL.Execute strSQL
            If Len(txtSebuCodeRe.Text) > 0 Then '�ڵ庯���̸�
               '�����ڵ庯�� ���̺� �߰�
               intRetVal = MsgBox(Left(cboMtGp(0).Text, 2) + Trim(Text1(0).Text) + " �ڵ带 " _
                         & Left(cboMtGpRe.Text, 2) + Trim(txtSebuCodeRe.Text) + " �� �����Ͻðڽ��ϱ� ?", vbCritical + vbYesNo + vbDefaultButton2, "ǰ���ڵ� ����")
               If intRetVal = vbYes Then
                  '�����ð��� ������
                  P_adoRec.CursorLocation = adUseClient
                  strSQL = "SELECT CONVERT(VARCHAR(8),GETDATE(), 112) AS ��������, " _
                          & "RIGHT(CONVERT(VARCHAR(23),GETDATE(), 121),12) AS �����ð� "
                  On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
                  P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
                  strServerDate = P_adoRec("��������")
                  strServerTime = Mid(P_adoRec("�����ð�"), 1, 2) + Mid(P_adoRec("�����ð�"), 4, 2) _
                                + Mid(P_adoRec("�����ð�"), 7, 2) + Mid(P_adoRec("�����ð�"), 10)
                  P_adoRec.Close
                  strSQL = "INSERT INTO �����ڵ庯��(�������з��ڵ�, �����������ڵ�, �����Ͻ�, " _
                                                  & "�����ĺз��ڵ�, �����ļ����ڵ�, ������ڵ�) VALUES(" _
                                                  & "'" & Left(Trim(cboMtGp(0).Text), 2) & "', " _
                                                  & "'" & Trim(Text1(0).Text) & "', " _
                                                  & "'" & strServerDate + strServerTime & "', " _
                                                  & "'" & Left(Trim(cboMtGpRe.Text), 2) & "', " _
                                                  & "'" & Trim(txtSebuCodeRe.Text) & "', " _
                                                  & "'" & PB_regUserinfoU.UserCode & "' )"
                  On Error GoTo ERROR_TABLE_UPDATE
                  PB_adoCnnSQL.Execute strSQL
                  '�����ڵ庯��(��ü����)
                  strSQL = "UPDATE ���� SET " _
                                & "�з��ڵ� = '" & Left(Trim(cboMtGpRe.Text), 2) & "', " _
                                & "�����ڵ� = '" & Trim(txtSebuCodeRe.Text) & "' " _
                          & "WHERE �з��ڵ� = '" & Left(Trim(cboMtGp(0).Text), 2) & "' " _
                            & "AND �����ڵ� = '" & Trim(Text1(0).Text) & "' "
                  On Error GoTo ERROR_TABLE_UPDATE
                  PB_adoCnnSQL.Execute strSQL
                  '���ֳ��� �����ڵ� ����
                  strSQL = "UPDATE ���ֳ��� SET " _
                                & "�����ڵ� = '" & Left(Trim(cboMtGpRe.Text), 2) + Trim(txtSebuCodeRe.Text) & "' " _
                          & "WHERE �����ڵ� = '" & Left(Trim(cboMtGp(0).Text), 2) + Trim(Text1(0).Text) & "' "
                  On Error GoTo ERROR_TABLE_UPDATE
                  PB_adoCnnSQL.Execute strSQL
                  '�������� �����ڵ� ����
                  strSQL = "UPDATE �������� SET " _
                                & "�����ڵ� = '" & Left(Trim(cboMtGpRe.Text), 2) + Trim(txtSebuCodeRe.Text) & "' " _
                          & "WHERE �����ڵ� = '" & Left(Trim(cboMtGp(0).Text), 2) + Trim(Text1(0).Text) & "' "
                  On Error GoTo ERROR_TABLE_UPDATE
                  PB_adoCnnSQL.Execute strSQL
                  '�������� �����ڵ� ����
               End If
            End If
            For lngC = Text1.LBound To Text1.UBound
                Select Case lngC
                       Case 0
                            If Len(txtSebuCodeRe.Text) > 0 And intRetVal = vbYes Then '�ڵ庯���̸�
                               cboMtGp(0).ListIndex = cboMtGpRe.ListIndex
                               .TextMatrix(.Row, 1) = Mid(Trim(cboMtGpRe.Text), 5)
                               Text1(0).Text = Trim(txtSebuCodeRe.Text)
                               txtSebuCodeRe.Text = ""
                            End If
                            .TextMatrix(.Row, 0) = Left(cboMtGp(0).Text, 2)
                            .TextMatrix(.Row, 1) = Mid(Trim(cboMtGp(0).Text), 5)
                            .TextMatrix(.Row, 2) = Trim(Text1(0).Text)
                            .TextMatrix(.Row, 4) = .TextMatrix(.Row, 0) + .TextMatrix(.Row, 2)
                            .Cell(flexcpData, .Row, 4, .Row, 4) = .TextMatrix(.Row, 4)
                            .TextMatrix(.Row, 18) = Vals(Left(cboTaxGbn.Text, 1)) '��������
                            .TextMatrix(.Row, 19) = cboTaxGbn.ListIndex
                            .TextMatrix(.Row, 20) = Right(Trim(cboTaxGbn.Text), Len(Trim(cboTaxGbn.Text)) - 3)
                            .TextMatrix(.Row, 21) = Vals(Left(cboState(1).Text, 1)) '��뱸��
                            .TextMatrix(.Row, 22) = cboState(1).ListIndex
                            .TextMatrix(.Row, 23) = Right(Trim(cboState(1).Text), Len(Trim(cboState(1).Text)) - 3)
                       Case 1 '1.ǰ��
                            .TextMatrix(.Row, 3) = Trim(Text1(1).Text)
                       Case 2 '2.���ڵ�
                            .TextMatrix(.Row, 16) = Trim(Text1(lngC).Text)
                       Case 3 To 4 '3.�԰�, 4.����
                            .TextMatrix(.Row, lngC + 2) = Trim(Text1(lngC).Text)
                       Case 5 '5.�����
                            .TextMatrix(.Row, 17) = Vals(Trim(Text1(lngC).Text))
                       Case 6 To 7 '6.�ָ���ó�ڵ�, 7.�ָ���ó��
                            .TextMatrix(.Row, lngC + 1) = Trim(Text1(lngC).Text)
                       Case 8 '8.����
                            .TextMatrix(.Row, 15) = Trim(Text1(lngC).Text)
                       Case 9 To 14
                            .TextMatrix(.Row, lngC) = Vals(Trim(Text1(lngC).Text))
                       Case Else
                End Select
            Next lngC
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
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "���� �б� ����"
    Unload Me
    Exit Sub
ERROR_TABLE_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "���� �߰� ����"
    Unload Me
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "���� ���� ����"
    Unload Me
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    PB_adoCnnSQL.RollbackTrans
    MsgBox PB_varErrCode & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�Ǹ� ���� �ý��� (�������� ���� ����)"
    Unload frmLogin
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
               '        & "WHERE �з��ڵ� = '" & Left(Trim(.TextMatrix(.Row, 0)), 2) & "' " _
               '          & "AND �����ڵ� = '" & Trim(.TextMatrix(.Row, 2)) & "' " _
               '          & "AND ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' "
               'On Error GoTo ERROR_TABLE_SELECT
               'P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
               'lngCnt = P_adoRec("�ش�Ǽ�")
               'P_adoRec.Close
               If lngCnt <> 0 Then
                  MsgBox "����ü�(" & Format(lngCnt, "#,#") & "��)�� �����Ƿ� ������ �� �����ϴ�.", vbCritical, "���� ���� �Ұ�"
                  cmdDelete.Enabled = True
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
               PB_adoCnnSQL.BeginTrans
               P_adoRec.CursorLocation = adUseClient
               strSQL = "DELETE FROM ������� " _
                       & "WHERE �з��ڵ� = '" & Left(Trim(.TextMatrix(.Row, 0)), 2) & "' " _
                         & "AND �����ڵ� = '" & Trim(.TextMatrix(.Row, 2)) & "' "
               On Error GoTo ERROR_TABLE_DELETE
               PB_adoCnnSQL.Execute strSQL
               strSQL = "DELETE FROM ���� " _
                       & "WHERE �з��ڵ� = '" & Left(Trim(.TextMatrix(.Row, 0)), 2) & "' " _
                         & "AND �����ڵ� = '" & Trim(.TextMatrix(.Row, 2)) & "' "
               On Error GoTo ERROR_TABLE_DELETE
               PB_adoCnnSQL.Execute strSQL
               PB_adoCnnSQL.CommitTrans
               .RemoveItem .Row
               cboMtGp(cboMtGp.LBound).Enabled = False
               cboMtGpRe.Enabled = True
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
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "���� ���� ����"
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
    Set frm�������� = Nothing
    frmMain.SBar.Panels(4).Text = ""
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
    Text1(Text1.LBound).Enabled = False '�����ڵ� FLASE
    With vsfg1                 'Rows 1, Cols 24, RowHeightMax(Min) 300
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
         .FixedCols = 5
         .Rows = 1             'Subvsfg1_Fill����ÿ� ����
         .Cols = 24
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1000   '�з��ڵ�
         .ColWidth(1) = 1150   '�з���
         .ColWidth(2) = 2000   '�����ڵ�
         .ColWidth(3) = 2500   '�����
         .ColWidth(4) = 3000   '�з��ڵ� + �����ڵ�
         .ColWidth(5) = 2500   '�԰�
         .ColWidth(6) = 800    '����
         .ColWidth(7) = 1200   '����ó�ڵ�
         .ColWidth(8) = 3000   '����ó��
         .ColWidth(9) = 1350   '�԰�ܰ�1
         .ColWidth(10) = 1350  '�԰�ܰ�2
         .ColWidth(11) = 1350  '�԰�ܰ�3
         .ColWidth(12) = 1350  '���ܰ�1
         .ColWidth(13) = 1350  '���ܰ�2
         .ColWidth(14) = 1350  '���ܰ�3
         .ColWidth(15) = 9400  '����
         .ColWidth(16) = 3000  '���ڵ�
         .ColWidth(17) = 900   '�����
         .ColWidth(18) = 1     '��������
         .ColWidth(19) = 1     '��������ListIndex
         .ColWidth(20) = 1000  '��������
         .ColWidth(21) = 1     '��뱸��
         .ColWidth(22) = 1     '��뱸��ListIndex
         .ColWidth(23) = 1000  '��뱸��
         
         .Cell(flexcpFontBold, 0, 0, .FixedRows - 1, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "�з��ڵ�"                'H
         .TextMatrix(0, 1) = "�з���"
         .TextMatrix(0, 2) = "�ڵ�"
         .TextMatrix(0, 3) = "ǰ��"
         .TextMatrix(0, 4) = "(�з��ڵ�+�����ڵ�)�ڵ�"  'H
         .TextMatrix(0, 5) = "�԰�"
         .TextMatrix(0, 6) = "����"
         .TextMatrix(0, 7) = "����ó�ڵ�"
         .TextMatrix(0, 8) = "����ó��"
         .TextMatrix(0, 9) = "���Դܰ�1"
         .TextMatrix(0, 10) = "���Դܰ�2"
         .TextMatrix(0, 11) = "���Դܰ�3"
         .TextMatrix(0, 12) = "����ܰ�1"
         .TextMatrix(0, 13) = "����ܰ�2"
         .TextMatrix(0, 14) = "����ܰ�3"
         .TextMatrix(0, 15) = "����"
         .TextMatrix(0, 16) = "���ڵ�"
         .TextMatrix(0, 17) = "�����"
         .TextMatrix(0, 18) = "��������"       'H
         .TextMatrix(0, 19) = "��������"       'H
         .TextMatrix(0, 20) = "��������"
         .TextMatrix(0, 21) = "��뱸��"       'H
         .TextMatrix(0, 22) = "��뱸��"       'H
         .TextMatrix(0, 23) = "��뱸��"
         
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 0, 4, 18, 19, 21, 22
                         .ColHidden(lngC) = True
             End Select
         Next lngC
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 9 To 14, 17
                         .ColFormat(lngC) = "#,#.00"
             End Select
         Next lngC
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 1, 2, 3, 5, 6, 7, 8, 15, 16
                         .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 0, 4, 18 To 23
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
Dim intIndex      As Integer
    cboMtGp(cboMtGp.LBound).Enabled = False
    Text1(Text1.LBound).Enabled = False
    cboMtGpRe.Enabled = True
    txtSebuCodeRe.Enabled = True
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.�з��ڵ� AS �з��ڵ�, ISNULL(T1.�з���,'') AS �з��� " _
             & "FROM ����з� T1 " _
            & "ORDER BY T1.�з��ڵ� "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       cboMtGp(0).ListIndex = -1
       cboMtGp(1).ListIndex = -1
       cboMtGpRe.ListIndex = -1
       cmdClear.Enabled = False: cmdFind.Enabled = False
       cmdSave.Enabled = False: cmdDelete.Enabled = False
       Exit Sub
    Else
       cboMtGp(1).AddItem "00. " & "��ü"
       Do Until P_adoRec.EOF
          cboMtGp(0).AddItem P_adoRec("�з��ڵ�") & ". " & P_adoRec("�з���")
          cboMtGp(1).AddItem P_adoRec("�з��ڵ�") & ". " & P_adoRec("�з���")
          cboMtGpRe.AddItem P_adoRec("�з��ڵ�") & ". " & P_adoRec("�з���")
          P_adoRec.MoveNext
       Loop
       P_adoRec.Close
       cboMtGp(0).ListIndex = 0
       cboMtGp(1).ListIndex = 0
       cboMtGpRe.ListIndex = 0
    End If
    With cboState(0)
         .AddItem "��    ü"
         .AddItem "��    ��"
         .AddItem "���Ұ�"
         .AddItem "��    Ÿ"
         .ListIndex = 1
    End With
    With cboTaxGbn
         .AddItem "0. �� �� ��"
         .AddItem "1. ��    ��"
         .ListIndex = 1
    End With
    With cboState(1)
         .AddItem "0. ��    ��"
         .AddItem "9. ���Ұ�"
         .ListIndex = 0
    End With
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
Dim lngRRR     As Long
     
    'If (Len(P_strFindString1) + Len(P_strFindString2) + Len(P_strFindString3) + Len(P_strFindString4)) = 0 Then
    '   txtFindCD.SetFocus
    '   Exit Sub
    'End If
    If Left(cboMtGp(1).Text, 2) <> "00" And (cboMtGp(1).ListCount > 0) Then
       cboMtGp(0).ListIndex = (cboMtGp(1).ListIndex) - 1
    End If
    '�˻����� ����з�
    Select Case Left(Trim(cboMtGp(1).Text), 2)
           Case "00" '��ü
                strWhere = ""
           Case Else
                strWhere = "WHERE T1.�з��ڵ� = '" & Mid(Trim(cboMtGp(1).Text), 1, 2) & "' "
    End Select
    '�˻����� ��뱸��
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
    If (Len(P_strFindString1) + Len(P_strFindString2) + Len(P_strFindString3) + Len(P_strFindString4)) = 0 Then   '�������� ��ȸ
       strOrderBy = "ORDER BY T1.����� "
    Else
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") + "T1.�����ڵ� LIKE '%" & P_strFindString1 & "%' " _
                & "AND T1.����� LIKE '%" & P_strFindString2 & "%' AND T1.�԰� LIKE '%" & P_strFindString3 & "%' " _
                & "AND T1.���ڵ� LIKE '%" & P_strFindString4 & "%' "
       strOrderBy = "ORDER BY T1.�з��ڵ�, T1.�����ڵ� "
    End If
    '??CODE????? �ε� ǰ�� ����
    If chkCodeException.Value = 1 Then
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                + "NOT (DATALENGTH(T1.�����ڵ�) = 9 AND UPPER(SUBSTRING(T1.�����ڵ�, 1, 4)) = 'CODE' " _
                + "AND T1.�����ڵ� LIKE 'CODE_____' " _
                + "AND ISNUMERIC(SUBSTRING(T1.�����ڵ�, 5, 5)) = 1) "
    End If
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT ISNULL(T1.�з��ڵ�,'') AS �з��ڵ�, ISNULL(T2.�з���,'') AS �з���, " _
                  & "ISNULL(T1.�����ڵ�,'') AS �����ڵ�, T1.����� AS �����, " _
                  & "T1.���ڵ� AS ���ڵ�, T1.�԰� AS �԰�, T1.���� AS ����, " _
                  & "T1.����� AS �����, T1.�������� AS ��������, " _
                  & "T1.��뱸�� AS ��뱸��, T3.�ָ���ó�ڵ� AS �ָ���ó�ڵ�, ISNULL(T4.����ó��, '') AS �ָ���ó��, " _
                  & "ISNULL(T3.�԰�ܰ�1,0) AS �԰�ܰ�1, ISNULL(T3.�԰�ܰ�2,0) AS �԰�ܰ�2, ISNULL(T3.�԰�ܰ�3,0) AS �԰�ܰ�3, " _
                  & "ISNULL(T3.���ܰ�1,0) AS ���ܰ�1, ISNULL(T3.���ܰ�2,0) AS ���ܰ�2, ISNULL(T3.���ܰ�3,0) AS ���ܰ�3, " _
                  & "T3.���� AS ���� " _
             & "FROM ���� T1 " _
             & "LEFT JOIN ����з� T2 " _
                    & "ON T2.�з��ڵ� = T1.�з��ڵ� " _
             & "LEFT JOIN ������� T3 ON T3.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                   & " AND T3.�з��ڵ� = T1.�з��ڵ� AND T3.�����ڵ� = T1.�����ڵ� " _
             & "LEFT JOIN ����ó T4 ON T4.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                   & " AND T4.����ó�ڵ� = T3.�ָ���ó�ڵ� " _
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
               .TextMatrix(lngR, 0) = IIf(IsNull(P_adoRec("�з��ڵ�")), "", P_adoRec("�з��ڵ�"))
               .TextMatrix(lngR, 1) = IIf(IsNull(P_adoRec("�з���")), "", P_adoRec("�з���"))
               .TextMatrix(lngR, 2) = IIf(IsNull(P_adoRec("�����ڵ�")), "", P_adoRec("�����ڵ�"))
               .TextMatrix(lngR, 3) = IIf(IsNull(P_adoRec("�����")), "", P_adoRec("�����"))
               'FindRow ����� ����
               .TextMatrix(lngR, 4) = .TextMatrix(lngR, 0) & P_adoRec("�����ڵ�")
               .Cell(flexcpData, lngR, 4, lngR, 4) = .TextMatrix(lngR, 4)
               .TextMatrix(lngR, 5) = IIf(IsNull(P_adoRec("�԰�")), "", P_adoRec("�԰�"))
               .TextMatrix(lngR, 6) = IIf(IsNull(P_adoRec("����")), "", P_adoRec("����"))
               .TextMatrix(lngR, 7) = IIf(IsNull(P_adoRec("�ָ���ó�ڵ�")), "", P_adoRec("�ָ���ó�ڵ�"))
               .TextMatrix(lngR, 8) = IIf(IsNull(P_adoRec("�ָ���ó��")), "", P_adoRec("�ָ���ó��"))
               .TextMatrix(lngR, 9) = IIf(IsNull(P_adoRec("�԰�ܰ�1")), 0, P_adoRec("�԰�ܰ�1"))
               .TextMatrix(lngR, 10) = IIf(IsNull(P_adoRec("�԰�ܰ�2")), 0, P_adoRec("�԰�ܰ�2"))
               .TextMatrix(lngR, 11) = IIf(IsNull(P_adoRec("�԰�ܰ�3")), 0, P_adoRec("�԰�ܰ�3"))
               .TextMatrix(lngR, 12) = IIf(IsNull(P_adoRec("���ܰ�1")), 0, P_adoRec("���ܰ�1"))
               .TextMatrix(lngR, 13) = IIf(IsNull(P_adoRec("���ܰ�2")), 0, P_adoRec("���ܰ�2"))
               .TextMatrix(lngR, 14) = IIf(IsNull(P_adoRec("���ܰ�3")), 0, P_adoRec("���ܰ�3"))
               .TextMatrix(lngR, 15) = IIf(IsNull(P_adoRec("����")), "", P_adoRec("����"))
               .TextMatrix(lngR, 16) = IIf(IsNull(P_adoRec("���ڵ�")), "", P_adoRec("���ڵ�"))
               .TextMatrix(lngR, 17) = IIf(IsNull(P_adoRec("�����")), 0, P_adoRec("�����"))
               .TextMatrix(lngR, 18) = IIf(IsNull(P_adoRec("��������")), 0, P_adoRec("��������"))
               'ListIndex
               For lngRRR = 0 To cboTaxGbn.ListCount - 1
                   If .ValueMatrix(lngR, 18) = Vals(Left(cboTaxGbn.List(lngRRR), 1)) Then
                      .TextMatrix(lngR, 19) = lngRRR
                      .TextMatrix(lngR, 20) = Right(Trim(cboTaxGbn.List(lngRRR)), Len(Trim(cboTaxGbn.List(lngRRR))) - 3)
                      Exit For
                   End If
               Next lngRRR
               .TextMatrix(lngR, 21) = IIf(IsNull(P_adoRec("��뱸��")), 0, P_adoRec("��뱸��"))
               'ListIndex
               For lngRRR = 0 To cboState(1).ListCount - 1
                   If .ValueMatrix(lngR, 21) = Vals(Left(cboState(1).List(lngRRR), 1)) Then
                      .TextMatrix(lngR, 22) = lngRRR
                      .TextMatrix(lngR, 23) = Right(Trim(cboState(1).List(lngRRR)), Len(Trim(cboState(1).List(lngRRR))) - 3)
                      Exit For
                   End If
               Next lngRRR
               'If .TextMatrix(lngR, 3) = P_strFindString2 Then
               '   lngRR = lngR
               'End If
               If P_adoRec.RecordCount = 1 Then
                  lngRR = 1
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
            If lngRR <> 0 Then
               vsfg1_AfterRowColChange 0, 0, 1, 1
            End If
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
    For lngC = Text1.LBound To Text1.UBound
        Select Case lngC
               Case 0  '�����ڵ�
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If LenH(Text1(lngC).Text) < 1 Or LenH(Text1(lngC).Text) > 16 Then
                       Exit Function
                    End If
               Case 1  '�����
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (LenH(Text1(lngC).Text) > 0 And LenH(Text1(lngC).Text) <= 30) Then
                       Exit Function
                    End If
               Case 2  '���ڵ�
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (LenH(Text1(lngC).Text) <= 13) Then
                       Exit Function
                    End If
               Case 3  '�԰�
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (LenH(Text1(lngC).Text) <= 30) Then
                       Exit Function
                    End If
               Case 4  '����
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (LenH(Text1(lngC).Text) <= 20) Then
                       Exit Function
                    End If
               Case 8  '����
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (LenH(Text1(lngC).Text) <= 100) Then
                       Exit Function
                    End If
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

Dim strForAppDate          As String  '��������       (Formula)
Dim strForBranchName       As String  '������       (Formula)
Dim strForPrtDateTime      As String  '����Ͻ�       (Formula)
Dim strParGroupCode        As Integer '��ǰ���Һз�   (Parameter)
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
            strExeFile = App.Path & ".\��������.rpt"
         Else
            strExeFile = App.Path & ".\��������T.rpt"
         End If
         varRetVal = Dir(strExeFile)
         If Len(varRetVal) = 0 Then
            intRetCHK = 0
         Else
            .ReportFileName = strExeFile
            On Error GoTo ERROR_CRYSTAL_REPORTS
            '--- Formula Fields ---
            .Formulas(0) = "ForAppDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'"
            .Formulas(1) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"
            .Formulas(2) = "ForPrtDateTime = '" & strForPrtDateTime & "'"
            '--- Parameter Fields ---
            .StoredProcParam(0) = Mid(cboMtGp(1).Text, 1, 2)    '����з�(�з��ڵ�)
            If cboState(0).ListIndex < 2 Then                   '��뱸��(0.��ü, 1.����, 2.����, 3.�� ��)
               .StoredProcParam(1) = 0
            Else
               .StoredProcParam(1) = 9
            End If
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
            .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "��������"
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
 
