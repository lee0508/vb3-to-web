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
      TabIndex        =   28
      Top             =   0
      Width           =   15195
      Begin VB.CommandButton cmdZero 
         Caption         =   "������� 0 ���� ����"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   67
         Top             =   195
         Width           =   2535
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   4320
         Top             =   195
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   375
         Left            =   7980
         Picture         =   "�������.frx":0000
         Style           =   1  '�׷���
         TabIndex        =   59
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   9120
         Picture         =   "�������.frx":0963
         Style           =   1  '�׷���
         TabIndex        =   33
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   13635
         Picture         =   "�������.frx":1308
         Style           =   1  '�׷���
         TabIndex        =   31
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   390
         Left            =   12510
         Picture         =   "�������.frx":1C56
         Style           =   1  '�׷���
         TabIndex        =   32
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   11385
         Picture         =   "�������.frx":25DA
         Style           =   1  '�׷���
         TabIndex        =   25
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Height          =   390
         Left            =   10260
         Picture         =   "�������.frx":2E61
         Style           =   1  '�׷���
         TabIndex        =   30
         Top             =   195
         Width           =   1095
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
         TabIndex        =   29
         Top             =   180
         Width           =   4650
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid vsfg1 
      Height          =   6317
      Left            =   60
      TabIndex        =   26
      Top             =   3645
      Width           =   15195
      _cx             =   26802
      _cy             =   11142
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
      Height          =   2949
      Left            =   60
      TabIndex        =   27
      Top             =   630
      Width           =   15195
      Begin VB.CheckBox chkCodeException 
         Caption         =   "CODE ǰ�� ����"
         Height          =   180
         Left            =   6360
         TabIndex        =   38
         Top             =   240
         Value           =   1  'Ȯ��
         Width           =   1575
      End
      Begin VB.TextBox txtFindBarCode 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Left            =   13400
         MaxLength       =   14
         TabIndex        =   44
         Top             =   600
         Width           =   1350
      End
      Begin VB.TextBox txtBarCode 
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         Left            =   4740
         MaxLength       =   14
         TabIndex        =   74
         Top             =   585
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Index           =   20
         Left            =   1515
         TabIndex        =   24
         Top             =   2550
         Width           =   9465
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   19
         Left            =   13420
         MaxLength       =   20
         TabIndex        =   23
         Top             =   2160
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   18
         Left            =   13420
         MaxLength       =   20
         TabIndex        =   22
         Top             =   1800
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   17
         Left            =   13420
         MaxLength       =   20
         TabIndex        =   21
         Top             =   1440
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   16
         Left            =   12000
         MaxLength       =   20
         TabIndex        =   20
         Top             =   2160
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   15
         Left            =   12000
         MaxLength       =   20
         TabIndex        =   19
         Top             =   1800
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   14
         Left            =   12000
         MaxLength       =   20
         TabIndex        =   18
         Top             =   1440
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '����
         Index           =   10
         Left            =   7320
         MaxLength       =   20
         TabIndex        =   12
         Top             =   1800
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '����
         Index           =   9
         Left            =   7320
         MaxLength       =   20
         TabIndex        =   11
         Top             =   1440
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '����
         Index           =   13
         Left            =   8925
         MaxLength       =   20
         TabIndex        =   17
         Top             =   2160
         Width           =   2610
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   12
         Left            =   10150
         MaxLength       =   8
         TabIndex        =   16
         Top             =   1800
         Width           =   825
      End
      Begin VB.TextBox txtFindCD 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Left            =   11355
         MaxLength       =   18
         TabIndex        =   41
         Top             =   225
         Width           =   1365
      End
      Begin VB.TextBox txtFindSZ 
         Appearance      =   0  '���
         Height          =   285
         Left            =   10725
         MaxLength       =   30
         TabIndex        =   43
         Top             =   600
         Width           =   1845
      End
      Begin VB.TextBox txtFindNM 
         Appearance      =   0  '���
         Height          =   285
         Left            =   7320
         MaxLength       =   30
         TabIndex        =   42
         Top             =   600
         Width           =   2685
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '����
         Index           =   8
         Left            =   7320
         MaxLength       =   20
         TabIndex        =   10
         Top             =   1080
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '����
         Index           =   7
         Left            =   4740
         MaxLength       =   20
         TabIndex        =   9
         Top             =   2160
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '����
         Index           =   4
         Left            =   1515
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   5
         Left            =   4740
         MaxLength       =   20
         TabIndex        =   7
         Top             =   1440
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   11
         Left            =   7320
         MaxLength       =   20
         TabIndex        =   13
         Top             =   2160
         Width           =   1350
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
         Top             =   1440
         Width           =   1095
      End
      Begin VB.ComboBox cboState 
         Height          =   300
         Index           =   1
         Left            =   4740
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   6
         Top             =   1080
         Width           =   1350
      End
      Begin VB.ComboBox cboTaxGbn 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1515
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   5
         Top             =   2160
         Width           =   1350
      End
      Begin VB.ComboBox cboMtGp 
         Height          =   300
         Left            =   8640
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   40
         Top             =   200
         Width           =   1695
      End
      Begin VB.ComboBox cboState 
         Height          =   300
         Index           =   0
         Left            =   13680
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   39
         Top             =   200
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '����
         Index           =   6
         Left            =   4740
         MaxLength       =   20
         TabIndex        =   8
         Top             =   1800
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   1  '�Է� ���� ����
         Index           =   2
         Left            =   1515
         MaxLength       =   30
         TabIndex        =   2
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Index           =   1
         Left            =   1515
         MaxLength       =   30
         TabIndex        =   1
         Top             =   585
         Width           =   2175
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
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker dtpInputDate 
         Height          =   270
         Left            =   10150
         TabIndex        =   14
         Top             =   1080
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   57540609
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin MSComCtl2.DTPicker dtpOutputDate 
         Height          =   270
         Left            =   10150
         TabIndex        =   15
         Top             =   1440
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
         Caption         =   "���ڵ�"
         Height          =   240
         Index           =   28
         Left            =   12670
         TabIndex        =   76
         Top             =   645
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "���ڵ�"
         Height          =   240
         Index           =   27
         Left            =   3645
         TabIndex        =   75
         Top             =   645
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   23
         Left            =   315
         TabIndex        =   73
         Top             =   2600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "3."
         Height          =   240
         Index           =   22
         Left            =   11650
         TabIndex        =   72
         Top             =   2220
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "2."
         Height          =   240
         Index           =   21
         Left            =   11650
         TabIndex        =   71
         Top             =   1860
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "1."
         Height          =   240
         Index           =   20
         Left            =   11650
         TabIndex        =   70
         Top             =   1485
         Width           =   255
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
         Left            =   13560
         TabIndex        =   69
         Top             =   1125
         Width           =   1095
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
         Left            =   12120
         TabIndex        =   68
         Top             =   1125
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����̵�"
         Height          =   240
         Index           =   17
         Left            =   6405
         TabIndex        =   66
         Top             =   1845
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�������"
         Height          =   240
         Index           =   16
         Left            =   6405
         TabIndex        =   65
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         Caption         =   "[Home]"
         Height          =   240
         Index           =   14
         Left            =   11025
         TabIndex        =   64
         Top             =   1860
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����ó�ڵ�"
         Height          =   240
         Index           =   12
         Left            =   8805
         TabIndex        =   63
         Top             =   1845
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�����ڵ�"
         Height          =   240
         Index           =   11
         Left            =   10440
         TabIndex        =   62
         Top             =   285
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�԰�"
         Height          =   240
         Index           =   10
         Left            =   10080
         TabIndex        =   61
         Top             =   645
         Width           =   495
      End
      Begin VB.Line Line2 
         X1              =   6360
         X2              =   6360
         Y1              =   480
         Y2              =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "ǰ��"
         Height          =   240
         Index           =   13
         Left            =   6405
         TabIndex        =   60
         Top             =   645
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         Caption         =   "[Home]"
         Height          =   240
         Index           =   1
         Left            =   3840
         TabIndex        =   58
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
         TabIndex        =   57
         Top             =   280
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�������"
         Height          =   240
         Index           =   15
         Left            =   6405
         TabIndex        =   56
         Top             =   1125
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "���Լ���"
         Height          =   240
         Index           =   9
         Left            =   3645
         TabIndex        =   55
         Top             =   2220
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�����������"
         Height          =   240
         Index           =   8
         Left            =   8805
         TabIndex        =   54
         Top             =   1485
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�����԰�����"
         Height          =   240
         Index           =   7
         Left            =   8805
         TabIndex        =   53
         Top             =   1125
         Width           =   1215
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   15000
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�����"
         Height          =   240
         Index           =   6
         Left            =   315
         TabIndex        =   52
         Top             =   1845
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�������"
         Height          =   240
         Index           =   5
         Left            =   3660
         TabIndex        =   51
         Top             =   1485
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�������"
         Height          =   240
         Index           =   4
         Left            =   5925
         TabIndex        =   50
         Top             =   2220
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��������"
         Height          =   240
         Index           =   3
         Left            =   315
         TabIndex        =   49
         Top             =   2220
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   2
         Left            =   315
         TabIndex        =   48
         Top             =   1485
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "ǰ��"
         Height          =   240
         Index           =   36
         Left            =   315
         TabIndex        =   47
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
         Left            =   5280
         TabIndex        =   46
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�з�"
         Height          =   240
         Index           =   34
         Left            =   7725
         TabIndex        =   45
         Top             =   255
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�̿����"
         Height          =   240
         Index           =   31
         Left            =   3660
         TabIndex        =   37
         Top             =   1845
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�԰�"
         Height          =   240
         Index           =   26
         Left            =   315
         TabIndex        =   36
         Top             =   1140
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��뱸��"
         Height          =   240
         Index           =   25
         Left            =   12720
         TabIndex        =   35
         Top             =   255
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��뱸��"
         Height          =   240
         Index           =   24
         Left            =   3645
         TabIndex        =   34
         Top             =   1125
         Width           =   975
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
'                  ������帶��, �������⳻��
' ��  ��  ��  �� :
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived       As Boolean
Private P_adoRec           As New ADODB.Recordset
Private P_adoRec1          As New ADODB.Recordset
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

    frmMain.SBar.Panels(4).Text = ""
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
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "������� (�������� ���� ����)"
    Unload Me
    Exit Sub
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
'+-------------------+
'/// dtpFirstDate ///
'+-------------------+
Private Sub dtpFirstDate_KeyDown(KeyCode As Integer, Shift As Integer)
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
Dim strSQL     As String
Dim strExeFile As String
Dim varRetVal  As Variant
    If (Index = 0 And (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn)) Then '����˻�
       PB_strCallFormName = "frm�������"
       PB_strMaterialsCode = Trim(Text1(0).Text)
       PB_strMaterialsName = Trim(Text1(1).Text)
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
       (Index = 12 And (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn)) Then '����ó�˻�
       PB_strFMCCallFormName = "frm�������"
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
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "������� �б� ����"
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
                     If Len(Trim(Text1(Index).Text)) = 0 Then
                        Text1(1).Text = "": txtBarCode.Text = "": Text1(2).Text = "": Text1(3).Text = "": Text1(4).Text = ""
                        Exit Sub
                     End If
                Case 4
                     .Text = Format(Vals(Trim(.Text)), "#,0.00")
                Case 5 To 11
                     .Text = Format(Fix(Vals(Trim(.Text))), "#,0")
                Case 12
                     If Len(Trim(.Text)) = 0 Then
                        Text1(Index + 1).Text = ""
                     End If
                Case 14 To 19
                     If Vals(Trim(.Text)) < 0 Then
                        .Text = Vals(Trim(.Text)) * -1
                     End If
                     .Text = Format(Vals(Trim(.Text)), "#,0.00")
         End Select
    End With
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "������� �б� ����"
    Unload Me
    Exit Sub
End Sub

'+---------------+
'/// �˻����� ///
'+---------------+
'+-------------------------------------------+
'/// chkCodeException(CODE ǰ������ �˻�) ///
'+-------------------------------------------+
Private Sub chkCodeException_KeyDown(KeyCode As Integer, Shift As Integer)
    With chkCodeException
         If KeyCode = vbKeyReturn Then cboMtGp.SetFocus
    End With
End Sub
'+-----------------------+
'/// cboMtGp(index) ///
'+-----------------------+
Private Sub cboMtGp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       txtFindCD.SetFocus
    End If
End Sub
'+----------------------------+
'/// txtFindCode(�ڵ�˻�) ///
'+----------------------------+
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
       txtFindNM.SetFocus
    End If
End Sub
Private Sub txtFindCD_LostFocus()
    With txtFindCD
         .Text = UPPER(Trim(.Text))
    End With
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
       txtFindSZ.SetFocus
    End If
End Sub

'+--------------------------+
'/// txtFindSZ(�԰ݰ˻�) ///
'+--------------------------+
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
       txtFindBarCode.SetFocus
    End If
End Sub

'+---------------------------------+
'/// txtFindBarCode(���ڵ�˻�) ///
'+---------------------------------+
Private Sub txtFindBarCode_GotFocus()
    With txtFindBarCode
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtFindBarCode_KeyDown(KeyCode As Integer, Shift As Integer)
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
         'P_strFindString1 = Trim(.Cell(flexcpData, .Row, 0)) 'Not Used
    End With
End Sub

Private Sub vsfg1_AfterSort(ByVal Col As Long, Order As Integer)
    With vsfg1
         'If .FindRow(P_strFindString1, , 0) > 0 Then
         '   .Row = .FindRow(P_strFindString1, , 0) 'Not Used
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
               'PB_strFMCCallFormName = "frm�������"
               'PB_strMaterialsCode = .TextMatrix(.Row, 4)
               'PB_strMaterialsName = .TextMatrix(.Row, 3)
               'PB_strSupplierCode = ""
               'frm����ü��˻�.Show vbModal
            End If
         End If
    End With
End Sub
Private Sub vsfg1_DblClick()
    With vsfg1
         If .Row >= .FixedRows Then
             vsfg1_KeyDown vbKeyF1, 0  '����ü��˻�
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
         Text1(Text1.LBound).Enabled = False: Text1(Text1.LBound + 1).Enabled = True
         If .Row >= .FixedRows And OldRow <> NewRow Then
            For lngC = 0 To .Cols - 1
                Select Case lngC
                       Case 0
                            Text1(0).Text = .TextMatrix(.Row, 4) '�����ڵ�
                       Case 3
                            Text1(1).Text = .TextMatrix(.Row, 3)
                       Case 5 To 6
                            Text1(lngC - 3).Text = .TextMatrix(.Row, lngC)
                       Case 7 To 13
                            Text1(lngC - 2).Text = Format(.ValueMatrix(.Row, lngC), "#,0")
                       Case 14
                            If Len(.TextMatrix(.Row, lngC)) = 10 Then
                               dtpInputDate.Value = .TextMatrix(.Row, lngC)
                            Else
                               dtpInputDate.Value = Format("19000101", "0000-00-00")
                            End If
                       Case 15
                            If Len(.TextMatrix(.Row, lngC)) = 10 Then
                               dtpOutputDate.Value = .TextMatrix(.Row, lngC)
                            Else
                               dtpOutputDate.Value = Format("19000101", "0000-00-00")
                            End If
                       Case 16, 17 '����ó
                            Text1(lngC - 4).Text = .TextMatrix(.Row, lngC)
                       Case 18 To 23  '�ܰ�
                            Text1(lngC - 4).Text = Format(.ValueMatrix(.Row, lngC), "#,0.00")
                       Case 24  '����
                            Text1(20).Text = .TextMatrix(.Row, lngC)
                       Case 25  '���ڵ�
                            txtBarCode.Text = .TextMatrix(.Row, lngC)
                       Case 26  '�����
                            Text1(4).Text = Format(.ValueMatrix(.Row, lngC), "#,0.00")
                       Case 27  '��������
                            cboTaxGbn.ListIndex = IIf(.TextMatrix(.Row, lngC) = "�� �� ��", 0, 1)
                       Case 29  '��뱸��
                            cboState(1).ListIndex = .ValueMatrix(.Row, lngC)
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
         'Text1(Text1.LBound).Enabled = False: Text1(Text1.LBound + 1).Enabled = True
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
    Text1(Text1.LBound).SetFocus
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
Dim strSQL         As String
Dim lngR           As Long
Dim lngC           As Long
Dim blnOK          As Boolean
Dim intRetVal      As Integer
Dim curAdjustAmt   As Currency
Dim lngLogCnt      As Long
Dim strServerTime  As String
Dim strTime        As String
Dim CurInputMny    As Currency '�԰�ܰ�
Dim CurInputVat    As Currency '�԰�ΰ�
Dim CurOutPutMny   As Currency '���ܰ�
Dim CurOutPutVat   As Currency '���ΰ�

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
       intRetVal = MsgBox("�Էµ� �ڷḦ �߰��Ͻðڽ��ϱ� ?", vbQuestion + vbYesNo + vbDefaultButton1, "�ڷ� �߰�")
    Else
       intRetVal = MsgBox("������ �ڷḦ �����Ͻðڽ��ϱ� ?", vbQuestion + vbYesNo + vbDefaultButton1, "�ڷ� ����")
    End If
    If intRetVal = vbNo Then
       vsfg1.SetFocus
       Exit Sub
    End If
    cmdSave.Enabled = False
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
    If Text1(Text1.LBound).Enabled = True Then '������� �߰��� �˻�
       curAdjustAmt = Vals(Trim(Text1(11).Text))
    Else
       curAdjustAmt = Vals(Trim(Text1(11).Text)) - vsfg1.ValueMatrix(vsfg1.Row, 13)
    End If
    '�ܰ�
    CurInputMny = Vals(Trim(Text1(14).Text)): CurInputVat = Fix(Vals(Trim(Text1(14).Text)) * (PB_curVatRate))
    CurOutPutMny = Vals(Trim(Text1(17).Text)): CurOutPutVat = Fix(Vals(Trim(Text1(17).Text)) * (PB_curVatRate))
    With vsfg1
         Screen.MousePointer = vbHourglass
         P_adoRec.CursorLocation = adUseClient
         If Text1(Text1.LBound).Enabled = True Then '������� �߰��� �˻�
            strSQL = "SELECT * FROM ������� T1 " _
                    & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                      & "AND T1.�з��ڵ� = '" & Mid(Text1(0).Text, 1, 2) & "' AND T1.�����ڵ� = '" & Mid(Text1(0).Text, 3) & "' "
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
         End If
         PB_adoCnnSQL.BeginTrans
         If curAdjustAmt > 0 Then '�԰�(+)
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
                      & "'" & PB_regUserinfoU.UserBranchCode & "', '" & Mid(Trim(Text1(0).Text), 1, 2) & "', " _
                      & "'" & Mid(Trim(Text1(0).Text), 3) & "', 5, " _
                      & "'" & PB_regUserinfoU.UserClientDate & "', '" & strServerTime & "', " _
                      & "" & curAdjustAmt & ", " & CurInputMny & ", " _
                      & "" & CurInputVat & ", 0, " _
                      & "0, 0, " _
                      & "'', '', " _
                      & "'" & PB_regUserinfoU.UserClientDate & "', 0, " _
                      & "'', 0, '" & PB_regUserinfoU.UserClientDate & "', " & lngLogCnt & ", " _
                      & "0, 0, 0, '', '', 0, 0, " _
                      & "0, '" & PB_regUserinfoU.UserServerDate & "', " _
                      & "'" & PB_regUserinfoU.UserCode & "', '' ) "
            On Error GoTo ERROR_TABLE_INSERT
            PB_adoCnnSQL.Execute strSQL
         ElseIf _
            curAdjustAmt < 0 Then '���(-)
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
                      & "'" & PB_regUserinfoU.UserBranchCode & "', '" & Mid(Trim(Text1(0).Text), 1, 2) & "', " _
                      & "'" & Mid(Trim(Text1(0).Text), 3) & "', 6, " _
                      & "'" & PB_regUserinfoU.UserClientDate & "', '" & strServerTime & "', " _
                      & "0, 0, " _
                      & "0, " & (curAdjustAmt * -1) & ", " _
                      & "" & CurOutPutMny & ", " & CurOutPutVat & ", " _
                      & "'', '', " _
                      & "'" & PB_regUserinfoU.UserClientDate & "', 0, " _
                      & "'', 0, '" & PB_regUserinfoU.UserClientDate & "', " & lngLogCnt & ", " _
                      & "0, 0, 0, '', '', 0, 0, " _
                      & "0, '" & PB_regUserinfoU.UserServerDate & "', " _
                      & "'" & PB_regUserinfoU.UserCode & "', '' ) "
            On Error GoTo ERROR_TABLE_INSERT
            PB_adoCnnSQL.Execute strSQL
         End If
    End With
    With vsfg1
         If Text1(Text1.LBound).Enabled = True Then '������� �߰�
            strSQL = "INSERT INTO �������(������ڵ�, �з��ڵ�, " _
                                        & "�����ڵ�, �������, " _
                                        & "�����԰�����, �����������, " _
                                        & "��뱸��, ��������, " _
                                        & "������ڵ�, ����, �ָ���ó�ڵ�, " _
                                        & "�԰�ܰ�1, �԰�ܰ�2, �԰�ܰ�3, " _
                                        & "���ܰ�1, ���ܰ�2, ���ܰ�3 ) Values( " _
                    & "'" & PB_regUserinfoU.UserBranchCode & "', '" & Mid(Trim(Text1(0).Text), 1, 2) & "', " _
                    & "'" & Mid(Trim(Text1(0).Text), 3) & "', " & Vals(Trim(Text1(5).Text)) & ", " _
                    & "'" & IIf(DTOS(dtpInputDate.Value) = "19000101", "", DTOS(dtpInputDate.Value)) & "', " _
                    & "'" & IIf(DTOS(dtpOutputDate.Value) = "19000101", "", DTOS(dtpOutputDate.Value)) & "', " _
                    & "" & Vals(Left(cboState(1).Text, 1)) & ",'" & PB_regUserinfoU.UserServerDate & "', " _
                    & "'" & PB_regUserinfoU.UserCode & "', '" & Trim(Text1(20).Text) & "', '" & Trim(Text1(12).Text) & "', " _
                    & "" & Vals(Trim(Text1(14).Text)) & ", " & Vals(Trim(Text1(15).Text)) & ", " & Vals(Trim(Text1(15).Text)) & ", " _
                    & "" & Vals(Trim(Text1(17).Text)) & ", " & Vals(Trim(Text1(18).Text)) & ", " & Vals(Trim(Text1(19).Text)) & " ) "
            On Error GoTo ERROR_TABLE_INSERT
            .AddItem .Rows
            For lngC = Text1.LBound To Text1.UBound
                Select Case lngC
                       Case 0
                            .TextMatrix(.Rows - 1, 0) = Left(Text1(0).Text, 2)
                            For lngR = 0 To cboMtGp.ListCount - 1
                                If Left(Text1(0).Text, 2) = Left(cboMtGp.List(lngR), 2) Then
                                   .TextMatrix(.Rows - 1, 1) = Trim(Mid(cboMtGp.List(lngR), 5))
                                   Exit For
                                End If
                            Next lngR
                            .TextMatrix(.Rows - 1, 2) = Mid(Text1(0).Text, 3)
                            .TextMatrix(.Rows - 1, 4) = .TextMatrix(.Rows - 1, 0) + .TextMatrix(.Rows - 1, 2)
                            .Cell(flexcpData, .Rows - 1, 4, .Rows - 1, 4) = .TextMatrix(.Rows - 1, 4)
                       Case 1      '1.ǰ��
                            .TextMatrix(.Rows - 1, 3) = Trim(Text1(1).Text)
                       Case 2, 3   '2.�԰�, 3.����
                            .TextMatrix(.Rows - 1, lngC + 3) = Trim(Text1(lngC).Text)
                       Case 4      '4,��������
                            .TextMatrix(.Rows - 1, 14) = Format(DTOS(dtpInputDate.Value), "0000-00-00")   '�����԰�����
                            .TextMatrix(.Rows - 1, 15) = Format(DTOS(dtpOutputDate.Value), "0000-00-00")  '�����������
                            .TextMatrix(.Rows - 1, 25) = Trim(txtBarCode.Text)
                            .TextMatrix(.Rows - 1, 26) = Vals(Trim(Text1(lngC).Text))                     '�����
                            .TextMatrix(.Rows - 1, 27) = Right(Trim(cboTaxGbn.Text), Len(Trim(cboTaxGbn.Text)) - 3) '��������
                            .TextMatrix(.Rows - 1, 28) = Vals(Left(cboState(1).Text, 1))                  '��뱸��
                            .TextMatrix(.Rows - 1, 29) = cboState(1).ListIndex
                            .TextMatrix(.Rows - 1, 30) = Right(Trim(cboState(1).Text), Len(Trim(cboState(1).Text)) - 3)
                       Case 5      '5.�������
                            .TextMatrix(.Rows - 1, 7) = Vals(Trim(Text1(lngC).Text))
                       Case 9      '9.�����������
                            .TextMatrix(.Rows - 1, 11) = Vals(Trim(Text1(lngC).Text)) + curAdjustAmt
                            Text1(lngC).Text = Format(Vals(Trim(Text1(lngC).Text)) + curAdjustAmt, "#,0")
                       Case 11     '11.�������
                            .TextMatrix(.Rows - 1, 13) = Vals(Trim(Text1(lngC).Text))
                       Case 12     '12.�ָ���ó�ڵ�
                            .TextMatrix(.Rows - 1, 16) = Trim(Text1(lngC).Text)
                       Case 13     '13.�ָ���ó��
                            .TextMatrix(.Rows - 1, 17) = Trim(Text1(lngC).Text)
                       Case 14 To 19  '�ܰ�
                            .TextMatrix(.Rows - 1, lngC + 4) = Vals(Trim(Text1(lngC).Text))
                       Case 20     '20.����
                            .TextMatrix(.Rows - 1, 24) = Trim(Text1(lngC).Text)
                       Case Else
                End Select
            Next lngC
            '�̴޻�ǰǥ��
            If .ValueMatrix(.Rows - 1, 13) < .ValueMatrix(.Rows - 1, 7) Then
               .Cell(flexcpForeColor, .Rows - 1, 13, .Rows - 1, 13) = vbRed
               .Cell(flexcpFontBold, .Rows - 1, 13, .Rows - 1, 13) = True
            Else
               .Cell(flexcpForeColor, .Rows - 1, 13, .Rows - 1, 13) = vbBlack
               .Cell(flexcpFontBold, .Rows - 1, 13, .Rows - 1, 13) = False
            End If
            If .Rows > PC_intRowCnt Then
               .ScrollBars = flexScrollBarBoth
               .TopRow = .Rows - 1
            End If
            Text1(0).Enabled = False: Text1(1).Enabled = True
            .Row = .Rows - 1          '�ڵ����� vsfg1_EnterCell Event �߻�
         Else
            strSQL = "UPDATE ������� SET " _
                          & "������� = " & Vals(Trim(Text1(5).Text)) & ", " _
                          & "�����԰����� = '" & IIf(DTOS(dtpInputDate.Value) = "19000101", "", DTOS(dtpInputDate.Value)) & "'," _
                          & "����������� = '" & IIf(DTOS(dtpOutputDate.Value) = "19000101", "", DTOS(dtpOutputDate.Value)) & "', " _
                          & "��뱸�� = " & Vals(Left(cboState(1).Text, 1)) & ", " _
                          & "�������� = '" & PB_regUserinfoU.UserServerDate & "', " _
                          & "������ڵ� = '" & PB_regUserinfoU.UserCode & "', " _
                          & "���� = '" & Trim(Text1(20).Text) & "', " _
                          & "�ָ���ó�ڵ� = '" & Trim(Text1(12).Text) & "', " _
                          & "�԰�ܰ�1 = " & Vals(Trim(Text1(14).Text)) & ", " _
                          & "�԰�ܰ�2 = " & Vals(Trim(Text1(15).Text)) & ", " _
                          & "�԰�ܰ�3 = " & Vals(Trim(Text1(16).Text)) & ", " _
                          & "���ܰ�1 = " & Vals(Trim(Text1(17).Text)) & ", " _
                          & "���ܰ�2 = " & Vals(Trim(Text1(18).Text)) & ", " _
                          & "���ܰ�3 = " & Vals(Trim(Text1(19).Text)) & " " _
                    & "WHERE ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                      & "AND �з��ڵ� = '" & Mid(Text1(0).Text, 1, 2) & "' " _
                      & "AND �����ڵ� = '" & Mid(Text1(0).Text, 3) & "' "
            On Error GoTo ERROR_TABLE_UPDATE
            For lngC = Text1.LBound To Text1.UBound
                Select Case lngC
                       Case 0
                            .TextMatrix(.Row, 0) = Left(Text1(0).Text, 2)                         '�з��ڵ�
                            .TextMatrix(.Row, 2) = Mid(Text1(0).Text, 3)                          '�����ڵ�
                            .TextMatrix(.Row, 4) = .TextMatrix(.Row, 0) + .TextMatrix(.Row, 2)    '�з��ڵ� + �����ڵ�
                            .Cell(flexcpData, .Row, 4, .Row, 4) = .TextMatrix(.Row, 4)
                       Case 1     'ǰ��
                            .TextMatrix(.Row, 3) = Trim(Text1(1).Text)
                       Case 2, 3  '�԰�, ����
                            .TextMatrix(.Row, lngC + 3) = Trim(Text1(lngC).Text)
                       Case 4     '��������
                            .TextMatrix(.Row, 14) = Format(DTOS(dtpInputDate.Value), "0000-00-00")   '�����԰�����
                            .TextMatrix(.Row, 15) = Format(DTOS(dtpOutputDate.Value), "0000-00-00")  '�����������
                            .TextMatrix(.Row, 25) = Trim(txtBarCode.Text)
                            .TextMatrix(.Row, 26) = Vals(Trim(Text1(lngC).Text))                     '�����
                            .TextMatrix(.Row, 27) = Right(Trim(cboTaxGbn.Text), Len(Trim(cboTaxGbn.Text)) - 3) '��������
                            .TextMatrix(.Row, 28) = Vals(Left(cboState(1).Text, 1))                  '��뱸��
                            .TextMatrix(.Row, 29) = cboState(1).ListIndex
                            .TextMatrix(.Row, 30) = Right(Trim(cboState(1).Text), Len(Trim(cboState(1).Text)) - 3)
                       Case 5     '�������
                            .TextMatrix(.Row, 7) = Vals(Trim(Text1(lngC).Text))
                       Case 9     '�����������
                            .TextMatrix(.Row, 11) = Vals(Trim(Text1(lngC).Text)) + curAdjustAmt
                            Text1(lngC).Text = Format(Vals(Trim(Text1(lngC).Text)) + curAdjustAmt, "#,0")
                       Case 11    '�������
                            .TextMatrix(.Row, 13) = Vals(Trim(Text1(lngC).Text))
                       Case 12    '�ָ���ó�ڵ�
                            .TextMatrix(.Row, 16) = Trim(Text1(lngC).Text)
                       Case 13    '�ָ���ó��
                            .TextMatrix(.Row, 17) = Trim(Text1(lngC).Text)
                       Case 14 To 19  '�ܰ�
                            .TextMatrix(.Row, lngC + 4) = Vals(Trim(Text1(lngC).Text))
                       Case 20    '����
                            .TextMatrix(.Row, 24) = Trim(Text1(lngC).Text)
                       Case Else
                End Select
            Next lngC
            '�̴޻�ǰǥ��
            If .ValueMatrix(.Row, 13) < .ValueMatrix(.Row, 7) Then
               .Cell(flexcpForeColor, .Row, 13, .Row, 13) = vbRed
               .Cell(flexcpFontBold, .Row, 13, .Row, 13) = True
            Else
               .Cell(flexcpForeColor, .Row, 13, .Row, 13) = vbBlack
               .Cell(flexcpFontBold, .Row, 13, .Row, 13) = False
            End If
         End If
         PB_adoCnnSQL.Execute strSQL
         PB_adoCnnSQL.CommitTrans
         .SetFocus
         Screen.MousePointer = vbDefault
         
    End With
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
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "������� �б� ����"
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
                       & "WHERE ��з��ڵ� = '" & Left(Trim(.TextMatrix(.Row, 0)), 2) & "' " _
                         & "AND �ߺз��ڵ� = '" & Mid(Trim(.TextMatrix(.Row, 0)), 3, 2) & "' " _
                         & "AND �Һз��ڵ� = '" & Right(Trim(.TextMatrix(.Row, 0)), 2) & "' " _
                         & "AND ���з��ڵ� = '" & Trim(.TextMatrix(.Row, 2)) & "' "
               'On Error GoTo ERROR_TABLE_SELECT
               'P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
               'lngCnt = P_adoRec("�ش�Ǽ�")
               'P_adoRec.Close
               If lngCnt <> 0 Then
                  MsgBox "����ü�(" & Format(lngCnt, "#,#") & "��)�� �����Ƿ� ������ �� �����ϴ�.", vbCritical, "������� ���� �Ұ�"
                  cmdDelete.Enabled = True
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
               PB_adoCnnSQL.BeginTrans
               P_adoRec.CursorLocation = adUseClient
               'strSQL = "DELETE FROM ������� " _
                       & "WHERE ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                         & "AND �з��ڵ� = '" & Mid(Trim(.TextMatrix(.Row, 0)), 1, 2) & "' " _
                         & "AND �����ڵ� = '" & Trim(.TextMatrix(.Row, 2)) & "' "
               strSQL = "UPDATE ������� SET ��뱸�� = 9 " _
                       & "WHERE ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                         & "AND �з��ڵ� = '" & Mid(Trim(.TextMatrix(.Row, 0)), 1, 2) & "' " _
                         & "AND �����ڵ� = '" & Trim(.TextMatrix(.Row, 2)) & "' "
               On Error GoTo ERROR_TABLE_DELETE
               PB_adoCnnSQL.Execute strSQL
               PB_adoCnnSQL.CommitTrans
               .RemoveItem .Row
               Text1(Text1.LBound).Enabled = False: Text1(1).Enabled = True
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
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "������� �б� ����"
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_DELETE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "������� ���� ����"
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
    Set frm������� = Nothing
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
    Text1(Text1.LBound).Enabled = False      '�����ڵ� FLASE
    Text1(Text1.LBound + 1).Enabled = True   '����� FLASE
    With vsfg1                 'Rows 1, Cols 31, RowHeightMax(Min) 300
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
         .FixedCols = 6
         .Rows = 1             'Subvsfg1_Fill����ÿ� ����
         .Cols = 31
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1000   '����з�(�з��ڵ�) 'H
         .ColWidth(1) = 1000   '�з���
         .ColWidth(2) = 1550   '�����ڵ�
         .ColWidth(3) = 2700   'ǰ��
         .ColWidth(4) = 1900   '�з��ڵ�+�����ڵ�  'H
         .ColWidth(5) = 2000   '�԰�
         .ColWidth(6) = 600    '����
         .ColWidth(7) = 1000   '�������
         .ColWidth(8) = 1000   '�̿����
         .ColWidth(9) = 1000   '�԰����
         .ColWidth(10) = 1000  '������
         .ColWidth(11) = 1000  '�����������
         .ColWidth(12) = 1000  '����̵�����
         .ColWidth(13) = 1000  '�������
         .ColWidth(14) = 1400  '�����԰�����
         .ColWidth(15) = 1400  '�����������
         .ColWidth(16) = 1200  '����ó�ڵ�
         .ColWidth(17) = 3000  '����ó��
         .ColWidth(18) = 1350  '�԰�ܰ�1
         .ColWidth(19) = 1350  '�԰�ܰ�2
         .ColWidth(20) = 1350  '�԰�ܰ�3
         .ColWidth(21) = 1350  '���ܰ�1
         .ColWidth(22) = 1350  '���ܰ�2
         .ColWidth(23) = 1350  '���ܰ�3
         .ColWidth(24) = 9400  '����
         .ColWidth(25) = 3000  '���ڵ�
         .ColWidth(26) = 1000  '�����
         .ColWidth(27) = 1000  '��������
         .ColWidth(28) = 1     '��뱸��
         .ColWidth(29) = 1     '��뱸��ListIndex
         .ColWidth(30) = 1000  '��뱸��
         
         .Cell(flexcpFontBold, 0, 0, .FixedRows - 1, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "�з��ڵ�"         'H
         .TextMatrix(0, 1) = "�з���"           'H
         .TextMatrix(0, 2) = "�ڵ�"             'H
         .TextMatrix(0, 3) = "ǰ��"
         .TextMatrix(0, 4) = "�ڵ�"             'H(�з��ڵ�+�����ڵ�)
         .TextMatrix(0, 5) = "�԰�"
         .TextMatrix(0, 6) = "����"
         .TextMatrix(0, 7) = "�������"
         .TextMatrix(0, 8) = "�̿����"
         .TextMatrix(0, 9) = "���Լ���"
         .TextMatrix(0, 10) = "�������"
         .TextMatrix(0, 11) = "�������"
         .TextMatrix(0, 12) = "����̵�"
         .TextMatrix(0, 13) = "�������"
         .TextMatrix(0, 14) = "������������"
         .TextMatrix(0, 15) = "������������"
         .TextMatrix(0, 16) = "����ó�ڵ�"
         .TextMatrix(0, 17) = "����ó��"
         .TextMatrix(0, 18) = "���Դܰ�1"
         .TextMatrix(0, 19) = "���Դܰ�2"
         .TextMatrix(0, 20) = "���Դܰ�3"
         .TextMatrix(0, 21) = "����ܰ�1"
         .TextMatrix(0, 22) = "����ܰ�2"
         .TextMatrix(0, 23) = "����ܰ�3"
         .TextMatrix(0, 24) = "����"
         .TextMatrix(0, 25) = "���ڵ�"
         .TextMatrix(0, 26) = "�����"
         .TextMatrix(0, 27) = "��������"
         .TextMatrix(0, 28) = "��뱸��"        'H
         .TextMatrix(0, 29) = "��뱸��"        'H
         .TextMatrix(0, 30) = "��뱸��"
         For lngC = 7 To 13
             .ColFormat(lngC) = "#,#"
         Next lngC
         For lngC = 18 To 23
             .ColFormat(lngC) = "#,#.00"
         Next lngC
         .ColFormat(26) = "#,#.00"
         .ColHidden(0) = True: .ColHidden(4) = True
         .ColHidden(28) = True: .ColHidden(29) = True
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 0, 1, 2, 3, 4, 5, 6, 16, 17, 24, 25
                         .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 14, 15, 27, 28, 29, 30
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
    Text1(Text1.LBound).Enabled = False: Text1(Text1.LBound + 1).Enabled = True
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.�з��ڵ� AS �з��ڵ�, " _
                  & "ISNULL(T1.�з���,'') AS �з��� " _
             & "FROM ����з� T1 " _
            & "ORDER BY T1.�з��ڵ� "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       cboMtGp.ListIndex = -1
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
    'dtpInputDate.Value = Format(Left(PB_regUserinfoU.UserClientDate, 4) + "0101", "0000-00-00")
    dtpInputDate.Value = Format("19000101", "0000-00-00")
    dtpOutputDate.Value = Format("19000101", "0000-00-00")
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

    'If (Len(P_strFindString1) + Len(P_strFindString2) + Len(P_strFindString3)) = 0 Then
    '   txtFindCD.SetFocus
    '   Exit Sub
    'End If
    With vsfg1
         '�˻����� ����з�
         Select Case Left(Trim(cboMtGp.Text), 2)
                Case "00" '��ü
                     strWhere = ""
                Case Else
                     strWhere = "WHERE T1.�з��ڵ� = '" & Mid(Trim(cboMtGp.Text), 1, 2) & "' "
         End Select
         strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                  & "T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' "
    End With
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
    '�������� ��ȸ
    If (Len(P_strFindString1) + Len(P_strFindString2) + Len(P_strFindString3) + Len(P_strFindString4)) = 0 Then
       strOrderBy = "ORDER BY T1.������ڵ�, T1.�з��ڵ�, T1.�����ڵ� "
    Else
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") + " " _
                     & "T1.�����ڵ� LIKE '%" & P_strFindString1 & "%' AND T3.����� LIKE '%" & P_strFindString2 & "%' " _
                & " AND T3.�԰� LIKE '%" & P_strFindString3 & "%' AND T3.���ڵ� LIKE '%" & P_strFindString4 & "%' "
       strOrderBy = "ORDER BY T1.������ڵ�, T1.�з��ڵ�, T1.�����ڵ� "
    End If
    '??CODE????? �ε� ǰ�� ����
    If chkCodeException.Value = 1 Then
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                + "NOT (DATALENGTH(T1.�����ڵ�) = 9 AND UPPER(SUBSTRING(T1.�����ڵ�, 1, 4)) = 'CODE' " _
                + "AND T1.�����ڵ� LIKE 'CODE_____' " _
                + "AND ISNUMERIC(SUBSTRING(T1.�����ڵ�, 5, 5)) = 1) "
    End If
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.������ڵ� AS ������ڵ�, T2.������ AS ������, " _
                  & "ISNULL(T1.�з��ڵ�,'') AS �з��ڵ�, ISNULL(T4.�з���,'') AS �з���, " _
                  & "ISNULL(T1.�����ڵ�,'') AS �����ڵ�, T3.����� AS �����, " _
                  & "T3.�԰� AS �԰�, T3.���� AS ����, T3.���ڵ� AS ���ڵ�, T3.����� AS �����, T3.�������� AS ��������, " _
                  & "T1.��뱸�� AS ��뱸��, T1.������� AS �������, " _
                  & "ISNULL(T1.�ָ���ó�ڵ�,'') AS �ָ���ó�ڵ�, ISNULL(T5.����ó��, '') AS �ָ���ó��, " _
                  & "ISNULL(T1.�����԰�����,'') AS �����԰�����, ISNULL(T1.�����������,'') AS �����������, " _
                  & "ISNULL(T1.�԰�ܰ�1,0) AS �԰�ܰ�1, ISNULL(T1.�԰�ܰ�2,0) AS �԰�ܰ�2, ISNULL(T1.�԰�ܰ�3,0) AS �԰�ܰ�3," _
                  & "ISNULL(T1.���ܰ�1,0) AS ���ܰ�1, ISNULL(T1.���ܰ�2,0) AS ���ܰ�2, ISNULL(T1.���ܰ�3,0) AS ���ܰ�3," _
                  & "ISNULL(T1.����, '') AS ����, " _
                  & "(SELECT ISNULL(SUM(�԰������-��������), 0) " _
                     & "FROM ������帶�� " _
                    & "WHERE �з��ڵ� = T1.�з��ڵ� AND �����ڵ� = T1.�����ڵ� " _
                      & "AND ������ڵ� = T1.������ڵ� " _
                      & "AND ������� >= '" & Left(PB_regUserinfoU.UserClientDate, 4) + "00" & "' " _
                      & "AND ������� < '" & Left(PB_regUserinfoU.UserClientDate, 6) & "') AS �̿����, "
    strSQL = strSQL _
                  & "(SELECT ISNULL(SUM(�԰����), 0) " _
                     & "FROM �������⳻�� " _
                    & "WHERE �з��ڵ� = T1.�з��ڵ� AND �����ڵ� = T1.�����ڵ� " _
                      & "AND ������ڵ� = T1.������ڵ� AND ��뱸�� = 0 " _
                      & "AND (������� = 1) " _
                      & "AND ��������� BETWEEN '" & Left(PB_regUserinfoU.UserClientDate, 6) + "01" & "' " _
                                         & "AND '" & PB_regUserinfoU.UserClientDate & "') AS �԰����, "
    strSQL = strSQL _
                  & "(SELECT ISNULL(SUM(������), 0) " _
                     & "FROM �������⳻�� " _
                    & "WHERE �з��ڵ� = T1.�з��ڵ� AND �����ڵ� = T1.�����ڵ� " _
                      & "AND ������ڵ� = T1.������ڵ� AND ��뱸�� = 0 " _
                      & "AND (������� = 2) " _
                      & "AND ��������� BETWEEN '" & Left(PB_regUserinfoU.UserClientDate, 6) + "01" & "' " _
                                         & "AND '" & PB_regUserinfoU.UserClientDate & "') AS ������, "
    strSQL = strSQL _
                  & "(SELECT ISNULL(SUM(�԰���� - ������), 0) " _
                     & "FROM �������⳻�� " _
                    & "WHERE �з��ڵ� = T1.�з��ڵ� AND �����ڵ� = T1.�����ڵ� " _
                      & "AND ������ڵ� = T1.������ڵ� AND ��뱸�� = 0 " _
                      & "AND (������� = 5 OR ������� = 6) " _
                      & "AND ��������� BETWEEN '" & Left(PB_regUserinfoU.UserClientDate, 6) + "01" & "' " _
                                         & "AND '" & PB_regUserinfoU.UserClientDate & "') AS �����������, "
    strSQL = strSQL _
                  & "(SELECT ISNULL(SUM(�԰���� - ������), 0) " _
                     & "FROM �������⳻�� " _
                    & "WHERE �з��ڵ� = T1.�з��ڵ� AND �����ڵ� = T1.�����ڵ� " _
                      & "AND ������ڵ� = T1.������ڵ� AND ��뱸�� = 0 " _
                      & "AND (������� = 11 OR ������� = 12) " _
                      & "AND ��������� BETWEEN '" & Left(PB_regUserinfoU.UserClientDate, 6) + "01" & "' " _
                                         & "AND '" & PB_regUserinfoU.UserClientDate & "') AS ����̵����� "
    strSQL = strSQL _
             & "FROM ������� T1 " _
             & "LEFT JOIN ����� T2 ON T2.������ڵ� = T1.������ڵ� " _
             & "LEFT JOIN ���� T3 " _
                    & "ON T3.�з��ڵ� = T1.�з��ڵ� AND T3.�����ڵ� = T1.�����ڵ� " _
             & "LEFT JOIN ����з� T4 ON T4.�з��ڵ� = T1.�з��ڵ� " _
             & "LEFT JOIN ����ó T5 ON T5.������ڵ� = T1.������ڵ� AND T5.����ó�ڵ� = T1.�ָ���ó�ڵ� "
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
               .TextMatrix(lngR, 0) = IIf(IsNull(P_adoRec("�з��ڵ�")), "", P_adoRec("�з��ڵ�"))
               .TextMatrix(lngR, 1) = IIf(IsNull(P_adoRec("�з���")), "", P_adoRec("�з���"))
               .TextMatrix(lngR, 2) = IIf(IsNull(P_adoRec("�����ڵ�")), "", P_adoRec("�����ڵ�"))
               .TextMatrix(lngR, 3) = IIf(IsNull(P_adoRec("�����")), "", P_adoRec("�����"))
               'FindRow ����� ����
               .TextMatrix(lngR, 4) = .TextMatrix(lngR, 0) & P_adoRec("�����ڵ�")
               .Cell(flexcpData, lngR, 4, lngR, 4) = .TextMatrix(lngR, 4)
               .TextMatrix(lngR, 5) = IIf(IsNull(P_adoRec("�԰�")), "", P_adoRec("�԰�"))
               .TextMatrix(lngR, 6) = IIf(IsNull(P_adoRec("����")), "", P_adoRec("����"))
               .TextMatrix(lngR, 7) = IIf(IsNull(P_adoRec("�������")), 0, P_adoRec("�������"))
               .TextMatrix(lngR, 8) = IIf(IsNull(P_adoRec("�̿����")), 0, P_adoRec("�̿����"))
               .TextMatrix(lngR, 9) = IIf(IsNull(P_adoRec("�԰����")), 0, P_adoRec("�԰����"))
               .TextMatrix(lngR, 10) = IIf(IsNull(P_adoRec("������")), 0, P_adoRec("������"))
               .TextMatrix(lngR, 11) = IIf(IsNull(P_adoRec("�����������")), 0, P_adoRec("�����������"))
               .TextMatrix(lngR, 12) = IIf(IsNull(P_adoRec("����̵�����")), 0, P_adoRec("����̵�����"))
               .TextMatrix(lngR, 13) = .ValueMatrix(lngR, 8) + .ValueMatrix(lngR, 9) - .ValueMatrix(lngR, 10) _
                                     + .ValueMatrix(lngR, 11) + .ValueMatrix(lngR, 12) '�������
               If .ValueMatrix(lngR, 7) <> 0 And .ValueMatrix(lngR, 13) < .ValueMatrix(lngR, 7) Then
                  .Cell(flexcpForeColor, lngR, 13, lngR, 13) = vbRed
                  .Cell(flexcpFontBold, lngR, 13, lngR, 13) = True
               End If
               If Len(P_adoRec("�����԰�����")) = 8 Then
                  .TextMatrix(lngR, 14) = Format(P_adoRec("�����԰�����"), "0000-00-00")
               End If
               If Len(P_adoRec("�����������")) = 8 Then
                  .TextMatrix(lngR, 15) = Format(P_adoRec("�����������"), "0000-00-00")
               End If
               .TextMatrix(lngR, 16) = P_adoRec("�ָ���ó�ڵ�")
               .TextMatrix(lngR, 17) = P_adoRec("�ָ���ó��")
               .TextMatrix(lngR, 18) = IIf(IsNull(P_adoRec("�԰�ܰ�1")), 0, P_adoRec("�԰�ܰ�1"))
               .TextMatrix(lngR, 19) = IIf(IsNull(P_adoRec("�԰�ܰ�2")), 0, P_adoRec("�԰�ܰ�2"))
               .TextMatrix(lngR, 20) = IIf(IsNull(P_adoRec("�԰�ܰ�3")), 0, P_adoRec("�԰�ܰ�3"))
               .TextMatrix(lngR, 21) = IIf(IsNull(P_adoRec("���ܰ�1")), 0, P_adoRec("���ܰ�1"))
               .TextMatrix(lngR, 22) = IIf(IsNull(P_adoRec("���ܰ�2")), 0, P_adoRec("���ܰ�2"))
               .TextMatrix(lngR, 23) = IIf(IsNull(P_adoRec("���ܰ�3")), 0, P_adoRec("���ܰ�3"))
               .TextMatrix(lngR, 24) = IIf(IsNull(P_adoRec("����")), "", P_adoRec("����"))
               .TextMatrix(lngR, 25) = IIf(IsNull(P_adoRec("���ڵ�")), "", P_adoRec("���ڵ�"))
               .TextMatrix(lngR, 26) = IIf(IsNull(P_adoRec("�����")), 0, P_adoRec("�����"))
               If P_adoRec("��������") = 0 Then
                  .TextMatrix(lngR, 27) = "�� �� ��"
               Else
                  .TextMatrix(lngR, 27) = "��    ��"
               End If
               .TextMatrix(lngR, 28) = IIf(IsNull(P_adoRec("��뱸��")), "", P_adoRec("��뱸��"))
               'ListIndex
               For lngRRR = 0 To cboState(1).ListCount - 1
                   If .ValueMatrix(lngR, 28) = Vals(Left(cboState(1).List(lngRRR), 1)) Then
                      .TextMatrix(lngR, 29) = lngRRR
                      .TextMatrix(lngR, 30) = Right(Trim(cboState(1).List(lngRRR)), Len(Trim(cboState(1).List(lngRRR))) - 3)
                      Exit For
                   End If
               Next lngRRR
               If .TextMatrix(lngR, 3) = P_strFindString1 Then
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
            If lngRR <> 0 Then
               vsfg1_AfterRowColChange 0, 0, 1, 1
            End If
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
    txtBarCode.Text = ""
    For lngC = Text1.LBound To Text1.UBound
        Text1(lngC).Text = ""
    Next lngC
    dtpInputDate.Value = Format("19000101", "0000-00-00")
    dtpOutputDate.Value = Format("19000101", "0000-00-00")
End Sub

'+-------------------------+
'/// Check text1(index) ///
'+-------------------------+
Private Function FncCheckTextBox(lngC As Long, blnOK As Boolean)
    For lngC = Text1.LBound To Text1.UBound
        Select Case lngC
               Case 0  '�����ڵ�
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (Len(Text1(lngC).Text) >= 1 And Len(Text1(lngC).Text) <= 18) Then
                       Text1(lngC).Text = ""
                       Exit Function
                    End If
               Case 1  '�����
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    'If Not Len(Text1(lngC).Text) > 0 Then
                    '   Text1(lngC).Text = ""
                    '   Exit Function
                    'End If
               Case 20  '����
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (LenH(Text1(lngC).Text) <= 100) Then
                       Exit Function
                    End If
        End Select
    Next lngC
    blnOK = True
End Function

'+---------------------------+
'/// ������� 0 ���� ���� ///
'+---------------------------+
Private Sub cmdZero_Click()
Dim strSQL         As String
Dim strGroupBy     As String
Dim strHaving      As String
Dim strWhere       As String
Dim strOrderBy     As String
Dim intRetVal      As Integer
Dim lngCnt         As Long
Dim curAdjustAmt   As Currency
Dim lngLogCnt      As Long
Dim strServerTime  As String
Dim strTime        As String
Dim CurInputMny    As Currency '�԰�ܰ�
Dim CurInputVat    As Currency '�԰�ΰ�
Dim CurOutPutMny   As Currency '���ܰ�
Dim CurOutPutVat   As Currency '���ΰ�

    intRetVal = MsgBox("������� 0 ���� �ڵ� ��������Ͻðڽ��ϱ� ?", vbCritical + vbYesNo + vbDefaultButton2, "�����������")
    If intRetVal = vbNo Then
       Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    cmdZero.Enabled = False
    With vsfg1
         strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                  & "T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " 'AND T1.��뱸�� = 0 AND T3.��뱸�� = 0 " 'T3(����)
    End With
    strOrderBy = "ORDER BY T1.������ڵ�, T1.�з��ڵ�, T1.�����ڵ�  "
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.������ڵ� AS ������ڵ�, ISNULL(T1.�з��ڵ�,'') AS �з��ڵ�, ISNULL(T1.�����ڵ�,'') AS �����ڵ�, " _
                  & "T1.�԰�ܰ�1 AS �԰�ܰ�1, T1.�԰�ܰ�2 AS �԰�ܰ�2, T1.�԰�ܰ�1 AS �԰�ܰ�3, " _
                  & "T1.�԰�ܰ�1 AS ���ܰ�1, T1.���ܰ�2 AS ���ܰ�2, T1.���ܰ�3 AS ���ܰ�3, " _
                  & "(SELECT ISNULL(SUM(�԰������ - ��������), 0) " _
                     & "FROM ������帶�� " _
                    & "WHERE �з��ڵ� = T1.�з��ڵ� AND �����ڵ� = T1.�����ڵ� " _
                      & "AND ������ڵ� = T1.������ڵ� " _
                      & "AND ������� >= '" & Left(PB_regUserinfoU.UserClientDate, 4) + "00" & "' " _
                      & "AND ������� < '" & Left(PB_regUserinfoU.UserClientDate, 6) & "') AS �̿����, "
    strSQL = strSQL _
                  & "(SELECT ISNULL(SUM(�԰����), 0) " _
                     & "FROM �������⳻�� " _
                    & "WHERE �з��ڵ� = T1.�з��ڵ� AND �����ڵ� = T1.�����ڵ� " _
                      & "AND ������ڵ� = T1.������ڵ� AND ��뱸�� = 0 " _
                      & "AND (������� = 1 OR ������� = 3) " _
                      & "AND ��������� BETWEEN '" & Left(PB_regUserinfoU.UserClientDate, 6) + "01" & "' " _
                                         & "AND '" & PB_regUserinfoU.UserClientDate & "') AS �԰����, "
    strSQL = strSQL _
                  & "(SELECT ISNULL(SUM(������), 0) " _
                     & "FROM �������⳻�� " _
                    & "WHERE �з��ڵ� = T1.�з��ڵ� AND �����ڵ� = T1.�����ڵ� " _
                      & "AND ������ڵ� = T1.������ڵ� AND ��뱸�� = 0 " _
                      & "AND (������� = 2 OR ������� = 4) " _
                      & "AND ��������� BETWEEN '" & Left(PB_regUserinfoU.UserClientDate, 6) + "01" & "' " _
                                         & "AND '" & PB_regUserinfoU.UserClientDate & "') AS ������, "
    strSQL = strSQL _
                  & "(SELECT ISNULL(SUM(�԰���� - ������), 0) " _
                     & "FROM �������⳻�� " _
                    & "WHERE �з��ڵ� = T1.�з��ڵ� AND �����ڵ� = T1.�����ڵ� " _
                      & "AND ������ڵ� = T1.������ڵ� AND ��뱸�� = 0 " _
                      & "AND (������� = 5 OR ������� = 6) " _
                      & "AND ��������� BETWEEN '" & Left(PB_regUserinfoU.UserClientDate, 6) + "01" & "' " _
                                         & "AND '" & PB_regUserinfoU.UserClientDate & "') AS �����������, "
    strSQL = strSQL _
                  & "(SELECT ISNULL(SUM(�԰���� - ������), 0) " _
                     & "FROM �������⳻�� " _
                    & "WHERE �з��ڵ� = T1.�з��ڵ� AND �����ڵ� = T1.�����ڵ� " _
                      & "AND ������ڵ� = T1.������ڵ� AND ��뱸�� = 0 " _
                      & "AND (������� = 11 OR ������� = 12) " _
                      & "AND ��������� BETWEEN '" & Left(PB_regUserinfoU.UserClientDate, 6) + "01" & "' " _
                                         & "AND '" & PB_regUserinfoU.UserClientDate & "') AS ����̵����� "
    strSQL = strSQL _
             & "FROM ������� T1 " _
             & "LEFT JOIN ����� T2 ON T2.������ڵ� = T1.������ڵ� " _
             & "LEFT JOIN ���� T3 " _
                    & "ON T3.�з��ڵ� = T1.�з��ڵ� AND T3.�����ڵ� = T1.�����ڵ� " _
             & "LEFT JOIN ����з� T4 ON T4.�з��ڵ� = T1.�з��ڵ� " _
             & "LEFT JOIN ����ó T5 ON T5.������ڵ� = T1.������ڵ� AND T5.����ó�ڵ� = T1.�ָ���ó�ڵ� "
    strSQL = strSQL _
           & "" & strWhere & " " _
           & "" & strOrderBy & " "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       cmdZero.Enabled = False
       Screen.MousePointer = vbDefault
       Exit Sub
    Else
       PB_adoCnnSQL.BeginTrans
       Do Until P_adoRec.EOF
          curAdjustAmt = (P_adoRec("�̿����") + P_adoRec("�԰����") - P_adoRec("������") _
                        + P_adoRec("�����������") + P_adoRec("����̵�����")) * -1
          If curAdjustAmt <> 0 Then
             '�����ð� ���ϱ�
             Screen.MousePointer = vbHourglass
             P_adoRec1.CursorLocation = adUseClient
             strSQL = "SELECT RIGHT(CONVERT(VARCHAR(23),GETDATE(), 121),12) AS �����ð� "
             On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
             P_adoRec1.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
             strServerTime = Mid(P_adoRec1("�����ð�"), 1, 2) + Mid(P_adoRec1("�����ð�"), 4, 2) + Mid(P_adoRec1("�����ð�"), 7, 2) _
                           + Mid(P_adoRec1("�����ð�"), 10)
             P_adoRec1.Close
             strTime = strServerTime
          End If
          If curAdjustAmt <> 0 Then
             CurInputMny = P_adoRec("�԰�ܰ�1"): CurInputVat = Fix(P_adoRec("�԰�ܰ�1") * (PB_curVatRate))
             CurOutPutMny = P_adoRec("���ܰ�1"): CurOutPutVat = Fix(P_adoRec("���ܰ�1") * (PB_curVatRate))
          End If
          If curAdjustAmt > 0 Then '�԰�(+)
             '�ŷ���ȣ ���ϱ�
             P_adoRec1.CursorLocation = adUseClient
             strSQL = "spLogCounter '�������⳻��', '" & PB_regUserinfoU.UserBranchCode + PB_regUserinfoU.UserClientDate + "5" & "', " _
                                  & "0, 0, '" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' "
             On Error GoTo ERROR_STORED_PROCEDURE
             P_adoRec1.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
             lngLogCnt = P_adoRec1(0)
             P_adoRec1.Close
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
                       & "'" & PB_regUserinfoU.UserBranchCode & "', '" & P_adoRec("�з��ڵ�") & "', " _
                       & "'" & P_adoRec("�����ڵ�") & "', 5, " _
                       & "'" & PB_regUserinfoU.UserClientDate & "', '" & strServerTime & "', " _
                       & "" & curAdjustAmt & ", " & CurInputMny & ", " _
                       & "" & CurInputVat & ", 0, " _
                       & "0, 0, " _
                       & "'', '', " _
                       & "'" & PB_regUserinfoU.UserClientDate & "', 0, " _
                       & "'', 0, '" & PB_regUserinfoU.UserClientDate & "', " & lngLogCnt & ", " _
                       & "0, 0, 0, '', '', 0, 0, " _
                       & "0, '" & PB_regUserinfoU.UserServerDate & "', " _
                       & "'" & PB_regUserinfoU.UserCode & "', '' ) "
             On Error GoTo ERROR_TABLE_INSERT
             PB_adoCnnSQL.Execute strSQL
          ElseIf _
             curAdjustAmt < 0 Then '���(-)
             '�ŷ���ȣ ���ϱ�
             P_adoRec1.CursorLocation = adUseClient
             strSQL = "spLogCounter '�������⳻��', '" & PB_regUserinfoU.UserBranchCode + PB_regUserinfoU.UserClientDate + "6" & "', " _
                                  & "0, 0, '" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' "
             On Error GoTo ERROR_STORED_PROCEDURE
             P_adoRec1.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
             lngLogCnt = P_adoRec1(0)
             P_adoRec1.Close
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
                       & "'" & PB_regUserinfoU.UserBranchCode & "', '" & P_adoRec("�з��ڵ�") & "', " _
                       & "'" & P_adoRec("�����ڵ�") & "', 6, " _
                       & "'" & PB_regUserinfoU.UserClientDate & "', '" & strServerTime & "', " _
                       & "0, 0, " _
                       & "0, " & (curAdjustAmt * -1) & ", " _
                       & "" & CurOutPutMny & ", " & CurOutPutVat & ", " _
                       & "'', '', " _
                       & "'" & PB_regUserinfoU.UserClientDate & "', 0, " _
                       & "'', 0, '" & PB_regUserinfoU.UserClientDate & "', " & lngLogCnt & ", " _
                       & "0, 0, 0, '', '', 0, 0, " _
                       & "0, '" & PB_regUserinfoU.UserServerDate & "', " _
                       & "'" & PB_regUserinfoU.UserCode & "', '' ) "
             On Error GoTo ERROR_TABLE_INSERT
             PB_adoCnnSQL.Execute strSQL
          End If
          P_adoRec.MoveNext
       Loop
       P_adoRec.Close
       PB_adoCnnSQL.CommitTrans
    End If
    cmdZero.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "(�������� ���� ����)"
    Unload Me
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "������� �б� ����"
    Unload Me
    Exit Sub
ERROR_TABLE_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�ڵ�������� ����"
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
            strExeFile = App.Path & ".\�������.rpt"
         Else
            strExeFile = App.Path & ".\�������T.rpt"
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
            .Formulas(3) = "ForAppDate = '�������� : ' & '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'"   '��������
            'DECLARE @ParAppPgDate VarChar(8), @ParMtGroupCode VarChar(6),  @ParMtName VarChar(20), @ParStateCode int
            '--- Parameter Fields ---
            '���α׷���������
            .StoredProcParam(0) = PB_regUserinfoU.UserClientDate
            '����з�(�з��ڵ�)
            If Mid(cboMtGp.Text, 1, 2) = "00" Then
               .StoredProcParam(1) = " "
            Else
               .StoredProcParam(1) = Mid(cboMtGp.Text, 1, 2)
            End If
            '�����
            If Len(txtFindNM.Text) = 0 Then
               .StoredProcParam(2) = " "
            Else
               .StoredProcParam(2) = Trim(txtFindNM.Text)
            End If
            '��뱸��(0.��ü, 1.����, 2.����, 3.�� ��)
            If cboState(0).ListIndex < 2 Then
               .StoredProcParam(3) = 0
            Else
               .StoredProcParam(3) = 9
            End If
            .StoredProcParam(4) = PB_regUserinfoU.UserBranchCode
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
            .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "�������"
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

