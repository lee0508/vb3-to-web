VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm��������� 
   BorderStyle     =   0  '����
   Caption         =   "���������"
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
   Icon            =   "���������.frx":0000
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
      TabIndex        =   48
      Top             =   0
      Width           =   15195
      Begin VB.CommandButton cmdPrint 
         Height          =   375
         Left            =   7980
         Picture         =   "���������.frx":0CCA
         Style           =   1  '�׷���
         TabIndex        =   57
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   9120
         Picture         =   "���������.frx":162D
         Style           =   1  '�׷���
         TabIndex        =   54
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   13635
         Picture         =   "���������.frx":1FD2
         Style           =   1  '�׷���
         TabIndex        =   52
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   390
         Left            =   12510
         Picture         =   "���������.frx":2920
         Style           =   1  '�׷���
         TabIndex        =   51
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   11385
         Picture         =   "���������.frx":32A4
         Style           =   1  '�׷���
         TabIndex        =   32
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Height          =   390
         Left            =   10260
         Picture         =   "���������.frx":3B2B
         Style           =   1  '�׷���
         TabIndex        =   50
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00008000&
         BorderStyle     =   1  '���� ����
         Caption         =   "�� �� �� �� �� �� ��"
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
         TabIndex        =   49
         Top             =   180
         Width           =   4650
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid vsfg1 
      Height          =   4935
      Left            =   60
      TabIndex        =   47
      Top             =   5100
      Width           =   15195
      _cx             =   26802
      _cy             =   8705
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
      Height          =   4395
      Left            =   60
      TabIndex        =   33
      Top             =   630
      Width           =   15195
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   30
         Left            =   11520
         MaxLength       =   1
         TabIndex        =   31
         Top             =   3960
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   29
         Left            =   6030
         MaxLength       =   1
         TabIndex        =   30
         Top             =   3960
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   28
         Left            =   14160
         MaxLength       =   3
         TabIndex        =   29
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   27
         Left            =   11505
         MaxLength       =   3
         TabIndex        =   28
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   26
         Left            =   8820
         MaxLength       =   3
         TabIndex        =   27
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   25
         Left            =   6030
         MaxLength       =   3
         TabIndex        =   26
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   24
         Left            =   3000
         MaxLength       =   3
         TabIndex        =   25
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   23
         Left            =   11505
         MaxLength       =   6
         TabIndex        =   24
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   22
         Left            =   8820
         MaxLength       =   6
         TabIndex        =   23
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   21
         Left            =   6030
         MaxLength       =   6
         TabIndex        =   22
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   20
         Left            =   3000
         MaxLength       =   6
         TabIndex        =   21
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   19
         Left            =   1275
         TabIndex        =   20
         Top             =   2520
         Width           =   5175
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   18
         Left            =   9300
         TabIndex        =   19
         Top             =   2160
         Width           =   5175
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   17
         Left            =   1275
         TabIndex        =   18
         Top             =   2160
         Width           =   5175
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   16
         Left            =   12720
         MaxLength       =   1
         TabIndex        =   17
         Top             =   1305
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   15
         Left            =   12720
         MaxLength       =   1
         TabIndex        =   16
         Top             =   945
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   14
         Left            =   12720
         MaxLength       =   7
         TabIndex        =   15
         Top             =   600
         Width           =   615
      End
      Begin MSComCtl2.DTPicker dtpOpenDate 
         Height          =   270
         Left            =   12720
         TabIndex        =   14
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   19857409
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Index           =   4
         Left            =   5910
         MaxLength       =   20
         TabIndex        =   4
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Index           =   13
         Left            =   9300
         TabIndex        =   13
         Top             =   1665
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   11
         Left            =   9315
         MaxLength       =   7
         TabIndex        =   11
         Top             =   1305
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Index           =   12
         Left            =   1275
         TabIndex        =   12
         Top             =   1665
         Width           =   6945
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   10
         Left            =   9300
         MaxLength       =   1
         TabIndex        =   10
         Top             =   945
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   9
         Left            =   9300
         MaxLength       =   14
         TabIndex        =   9
         Top             =   585
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   8
         Left            =   9300
         MaxLength       =   20
         TabIndex        =   8
         Top             =   233
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Index           =   7
         Left            =   5910
         TabIndex        =   7
         Top             =   1305
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Index           =   6
         Left            =   5910
         TabIndex        =   6
         Top             =   945
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   5
         Left            =   5910
         MaxLength       =   14
         TabIndex        =   5
         Top             =   585
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Index           =   3
         Left            =   1275
         MaxLength       =   14
         TabIndex        =   3
         Top             =   1305
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Index           =   2
         Left            =   1275
         MaxLength       =   50
         TabIndex        =   2
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
         TabIndex        =   1
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
         MaxLength       =   4
         TabIndex        =   0
         Top             =   225
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   ")"
         Height          =   255
         Index           =   5
         Left            =   15000
         TabIndex        =   86
         Top             =   4020
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "���� ����ܰ� �ڵ�����"
         Height          =   240
         Index           =   38
         Left            =   9240
         TabIndex        =   85
         Top             =   4020
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "���� ���Դܰ� �ڵ�����"
         Height          =   240
         Index           =   37
         Left            =   3720
         TabIndex        =   84
         Top             =   4020
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "("
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   83
         Top             =   4020
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�ܰ��ڵ�����"
         Height          =   240
         Index           =   36
         Left            =   75
         TabIndex        =   82
         ToolTipText     =   "�ŷ�����/���ݰ�꼭"
         Top             =   4020
         Width           =   1095
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   120
         X2              =   15015
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Label Label3 
         Caption         =   ")"
         Height          =   255
         Index           =   3
         Left            =   14985
         TabIndex        =   81
         Top             =   3555
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   ")"
         Height          =   255
         Index           =   2
         Left            =   12360
         TabIndex        =   80
         Top             =   3060
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "���Ÿ�Ա���"
         Height          =   240
         Index           =   35
         Left            =   1320
         TabIndex        =   79
         Top             =   3555
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "("
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   78
         Top             =   3555
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "("
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   77
         Top             =   3060
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "���ʸ������"
         Height          =   240
         Index           =   34
         Left            =   75
         TabIndex        =   76
         Top             =   3060
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "���ݰ�꼭���ʸ���"
         Height          =   240
         Index           =   33
         Left            =   12360
         TabIndex        =   75
         Top             =   3555
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "���ݰ�꼭��ܸ���"
         Height          =   240
         Index           =   32
         Left            =   9720
         TabIndex        =   74
         Top             =   3555
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�ŷ��������ʸ���"
         Height          =   240
         Index           =   31
         Left            =   6900
         TabIndex        =   73
         Top             =   3555
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�ŷ�������ܸ���"
         Height          =   240
         Index           =   30
         Left            =   4200
         TabIndex        =   72
         Top             =   3555
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�ŷ�/��꼭"
         Height          =   240
         Index           =   29
         Left            =   75
         TabIndex        =   71
         ToolTipText     =   "�ŷ�����/���ݰ�꼭"
         Top             =   3550
         Width           =   1095
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   120
         X2              =   15015
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "ȸ�� ���ʸ������"
         Height          =   240
         Index           =   28
         Left            =   9840
         TabIndex        =   70
         Top             =   3060
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�̼��� ���ʸ������"
         Height          =   240
         Index           =   27
         Left            =   6900
         TabIndex        =   69
         Top             =   3060
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�����ޱ� ���ʸ������"
         Height          =   240
         Index           =   26
         Left            =   3960
         TabIndex        =   68
         Top             =   3060
         Width           =   1935
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   120
         X2              =   15015
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "���� ���ʸ������"
         Height          =   240
         Index           =   25
         Left            =   1320
         TabIndex        =   67
         Top             =   3060
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�������"
         Height          =   240
         Index           =   24
         Left            =   75
         TabIndex        =   66
         ToolTipText     =   "������ ��������Դϴ�."
         Top             =   2600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "Ȩ�������ּ�"
         Height          =   240
         Index           =   23
         Left            =   8100
         TabIndex        =   65
         Top             =   2205
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�̸����ּ�"
         Height          =   240
         Index           =   22
         Left            =   75
         TabIndex        =   64
         Top             =   2200
         Width           =   1095
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   120
         X2              =   15015
         Y1              =   2055
         Y2              =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "(1.��ǥ, 2.��꼭)"
         Height          =   240
         Index           =   21
         Left            =   13440
         TabIndex        =   63
         Top             =   1365
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "(1.��ǥ, 2.��꼭)"
         Height          =   240
         Index           =   20
         Left            =   13440
         TabIndex        =   62
         Top             =   1005
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�̼��ݹ߻�"
         Height          =   240
         Index           =   19
         Left            =   11475
         TabIndex        =   61
         ToolTipText     =   "(1.��ǥ, 2.��꼭)"
         Top             =   1365
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�����ޱݹ߻�"
         Height          =   240
         Index           =   17
         Left            =   11475
         TabIndex        =   60
         ToolTipText     =   "(1.��ǥ, 2.��꼭)"
         Top             =   1005
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "(0.����)"
         Height          =   240
         Index           =   18
         Left            =   10200
         TabIndex        =   59
         Top             =   1005
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�ΰ�����(%)"
         Height          =   240
         Index           =   16
         Left            =   11475
         TabIndex        =   58
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "[Home]"
         Height          =   240
         Index           =   15
         Left            =   10440
         TabIndex        =   56
         ToolTipText     =   "300"
         Top             =   1365
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��������"
         Height          =   240
         Index           =   14
         Left            =   11475
         TabIndex        =   55
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��뱸��"
         Height          =   240
         Index           =   9
         Left            =   8100
         TabIndex        =   53
         ToolTipText     =   "0.����, ��Ÿ.�ÿ�Ұ�"
         Top             =   1005
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   13
         Left            =   8100
         TabIndex        =   46
         Top             =   1725
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�ѽ���ȣ"
         Height          =   240
         Index           =   12
         Left            =   8100
         TabIndex        =   45
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�����ȣ"
         Height          =   240
         Index           =   11
         Left            =   8100
         TabIndex        =   44
         Top             =   1365
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�ּ�"
         Height          =   240
         Index           =   10
         Left            =   75
         TabIndex        =   43
         Top             =   1725
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��ȭ��ȣ"
         Height          =   240
         Index           =   8
         Left            =   8100
         TabIndex        =   42
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   7
         Left            =   4710
         TabIndex        =   41
         Top             =   1365
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   6
         Left            =   4710
         TabIndex        =   40
         Top             =   1005
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�ֹι�ȣ"
         Height          =   240
         Index           =   5
         Left            =   4710
         TabIndex        =   39
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��ǥ�ڸ�"
         Height          =   240
         Index           =   4
         Left            =   4710
         TabIndex        =   38
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "���ι�ȣ"
         Height          =   240
         Index           =   3
         Left            =   75
         TabIndex        =   37
         Top             =   1365
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����ڹ�ȣ"
         Height          =   240
         Index           =   2
         Left            =   75
         TabIndex        =   36
         Top             =   1005
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "������"
         Height          =   240
         Index           =   1
         Left            =   75
         TabIndex        =   35
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "������ڵ�"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   75
         TabIndex        =   34
         Top             =   285
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frm���������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' ���α׷� �� �� : ���������
' ���� Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' ������ Table   : �����
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
       Subvsfg1_INIT
       Subvsfg1_FILL
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
       dtpOpenDate.Value = Format("19000101", "0000-00-00")
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
    If (Index = 11 And (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn)) Then
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
       Exit Sub
    End If
    If KeyCode = vbKeyReturn Then
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
                     .Text = Format(Val(Trim(.Text)), "00")
                     If Trim(Text1(Index).Text) = "00" Then
                        Text1(Index).Text = ""
                     End If
                     If Text1(Index).Enabled = True Then
                        P_adoRec.CursorLocation = adUseClient
                        strSQL = "SELECT * FROM ����� " _
                                & "WHERE ������ڵ� = '" & Trim(.Text) & "' "
                        On Error GoTo ERROR_TABLE_SELECT
                        P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
                        If P_adoRec.RecordCount <> 0 Then
                           P_adoRec.Close
                           .Text = ""
                           .SetFocus
                           Exit Sub
                        End If
                        P_adoRec.Close
                     End If
                Case 2 '����ڵ�Ϲ�ȣ
                     If Len(Trim(.Text)) = 10 Then
                        .Text = Format(Trim(.Text), "###-##-#####")
                     End If
                Case 5 '�ֹε�Ϲ�ȣ
                     If Len(Trim(.Text)) = 13 Then
                        .Text = Format(Trim(.Text), "######-#######")
                     End If
                Case 11 '�����ȣ
                     If Len(Trim(.Text)) = 6 Then
                        .Text = Format(Trim(.Text), "###-###")
                     End If
                Case 14 '�ΰ�����
                     .Text = Format(Vals(Trim(.Text)), "#00.00")
                Case 15, 16 '15.�����ޱݹ߻�, 16.�̼��ݹ߻�
                     If Len(.Text) = 0 Or Trim(.Text) = "0" Then .Text = "2"
                     .Text = Format(Val(Trim(.Text)), "0")
                Case 24     '24.���Ÿ�Ա���
                     .Text = Fix(Val(Trim(.Text)))
                     If Len(.Text) = 0 Or Trim(.Text) = "0" Then .Text = "1"
                     .Text = Fix(Val(Trim(.Text)))
                Case 25 To 28   '���, ����
                     .Text = Fix(Val(Trim(.Text)))
                Case 29 To 30   '�ܰ��ڵ�����
                     If Val(Trim(.Text)) <> 1 Then
                        .Text = "0"
                     End If
                     .Text = Fix(Val(Trim(.Text)))
         End Select
    End With
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����� �б� ����"
    Unload Me
    Exit Sub
End Sub

Private Sub dtpOpenDate_KeyDown(KeyCode As Integer, Shift As Integer)
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
            '.Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlack '.ForeColorSel
            '.Cell(flexcpForeColor, 0, .MouseCol, 0, .MouseCol) = vbRed
            'strData = Trim(.Cell(flexcpData, .Row, 0))
            'If P_intButton = 1 Then
            '   .Sort = flexSortGenericAscending
            'Else
            '   .Sort = flexSortGenericDescending
            'End If
            'If .FindRow(strData, , 0) > 0 Then
            '   .Row = .FindRow(strData, , 0)
            'End If
            'If PC_intRowCnt < .Rows Then
            '   .TopRow = .Row
            'End If
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
                       Case Is <= 10
                            Text1(lngC).Text = .TextMatrix(.Row, lngC)
                       Case 12
                            If Len(.TextMatrix(.Row, lngC)) = 10 Then
                               dtpOpenDate.Value = .TextMatrix(.Row, lngC)
                            Else
                               dtpOpenDate.Value = Format("19000101", "0000-00-00")
                            End If
                       Case 13 To 15
                            Text1(lngC - 2).Text = .TextMatrix(.Row, lngC)
                       Case 17
                            Text1(14).Text = Format(.ValueMatrix(.Row, lngC), "#00.00")
                       Case 18, 19
                            Text1(lngC - 3).Text = Format(.ValueMatrix(.Row, lngC), "0")
                       Case 20 To 26
                            Text1(lngC - 3).Text = .TextMatrix(.Row, lngC)
                       Case 27 To 33
                            Text1(lngC - 3).Text = .TextMatrix(.Row, lngC)
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
         If .Row = 0 Then
            Text1(Text1.LBound).Enabled = True
         Else
            Text1(Text1.LBound).Enabled = False
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
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����� �б� ����"
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
    SubClearText
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
         If Text1(Text1.LBound).Enabled = True Then '����� �߰��� �˻�
            strSQL = "SELECT * FROM ����� " _
                    & "WHERE ������ڵ� = '" & Trim(Text1(Text1.LBound).Text) & "' "
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
         If Text1(Text1.LBound).Enabled = True Then '����� �߰�
            strSQL = "INSERT INTO �����(������ڵ�, ������, ����ڹ�ȣ," _
                                       & "���ι�ȣ, ��ǥ�ڸ�, ��ǥ���ֹι�ȣ," _
                                       & "��������," _
                                       & "�����ȣ, �ּ�, ����," _
                                       & "����, ����, ��ȭ��ȣ," _
                                       & "�ѽ���ȣ, �ΰ�����, �����ޱݹ߻�����, �̼��ݹ߻�����, " _
                                       & "��뱸��, ��������, ������ڵ�, " _
                                       & "�̸����ּ�, Ȩ�������ּ�, �������, " _
                                       & "������ʸ������,�����ޱݱ��ʸ������,�̼��ݱ��ʸ������, ȸ����ʸ������, " _
                                       & "���Ÿ�Ա���, �ŷ�������ܸ���, �ŷ��������ʸ���, " _
                                       & "���ݰ�꼭��ܸ���, ���ݰ�꼭���ʸ���, " _
                                       & "�����԰�ܰ��ڵ����ű���, �����԰�ܰ��ڵ����ű��� ) VALUES( " _
                    & "'" & Trim(Text1(0).Text) & "','" & Trim(Text1(1).Text) & "','" & Trim(Text1(2).Text) & "', " _
                    & "'" & Trim(Text1(3).Text) & "','" & Trim(Text1(4).Text) & "','" & Trim(Text1(5).Text) & "', " _
                    & "'" & IIf(DTOS(dtpOpenDate.Value) = "19000101", "", DTOS(dtpOpenDate.Value)) & "', " _
                    & "'" & Trim(Text1(11).Text) & "','" & Trim(Text1(12).Text) & "','" & Trim(Text1(13).Text) & "', " _
                    & "'" & Trim(Text1(6).Text) & "','" & Trim(Text1(7).Text) & "','" & Trim(Text1(8).Text) & "', " _
                    & "'" & Trim(Text1(9).Text) & "', " & Vals(Text1(14).Text) & ", " & Vals(Text1(15).Text) & ", " & Vals(Text1(16).Text) & ", " _
                    & "" & Val(Text1(10).Text) & ", '" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "', " _
                    & "'" & Trim(Text1(17).Text) & "', '" & Trim(Text1(18).Text) & "', '" & Trim(Text1(19).Text) & "', " _
                    & "'" & Trim(Text1(20).Text) & "', '" & Trim(Text1(21).Text) & "', " _
                    & "'" & Trim(Text1(22).Text) & "', '" & Trim(Text1(23).Text) & "', " _
                    & "" & Vals(Trim(Text1(24).Text)) & ", " & Vals(Trim(Text1(25).Text)) & ", " & Vals(Trim(Text1(26).Text)) & ", " _
                    & "" & Vals(Trim(Text1(27).Text)) & ", " & Vals(Trim(Text1(28).Text)) & ", " _
                    & "" & Vals(Trim(Text1(29).Text)) & ", " & Vals(Trim(Text1(30).Text)) & " ) "
            On Error GoTo ERROR_TABLE_INSERT
            PB_adoCnnSQL.Execute strSQL
            '����������(����->�������)
            strSQL = "INSERT INTO ������� " _
                        & "SELECT '" & Trim(Text1(0).Text) & "' AS ������ڵ�, " _
                               & "ISNULL(T1.�з��ڵ� , '') AS �з��ڵ�, ISNULL(T1.�����ڵ� , '') AS �����ڵ�, " _
                               & "0 AS �������, 0 AS �������, '' AS �����԰�����, '' AS �����������, " _
                               & "0 AS ��뱸��, " _
                               & "'" & PB_regUserinfoU.UserServerDate & "' AS ��������, '" & PB_regUserinfoU.UserCode & "' AS ������ڵ�, " _
                               & "'' AS ����, '' AS �ָ���ó�ڵ�, " _
                               & "ISNULL(T2.�԰�ܰ�1, 0) AS �԰�ܰ�1, ISNULL(T2.�԰�ܰ�2, 0) AS �԰�ܰ�2, " _
                               & "ISNULL(T2.�԰�ܰ�3, 0) AS �԰�ܰ�3, ISNULL(T2.���ܰ�1, 0) AS ���ܰ�1, " _
                               & "ISNULL(T2.���ܰ�3, 0) AS ���ܰ�2, ISNULL(T2.���ܰ�3, 0) AS ���ܰ�3 " _
                          & "FROM ���� T1 " _
                         & "INNER JOIN ������� T2 ON T2.������ڵ� = '01' AND T2.�з��ڵ� = T1.�з��ڵ� AND T2.�����ڵ� = T1.�����ڵ� " _
                         & "ORDER BY T1.�з��ڵ�, T1.�����ڵ� "
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
                       Case 11 To 13
                            .TextMatrix(.Rows - 1, lngC + 2) = Text1(lngC).Text
                       Case 14
                            .TextMatrix(.Rows - 1, 17) = Vals(Text1(14).Text)
                       Case 15, 16
                            .TextMatrix(.Rows - 1, lngC + 3) = Val(Text1(lngC).Text)
                       Case 17 To 23, 24 To 28
                            .TextMatrix(.Rows - 1, lngC + 3) = Trim(Text1(lngC).Text)
                       Case 29 To 30
                            .TextMatrix(.Rows - 1, lngC + 3) = Trim(Text1(lngC).Text)
                       Case Else
                End Select
            Next lngC
            .TextMatrix(.Rows - 1, 12) = Format(IIf(DTOS(dtpOpenDate.Value) = "19000101", "", DTOS(dtpOpenDate.Value)), "0000-00-00")
            .TextMatrix(.Rows - 1, 16) = Trim(.TextMatrix(.Rows - 1, 14)) & Space(1) & Trim(.TextMatrix(.Rows - 1, 15))
            If .Rows > PC_intRowCnt Then
               .ScrollBars = flexScrollBarBoth
               .TopRow = .Rows - 1
            End If
            Text1(Text1.LBound).Enabled = False
            .Row = .Rows - 1          '�ڵ����� vsfg1_EnterCell Event �߻�
         Else                                          '����� ����
            strSQL = "UPDATE ����� SET " _
                          & "������ = '" & Trim(Text1(1).Text) & "', " _
                          & "����ڹ�ȣ = '" & Trim(Text1(2).Text) & "', " _
                          & "���ι�ȣ = '" & Trim(Text1(3).Text) & "', " _
                          & "��ǥ�ڸ� = '" & Trim(Text1(4).Text) & "', " _
                          & "��ǥ���ֹι�ȣ = '" & Trim(Text1(5).Text) & "', " _
                          & "�������� = '" & IIf(DTOS(dtpOpenDate.Value) = "19000101", "", DTOS(dtpOpenDate.Value)) & "', " _
                          & "�����ȣ = '" & Trim(Text1(11).Text) & "', " _
                          & "�ּ� = '" & Trim(Text1(12).Text) & "', ���� = '" & Trim(Text1(13).Text) & "', " _
                          & "���� = '" & Trim(Text1(6).Text) & "', ���� = '" & Trim(Text1(7).Text) & "', " _
                          & "��ȭ��ȣ = '" & Trim(Text1(8).Text) & "', �ѽ���ȣ = '" & Trim(Text1(9).Text) & "', " _
                          & "�ΰ����� = " & Vals(Text1(14).Text) & ", " _
                          & "�����ޱݹ߻����� = " & Val(Text1(15).Text) & ", �̼��ݹ߻����� = " & Val(Text1(16).Text) & ", " _
                          & "��뱸�� = " & Val(Text1(10).Text) & ", " _
                          & "�������� = '" & PB_regUserinfoU.UserServerDate & "', ������ڵ� = '" & PB_regUserinfoU.UserCode & "', " _
                          & "�̸����ּ� = '" & Trim(Text1(17).Text) & "', Ȩ�������ּ� = '" & Trim(Text1(18).Text) & "', " _
                          & "������� = '" & Trim(Text1(19).Text) & "', ������ʸ������ = '" & Trim(Text1(20).Text) & "', " _
                          & "�����ޱݱ��ʸ������ = '" & Trim(Text1(21).Text) & "', " _
                          & "�̼��ݱ��ʸ������ = '" & Trim(Text1(22).Text) & "', ȸ����ʸ������ = '" & Trim(Text1(23).Text) & "', " _
                          & "���Ÿ�Ա��� = " & Vals(Trim(Text1(24).Text)) & ", " _
                          & "�ŷ�������ܸ��� = " & Vals(Trim(Text1(25).Text)) & ",�ŷ��������ʸ��� = " & Vals(Trim(Text1(26).Text)) & ",  " _
                          & "���ݰ�꼭��ܸ��� = " & Vals(Trim(Text1(27).Text)) & ",���ݰ�꼭���ʸ��� = " & Vals(Trim(Text1(28).Text)) & ", " _
                          & "�����԰�ܰ��ڵ����ű��� = " & Vals(Trim(Text1(29).Text)) & ", " _
                          & "�������ܰ��ڵ����ű��� = " & Vals(Trim(Text1(30).Text)) & " " _
                    & "WHERE ������ڵ� = '" & Trim(Text1(Text1.LBound).Text) & "' "
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
                       Case 11 To 13
                            .TextMatrix(.Row, lngC + 2) = Text1(lngC).Text
                       Case 14
                            .TextMatrix(.Row, 17) = Vals(Trim(Text1(lngC).Text))
                       Case 15, 16
                            .TextMatrix(.Row, lngC + 3) = Val(Trim(Text1(lngC).Text))
                       Case 17 To 23, 24 To 30
                            .TextMatrix(.Row, lngC + 3) = Trim(Text1(lngC).Text)
                       Case Else
                End Select
            Next lngC
            .TextMatrix(.Row, 12) = Format(IIf(DTOS(dtpOpenDate.Value) = "19000101", "", DTOS(dtpOpenDate.Value)), "0000-00-00")
            .TextMatrix(.Row, 16) = Trim(.TextMatrix(.Row, 14)) & Space(1) & Trim(.TextMatrix(.Row, 15))
         End If
         PB_adoCnnSQL.CommitTrans
         
         '+--------+
         ' ��������
         '+--------+
         '(�ΰ����� : ���������̰� �ΰ������� ����Ǹ�)
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And _
            Vals(PB_curVatRate) <> Vals(Text1(14).Text) Then
            PB_curVatRate = Vals(Text1(14).Text) / 100
         End If
         '(�����԰�ܰ��ڵ����ű��� : ���������̰� �����԰�ܰ��ڵ����ű����� ����Ǹ�)
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And _
            Vals(PB_intIAutoPriceGbn) <> Vals(Text1(29).Text) Then
            PB_intIAutoPriceGbn = Vals(Trim(Text1(29).Text))
         End If
         '(�������ܰ��ڵ����ű��� : ���������̰� �������ܰ��ڵ����ű����� ����Ǹ�)
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And _
            Vals(PB_intOAutoPriceGbn) <> Vals(Text1(30).Text) Then
            PB_intOAutoPriceGbn = Vals(Trim(Text1(30).Text))
         End If
         
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And PB_strEnterNo <> Trim(Text1(2).Text) Then
            PB_strEnterNo = Trim(Text1(2).Text)
         End If
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And PB_strEnterName <> Trim(Text1(1).Text) Then
            PB_strEnterName = Trim(Text1(1).Text)
         End If
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And PB_strRepName <> Trim(Text1(4).Text) Then
            PB_strRepName = Trim(Text1(4).Text)
         End If
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And _
            PB_strEnterAddress <> (Trim(Text1(12).Text) + Space(1) + Trim(Text1(13).Text)) Then
            PB_strEnterAddress = Trim(Text1(12).Text) + Space(1) + Trim(Text1(13).Text)
         End If
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And PB_strUptae <> Trim(Text1(6).Text) Then
            PB_strUptae = Trim(Text1(6).Text)
         End If
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And PB_strUpjong <> Trim(Text1(7).Text) Then
            PB_strUpjong = Trim(Text1(7).Text)
         End If
                  
         '(���Ÿ�Ա��� : ���������̰� ���Ÿ���� ����Ǹ�)
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And Vals(PB_intPrtTypeGbn) <> Vals(Text1(24).Text) Then
            PB_intPrtTypeGbn = Vals(Text1(24).Text)
         End If
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And Vals(PB_intDTopMargin) <> Vals(Text1(25).Text) Then
            PB_intDTopMargin = Vals(Text1(25).Text)
         End If
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And Vals(PB_intDLeftMargin) <> Vals(Text1(26).Text) Then
            PB_intDLeftMargin = Vals(Text1(26).Text)
         End If
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And Vals(PB_intTTopMargin) <> Vals(Text1(27).Text) Then
            PB_intTTopMargin = Vals(Text1(27).Text)
         End If
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And Vals(PB_intTLeftMargin) <> Vals(Text1(28).Text) Then
            PB_intTLeftMargin = Vals(Text1(28).Text)
         End If
         
         '+----------+
         ' ������Ʈ��
         '+----------+
         '������̸��� �ٲ� ��� ������Ʈ�� ����
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And _
            PB_regUserinfoU.UserBranchName <> Trim(Text1(1).Text) Then
            frmMain.Caption = PB_strSystemName & " - " & Trim(Text1(1).Text)
            PB_regUserinfoU.UserBranchName = Trim(Text1(1).Text)
            UserinfoU_Save PB_regUserinfoU
         End If
         '�����ޱݹ߻������� �ٲ� ��� ������Ʈ�� ����
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And _
            PB_regUserinfoU.UserMJGbn <> Trim(Text1(15).Text) Then
            PB_regUserinfoU.UserMJGbn = Trim(Text1(15).Text)
            UserinfoU_Save PB_regUserinfoU
         End If
         '�̼��ݹ߻������� �ٲ� ��� ������Ʈ�� ����
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And _
            PB_regUserinfoU.UserMSGbn <> Trim(Text1(16).Text) Then
            PB_regUserinfoU.UserMSGbn = Trim(Text1(16).Text)
            UserinfoU_Save PB_regUserinfoU
         End If
         .SetFocus
         Screen.MousePointer = vbDefault
    End With
    cmdSave.Enabled = True
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����� �б� ����"
    Unload Me
    Exit Sub
ERROR_TABLE_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����� �߰� ����"
    Unload Me
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����� ���� ����"
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
               '        & "WHERE ����屸�� = " & .TextMatrix(.Row, 0) & " "
               'On Error GoTo ERROR_TABLE_SELECT
               'P_adoRec.Open SQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
               'lngCnt = P_adoRec("�ش�Ǽ�")
               'P_adoRec.Close
               If lngCnt <> 0 Then
                  MsgBox "XXX(" & Format(lngCnt, "#,#") & "��)�� �����Ƿ� ������ �� �����ϴ�.", vbCritical, "����� ���� �Ұ�"
                  cmdDelete.Enabled = True
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
               PB_adoCnnSQL.BeginTrans
               P_adoRec.CursorLocation = adUseClient
               strSQL = "DELETE FROM ������� WHERE ������ڵ� = " & .TextMatrix(.Row, 0) & " "
               On Error GoTo ERROR_TABLE_DELETE
               PB_adoCnnSQL.Execute strSQL
               strSQL = "DELETE FROM ����� WHERE ������ڵ� = " & .TextMatrix(.Row, 0) & " "
               On Error GoTo ERROR_TABLE_DELETE
               PB_adoCnnSQL.Execute strSQL
               PB_adoCnnSQL.CommitTrans
               .RemoveItem .Row
               Text1(Text1.LBound).Enabled = False
               If .Rows <= PC_intRowCnt Then
                  .ScrollBars = flexScrollBarHorizontal
               End If
               Screen.MousePointer = vbDefault
               If .Rows = 1 Then
                  SubClearText
                  .Row = 0
                  Text1(Text1.LBound).Enabled = True
                  Text1(Text1.LBound).SetFocus
                  Exit Sub
               End If
               cmdDelete.Enabled = True
               vsfg1_EnterCell
               vsfg1.SetFocus
               vsfg1_AfterRowColChange 0, 0, 1, 1
            End If
         End If
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_DELETE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����� ���� ����"
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
    Set frm��������� = Nothing
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
    Text1(Text1.LBound).Enabled = False                '������ڵ� FLASE
    With vsfg1              'Rows 0, Cols 34, RowHeightMax(Min) 300
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
         .Cols = 34
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 850    '������ڵ�
         .ColWidth(1) = 3000   '������
         .ColWidth(2) = 1600   '����ڹ�ȣ
         .ColWidth(3) = 1600   '���ι�ȣ
         .ColWidth(4) = 1000   '��ǥ�ڸ�
         .ColWidth(5) = 1500   '��ǥ���ֹι�ȣ
         .ColWidth(6) = 2500   '����
         .ColWidth(7) = 2500   '����
         .ColWidth(8) = 2000   '��ȭ��ȣ
         .ColWidth(9) = 1600   '�ѽ���ȣ
         .ColWidth(10) = 1     '��뱸��
         .ColWidth(11) = 1000  '��뱸��
         
         .ColWidth(12) = 1000  '��������
         .ColWidth(13) = 1000  '�����ȣ
         .ColWidth(14) = 1     '������ּ�
         .ColWidth(15) = 1     '��������
         .ColWidth(16) = 7000  '������ּ�(�ּ�+����)
         .ColWidth(17) = 1000  '�ΰ�����
         .ColWidth(18) = 1000  '�����ޱݹ߻�����
         .ColWidth(19) = 1000  '�̼��ݹ߻�����
         
         .ColWidth(20) = 4600  '�̸����ּ�
         .ColWidth(21) = 4600  'Ȩ�������ּ�
         .ColWidth(22) = 4600  '�������
         .ColWidth(23) = 2000  '������ʸ������
         .ColWidth(24) = 2000  '�����ޱݱ��ʸ������
         .ColWidth(25) = 2000  '�̼��ݱ��ʸ������
         .ColWidth(26) = 2000  'ȸ����ʸ������
         .ColWidth(27) = 2000  '���Ÿ�Ա���
         .ColWidth(28) = 2000  '�ŷ�������ܸ���
         .ColWidth(29) = 2000  '�ŷ��������ʸ���
         .ColWidth(30) = 2000  '���ݰ�꼭��ܸ���
         .ColWidth(31) = 2000  '���ݰ�꼭���ʸ���
         .ColWidth(32) = 2000  '�����԰�ܰ��ڵ����ű���
         .ColWidth(33) = 2000  '�������ܰ��ڵ����ű���
         
         .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "�ڵ�"
         .TextMatrix(0, 1) = "������"
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
         .TextMatrix(0, 14) = "������ּ�"
         .TextMatrix(0, 15) = "��������"
         .TextMatrix(0, 16) = "������ּ�" '�ּ�+����
         .TextMatrix(0, 17) = "�ΰ�����"
         .TextMatrix(0, 18) = "�����ޱ�"
         .TextMatrix(0, 19) = "�̼���"
         .TextMatrix(0, 20) = "�̸����ּ�"
         .TextMatrix(0, 21) = "Ȩ�������ּ�"
         .TextMatrix(0, 22) = "�������"
         .TextMatrix(0, 23) = "������ʸ������"
         .TextMatrix(0, 24) = "�����ޱݱ��ʸ������"
         .TextMatrix(0, 25) = "�̼��ݱ��ʸ������"
         .TextMatrix(0, 26) = "ȸ����ʸ������"
         .TextMatrix(0, 27) = "���Ÿ��"
         .TextMatrix(0, 28) = "�ŷ�������ܸ���"
         .TextMatrix(0, 29) = "�ŷ��������ʸ���"
         .TextMatrix(0, 30) = "���ݰ�꼭��ܸ���"
         .TextMatrix(0, 31) = "���ݰ�꼭���ʸ���"
         .TextMatrix(0, 32) = "���Դܰ�����"
         .TextMatrix(0, 33) = "����ܰ�����"
         .ColHidden(10) = True: .ColHidden(14) = True: .ColHidden(15) = True
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 1 To 9, 16
                        .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 0, 11 To 13, 17 To 19, 23 To 26, 27, 32, 33
                         .ColAlignment(lngC) = flexAlignCenterCenter
                    Case Else
             End Select
         Next lngC
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 17
                        .ColFormat(17) = "#0.00"
             End Select
         Next lngC
         
    End With
End Sub

'+---------------------------------+
'/// VsFlexGrid(vsfg1) ä���///
'+---------------------------------+
Private Sub Subvsfg1_FILL()
Dim SQL        As String
Dim strWhere   As String
Dim strOrderBy As String
Dim lngR       As Long
Dim lngC       As Long
Dim lngRR      As Long
    P_adoRec.CursorLocation = adUseClient
    SQL = "SELECT * " _
          & "FROM ����� T1 " _
         & "ORDER BY T1.������ڵ� "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open SQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
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
               .TextMatrix(lngR, 0) = P_adoRec("������ڵ�")
               .Cell(flexcpData, lngR, 0, lngR, 0) = Trim(.TextMatrix(lngR, 0)) 'FindRow ����� ����
               .TextMatrix(lngR, 1) = IIf(IsNull(P_adoRec("������")), "", P_adoRec("������"))
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
               .TextMatrix(lngR, 17) = P_adoRec("�ΰ�����")
               .TextMatrix(lngR, 18) = IIf(IsNull(P_adoRec("�����ޱݹ߻�����")), "", P_adoRec("�����ޱݹ߻�����"))
               .TextMatrix(lngR, 19) = IIf(IsNull(P_adoRec("�̼��ݹ߻�����")), "", P_adoRec("�̼��ݹ߻�����"))
               .TextMatrix(lngR, 20) = IIf(IsNull(P_adoRec("�̸����ּ�")), "", P_adoRec("�̸����ּ�"))
               .TextMatrix(lngR, 21) = IIf(IsNull(P_adoRec("Ȩ�������ּ�")), "", P_adoRec("Ȩ�������ּ�"))
               .TextMatrix(lngR, 22) = IIf(IsNull(P_adoRec("�������")), "", P_adoRec("�������"))
               .TextMatrix(lngR, 23) = IIf(IsNull(P_adoRec("������ʸ������")), "", P_adoRec("������ʸ������"))
               .TextMatrix(lngR, 24) = IIf(IsNull(P_adoRec("�����ޱݱ��ʸ������")), "", P_adoRec("�����ޱݱ��ʸ������"))
               .TextMatrix(lngR, 25) = IIf(IsNull(P_adoRec("�̼��ݱ��ʸ������")), "", P_adoRec("�̼��ݱ��ʸ������"))
               .TextMatrix(lngR, 26) = IIf(IsNull(P_adoRec("ȸ����ʸ������")), "", P_adoRec("ȸ����ʸ������"))
               
               .TextMatrix(lngR, 27) = P_adoRec("���Ÿ�Ա���")
               .TextMatrix(lngR, 28) = P_adoRec("�ŷ�������ܸ���")
               .TextMatrix(lngR, 29) = P_adoRec("�ŷ��������ʸ���")
               .TextMatrix(lngR, 30) = P_adoRec("���ݰ�꼭��ܸ���")
               .TextMatrix(lngR, 31) = P_adoRec("���ݰ�꼭���ʸ���")
               .TextMatrix(lngR, 32) = P_adoRec("�����԰�ܰ��ڵ����ű���")
               .TextMatrix(lngR, 33) = P_adoRec("�������ܰ��ڵ����ű���")
               
               If .TextMatrix(lngR, 0) = PB_regUserinfoU.UserBranchCode Then
                  lngRR = lngR
               End If
               P_adoRec.MoveNext
            Loop
            P_adoRec.Close
            If lngRR = 0 Then
               .Row = lngRR       'vsfg1_EnterCell �ڵ�����(���� �Ѱ� �϶��� �ڵ����� ����)
               If .Rows > PC_intRowCnt Then
                  .TopRow = .Rows - PC_intRowCnt + .FixedRows
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
            If lngRR <> 0 Then
               vsfg1_AfterRowColChange 0, 0, 1, 1
            End If
       End With
    End If
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����� �б� ����"
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
    dtpOpenDate.Value = Format("19000101", "0000-00-00")
End Sub

'+-------------------------+
'/// Check text1(index) ///
'+-------------------------+
Private Function FncCheckTextBox(lngC As Long, blnOK As Boolean)
    For lngC = Text1.LBound To Text1.UBound
        Select Case lngC
               Case 0  '������ڵ�
                    Text1(lngC).Text = Format(Val(Trim(Text1(lngC).Text)), "00")
                    If Not (Text1(lngC).Text >= "01" And Text1(lngC).Text <= "99") Then
                       Exit Function
                    End If
               Case 5  '�ֹι�ȣ
                    If Len(Trim(Text1(lngC).Text)) > 14 Then
                       Exit Function
                    End If
               Case 10  '��뱸��
                    Text1(lngC).Text = Format(Val(Trim(Text1(lngC).Text)), "0")
                    If Not (Text1(lngC).Text >= "0" And Text1(lngC).Text <= "9") Then
                       Exit Function
                    End If
               Case 15, 16  '�����ޱݹ߻�, �̼��ݹ߻�
                    Text1(lngC).Text = Format(Val(Trim(Text1(lngC).Text)), "0")
                    If Not (Text1(lngC).Text >= "1" And Text1(lngC).Text <= "2") Then
                       Exit Function
                    End If
               Case 17, 19  '17.�̸����ּ�, 18.Ȩ�������ּ�, 19.�������
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (LenH(Text1(lngC).Text) <= 50) Then
                       Exit Function
                    End If
               Case 20 To 23  '���ʸ������
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (LenH(Text1(lngC).Text) = 6) Then
                       Exit Function
                    End If
               Case 24 To 28  '���Ÿ�Ա���
                    Text1(lngC).Text = Val(Trim(Text1(lngC).Text))
                    If Len(Trim(Text1(lngC).Text)) > 2 Then
                       Exit Function
                    End If
               Case 29 To 30  '�ܰ�����
                    Text1(lngC).Text = Val(Trim(Text1(lngC).Text))
                    If Val(Trim(Text1(lngC).Text)) > 1 Then
                       Exit Function
                    End If
               Case Else
        End Select
    Next lngC
    blnOK = True
End Function

