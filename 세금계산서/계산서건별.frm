VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm���ݰ�꼭 
   BorderStyle     =   0  '����
   Caption         =   "���ݰ�꼭"
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
      Begin VB.ComboBox cboPrinter 
         Height          =   300
         Left            =   4920
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   43
         Top             =   240
         Width           =   2235
      End
      Begin VB.OptionButton optPrtChk0 
         Caption         =   "�Ǻ�"
         Height          =   255
         Left            =   7200
         TabIndex        =   25
         Top             =   150
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optPrtChk1 
         Caption         =   "��ü"
         Height          =   255
         Left            =   7200
         TabIndex        =   24
         Top             =   390
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   375
         Left            =   7980
         Picture         =   "��꼭�Ǻ�.frx":0000
         Style           =   1  '�׷���
         TabIndex        =   20
         Top             =   195
         Width           =   1095
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   4320
         Top             =   195
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowBorderStyle=   0
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowGroupTree=   -1  'True
      End
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   9120
         Picture         =   "��꼭�Ǻ�.frx":0963
         Style           =   1  '�׷���
         TabIndex        =   11
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   13635
         Picture         =   "��꼭�Ǻ�.frx":1308
         Style           =   1  '�׷���
         TabIndex        =   13
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   390
         Left            =   12510
         Picture         =   "��꼭�Ǻ�.frx":1C56
         Style           =   1  '�׷���
         TabIndex        =   12
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   11385
         Picture         =   "��꼭�Ǻ�.frx":25DA
         Style           =   1  '�׷���
         TabIndex        =   14
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Height          =   390
         Left            =   10260
         Picture         =   "��꼭�Ǻ�.frx":2E61
         Style           =   1  '�׷���
         TabIndex        =   0
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00008000&
         BorderStyle     =   1  '���� ����
         Caption         =   "���ݰ�꼭"
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
      Height          =   7875
      Left            =   60
      TabIndex        =   10
      Top             =   2055
      Width           =   15195
      _cx             =   26802
      _cy             =   13891
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
      Height          =   1395
      Left            =   60
      TabIndex        =   15
      Top             =   630
      Width           =   15195
      Begin VB.ComboBox cboUsage 
         Height          =   300
         Left            =   12840
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   41
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox cboCredit 
         Height          =   300
         Left            =   11040
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   9
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox cboRS 
         Height          =   300
         Left            =   9240
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   8
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox cboMny 
         Height          =   300
         Left            =   7440
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   7
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox cboMake 
         Height          =   300
         Left            =   5760
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   6
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox cboPrint 
         Height          =   300
         Left            =   3720
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
      Begin VB.ListBox lstPort 
         Height          =   240
         Left            =   7800
         TabIndex        =   32
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "������ �μ�"
         Height          =   375
         Left            =   13440
         TabIndex        =   31
         Top             =   195
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '����
         Index           =   1
         Left            =   4200
         TabIndex        =   2
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   0
         Left            =   3000
         MaxLength       =   4
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpF_Date 
         Height          =   270
         Left            =   3000
         TabIndex        =   3
         Top             =   600
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   19660801
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin MSComCtl2.DTPicker dtpT_Date 
         Height          =   270
         Left            =   5160
         TabIndex        =   4
         Top             =   600
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   19660801
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin VB.Label Label1 
         Caption         =   "]"
         Height          =   240
         Index           =   16
         Left            =   13920
         TabIndex        =   42
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "��� :"
         Height          =   240
         Index           =   14
         Left            =   12120
         TabIndex        =   40
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "�̼� :"
         Height          =   240
         Index           =   13
         Left            =   10320
         TabIndex        =   39
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "��û :"
         Height          =   240
         Index           =   12
         Left            =   8520
         TabIndex        =   38
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "["
         Height          =   240
         Index           =   11
         Left            =   2520
         TabIndex        =   37
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   10
         Left            =   1680
         TabIndex        =   36
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "�ݾ� :"
         Height          =   240
         Index           =   9
         Left            =   6720
         TabIndex        =   35
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "�ۼ� :"
         Height          =   240
         Index           =   8
         Left            =   5040
         TabIndex        =   34
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "���� :"
         Height          =   240
         Index           =   7
         Left            =   3000
         TabIndex        =   33
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   6
         Left            =   6600
         TabIndex        =   30
         Top             =   645
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   5
         Left            =   4440
         TabIndex        =   29
         Top             =   645
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "]"
         Height          =   240
         Index           =   4
         Left            =   7520
         TabIndex        =   28
         Top             =   645
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "["
         Height          =   240
         Index           =   3
         Left            =   2520
         TabIndex        =   27
         Top             =   645
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��������"
         Height          =   240
         Index           =   2
         Left            =   1200
         TabIndex        =   26
         Top             =   650
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��ü�ݾ�"
         Height          =   240
         Index           =   1
         Left            =   8760
         TabIndex        =   23
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label lblTotMny 
         Alignment       =   1  '������ ����
         Caption         =   "0.00"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   10080
         TabIndex        =   22
         Top             =   285
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "[F1]"
         Height          =   240
         Index           =   15
         Left            =   2520
         TabIndex        =   21
         Top             =   285
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����ó�ڵ�"
         Height          =   240
         Index           =   0
         Left            =   1200
         TabIndex        =   19
         Top             =   285
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
         Left            =   120
         TabIndex        =   18
         Top             =   285
         Width           =   975
      End
   End
End
Attribute VB_Name = "frm���ݰ�꼭"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' ���α׷� �� �� : ���ݰ�꼭
' ���� Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' ������ Table   : �����, ����ó, �������⳻��, ���ݰ�꼭
' ��  ��  ��  �� :
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived         As Boolean
Private P_adoRec             As New ADODB.Recordset
Private P_adoRecW            As New ADODB.Recordset
Private P_intButton          As Integer
Private P_strFindString2     As String
Private Const PC_intRowCnt1  As Integer = 25   '�׸���1�� �� ������ �� ���(FixedRows ����)

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
Dim strSQL             As String
Dim inti               As Integer

Dim P                  As Printer
Dim strDefaultPrinter  As String
Dim aryPrinter()       As String
Dim strBuffer          As String

    If P_blnActived = False Then
       Screen.MousePointer = vbHourglass
       On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
       
       'Printer Setting and Seaching(API)
       strBuffer = Space(1024)
       inti = GetProfileString("windows", "Device", "", strBuffer, Len(strBuffer))
       aryPrinter = Split(strBuffer, ",")
       strDefaultPrinter = Trim(Trim(aryPrinter(0)))
       For Each P In Printers
           cboPrinter.AddItem Trim(P.DeviceName)
           lstPort.AddItem P.Port
       Next
       For inti = 0 To cboPrinter.ListCount - 1
           cboPrinter.ListIndex = inti
           If UCase(Trim(cboPrinter.Text)) = UCase(Trim(strDefaultPrinter)) Then
              Exit For
           End If
       Next inti
       '---
       Subvsfg1_INIT  '���ݰ�꼭
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
                   cmdSave.Enabled = True: cmdDelete.Enabled = True: cmdExit.Enabled = True
              Case Is <= 99 '����, �߰�, ����, ��ȸ, �μ�
                   cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = True: cmdDelete.Enabled = True: cmdExit.Enabled = True
              Case Else
                   cmdPrint.Enabled = False: cmdClear.Enabled = False: cmdFind.Enabled = False
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
       End Select
       SubOther_FILL
       P_blnActived = True
       Screen.MousePointer = vbDefault
    End If
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�������(�������� ���� ����)"
    Unload Me
    Exit Sub
End Sub

'+--------------------+
'--- Select Printer ---
'+--------------------+
Private Sub cboPrinter_Click()
    lstPort.ListIndex = cboPrinter.ListIndex
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
    If KeyCode = vbKeyF1 Then '����ó�˻�
       If Index = 0 Then      '
          PB_strSupplierCode = Trim(Text1(Index).Text)
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
       End If
    End If
    If KeyCode = vbKeyReturn Then
       Select Case Index
       End Select
       SendKeys "{tab}"
    End If
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����ó �б� ����"
    Unload Me
    Exit Sub
End Sub

'+-------------+
'/// ����ó ///
'+-------------+
Private Sub Text1_LostFocus(Index As Integer)
Dim strSQL As String
Dim lngR   As Long
    With Text1(Index)
         Select Case Index
                Case 0
                     .Text = Trim(.Text)
                     If Len(.Text) < 1 Then
                        Text1(Index).Text = ""
                        Text1(Index + 1).Text = ""
                        Exit Sub
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

'+-------------------+
'/// �ŷ����ڼ��� ///
'+-------------------+
Private Sub dtpF_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub dtpT_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    End If
End Sub
'+---------------+
'/// ���м��� ///
'+---------------+
Private Sub cboPrint_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub cboMake_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub cboMny_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub cboRS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub cboCredit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       If cmdFind.Enabled = True Then
          cmdFind_Click
       Else
          SendKeys "{tab}"
       End If
    End If
End Sub

'+------------+
'/// vsfg1 ///
'+------------+
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
    With vsfg1
         P_intButton = Button
         'If (.MouseRow >= .FixedRows) Then
         If (.MouseRow >= .FixedRows And _
            .TextMatrix(.MouseRow, 20) = "����" And .TextMatrix(.MouseRow, 16) = "����" And .TextMatrix(.MouseRow, 19) = "�̼�") Then
            If cmdSave.Enabled = False Then Exit Sub
            If (.MouseCol = 8) Then      '���ް���
                If Button = vbLeftButton Then
                  .Select .MouseRow, .MouseCol
                  .EditCell
                End If
            'ElseIf _
            '   (.MouseCol = 9) Then      '����
            '   If Button = vbLeftButton Then
            '      .Select .MouseRow, .MouseCol
            '      .EditCell
            '    End If
            ElseIf _
               (.MouseCol = 11) Then     'ǰ��ױ԰�
               If Button = vbLeftButton Then
                  .Select .MouseRow, .MouseCol
                  .EditCell
                End If
            ElseIf _
                (.MouseCol = 13) Then    '����(��)
                If Button = vbLeftButton Then
                  .Select .MouseRow, .MouseCol
                  .EditCell
                End If
            End If
         End If
    End With
End Sub
Sub vsfg1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim blnModify As Boolean
Dim curTmpMny As Currency
    With vsfg1
         If Row >= .FixedRows Then
            If (Col = 8) Then   '���ް���
               If .TextMatrix(Row, Col) <> .EditText Then
                  If (IsNumeric(.EditText) = False Or Vals(.EditText) < 0 Or _
                     Int(Vals(.EditText)) < Vals(.EditText) Or IsNumeric(Right(.EditText, 1)) = False) Then '�Ҽ������� ���Ұ�
                     'IsNumeric(Right(.EditText, 1)) = False) Then                                            '�Ҽ������� ��밡
                     Beep
                     Cancel = True
                  Else
                     blnModify = True
                     curTmpMny = .ValueMatrix(Row, 8)
                     .TextMatrix(Row, 8) = Vals(.EditText)
                     .TextMatrix(Row, 9) = Int(Vals(.EditText) * 0.1)  '�ΰ���
                     .TextMatrix(Row, 10) = Vals(.EditText) + .ValueMatrix(Row, 9)
                  End If
               End If
            ElseIf _
               (Col = 9) Then  '����
               If .TextMatrix(Row, Col) <> .EditText Then
                  If (IsNumeric(.EditText) = False Or Vals(.EditText) < 0 Or _
                     Int(Vals(.EditText)) < Vals(.EditText) Or IsNumeric(Right(.EditText, 1)) = False) Then
                     Beep
                     Cancel = True
                  Else
                     blnModify = True
                     curTmpMny = .ValueMatrix(Row, 8)
                     .TextMatrix(Row, 10) = .ValueMatrix(Row, 8) + Vals(.EditText)
                  End If
               End If
            ElseIf _
               (Col = 11) Then  'ǰ��ױ԰�
               If .TextMatrix(Row, Col) <> .EditText Then
                  If Not (LenH(Trim(.EditText)) <= 50) Then
                     Beep
                     .TextMatrix(Row, Col) = .EditText
                     Cancel = True
                  Else
                     blnModify = True
                  End If
               End If
            ElseIf _
               (Col = 13) Then  '����(��)
               If .TextMatrix(Row, Col) <> .EditText Then
                  If (IsNumeric(.EditText) = False Or Vals(.EditText) < 0 Or _
                     Int(Vals(.EditText)) < Vals(.EditText) Or IsNumeric(Right(.EditText, 1)) = False) Then
                     Beep
                     Cancel = True
                  Else
                     blnModify = True
                     curTmpMny = .ValueMatrix(Row, 8)
                  End If
               End If
            'ElseIf _
            '   (Col = 18) Then  '��꼭���࿩��
            '   If (Len(.TextMatrix(Row, 9)) = 0) Then '����ó�� ����
            '      .Cell(flexcpChecked, Row, 18, Row, 18) = flexUnchecked
            '      Beep
            '      Cancel = True
            '      Exit Sub
            '   End If
            '   If .Cell(flexcpChecked, Row, Col) <> .EditText Then
            '      blnModify = True
            '   End If
            End If
            '����ǥ�� + �ݾ�����
            If blnModify = True Then
               If .TextMatrix(Row, 21) = "" Then
                  .TextMatrix(Row, 21) = "U"
               End If
               .Cell(flexcpBackColor, Row, Col, Row, Col) = vbRed
               .Cell(flexcpForeColor, Row, Col, Row, Col) = vbWhite
               Select Case Col
                      Case 8, 9, 13
                           lblTotMny.Caption = Format(Vals(lblTotMny.Caption) - curTmpMny + .ValueMatrix(Row, 8), "#,#.00")
                      Case Else
               End Select
            End If
         End If
    End With
End Sub
Private Sub vsfg1_Click()
Dim strData As String
Dim lngR1   As Long
Dim lngRH1  As Long
Dim lngR2   As Long
Dim lngRR2  As Long
Dim lngC    As Long
    With vsfg1
         If (.MouseRow >= 0 And .MouseRow < .FixedRows) Then
            .Col = .MouseCol
            .Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlack
            .Cell(flexcpForeColor, 0, .MouseCol, 0, .MouseCol) = vbRed
            strData = Trim(.Cell(flexcpData, .Row, 4))
            Select Case .MouseCol
                   Case 4
                        '.ColSel = 4
                        .Select 0, 0, 0, 4
                        .ColSort(0) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .ColSort(1) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .ColSort(2) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .ColSort(3) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .Sort = flexSortUseColSort
                   Case 5
                        .ColSel = 5
                        .ColSort(0) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .ColSort(1) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .ColSort(2) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .ColSort(3) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .ColSort(5) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .Sort = flexSortUseColSort
                   Case Else
                        .Sort = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            End Select
            If .FindRow(strData, , 4) > 0 Then
               .Row = .FindRow(strData, , 4)
            End If
            If PC_intRowCnt1 < .Rows Then
               .TopRow = .Row
            End If
         End If
    End With
End Sub
Private Sub vsfg1_EnterCell()
Dim lngR1  As Long
Dim lngR2  As Long
Dim lngC   As Long
Dim lngCnt As Long
    With vsfg1
         If .Row >= .FixedRows Then
         End If
    End With
End Sub
Private Sub vsfg1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
Dim lngR1  As Long
Dim lngR2  As Long
Dim lngC   As Long
Dim lngCnt As Long
    With vsfg1
         If NewRow <> OldRow Then
            'For lngR2 = 1 To vsfg2.Rows - 1
            '    vsfg2.RowHidden(lngR2) = True
            'Next lngR2
            'If NewRow > 0 Then 'Add 20041002
            '   For lngR1 = .ValueMatrix(.Row, 14) To .ValueMatrix(.Row, 15)
            '       vsfg2.RowHidden(lngR1) = False
            '       lngCnt = lngCnt + 1
            '   Next lngR1
            'End If
            'If PC_intRowCnt2 < lngCnt Then
            '   vsfg2.TopRow = vsfg2.Row
            'End If
         End If
    End With
End Sub
Private Sub vsfg1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lngR     As Long
Dim blnDupOK As Boolean
    With vsfg1
         If .Row >= .FixedRows Then
         End If
    End With
End Sub

'+-----------+
'/// ��� ///
'+-----------+
Private Sub cmdPrint_Click()
Dim strSQL         As String
Dim lngR           As Long
Dim lngLogCnt      As Long
Dim strMakeYear    As String
Dim lngLogCnt1     As Long
Dim lngLogCnt2     As Long
Dim strServerTime  As String
    If (vsfg1.Rows = 1) Or (vsfg1.Row < 1) Then
       Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    strSQL = "SELECT CONVERT(VARCHAR(19),GETDATE(), 120) AS �����ð� "
    On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    strServerTime = Format(Right(P_adoRec("�����ð�"), 8), "hhmmss")
    P_adoRec.Close
    If optPrtChk0.Value = True Then '���ݰ�꼭�μ�(�Ǻ�)
       With vsfg1
            If .TextMatrix(.Row, 20) = "����" Then
               If .Cell(flexcpChecked, .Row, 15) = flexUnchecked Then
                  PB_adoCnnSQL.BeginTrans
                  strSQL = "UPDATE ���ݰ�꼭 SET " _
                                & "���࿩�� = 1," _
                                & "�������� = '" & PB_regUserinfoU.UserServerDate & "', " _
                                & "������ڵ� = '" & PB_regUserinfoU.UserCode & "' " _
                          & "WHERE ������ڵ� = '" & .TextMatrix(.Row, 0) & "' " _
                            & "AND �ۼ��⵵ = '" & .TextMatrix(.Row, 1) & "' AND å��ȣ = " & .ValueMatrix(.Row, 2) & " " _
                            & "AND �Ϸù�ȣ = " & .ValueMatrix(.Row, 3) & " "
                   On Error GoTo ERROR_TABLE_UPDATE
                   PB_adoCnnSQL.Execute strSQL
                   PB_adoCnnSQL.CommitTrans
                   .Cell(flexcpChecked, .Row, 15) = flexChecked
                   .Cell(flexcpText, .Row, 15) = "�� ��"
               End If
               SubPrint_TaxBill .TextMatrix(.Row, 0), .TextMatrix(.Row, 1), .ValueMatrix(.Row, 2), .ValueMatrix(.Row, 3)
            End If
       End With
    End If
    If optPrtChk1.Value = True Then '���ݰ�꼭�μ�(��ü)
       With vsfg1
            SubPrint_TaxBill .TextMatrix(.Row, 0), .TextMatrix(lngR, 1), .ValueMatrix(lngR, 2), .ValueMatrix(lngR, 3)
       End With
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "(�������� ���� ����)"
    Unload Me
    Exit Sub
ERROR_TABLE_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�߰� ����"
    Unload Me
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "���� ����"
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
'/// �߰� ///
'+-----------+
Private Sub cmdClear_Click()
    '
End Sub
'+-----------+
'/// ��ȸ ///
'+-----------+
Private Sub cmdFind_Click()
    Screen.MousePointer = vbHourglass
    cmdFind.Enabled = False
    lblTotMny.Caption = "0.00"
    Subvsfg1_FILL
    'Subvsfg2_FILL
    cmdFind.Enabled = True
    Screen.MousePointer = vbDefault
End Sub

'+-----------+
'/// ���� ///
'+-----------+
Private Sub cmdSave_Click()
Dim blnSaveOK      As Boolean
Dim strSQL         As String
Dim lngR           As Long
Dim lngRR          As Long
Dim lngRRR         As Long
Dim lngC           As Long
Dim blnOK          As Boolean
Dim intRetVal      As Integer
Dim lngChkCnt      As Long
Dim lngDelCntS     As Long
Dim lngDelCntE     As Long
Dim lngDelCnt      As Long
Dim lngLogCnt      As Long
Dim intAddTax      As Integer '���ݰ�꼭(0.���ۼ�, 1.����, 2.�߰�, 3.�������߰�)
Dim intCreateTax   As Integer '���ݰ�꼭���࿩��(1.����)
Dim strOldMakeYear As String
Dim lngOldLogCnt1  As Long
Dim lngOldLogCnt2  As Long
Dim strNewMakeYear As String
Dim lngNewLogCnt1  As Long
Dim lngNewLogCnt2  As Long
Dim curInputMoney  As Long
Dim curOutputMoney As Long
Dim strServerTime  As String
    If vsfg1.Row >= vsfg1.FixedRows Then
       With vsfg1
            If (.TextMatrix(lngRR, 21) = "U") Then
               blnSaveOK = True
            End If
            If blnSaveOK = False Then '������(�����) ���̾�����
               Exit Sub
            End If
       End With
       intRetVal = MsgBox("����ó���� �ڷḦ �����Ͻðڽ��ϱ� ?", vbQuestion + vbYesNo + vbDefaultButton2, "�ڷ� ����")
       If intRetVal = vbNo Then
          vsfg1.SetFocus
          Exit Sub
       End If
       cmdSave.Enabled = False
       Screen.MousePointer = vbHourglass
       strSQL = "SELECT CONVERT(VARCHAR(19),GETDATE(), 120) AS �����ð� "
       On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
       P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
       strServerTime = Format(Right(P_adoRec("�����ð�"), 8), "hhmmss")
       P_adoRec.Close
       PB_adoCnnSQL.BeginTrans
       P_adoRec.CursorLocation = adUseClient
       With vsfg1
            strSQL = "UPDATE �������⳻�� SET " _
                          & "��꼭���࿩�� = 1, " _
                          & "���ܰ� = " & .ValueMatrix(.Row, 8) & "," _
                          & "���ΰ� = " & .ValueMatrix(.Row, 9) & ", " _
                          & "���� = '" & .TextMatrix(.Row, 11) + " " + CStr(.ValueMatrix(.Row, 13)) + "��" & "', " _
                          & "�������� = '" & PB_regUserinfoU.UserServerDate & "', " _
                          & "������ڵ� = '" & PB_regUserinfoU.UserCode & "' " _
                    & "WHERE ������ڵ� = '" & .TextMatrix(.Row, 0) & "' " _
                      & "AND �ۼ��⵵ = '" & .TextMatrix(.Row, 1) & "' " _
                      & "AND å��ȣ = " & .ValueMatrix(.Row, 2) & " " _
                      & "AND �Ϸù�ȣ = " & .ValueMatrix(.Row, 3) & " "
             On Error GoTo ERROR_TABLE_INSERT
             PB_adoCnnSQL.Execute strSQL
             strSQL = "UPDATE ���ݰ�꼭 SET " _
                          & "���࿩�� = " & chkPrint.Value & ", " _
                          & "ǰ��ױ԰� = '" & .TextMatrix(.Row, 11) & "', " _
                          & "���� = " & .ValueMatrix(.Row, 13) & ", " _
                          & "���ް��� = " & .ValueMatrix(.Row, 8) & "," _
                          & "���� = " & .ValueMatrix(.Row, 9) & ", " _
                          & "�������� = '" & PB_regUserinfoU.UserServerDate & "', " _
                          & "������ڵ� = '" & PB_regUserinfoU.UserCode & "' " _
                    & "WHERE ������ڵ� = '" & .TextMatrix(.Row, 0) & "' " _
                      & "AND �ۼ��⵵ = '" & .TextMatrix(.Row, 1) & "' " _
                      & "AND å��ȣ = " & .ValueMatrix(.Row, 2) & " " _
                      & "AND �Ϸù�ȣ = " & .ValueMatrix(.Row, 3) & " "
             On Error GoTo ERROR_TABLE_INSERT
             PB_adoCnnSQL.Execute strSQL
             '(���� ����ġ)
             .Cell(flexcpBackColor, .Row, .FixedCols, .Row, .Cols - 1) = vbWhite
             .Cell(flexcpForeColor, .Row, .FixedCols, .Row, .Cols - 1) = vbBlack
             .Cell(flexcpText, .Row, 21, .Row, 21) = ""  'SQL���� ����
       End With
       PB_adoCnnSQL.CommitTrans
       If (chkTaxBillPrint.Value = 1) And (cboPrinter.ListIndex >= 0) Then                        '���ݰ�꼭 ���
          With vsfg1
               SubPrint_TaxBill .TextMatrix(.Row, 0), .TextMatrix(.Row, 1), .TextMatrix(.Value, 2), .ValueMatrix(.Row, 3)
          End With
       End If
       Screen.MousePointer = vbDefault
    End If
    cmdSave.Enabled = True
    vsfg1.SetFocus
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
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�˻� ����"
    Unload Me
    Exit Sub
ERROR_TABLE_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�߰� ����"
    Unload Me
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "���� ����"
    Unload Me
    Exit Sub
ERROR_TABLE_DELETE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "���� ����"
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
Dim strSQL     As String
Dim lngR       As Long
Dim lngRR      As Long
Dim lngRRR     As Long
Dim lngC       As Long
Dim blnOK      As Boolean
Dim intRetVal  As Integer
Dim lngChkCnt  As Long
Dim lngDelCntS As Long
Dim lngDelCntE As Long
Dim lngLogCnt  As Long
    If (vsfg1.Row >= vsfg1.FixedRows And vsfg1.TextMatrix(vsfg1.Row, 20) = "����" And _
       vsfg1.TextMatrix(vsfg1.Row, 16) = "����" And vsfg1.TextMatrix(vsfg1.Row, 19) = "�̼�") Then
       intRetVal = MsgBox("���ݰ�꼭�� �����Ͻðڽ��ϱ� ?", vbQuestion + vbYesNo + vbDefaultButton2, "���ݰ�꼭 ����")
       If intRetVal = vbNo Then
          vsfg1.SetFocus
          Exit Sub
       End If
       cmdDelete.Enabled = False
       Screen.MousePointer = vbHourglass
       With vsfg1
            lblTotMny.Caption = Format(Vals(lblTotMny.Caption) - .ValueMatrix(.Row, 8), "#,#.00") '��ü�ݾ׿��� ����
            PB_adoCnnSQL.BeginTrans
            strSQL = "UPDATE �������⳻�� SET " _
                          & "��뱸�� = 9, " _
                          & "�������� = '" & PB_regUserinfoU.UserServerDate & "', " _
                          & "������ڵ� = '" & PB_regUserinfoU.UserCode & "' " _
                    & "WHERE ������ڵ� = '" & .TextMatrix(.Row, 0) & "' AND �ۼ��⵵ = '" & .TextMatrix(.Row, 1) & "' " _
                      & "AND å��ȣ = " & .ValueMatrix(.Row, 2) & " AND �Ϸù�ȣ = " & .ValueMatrix(.Row, 3) & " "
            On Error GoTo ERROR_TABLE_DELETE
            PB_adoCnnSQL.Execute strSQL
            strSQL = "UPDATE ���ݰ�꼭 SET " _
                          & "��뱸�� = 9, " _
                          & "�������� = '" & PB_regUserinfoU.UserServerDate & "', " _
                          & "������ڵ� = '" & PB_regUserinfoU.UserCode & "' " _
                    & "WHERE ������ڵ� = '" & .TextMatrix(.Row, 0) & "' AND �ۼ��⵵ = '" & .TextMatrix(.Row, 1) & "' " _
                      & "AND å��ȣ = " & .ValueMatrix(.Row, 2) & " AND �Ϸù�ȣ = " & .ValueMatrix(.Row, 3) & " "
            On Error GoTo ERROR_TABLE_DELETE
            PB_adoCnnSQL.Execute strSQL
            PB_adoCnnSQL.CommitTrans
           .RemoveItem .Row
           .Row = 0
       End With
       cmdFind.SetFocus
       Screen.MousePointer = vbDefault
    End If
    cmdDelete.Enabled = True
    Exit Sub
ERROR_TABLE_DELETE:
    PB_adoCnnSQL.RollbackTrans
    Clipboard.Clear: Clipboard.SetText strSQL
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "���ݰ�꼭 ���� ����"
    Unload Me
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
    Set frm���ݰ�꼭 = Nothing
    frmMain.SBar.Panels(4).Text = ""
End Sub

'+--------------------+
'/// Sub Procedure ///
'+--------------------+
Private Sub SubOther_FILL()
Dim strSQL        As String
Dim intIndex      As Integer
    Text1(0).Text = "": Text1(1).Text = ""
    dtpF_Date.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00") 'Format(Mid(PB_regUserinfoU.UserClientDate, 1, 6) & "01", "0000-00-00")
    dtpT_Date.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
    cboPrint.AddItem "��  ü"
    cboPrint.AddItem "�̹���"
    cboPrint.AddItem "��  ��"
    cboPrint.ListIndex = 0
    cboMake.AddItem "��ü"
    cboMake.AddItem "�ŷ�"
    cboMake.AddItem "����"
    cboMake.AddItem "�ϰ�"
    cboMake.ListIndex = 0
    cboMny.AddItem "��ü"
    cboMny.AddItem "����"
    cboMny.AddItem "��ǥ"
    cboMny.AddItem "����"
    cboMny.AddItem "�ܻ�"
    cboMny.ListIndex = 0
    cboRS.AddItem "��ü"
    cboRS.AddItem "����"
    cboRS.AddItem "û��"
    cboRS.ListIndex = 0
    cboCredit.AddItem "��ü"
    cboCredit.AddItem "�Ϲ�"
    cboCredit.AddItem "�̼�"
    cboCredit.ListIndex = 0
    cboUsage.AddItem "��ü"
    cboUsage.AddItem "����"
    cboUsage.AddItem "����"
    cboUsage.ListIndex = 1
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����� �б� ����"
    Unload Me
    Exit Sub
End Sub
'+----------------------------------+
'/// VsFlexGrid(vsfg1) �ʱ�ȭ ///
'+----------------------------------+
Private Sub Subvsfg1_INIT()
Dim lngR    As Long
Dim lngC    As Long
    With vsfg1              'Rows 1, Cols 22, RowHeightMax(Min) 300
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
         .Rows = 1             'Subvsfg1_Fill����ÿ� ����
         .Cols = 22
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1200   '������ڵ� 'H
         .ColWidth(1) = 1000   '�ۼ��⵵   'H
         .ColWidth(2) = 1000   'å��ȣ     'H
         .ColWidth(3) = 1000   '�Ϸù�ȣ   'H
         .ColWidth(4) = 1500   '���ݰ�꼭��ȣ(�ۼ��⵵-å��ȣ-�Ϸù�ȣ)
         .ColWidth(5) = 1100   '��������(�ۼ�����)
         .ColWidth(6) = 1000   '����ó�ڵ�
         .ColWidth(7) = 2000   '����ó��
         .ColWidth(8) = 1500   '���ް���(�ܰ�)
         .ColWidth(9) = 1300   '����(�ΰ�)
         .ColWidth(10) = 1500  '�հ�       'H
         .ColWidth(11) = 3700  'ǰ��ױ԰�
         .ColWidth(12) = 300   '��
         .ColWidth(13) = 500   '����
         .ColWidth(14) = 300   '��
         .ColWidth(15) = 750   '���࿩��
         .ColWidth(16) = 750   '�ۼ�����
         .ColWidth(17) = 750   '�ݾױ���
         .ColWidth(18) = 750   '��û����
         .ColWidth(19) = 750   '�̼�����
         .ColWidth(20) = 750   '��뱸��
         .ColWidth(21) = 750   'SQL����
         
         .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "������ڵ�" 'H
         .TextMatrix(0, 1) = "�ۼ��⵵"   'H
         .TextMatrix(0, 2) = "å��ȣ"     'H
         .TextMatrix(0, 3) = "�Ϸù�ȣ"   'H
         .TextMatrix(0, 4) = "���ݰ�꼭��ȣ"
         .TextMatrix(0, 5) = "��������"
         .TextMatrix(0, 6) = "����ó�ڵ�" 'H
         .TextMatrix(0, 7) = "����ó��"
         .TextMatrix(0, 8) = "���ް���"
         .TextMatrix(0, 9) = "����"
         .TextMatrix(0, 10) = "�հ�ݾ�"  'H
         .TextMatrix(0, 11) = "ǰ��ױ԰�"
         .TextMatrix(0, 12) = "��"
         .TextMatrix(0, 13) = "����"
         .TextMatrix(0, 14) = "��"
         .TextMatrix(0, 15) = "����"
         .TextMatrix(0, 16) = "�ۼ�"
         .TextMatrix(0, 17) = "����"
         .TextMatrix(0, 18) = "��û"
         .TextMatrix(0, 19) = "�̼�"
         .TextMatrix(0, 20) = "���"
         .TextMatrix(0, 21) = "SQL"
         
         .ColHidden(0) = True: .ColHidden(1) = True: .ColHidden(2) = True: .ColHidden(3) = True:
         .ColHidden(6) = True: .ColHidden(10) = True: .ColHidden(21) = True
         .ColFormat(8) = "#,#.00": .ColFormat(9) = "#,#.00": .ColFormat(10) = "#,#.00"
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 7, 11, 12, 14
                         .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 0, 1, 2, 3, 4, 5, 6, 15, 16, 17, 18, 19, 20, 21
                         .ColAlignment(lngC) = flexAlignCenterCenter
                    Case Else
                         .ColAlignment(lngC) = flexAlignRightCenter
             End Select
         Next lngC
         
         .ColComboList(16) = "�ŷ�|����|�ϰ�"
         .ColComboList(17) = "����|��ǥ|����|�ܻ�"
         .ColComboList(18) = "����|û��"
         .ColComboList(20) = "����|����"
         
         '.MergeCells = flexMergeRestrictRows  'flexMergeFixedOnly
         '.MergeRow(0) = True
         'For lngC = 0 To 4
         '    .MergeCol(lngC) = True
         'Next lngC
    End With
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
    vsfg1.Rows = 1
    P_adoRec.CursorLocation = adUseClient
    If Len(Text1(0).Text) > 0 Then
       strWhere = "AND T1.����ó�ڵ� = '" & Trim(Text1(0).Text) & "' "
    End If
    If cboPrint.ListIndex > 0 Then
       strWhere = strWhere + "AND T1.���࿩�� = " & (cboPrint.ListIndex - 1) & " "
    End If
    If cboMake.ListIndex > 0 Then
       strWhere = strWhere + "AND T1.�ۼ����� = " & (cboMake.ListIndex - 1) & " "
    End If
    If cboMny.ListIndex > 0 Then
       strWhere = strWhere + "AND T1.�ݾױ��� = " & (cboMny.ListIndex - 1) & " "
    End If
    If cboRS.ListIndex > 0 Then
       strWhere = strWhere + "AND T1.��û���� = " & (cboRS.ListIndex - 1) & " "
    End If
    If cboCredit.ListIndex > 0 Then
       strWhere = strWhere + "AND T1.�̼����� = " & (cboCredit.ListIndex - 1) & " "
    End If
    If cboUsage.ListIndex = 0 Then
       vsfg1.ColHidden(20) = False
    Else
       vsfg1.ColHidden(20) = True
    End If
    If cboUsage.ListIndex > 0 Then
       strWhere = strWhere + "AND T1.��뱸�� = " & IIf(cboUsage.ListIndex = 1, 0, 9) & " "
    End If
    strSQL = "SELECT T1.������ڵ� AS ������ڵ�, T1.�ۼ��⵵ AS �ۼ��⵵, " _
                  & "T1.å��ȣ AS å��ȣ, T1.�Ϸù�ȣ AS �Ϸù�ȣ, " _
                  & "T1.�ۼ����� AS �ۼ�����, T1.����ó�ڵ� AS ����ó�ڵ�, T2.����ó�� AS ����ó��, " _
                  & "T1.���ް��� AS ���ް���, T1.���� AS ����, " _
                  & "T1.ǰ��ױ԰� AS ǰ��ױ԰�, T1.���� AS ����, " _
                  & "T1.�ݾױ��� AS �ݾױ���, T1.��û���� AS ��û����, T1.���࿩�� AS ���࿩��, " _
                  & "T1.�ۼ����� AS �ۼ�����, T1.�̼����� AS �̼�����, T1.��뱸�� AS ��뱸�� " _
             & "FROM ���ݰ�꼭 T1 " _
             & "LEFT JOIN ����ó T2 ON T2.������ڵ� = T1.������ڵ� AND T2.����ó�ڵ� = T1.����ó�ڵ� " _
            & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
              & "AND T1.�ۼ����� BETWEEN '" & DTOS(dtpF_Date.Value) & "' AND '" & DTOS(dtpT_Date.Value) & "' " _
              & "" & strWhere & " " _
            & "ORDER BY T1.������ڵ�, T1.�ۼ��⵵, T1.å��ȣ, T1.�Ϸù�ȣ "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    vsfg1.Rows = P_adoRec.RecordCount + 1
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       cmdExit.SetFocus
       Exit Sub
    Else
       With vsfg1
            .Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlack
            If .Rows <= PC_intRowCnt1 Then
               .ScrollBars = flexScrollBarHorizontal
            Else
               .ScrollBars = flexScrollBarBoth
            End If
            Do Until P_adoRec.EOF
               lngR = lngR + 1
               .TextMatrix(lngR, 0) = IIf(IsNull(P_adoRec("������ڵ�")), "", P_adoRec("������ڵ�"))
               .TextMatrix(lngR, 1) = P_adoRec("�ۼ��⵵")
               .TextMatrix(lngR, 2) = P_adoRec("å��ȣ")
               .TextMatrix(lngR, 3) = P_adoRec("�Ϸù�ȣ")
               .TextMatrix(lngR, 4) = P_adoRec("������ڵ�") & "-" & P_adoRec("�ۼ��⵵") _
                                   & "-" & CStr(P_adoRec("å��ȣ")) & "-" & CStr(P_adoRec("�Ϸù�ȣ"))
               .Cell(flexcpData, lngR, 4, lngR, 4) = Trim(.TextMatrix(lngR, 4)) 'FindRow ����� ����
               .TextMatrix(lngR, 5) = Format(P_adoRec("�ۼ�����"), "0000-00-00")
               .TextMatrix(lngR, 6) = IIf(IsNull(P_adoRec("����ó�ڵ�")), "", P_adoRec("����ó�ڵ�"))
               .TextMatrix(lngR, 7) = IIf(IsNull(P_adoRec("����ó��")), "", P_adoRec("����ó��"))
               .TextMatrix(lngR, 8) = IIf(IsNull(P_adoRec("���ް���")), 0, P_adoRec("���ް���"))
               .TextMatrix(lngR, 9) = IIf(IsNull(P_adoRec("����")), 0, P_adoRec("����"))
               .TextMatrix(lngR, 10) = .ValueMatrix(lngR, 8) + .ValueMatrix(lngR, 9)
               .TextMatrix(lngR, 11) = IIf(IsNull(P_adoRec("ǰ��ױ԰�")), 0, P_adoRec("ǰ��ױ԰�"))
               .TextMatrix(lngR, 12) = "��"
               .TextMatrix(lngR, 13) = IIf(IsNull(P_adoRec("����")), 0, P_adoRec("����"))
               .TextMatrix(lngR, 14) = "��"
               If P_adoRec("���࿩��") = 1 Then
                  .Cell(flexcpChecked, lngR, 15) = flexChecked    '1
               Else
                  .Cell(flexcpChecked, lngR, 15) = flexUnchecked  '2
               End If
               .Cell(flexcpText, lngR, 15) = "����"
               Select Case P_adoRec("�ۼ�����")
                      Case 0: .Cell(flexcpText, lngR, 16) = "�ŷ�"
                      Case 1: .Cell(flexcpText, lngR, 16) = "����"
                      Case 2: .Cell(flexcpText, lngR, 16) = "�ϰ�"
               End Select
               Select Case P_adoRec("�ݾױ���")
                      Case 0: .Cell(flexcpText, lngR, 17) = "����"
                      Case 1: .Cell(flexcpText, lngR, 17) = "��ǥ"
                      Case 2: .Cell(flexcpText, lngR, 17) = "����"
                      Case 3: .Cell(flexcpText, lngR, 17) = "�ܻ�"
                      Case Else: .Cell(flexcpText, lngR, 17) = "����"
               End Select
               Select Case P_adoRec("��û����")
                      Case 0: .Cell(flexcpText, lngR, 18) = "����"
                      Case 1: .Cell(flexcpText, lngR, 18) = "û��"
               End Select
               If P_adoRec("�̼�����") = 1 Then
                  .Cell(flexcpChecked, lngR, 19) = flexChecked    '1
               Else
                  .Cell(flexcpChecked, lngR, 19) = flexUnchecked  '2
               End If
               .Cell(flexcpText, lngR, 19) = "�̼�"
               Select Case P_adoRec("��뱸��")
                      Case 0: .Cell(flexcpText, lngR, 20) = "����"
                      Case 9: .Cell(flexcpText, lngR, 20) = "����"
               End Select
               'If .TextMatrix(lngR, 0) = PB_regUserinfoU.UserBranchCode Then
               '   lngRR = lngR
               'End If
               '��꼭 �հ�ݾ� ���
               lblTotMny.Caption = Format(Vals(lblTotMny.Caption) + .ValueMatrix(lngR, 8), "#,#.00")
               P_adoRec.MoveNext
            Loop
            P_adoRec.Close
            If lngRR = 0 Then
               .Row = lngRR       'vsfg1_EnterCell �ڵ�����(���� �Ѱ� �϶��� �ڵ����� ����)
               If .Rows > PC_intRowCnt1 Then
                  '.TopRow = .Rows - PC_intRowCnt1 + .FixedRows
                  .TopRow = 1
               End If
            Else
               .Row = lngRR       'vsfg1_EnterCell �ڵ�����(���� �Ѱ� �϶��� �ڵ����� ����)
               If .Rows > PC_intRowCnt1 Then
                  .TopRow = .Row
               End If
            End If
            'vsfg1_EnterCell       'vsfg1_EnterCell �ڵ�����(���� �Ѱ� �϶��� ������ �ڵ�����)
            '.SetFocus
       End With
    End If
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "��꼭 �б� ����"
    Unload Me
    Exit Sub
End Sub

'+----------------------+
'/// ���ݰ�꼭 ��� ///
'+----------------------+
Private Sub SubPrint_TaxBill(strBranchCode As String, strMakeYear As String, lngLogCnt1 As Long, lngLogCnt2 As Long)
Dim strSQL               As String
Dim P                    As Printer
Dim strPort              As String
Dim intFile              As Integer
Dim blnEof               As Boolean
Dim intPrtCnt            As Integer
Dim strPrtLine           As String
Dim inti                 As Integer
Dim C_TMargin            As Integer  'Top Margin
Dim C_LMargin            As Integer  'Left Margin
Dim intA                 As Integer
Dim SW_A                 As Integer

Dim C_intCntPerPage      As Integer
Dim intTotCnt            As Integer
Dim strBuyerCode         As String   '����ó�ڵ�

Dim A()                  As String   '��

Dim lngR                 As Long
Dim lngC                 As Long

Dim strBookNo            As String   '���ݰ�꼭 å��ȣ
Dim lngSeqNo             As Long     '���ݰ�꼭 �Ϸù�ȣ
    
    For Each P In Printers
        If Trim(P.DeviceName) = Trim(cboPrinter.Text) And lstPort.List(cboPrinter.ListIndex) = P.Port Then
           Set Printer = P
           Exit For
        End If
    Next
    Screen.MousePointer = vbHourglass
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.������ڵ� AS ������ڵ�, T1.�ۼ��⵵ AS �ۼ��⵵, T1.å��ȣ AS å��ȣ, T1.�Ϸù�ȣ AS �Ϸù�ȣ, " _
                  & "T1.����ó�ڵ� AS ����ó�ڵ�, ISNULL(T2.����ڹ�ȣ, '') AS ��Ϲ�ȣ, ISNULL(T2.����ó��, '') AS ��ȣ���θ�, " _
                  & "ISNULL(T2.��ǥ�ڸ�, '') AS ����, (ISNULL(T2.�ּ�, '') + SPACE(1) + ISNULL(T2.����, '')) AS ������ּ�, " _
                  & "ISNULL(����, '') AS ����, ISNULL(����, '') AS ����, " _
                  & "T1.�ۼ����� AS �ۼ�����, T1.���ް��� AS ���ް���, T1.���� AS ����, T1.ǰ��ױ԰� AS ǰ��ױ԰�, T1.���� AS ����, " _
                  & "T1.�ݾױ��� AS �ݾױ���, T1.��û���� AS ��û���� " _
             & "FROM ���ݰ�꼭 T1 " _
             & "LEFT JOIN ����ó T2 ON T2.����ó�ڵ� = T1.����ó�ڵ� " _
            & "WHERE T1.������ڵ� = '" & strBranchCode & "' AND T1.�ۼ��⵵ = '" & strMakeYear & "' " _
              & "AND T1.å��ȣ = " & lngLogCnt1 & " AND T1.�Ϸù�ȣ = " & lngLogCnt2 & " AND T1.��뱸�� = 0 " _
            & "ORDER BY T1.������ڵ�, T1.�ۼ��⵵, T1.å��ȣ, T1.�Ϸù�ȣ "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       Screen.MousePointer = vbDefault
       Exit Sub
    Else
       intTotCnt = P_adoRec.RecordCount
       C_TMargin = 3
       C_LMargin = 20
       
       ReDim A(intTotCnt, 23)
       
       Do Until P_adoRec.EOF
          A(intA, 0) = P_adoRec("å��ȣ")
          A(intA, 1) = P_adoRec("�Ϸù�ȣ")
          A(intA, 2) = P_adoRec("����ó�ڵ�")
          A(intA, 3) = P_adoRec("��Ϲ�ȣ")
          A(intA, 4) = P_adoRec("��ȣ���θ�")
          A(intA, 5) = P_adoRec("����")
          A(intA, 6) = P_adoRec("������ּ�")
          A(intA, 7) = P_adoRec("����")
          A(intA, 8) = P_adoRec("����")
          A(intA, 9) = P_adoRec("�ۼ�����")
          A(intA, 10) = Mid(P_adoRec("�ۼ�����"), 5, 2)         '��
          A(intA, 11) = Mid(P_adoRec("�ۼ�����"), 7, 2)         '��
          A(intA, 12) = PADR(P_adoRec("ǰ��ױ԰�"), 20, "") & " ��"  'ǰ�� �� �԰�
          A(intA, 13) = Format(P_adoRec("����"), "#") & "��"    '����
          A(intA, 14) = ""                                      '�ܰ�
          A(intA, 15) = P_adoRec("���ް���")                    '���ް���
          A(intA, 16) = P_adoRec("����")                        '����
          A(intA, 17) = P_adoRec("���ް���") + P_adoRec("����") '�հ�ݾ�
          A(intA, 18) = 0                                       '����
          A(intA, 19) = 0                                       '��ǥ
          A(intA, 20) = 0                                       '����
          A(intA, 21) = 0                                       '�ܻ�̼���
          A(intA, 22) = P_adoRec("��û����")                    '0.������, 1.û����
          intA = intA + 1
          P_adoRec.MoveNext
       Loop
       P_adoRec.Close
    End If
    'strPort = x.Port                  '��)\\Gp202\hp 'Print On Printer
    'strPort = "C:\Documents\���ݰ�꼭.TXT"
    'intFile = FreeFile
    'Open strPort For Output As #intFile
    Printer.PaperSize = vbPRPSA4          '��������
    Printer.Orientation = vbPRORPortrait  '�������� [ vbPRORPortrait(����), vbPRORLandscape(����) ]
    Printer.FontName = "����ü"
    Printer.FontUnderline = False
    Printer.FontSize = 8
    Printer.FontBold = False
    For intA = LBound(A, 1) To UBound(A, 1) - 1
        '��
        'HEAD
        SubPrint_TaxBill_HEAD_1 C_TMargin, C_LMargin, A(intA, 0), A(intA, 1), A(intA, 3), A(intA, 4), A(intA, 5), _
                                                      A(intA, 6), A(intA, 7), A(intA, 8), A(intA, 9), A(intA, 15), A(intA, 16)
        'BODY
        '10.��, 11.��, 12.ǰ��ױ԰�, 13.����, 14.�ܰ�, 15.���ް���, 16.����
        Printer.Print Space(C_LMargin + 10) & A(intA, 10) & A(intA, 11) & Space(1) _
                     & PADR(A(intA, 12), 40, "") & PADL(A(intA, 13), 8, "") & Space(10) _
                     & PADL(Format(Vals(A(intA, 15)), "#,0"), 12, "") & Space(1) _
                     & PADL(Format(Vals(A(intA, 16)), "#,0"), 12, "")
        Printer.Print Space(C_LMargin + 30) & "--- �� �� �� �� ---"
        For inti = 1 To 4: Printer.Print "": Next inti
        'FOOT
        Printer.Print Space(C_LMargin + 100) & "***"
        Printer.Print Space(C_LMargin + 10) & PADL(Format(Vals(A(intA, 17)), "#,0"), 12, "") & Space(30) _
                                            & PADL(Format(Vals(A(intA, 18)), "#,0"), 12, "")
        For inti = 1 To 2: Printer.Print "": Next inti
        '��
        'HEAD
        SubPrint_TaxBill_HEAD_1 C_TMargin, C_LMargin, A(intA, 0), A(intA, 1), A(intA, 3), A(intA, 4), A(intA, 5), _
                                                      A(intA, 6), A(intA, 7), A(intA, 8), A(intA, 9), A(intA, 15), A(intA, 16)
        'BODY
        '10.��, 11.��, 12.ǰ��ױ԰�, 13.����, 14.�ܰ�, 15.���ް���, 16.����
        Printer.Print Space(C_LMargin + 10) & A(intA, 10) & A(intA, 11) & Space(1) _
                     & PADR(A(intA, 12), 40, "") & PADL(A(intA, 13), 8, "") & Space(10) _
                     & PADL(Format(Vals(A(intA, 15)), "#,0"), 12, "") & Space(1) _
                     & PADL(Format(Vals(A(intA, 16)), "#,0"), 12, "")
        Printer.Print Space(C_LMargin + 30) & "--- �� �� �� �� ---"
        For inti = 1 To 4: Printer.Print "": Next inti
        'FOOT(17.�հ�ݾ�, 18.����, 19.��ǥ, 20.����, 21.�ܻ�̼���)
        If A(intA, 22) = "0" Then '������
           Printer.Print Space(C_LMargin + 100) & "***"
           Printer.Print Space(C_LMargin + 10) & PADL(Format(Vals(A(intA, 17)), "#,0"), 12, "") & Space(30) _
                                            & PADL(Format(Vals(A(intA, 21)), "#,0"), 12, "")
        Else
           Printer.Print ""
           Printer.Print Space(C_LMargin + 10) & PADL(Format(Vals(A(intA, 17)), "#,0"), 12, "") & Space(30) _
                                            & PADL(Format(Vals(A(intA, 21)), "#,0"), 12, "") & Space(2) & "***"
        End If
        Printer.NewPage
    Next intA
    Erase A
    Printer.EndDoc
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_CRYSTAL_REPORTS:
    MsgBox Err.Number & Space(1) & Err.Description
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "���⳻�� �б� ����"
    Unload Me
    Exit Sub
End Sub

Private Sub SubPrint_TaxBill_HEAD_1(C_TMargin As Integer, C_LMargin As Integer, _
                                    A0 As String, A1 As String, _
                                    A3 As String, A4 As String, A5 As String, _
                                    A6 As String, A7 As String, A8 As String, _
                                    A9 As String, A15 As String, A16 As String)
Dim aryEnterNo12(12) As String
Dim aryEnterNo24(24) As String
Dim strEnterNo23     As String
Dim inti             As Integer
Dim intBlankCnt      As Integer '������
    'A0.å��ȣ, A1.�Ϸù�ȣ, A3.��Ϲ�ȣ, A4.��ȣ, A5.����, A6.�ּ�, A7.����, A8.����, A9.�ŷ�����, A15.���ް���, A16.����
    For inti = 1 To C_TMargin: Printer.Print "": Next inti
    ' ��Ϲ�ȣ ����
    For inti = 1 To 12
        aryEnterNo12(inti) = Mid(A3, inti, 1)
    Next inti
    For inti = 1 To 12
        If inti = 1 Then
           aryEnterNo24(inti) = aryEnterNo12(inti): aryEnterNo24(inti + 1) = " "
        Else
           aryEnterNo24(inti * 2 - 1) = aryEnterNo12(inti): aryEnterNo24(inti * 2) = " "
        End If
    Next inti
    For inti = 1 To 23
        strEnterNo23 = strEnterNo23 + aryEnterNo24(inti)
    Next inti
    For inti = 1 To 1: Printer.Print "": Next inti
    'å��ȣ
    Printer.Print Space(C_LMargin + 80) & PADR(A0, 6, "")
    '�Ϸù�ȣ
    Printer.Print Space(C_LMargin + 80) & PADR(A1, 6, "")
    '��Ϲ�ȣ
    Printer.Print Space(C_LMargin + 50) & PADR(strEnterNo23, 23, "")
    For inti = 1 To 1: Printer.Print "": Next inti
    'Printer.Print Space(C_LMargin + 50) & Chr(27) & "W1" & PADC(strEnterNo, 14, "") & Chr(27) & "W0"
    '��ȣ, ����
    Printer.Print Space(C_LMargin + 50) & PADR(A4, 14, "") & Space(8) & PADR(A5, 10, "")
    '�ּ�(�۰�)
    Printer.FontSize = 6
    Printer.Print Space(C_LMargin + 70) & PADR(A6, 70, "")
    For inti = 1 To 1: Printer.Print "": Next inti
    '����, ����(�۰�)
    Printer.FontSize = 6
    Printer.Print Space(C_LMargin + 70) & PADR(A7, 14, "") & Space(3) & PADR(A8, 14, "")
    For inti = 1 To 1: Printer.Print "": Next inti
    '�ۼ������, ������, ���ް���, ����
    Printer.FontSize = 8
    intBlankCnt = 11 - Len(Trim(A15))
    Printer.Print Space(C_LMargin + 10) & Mid(A9, 1, 4) & Mid(A9, 5, 2) & Mid(A9, 7, 2) _
                      & PADC(intBlankCnt, 3, "") & PADL(A15, 10, "") & Space(1) & PADL(A16, 10, "")
    For inti = 1 To 3: Printer.Print "": Next inti
End Sub


