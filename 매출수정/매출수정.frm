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
   Begin VSFlex7Ctl.VSFlexGrid vsfg2 
      Height          =   5777
      Left            =   60
      TabIndex        =   6
      Top             =   4248
      Width           =   15195
      _cx             =   26802
      _cy             =   10190
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
      FocusRect       =   2
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
   Begin VB.Frame Frame1 
      Appearance      =   0  '���
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   60
      TabIndex        =   12
      Top             =   0
      Width           =   15195
      Begin VB.OptionButton optPrtChk0 
         Caption         =   "�ŷ�����"
         Height          =   255
         Left            =   6600
         TabIndex        =   21
         Top             =   150
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optPrtChk1 
         Caption         =   "���ݰ�꼭"
         Height          =   255
         Left            =   6600
         TabIndex        =   20
         Top             =   390
         Width           =   1335
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   375
         Left            =   7980
         Picture         =   "�������.frx":0000
         Style           =   1  '�׷���
         TabIndex        =   16
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
         Picture         =   "�������.frx":0963
         Style           =   1  '�׷���
         TabIndex        =   7
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   13635
         Picture         =   "�������.frx":1308
         Style           =   1  '�׷���
         TabIndex        =   9
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   390
         Left            =   12510
         Picture         =   "�������.frx":1C56
         Style           =   1  '�׷���
         TabIndex        =   8
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
         TabIndex        =   0
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00008000&
         BorderStyle     =   1  '���� ����
         Caption         =   "�ŷ����� ��ȸ�� ����"
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
         TabIndex        =   13
         Top             =   180
         Width           =   4650
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid vsfg1 
      Height          =   2475
      Left            =   60
      TabIndex        =   5
      Top             =   1695
      Width           =   15195
      _cx             =   26802
      _cy             =   4366
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
      FocusRect       =   2
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
      Height          =   1035
      Left            =   60
      TabIndex        =   11
      Top             =   630
      Width           =   15195
      Begin VB.CheckBox chkTaxBillPrint 
         Caption         =   "���ݰ�꼭"
         Height          =   255
         Left            =   12120
         TabIndex        =   30
         Top             =   600
         Width           =   1215
      End
      Begin VB.ListBox lstPort 
         Height          =   240
         Left            =   7800
         TabIndex        =   29
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox cboPrinter 
         Height          =   300
         Left            =   12120
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   28
         Top             =   200
         Width           =   2600
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "������ �ŷ����� �μ�"
         Height          =   375
         Left            =   13440
         TabIndex        =   27
         Top             =   550
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '����
         Index           =   1
         Left            =   4440
         TabIndex        =   2
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   0
         Left            =   3240
         MaxLength       =   8
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
         Format          =   57737217
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
         Format          =   57737217
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   6
         Left            =   6600
         TabIndex        =   26
         Top             =   645
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   5
         Left            =   4440
         TabIndex        =   25
         Top             =   645
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "]"
         Height          =   240
         Index           =   4
         Left            =   7520
         TabIndex        =   24
         Top             =   645
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "["
         Height          =   240
         Index           =   3
         Left            =   2520
         TabIndex        =   23
         Top             =   645
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�ŷ�����"
         Height          =   240
         Index           =   2
         Left            =   1200
         TabIndex        =   22
         Top             =   650
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��ü�ݾ�"
         Height          =   240
         Index           =   1
         Left            =   8040
         TabIndex        =   19
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
         Left            =   9360
         TabIndex        =   18
         Top             =   285
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "[Home]"
         Height          =   240
         Index           =   15
         Left            =   2520
         TabIndex        =   17
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����ó�ڵ�"
         Height          =   240
         Index           =   0
         Left            =   1200
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   285
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
' ������ Table   : �����, ����ó, �������⳻��
' ��  ��  ��  �� :
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived         As Boolean
Private P_adoRec             As New ADODB.Recordset
Private P_adoRecW            As New ADODB.Recordset
Private P_intButton          As Integer
Private P_strFindString2     As String
Private Const PC_intRowCnt1  As Integer = 8   '�׸���1�� �� ������ �� ���(FixedRows ����)
Private Const PC_intRowCnt2  As Integer = 18  '�׸���2�� �� ������ �� ���(FixedRows ����)

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

Dim p                  As Printer
Dim strDefaultPrinter  As String
Dim aryPrinter()       As String
Dim strBuffer          As String

    frmMain.SBar.Panels(4).Text = "������ ��� ����ó��(����ŷ�ó������ó��)�Ͻð�, ��꼭����(�߰���)�� ���⼼�ݰ�꼭��ο��� ����˴ϴ�."
    If P_blnActived = False Then
       Screen.MousePointer = vbHourglass
       On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
       
       'Printer Setting and Seaching(API)
       strBuffer = Space(1024)
       inti = GetProfileString("windows", "Device", "", strBuffer, Len(strBuffer))
       aryPrinter = Split(strBuffer, ",")
       strDefaultPrinter = Trim(Trim(aryPrinter(0)))
       For Each p In Printers
           cboPrinter.AddItem Trim(p.DeviceName)
           lstPort.AddItem p.Port
       Next
       For inti = 0 To cboPrinter.ListCount - 1
           cboPrinter.ListIndex = inti
           If UCase(Trim(cboPrinter.Text)) = UCase(Trim(strDefaultPrinter)) Then
              Exit For
           End If
       Next inti
       '---
       Subvsfg1_INIT  '�ŷ��հ�
       Subvsfg2_INIT  '�ŷ�����
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
    If (Index = 0 And (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn)) Then  '����ó�˻�
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
                     .Text = UPPER(Trim(.Text))
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
         If .MouseRow >= .FixedRows Then
            If (.MouseCol = 14) Then     '���ݱ���
                If Button = vbLeftButton Then
                  .Select .MouseRow, .MouseCol
                  .EditCell
                End If
            ElseIf _
               (.MouseCol = 17) Then     '��������
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
    With vsfg1
         If Row >= .FixedRows Then
            If (Col = 14) Then '���ݱ���
               If .Cell(flexcpChecked, Row, Col) <> .EditText Then
                  blnModify = True
               End If
            ElseIf _
               (Col = 17) Then '��������
               If .TextMatrix(Row, 1) <> .EditText Then
                  .EditText = Format(Replace(.EditText, "-", ""), "0000-00-00")
                  If Not ((Len(Trim(.EditText)) = 10) And IsDate(.EditText) And Val(.EditText) > 2000) Then
                     Beep
                     Cancel = True
                  Else
                     blnModify = True
                  End If
               End If
            End If
            '����ǥ��
            If blnModify = True Then
               .Cell(flexcpBackColor, Row, Col, Row, Col) = vbRed
               .Cell(flexcpForeColor, Row, Col, Row, Col) = vbWhite
            Else
               .Cell(flexcpBackColor, Row, Col, Row, Col) = vbWhite
               .Cell(flexcpForeColor, Row, Col, Row, Col) = vbBlack
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
            '.Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlack
            '.Cell(flexcpForeColor, 0, .MouseCol, 0, .MouseCol) = vbRed
            'strData = Trim(.Cell(flexcpData, .Row, 3))
            'Select Case .MouseCol
            '       Case 0
            '            .ColSel = 3
            '            .ColSort(0) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .ColSort(1) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .ColSort(2) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .Sort = flexSortUseColSort
            '       Case 1
            '            .ColSel = 3
            '            .ColSort(0) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .ColSort(1) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .ColSort(2) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .Sort = flexSortUseColSort
            '       Case 12
            '            .ColSel = 3
            '            For lngC = 0 To 8: .ColSort(lngC) = flexSortNone: Next lngC
            '            .ColSort(9) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .ColSort(10) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .ColSort(11) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .Sort = flexSortUseColSort
            '       Case Else
            '            .Sort = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            'End Select
            'If .FindRow(strData, , 3) > 0 Then
            '   .Row = .FindRow(strData, , 3)
            'End If
            'If PC_intRowCnt1 < .Rows Then
            '   .TopRow = .Row
            'End If
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
         If NewRow < 1 Then Exit Sub
         If NewRow <> OldRow Then
            For lngR2 = 1 To vsfg2.Rows - 1
                vsfg2.RowHidden(lngR2) = True
            Next lngR2
            If NewRow > 0 Then
               For lngR1 = .ValueMatrix(.Row, 15) To .ValueMatrix(.Row, 16)
                   If vsfg2.TextMatrix(lngR1, 28) = "D" Then
                      vsfg2.RowHidden(lngR1) = True
                   Else
                      vsfg2.RowHidden(lngR1) = False
                      lngCnt = lngCnt + 1
                      vsfg2.TextMatrix(lngR1, 0) = lngCnt '����
                   End If
               Next lngR1
            End If
            If PC_intRowCnt2 < lngCnt Then
               vsfg2.TopRow = vsfg2.Row
            End If
            vsfg2.Row = 0
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
End Sub

'+------------+
'/// vsfg2 ///
'+------------+
Private Sub vsfg2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    With vsfg2
         .ToolTipText = ""
         If .MouseRow < .FixedRows Or .MouseCol < 0 Then
            Exit Sub
         End If
         .ToolTipText = .TextMatrix(.MouseRow, .MouseCol)
    End With
End Sub
Private Sub vsfg2_DblClick()
    With vsfg1
         If .Row >= .FixedRows Then
             vsfg2_KeyDown vbKeyF1, 0  '����ü��˻�
         End If
    End With
End Sub
Private Sub vsfg2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    With vsfg2
         P_intButton = Button
         If .Row >= .FixedRows Then
            If (.Col = 16) Then      '����
                If Button = vbLeftButton Then
                  .Select .Row, .Col
                  .EditCell
                End If
            ElseIf _
               (.Col = 18) Then      '��꼭���࿩��
               If Button = vbLeftButton Then
                  .Select .Row, .Col
                  .EditCell
                End If
            ElseIf _
                (.Col = 23) Then     '���ܰ�
                If Button = vbLeftButton Then
                  .Select .Row, .Col
                  .EditCell
                End If
            ElseIf _
                (.Col = 24) Then     '���ΰ�
                If Button = vbLeftButton Then
                  .Select .Row, .Col
                  .EditCell
                End If
            ElseIf _
                (.Col = 27) Then     '����
                If Button = vbLeftButton Then
                  .Select .Row, .Col
                  .EditCell
                End If
            End If
         End If
    End With
End Sub
Sub vsfg2_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim blnModify As Boolean
Dim curTmpMny As Currency
    With vsfg2
         If Row >= .FixedRows Then
            If (Col = 16) Then  '����
               If .TextMatrix(Row, Col) <> .EditText Then
                  If (IsNumeric(.EditText) = False Or _
                     Fix(Vals(.EditText)) < Vals(.EditText) Or IsNumeric(Right(.EditText, 1)) = False) Then
                     Beep
                     Cancel = True
                  Else
                     blnModify = True
                     curTmpMny = .ValueMatrix(Row, 26)
                     .TextMatrix(Row, 26) = Vals(.EditText) * .ValueMatrix(Row, 23)
                  End If
               End If
            ElseIf _
               (Col = 18) Then  '��꼭���࿩��
               If (Len(.TextMatrix(Row, 9)) = 0) Then '����ó�� ����
                  .Cell(flexcpChecked, Row, 18, Row, 18) = flexUnchecked
                  Beep
                  Cancel = True
                  Exit Sub
               End If
               If .Cell(flexcpChecked, Row, Col) <> .EditText Then
                  blnModify = True
               End If
            ElseIf _
               (Col = 23) Then  '���ܰ�
               If .TextMatrix(Row, Col) <> .EditText Then
                  If (IsNumeric(.EditText) = False Or Vals(.EditText) < 0 Or _
                     IsNumeric(Right(.EditText, 1)) = False) Then                                            '�Ҽ������� ��밡
                     'fix(Vals(.EditText)) < Vals(.EditText) Or IsNumeric(Right(.EditText, 1)) = False) Then '�Ҽ������� ���Ұ�
                     Beep
                     Cancel = True
                  Else
                     blnModify = True
                     .TextMatrix(Row, 24) = Fix(Vals(.EditText) * (PB_curVatRate))  '�ΰ���
                     curTmpMny = .ValueMatrix(Row, 26)
                     .TextMatrix(Row, 25) = Vals(.EditText) + .ValueMatrix(Row, 24)
                     .TextMatrix(Row, 26) = .ValueMatrix(Row, 16) * Vals(.EditText)
                  End If
               End If
            ElseIf _
               (Col = 24) Then  '���ΰ�
               If .TextMatrix(Row, Col) <> .EditText Then
                  If (IsNumeric(.EditText) = False Or Vals(.EditText) < 0 Or _
                     Fix(Vals(.EditText)) < Vals(.EditText) Or IsNumeric(Right(.EditText, 1)) = False) Then
                     Beep
                     Cancel = True
                  Else
                     blnModify = True
                     curTmpMny = .ValueMatrix(Row, 26)
                     .TextMatrix(Row, 25) = .ValueMatrix(Row, 23) + Vals(.EditText)
                     .TextMatrix(Row, 26) = .ValueMatrix(Row, 16) * .ValueMatrix(Row, 23)
                  End If
               End If
            ElseIf _
               (Col = 27) Then '���� ���� �˻�
               If .TextMatrix(Row, Col) <> .EditText Then
                  If Not (LenH(Trim(.EditText)) <= 50) Then
                     Beep
                     Cancel = True
                  Else
                     blnModify = True
                  End If
               End If
            End If
            '����ǥ�� + �ݾ�����
            If blnModify = True Then
               If .TextMatrix(Row, 28) = "" Then
                  .TextMatrix(Row, 28) = "U"
               End If
               .Cell(flexcpBackColor, Row, Col, Row, Col) = vbRed
               .Cell(flexcpForeColor, Row, Col, Row, Col) = vbWhite
               Select Case Col
                      Case 16, 23, 24
                           vsfg1.TextMatrix(vsfg1.Row, 6) = vsfg1.ValueMatrix(vsfg1.Row, 6) - curTmpMny + .ValueMatrix(Row, 26)
                           lblTotMny.Caption = Format(Vals(lblTotMny.Caption) - curTmpMny + .ValueMatrix(Row, 26), "#,#.00")
                      Case Else
               End Select
            End If
         End If
    End With
End Sub
Private Sub vsfg2_EnterCell()
Dim lngR As Long
Dim lngC As Long
    With vsfg2
         .Editable = flexEDNone
         If .Row >= .FixedRows Then
             Select Case .Col
                    Case 16, 23, 27
                         .Editable = flexEDKbdMouse
                         vsfg2_MouseUp vbLeftButton, 0, 0, 0
             End Select
         End If
    End With
End Sub
Private Sub vsfg2_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
Dim lngR1  As Long
Dim lngR2  As Long
Dim lngC   As Long
Dim lngCnt As Long
    With vsfg2
         'If NewRow < 1 Then Exit Sub
         'If NewRow <> OldRow Then
         '   .Row = NewRow
         'End If
         'If NewCol <> OldCol Then
         '   .Col = NewCol
         'End If
    End With
End Sub
Private Sub vsfg2_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
Dim lngR    As Long
    With vsfg2
         If KeyCode = vbKeyReturn Then
            If Col = 16 Then
               .Col = 23
            ElseIf _
               Col = 23 Then
               .Col = 27
            ElseIf _
               Col = 27 Then
               For lngR = Row To vsfg1.ValueMatrix(vsfg1.Row, 16)
                   If .RowHidden(lngR) = False Then
                      Exit For
                   End If
               Next lngR
               If lngR <> vsfg1.ValueMatrix(vsfg1.Row, 16) Then
                  .Col = 16: .LeftCol = 15
                  .Row = .Row + 1
                  If .ValueMatrix(lngR, 0) > PC_intRowCnt2 Then
                     .TopRow = .TopRow + 1
                  End If
               End If
            End If
         End If
    End With
End Sub
Private Sub vsfg2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lngR     As Long
Dim lngPos   As Long
Dim blnDupOK As Boolean
Dim strTime  As String
Dim strHH    As String
Dim strMM    As String
Dim strSS    As String
Dim strMS    As String
Dim intRetVal As Integer
Dim CtrlDown  As Variant
    With vsfg2
         If (.Row >= .FixedRows) Then     '�����ü��˻�
            If KeyCode = vbKeyF2 And (Len(vsfg1.TextMatrix(vsfg1.Row, 5)) > 0) And _
              (Len(.TextMatrix(.Row, 4)) > 0) Then
               PB_strFMCCallFormName = "frm�������"
               PB_strMaterialsCode = .TextMatrix(.Row, 4)
               PB_strMaterialsName = .TextMatrix(.Row, 5)
               PB_strSupplierCode = vsfg1.TextMatrix(vsfg1.Row, 4)
               frm�����ü��˻�.Show vbModal
               If Len(PB_strMaterialsCode) <> 0 Then
                  PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
               End If
            End If
         End If
    End With
    With vsfg2
         '������ ���� ��� �߰�
         If .Row = 0 And KeyCode = vbKeyInsert Then
            For lngR = vsfg1.ValueMatrix(vsfg1.Row, 15) To vsfg1.ValueMatrix(vsfg1.Row, 16)
                If .RowHidden(lngR) = False Then
                   lngPos = lngPos + 1
                End If
            Next lngR
            If lngPos = 0 Then .Row = vsfg1.ValueMatrix(vsfg1.Row, 16)
         End If
         If .Row >= .FixedRows Then
            If KeyCode = vbKeyF1 Then  '����ü��˻�
               'If (.MouseCol = 5) Then
                  PB_strFMCCallFormName = "frm�������"
                  PB_strMaterialsCode = .TextMatrix(.Row, 4)
                  PB_strMaterialsName = .TextMatrix(.Row, 5)
                  PB_strSupplierCode = .TextMatrix(.Row, 8)
                  frm����ü��˻�.Show vbModal
                  If Len(PB_strMaterialsCode) <> 0 Then
                     PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                  End If
               'ElseIf _
               '   (.Col = 9) Then      '����ó�˻�(Not Used)
               '   PB_strSupplierCode = .TextMatrix(.Row, 8)
               '   PB_strSupplierName = .TextMatrix(.Row, 9)
               '   frm����ó�˻�.Show vbModal
               '   If Len(PB_strSupplierCode) <> 0 Then
               '      'For lngR = vsfg1.ValueMatrix(vsfg1.Row, 15) To vsfg1.ValueMatrix(vsfg1.Row, 16) '�����ڵ�(������) + ����ó�ڵ�
               '      '    If .Row <> lngR And .TextMatrix(lngR, 0) = .TextMatrix(.Row, 0) And _
               '      '       .TextMatrix(lngR, 4) = PB_strSupplierCode Then
               '      '       blnDupOK = True
               '      '       Exit For
               '      '    End If
               '      'Next lngR
               '      If blnDupOK = False Then
               '         If PB_strSupplierCode <> .TextMatrix(.Row, 8) Then
               '            .Cell(flexcpBackColor, .Row, 9, .Row, 9) = vbRed
               '            .Cell(flexcpForeColor, .Row, 9, .Row, 9) = vbWhite
               '         End If
               '         .TextMatrix(.Row, 8) = PB_strSupplierCode
               '         .TextMatrix(.Row, 9) = PB_strSupplierName
               '         If .TextMatrix(.Row, 28) = "" Then
               '            .TextMatrix(.Row, 28) = "U"
               '         End If
               '      End If
               '   End If
               'End If
            ElseIf _
               KeyCode = vbKeyInsert And cmdSave.Enabled = True Then '�ŷ����� �߰�
               .AddItem "", .Row + 1
               .Row = .Row + 1
               .TopRow = .Row
               .TextMatrix(.Row, 0) = .ValueMatrix(.Row - 1, 0) + 1 '����
               .TextMatrix(.Row, 1) = .TextMatrix(.Row - 1, 1)   '������ڵ�
               .TextMatrix(.Row, 2) = .TextMatrix(.Row - 1, 2)   '�ŷ�����
               .TextMatrix(.Row, 3) = .TextMatrix(.Row - 1, 3)   '�ŷ���ȣ
               .TextMatrix(.Row, 8) = vsfg1.TextMatrix(vsfg1.Row, 4) '�ŷ��� ����ó�ڵ�
               .TextMatrix(.Row, 9) = vsfg1.TextMatrix(vsfg1.Row, 5) '�ŷ��� ����ó��
               .Cell(flexcpChecked, .Row, 18) = flexChecked      '��꼭���࿩��
               .Cell(flexcpText, .Row, 18) = "�� ��"
               .Cell(flexcpAlignment, .Row, 18, .Row, 18) = flexAlignLeftCenter
               .TextMatrix(.Row, 28) = "I"                       'SQL����
               '��������ð�
               strTime = .TextMatrix(.Row - 1, 29)
               If .Row <= vsfg1.ValueMatrix(vsfg1.Row, 16) Then  '�ŷ���ȣ�� ������ �ƴϸ�
                  strTime = Format(Fix((.ValueMatrix(.Row - 1, 29) + .ValueMatrix(.Row + 1, 29)) / 2), "000000000")
                  '�߰� �������� �˻�
                  If (strTime = .TextMatrix(.Row - 1, 29)) Or (strTime = .TextMatrix(.Row - 1, 29)) Then
                     MsgBox "�� �࿡�� �� �̻� �߰� �� �� �����ϴ�. �ٸ� �࿡ �߰��ϼ���.", vbCritical + vbDefaultButton1, "�߰�"
                     .RemoveItem (.Row)
                     Exit Sub
                  End If
               Else                                              '�ŷ���ȣ�� �������̸�
                  strTime = Format((Val(strTime) + 10000), "000000000")
                  strHH = Mid(strTime, 1, 2): strMM = Mid(strTime, 3, 2): strSS = Mid(strTime, 5, 2): strMS = Mid(strTime, 7, 3)
                  If Val(strMS) > 999 Then
                     strMS = Format(0, "000")
                     strSS = Format(Val(strMM) + 1, "00")
                  End If
                  If Val(strSS) > 59 Then
                     strSS = Format(Val(strSS) - 60, "00")
                     strMM = Format(Val(strMM) + 1, "00")
                  End If
                  If Val(strMM) > 59 Then
                     strMM = Format(Val(strMM) - 60, "00")
                     strHH = Format(Val(strHH) + 1, "00")
                  End If
                  strTime = strHH & strMM & strSS & strMS
               End If
               .TextMatrix(.Row, 29) = strTime                   '��������ð�
               PB_strFMCCallFormName = "frm�������"
               PB_strMaterialsCode = .TextMatrix(.Row, 4)
               PB_strMaterialsName = .TextMatrix(.Row, 5)
               PB_strSupplierCode = .TextMatrix(.Row, 8)
               frm����ü��˻�.Show vbModal
               If Len(PB_strMaterialsCode) <> 0 Then '�������� �����̸�
                  PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                  '����
                  For lngR = (.Row + 1) To (vsfg1.ValueMatrix(vsfg1.Row, 16) + 1)
                      If .TextMatrix(lngR, 28) <> "D" Then
                         .TextMatrix(lngR, 0) = .ValueMatrix(lngR, 0) + 1
                      End If
                  Next lngR
                  
                  vsfg1.TextMatrix(vsfg1.Row, 16) = vsfg1.ValueMatrix(vsfg1.Row, 16) + 1
                  For lngR = 1 To vsfg1.Rows - 1
                      If (lngR <> .Row) Then
                         If (vsfg1.ValueMatrix(lngR, 15) >= .Row) Then
                            vsfg1.TextMatrix(lngR, 15) = vsfg1.ValueMatrix(lngR, 15) + 1
                            vsfg1.TextMatrix(lngR, 16) = vsfg1.ValueMatrix(lngR, 16) + 1
                         End If
                      End If
                  Next lngR
               Else
                  .RemoveItem (.Row)
                  .Row = .Row - 1
               End If
            ElseIf _
               KeyCode = vbKeyDelete And .Col = 9 Then  '����ó ����
               If (Len(.TextMatrix(.Row, 9)) <> 0) Then
                  .TextMatrix(.Row, 8) = "": .TextMatrix(.Row, 9) = ""
                  .Cell(flexcpBackColor, .Row, 9, .Row, 9) = vbRed
                  .Cell(flexcpForeColor, .Row, 9, .Row, 9) = vbWhite
                  If .TextMatrix(.Row, 28) = "" Then
                     .TextMatrix(.Row, 28) = "U"
                  End If
               End If
               'If .Cell(flexcpChecked, .Row, 18, .Row, 18) = flexChecked Then
               '   .Cell(flexcpChecked, .Row, 18, .Row, 18) = flexUnchecked
               '   .Cell(flexcpBackColor, .Row, 18, .Row, 18) = vbRed
               '   .Cell(flexcpForeColor, .Row, 18, .Row, 18) = vbWhite
               'End If
            ElseIf _
               KeyCode = vbKeyDelete And (.Col <> 9) And (.Row > 0) And .RowHidden(.Row) = False Then
               intRetVal = MsgBox("�Է��� �ŷ������� �����Ͻðڽ��ϱ� ?", vbQuestion + vbYesNo + vbDefaultButton2, "�ŷ���������")
               If intRetVal = vbYes Then
                  .TextMatrix(.Row, 28) = "D": .TextMatrix(.Row, 0) = "0"
                  vsfg1.TextMatrix(vsfg1.Row, 6) = vsfg1.ValueMatrix(vsfg1.Row, 6) - .ValueMatrix(.Row, 26)
                  lblTotMny.Caption = Format(Vals(lblTotMny.Caption) - .ValueMatrix(.Row, 26), "#,#.00")
                  .RowHidden(.Row) = True
                  For lngR = .Row + 1 To vsfg1.ValueMatrix(vsfg1.Row, 16)
                      If .TextMatrix(lngR, 28) <> "D" Then
                         .TextMatrix(lngR, 0) = .ValueMatrix(lngR, 0) - 1
                         If lngPos = 0 Then
                            lngPos = lngR
                         End If
                      End If
                  Next lngR
                  If lngPos = 0 Then
                     For lngR = vsfg1.ValueMatrix(vsfg1.Row, 15) To .Row
                         If .TextMatrix(lngR, 28) <> "D" And lngR < .Row Then
                            lngPos = lngR
                         End If
                     Next lngR
                  End If
                  .Row = lngPos
               End If
            End If
         End If
    End With
End Sub

'+-----------+
'/// ��� ///
'+-----------+
Private Sub cmdPrint_Click()
Dim p              As Printer
Dim strSQL         As String
Dim lngLogCnt      As Long
Dim strMakeYear    As String
Dim lngLogCnt1     As Long
Dim lngLogCnt2     As Long
Dim strServerTime  As String
Dim strTime        As String
Dim lngR           As Long
Dim intChkCash     As Integer

    For Each p In Printers
        If Trim(p.DeviceName) = Trim(cboPrinter.Text) And lstPort.List(cboPrinter.ListIndex) = p.Port Then
           Set Printer = p
           Exit For
        End If
    Next
    If ((vsfg1.Rows + vsfg2.Rows) = 2) Or (vsfg1.Row < 1) Then
       Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    With vsfg1
         If .Cell(flexcpChecked, .Row, 14) = flexChecked Then '���ݸ���
            intChkCash = 1
         End If
    End With
    If optPrtChk0.Value = True Then '�ŷ������μ�
       SubPubPrint_DealBill p, PB_intPrtTypeGbn, DTOS(vsfg1.TextMatrix(vsfg1.Row, 1)), vsfg1.ValueMatrix(vsfg1.Row, 2)
    End If
    If optPrtChk1.Value = True Then '���ݰ�꼭�μ�(�ݾ��� +/- �̸�)
       If (vsfg1.ValueMatrix(vsfg1.Row, 18) = 1) And (Len(vsfg1.TextMatrix(vsfg1.Row, 12)) = 0) Then
          MsgBox "�̹� ��꼭�� �ۼ��Ǿ����� ��꼭��ȣ�� �˼� �����ϴ�. ���ݰ�꼭���� ������ϼ���.!", vbCritical + vbOKOnly, "��꼭����"
          Screen.MousePointer = vbDefault
          Exit Sub
       End If
       If Len(vsfg1.TextMatrix(vsfg1.Row, 12)) = 0 Then '���ݰ�꼭 ��ȣ�� ������
          '�����ð� ���ϱ�
          P_adoRec.CursorLocation = adUseClient
          strSQL = "SELECT RIGHT(CONVERT(VARCHAR(23),GETDATE(), 121), 12) AS �����ð� "
          On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
          P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
          strServerTime = Mid(P_adoRec("�����ð�"), 1, 2) + Mid(P_adoRec("�����ð�"), 4, 2) + Mid(P_adoRec("�����ð�"), 7, 2) _
                        + Mid(P_adoRec("�����ð�"), 10)
          P_adoRec.Close
          strTime = strServerTime
          PB_adoCnnSQL.BeginTrans
          P_adoRec.CursorLocation = adUseClient
          'å��ȣ, �Ϸù�ȣ ���ϱ�
          strMakeYear = Mid(PB_regUserinfoU.UserClientDate, 1, 4)
          strSQL = "spLogCounter '���ݰ�꼭', '" & PB_regUserinfoU.UserBranchCode + strMakeYear & "', " _
                               & "0, 0, '" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' "
          On Error GoTo ERROR_STORED_PROCEDURE
          P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
          lngLogCnt1 = P_adoRec(0)
          lngLogCnt2 = P_adoRec(1)
          P_adoRec.Close
          With vsfg1
               '���ݰ�꼭 �ۼ�(1.���ݰ�꼭, 2.���⼼�ݰ�꼭���)
               '1.���ݰ�꼭
               strSQL = "INSERT INTO ���ݰ�꼭 " _
                      & "SELECT T1.������ڵ� AS ������ڵ�, '" & strMakeYear & "' AS �ۼ��⵵, " _
                             & "" & lngLogCnt1 & " AS å��ȣ, " & lngLogCnt2 & " AS �Ϸù�ȣ," _
                             & "T1.����ó�ڵ� AS ����ó�ڵ�, T1.�ŷ����� AS �ۼ�����, " _
                            & "(SELECT TOP 1 (S2.����� + SPACE(1) + S2.�԰�)  FROM �������⳻�� S1 " _
                               & "LEFT JOIN ���� S2 ON S1.�з��ڵ� = S2.�з��ڵ� AND S1.�����ڵ� = S2.�����ڵ� " _
                              & "WHERE S1.������ڵ� = T1.������ڵ� " _
                                & "AND S1.�ŷ����� = T1.�ŷ����� AND S1.�ŷ���ȣ = " & .ValueMatrix(.Row, 2) & " " _
                                & "AND S1.������� = 2 AND S1.��뱸�� = 0 AND S1.��꼭���࿩�� = 0 " _
                                & "AND S1.����ó�ڵ� = T1.����ó�ڵ� " _
                              & "ORDER BY S1.���������, S1.�����ð�) AS ǰ��ױ԰�, "
               strSQL = strSQL + _
                              "(SELECT (COUNT(S1.�����ڵ�) - 1) FROM �������⳻�� S1 " _
                              & "WHERE S1.������ڵ� = T1.������ڵ� " _
                                & "AND S1.�ŷ����� = T1.�ŷ����� AND S1.�ŷ���ȣ = " & .ValueMatrix(.Row, 2) & " " _
                                & "AND S1.������� = 2 AND S1.��뱸�� = 0 AND S1.��꼭���࿩�� = 0 " _
                                & "AND S1.����ó�ڵ� = T1.����ó�ڵ�) AS ����, " _
                             & "SUM(T1.������ * T1.���ܰ�) AS ���ް���, SUM(T1.������ * T1.���ΰ�) AS ����, " _
                             & "" & IIf(intChkCash = 1, 0, 3) & " AS �ݾױ���, " & IIf(intChkCash = 1, 0, 1) & " AS ��û����, " _
                             & "1 AS ���࿩��, 0 AS �ۼ�����, 1 AS �̼�����, '' AS ����, 0 AS ��뱸��, " _
                             & "'" & PB_regUserinfoU.UserServerDate & "' AS ��������, '" & PB_regUserinfoU.UserCode & "' AS ������ڵ� " _
                        & "FROM �������⳻�� T1 " _
                        & "LEFT JOIN ���� T2 ON T2.�з��ڵ� = T1.�з��ڵ� AND T2.�����ڵ� = T1.�����ڵ� " _
                        & "LEFT JOIN ����ó T3 ON T3.����ó�ڵ� = T1.����ó�ڵ� " _
                       & "WHERE T1.������ڵ� = '" & .TextMatrix(.Row, 0) & "' " _
                         & "AND T1.�ŷ����� = '" & DTOS(.TextMatrix(.Row, 1)) & "' AND T1.�ŷ���ȣ = " & .ValueMatrix(.Row, 2) & " " _
                         & "AND T1.������� = 2 AND T1.��뱸�� = 0 " _
                       & "GROUP BY T1.������ڵ�, T1.�ۼ��⵵, T1.å��ȣ, T1.�Ϸù�ȣ, T1.����ó�ڵ�, T1.�ŷ����� " _
                       & "ORDER BY T1.������ڵ�, T1.�ۼ��⵵, T1.å��ȣ, T1.�Ϸù�ȣ "
               On Error GoTo ERROR_TABLE_INSERT
               PB_adoCnnSQL.Execute strSQL
               '2.���⼼�ݰ�꼭���(�ۼ��ð�:strTime)
               strSQL = "INSERT INTO ���⼼�ݰ�꼭��� " _
                      & "SELECT T1.������ڵ�, T1.�ۼ�����, '" & strTime & "', T1.����ó�ڵ�, " _
                             & "T1.ǰ��ױ԰�, T1.����, T1.���ް���, T1.����, " _
                             & "T1.�ݾױ���, T1.��û����, T1.���࿩��, T1.�ۼ�����, " _
                             & "T1.�̼�����, T1.����, T1.��뱸��, T1.��������, " _
                             & "T1.������ڵ�, T1.�ۼ��⵵, T1.å��ȣ, T1.�Ϸù�ȣ " _
                        & "FROM ���ݰ�꼭 T1 " _
                       & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                         & "AND T1.�ۼ��⵵  = '" & strMakeYear & "' AND å��ȣ = " & lngLogCnt1 & " AND T1.�Ϸù�ȣ = " & lngLogCnt2 & " "
               On Error GoTo ERROR_TABLE_INSERT
               PB_adoCnnSQL.Execute strSQL
               strSQL = "UPDATE �������⳻�� SET " _
                             & "��꼭���࿩�� = 1," _
                             & "�ۼ��⵵ = " & strMakeYear & ", " _
                             & "å��ȣ = " & lngLogCnt1 & ", " _
                             & "�Ϸù�ȣ = " & lngLogCnt2 & ", " _
                             & "�������� = '" & PB_regUserinfoU.UserServerDate & "', " _
                             & "������ڵ� = '" & PB_regUserinfoU.UserCode & "' " _
                       & "WHERE ������ڵ� = '" & .TextMatrix(.Row, 0) & "' " _
                         & "AND �ŷ����� = '" & DTOS(.TextMatrix(.Row, 1)) & "' AND �ŷ���ȣ = " & .ValueMatrix(.Row, 2) & " " _
                         & "AND ������� = 2 AND ��뱸�� = 0 AND ��꼭���࿩�� = 0 "
               On Error GoTo ERROR_TABLE_UPDATE
               PB_adoCnnSQL.Execute strSQL
               .TextMatrix(.Row, 9) = Mid(PB_regUserinfoU.UserClientDate, 1, 4)
               .TextMatrix(.Row, 10) = lngLogCnt1
               .TextMatrix(.Row, 11) = lngLogCnt2
               .TextMatrix(.Row, 12) = .TextMatrix(.Row, 9) + "-" + CStr(lngLogCnt1) + "-" + CStr(lngLogCnt2)
               .Cell(flexcpChecked, .Row, 13) = flexChecked
               .Cell(flexcpText, .Row, 13) = "�� ��"
          End With
          PB_adoCnnSQL.CommitTrans
       Else    '��꼭��ȣ �ְ�
          If vsfg1.Cell(flexcpChecked, vsfg1.Row, 13) = flexUnchecked Then '�̹���
             PB_adoCnnSQL.BeginTrans
             strSQL = "UPDATE ���ݰ�꼭 SET " _
                           & "���࿩�� = 1, " _
                           & "�������� = '" & PB_regUserinfoU.UserServerDate & "', " _
                           & "������ڵ� = '" & PB_regUserinfoU.UserCode & "' " _
                     & "WHERE ������ڵ� = '" & vsfg1.TextMatrix(vsfg1.Row, 0) & "' " _
                       & "AND �ۼ��⵵ = '" & vsfg1.TextMatrix(vsfg1.Row, 9) & "' " _
                       & "AND å��ȣ = " & vsfg1.ValueMatrix(vsfg1.Row, 10) & " " _
                       & "AND �Ϸù�ȣ = " & vsfg1.ValueMatrix(vsfg1.Row, 11) & " "
             On Error GoTo ERROR_TABLE_UPDATE
             PB_adoCnnSQL.Execute strSQL
             vsfg1.Cell(flexcpChecked, vsfg1.Row, 13) = flexChecked
             vsfg1.Cell(flexcpText, vsfg1.Row, 13) = "�� ��"
             PB_adoCnnSQL.CommitTrans
             For lngR = 1 To vsfg1.Rows - 1
                 If vsfg1.TextMatrix(lngR, 9) = vsfg1.TextMatrix(vsfg1.Row, 9) And _
                    vsfg1.TextMatrix(lngR, 10) = vsfg1.TextMatrix(vsfg1.Row, 10) And _
                    vsfg1.TextMatrix(lngR, 11) = vsfg1.TextMatrix(vsfg1.Row, 11) And (lngR <> vsfg1.Row) Then
                    vsfg1.Cell(flexcpChecked, lngR, 13) = flexChecked
                    vsfg1.Cell(flexcpText, lngR, 13) = "�� ��"
                 End If
             Next lngR
             vsfg1.Cell(flexcpChecked, vsfg1.Row, 13) = flexChecked
             vsfg1.Cell(flexcpText, vsfg1.Row, 13) = "�� ��"
          End If
       End If
       SubPubPrint_TaxBill p, PB_intPrtTypeGbn, vsfg1.TextMatrix(vsfg1.Row, 0), vsfg1.TextMatrix(vsfg1.Row, 9), _
                                                vsfg1.ValueMatrix(vsfg1.Row, 10), vsfg1.ValueMatrix(vsfg1.Row, 11), 0, "", "", ""
    End If
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
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�б� ����"
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
    Subvsfg2_FILL
    cmdFind.Enabled = True
    Screen.MousePointer = vbDefault
End Sub

'+-----------+
'/// ���� ///
'+-----------+
Private Sub cmdSave_Click()
Dim p              As Printer
Dim blnASaveOK     As Boolean
Dim blnBSaveOK     As Boolean
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
Dim lngLogCnt      As Long    '������ �ŷ����� ��ȣ
Dim lngOrgLogCnt   As Long    '������ �ŷ����� ��ȣ
Dim intAddTax      As Integer '���ݰ�꼭(0.���ۼ�, 1.����, 2.�߰�, 3.�������߰�)
Dim intTaxPrt      As Integer
Dim strOldMakeYear As String
Dim lngOldLogCnt1  As Long
Dim lngOldLogCnt2  As Long
Dim strNewMakeYear As String
Dim lngNewLogCnt1  As Long
Dim lngNewLogCnt2  As Long
Dim strMakeDate    As String
Dim intGbn1        As Integer '�ݾױ���
Dim intGbn2        As Integer '��û����
Dim intGbn3        As Integer '�ۼ�����
Dim intGbn4        As Integer '�̼�����
Dim curInputMoney  As Long    'Not Used
Dim curOutputMoney As Long    'Not Used
Dim strServerTime  As String
Dim strTime        As String
Dim strHH          As String
Dim strMM          As String
Dim strSS          As String
Dim strMS          As String
Dim intChkCash     As Integer '���ݱ���(���ݸ���)
Dim intChgDate     As Integer '�������� ���� �˻�
Dim strOrgDate     As String  '������ ��������
Dim strChgDate     As String  '������ ��������

    For Each p In Printers
        If Trim(p.DeviceName) = Trim(cboPrinter.Text) And lstPort.List(cboPrinter.ListIndex) = p.Port Then
           Set Printer = p
           Exit For
        End If
    Next
    
    If vsfg1.Row >= vsfg1.FixedRows Then
       With vsfg1
            '���ݸ��� �κ� ����, �������� ����
            If (.Cell(flexcpBackColor, .Row, 14, .Row, 14) = vbRed) Or (.Cell(flexcpBackColor, .Row, 17, .Row, 17) = vbRed) Then
               blnASaveOK = True
            End If
            If .Cell(flexcpChecked, .Row, 14) = flexChecked Then         '���ݸ���
               intChkCash = 1
            End If
            If (.Cell(flexcpBackColor, .Row, 17, .Row, 17) = vbRed) Then '��������
               intChgDate = 1
            End If
       End With
       With vsfg2
            'If .RowHidden(lngRR) = False Then
                For lngRR = vsfg1.ValueMatrix(vsfg1.Row, 15) To vsfg1.ValueMatrix(vsfg1.Row, 16)
                    If (.TextMatrix(lngRR, 28) <> "") Then
                       blnBSaveOK = True
                       Exit For
                    End If
                Next lngRR
            'End If
            If (blnASaveOK = False And blnBSaveOK = False) Then '������(�����) ���̾�����
               Exit Sub
            End If
       End With
       '�������⳻���� ��꼭���࿩��(1) and ���ݰ�꼭��ȣ(������)
       If (vsfg1.ValueMatrix(vsfg1.Row, 18) = 1) And (Len(vsfg1.TextMatrix(vsfg1.Row, 12)) = 0) Then
          MsgBox "�̹� ��꼭�� �ۼ��Ǿ����� ��꼭��ȣ�� �˼� �����ϴ�. ���ݰ�꼭������ �ŷ������� �����ϼ���.!", vbCritical + vbOKOnly, "��꼭����"
          Screen.MousePointer = vbDefault
          Exit Sub
       End If
       intRetVal = MsgBox("������ �����ڷḦ �����Ͻðڽ��ϱ� ?", vbQuestion + vbYesNo + vbDefaultButton1, "�ڷ� ����")
       If intRetVal = vbNo Then
          vsfg1.SetFocus
          Exit Sub
       End If
       Screen.MousePointer = vbHourglass
       If vsfg1.Cell(flexcpChecked, vsfg1.Row, 14) = flexChecked Then '���ݱ���
          intChkCash = 1
       End If
       If Len(vsfg1.TextMatrix(vsfg1.Row, 12)) <> 0 Then '���ݰ�꼭�� ������
          intRetVal = MsgBox("�̹� ����� ���ݰ�꼭�� ������ �ٽ� �ۼ� �Ͻðڽ��ϱ� ?", vbQuestion + vbYesNo + vbDefaultButton1, "���ݰ�꼭")
          If intRetVal = vbYes Then
             intAddTax = 3 '�������ݰ�꼭 ������ �ٽ��߰�(����)
          Else
             intAddTax = 1 '�������ݰ�꼭 ����
          End If
       Else
          If chkTaxBillPrint.Value = 1 Then
             intAddTax = 2 '���ݰ�꼭 �߰�(����)
          End If
       End If
       Select Case intAddTax
              Case 1 '���ݰ�꼭 ���븸 ����
                   intTaxPrt = IIf(vsfg1.Cell(flexcpChecked, vsfg1.Row, 13) = flexUnchecked, 0, 1) '��꼭����(���)���� ����
              Case 2, 3 '2.���ݰ�꼭 �߰�, 3.������ �߰�
                   intTaxPrt = chkTaxBillPrint.Value
       End Select
       '�����ð� ���ϱ�
       P_adoRec.CursorLocation = adUseClient
       strSQL = "SELECT RIGHT(CONVERT(VARCHAR(23),GETDATE(), 121),12) AS �����ð� "
       On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
       P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
       strServerTime = Mid(P_adoRec("�����ð�"), 1, 2) + Mid(P_adoRec("�����ð�"), 4, 2) + Mid(P_adoRec("�����ð�"), 7, 2) _
                     + Mid(P_adoRec("�����ð�"), 10)
       P_adoRec.Close
       strTime = strServerTime
       PB_adoCnnSQL.BeginTrans
       '���� �ŷ�����, �ŷ���ȣ ����
       strOrgDate = DTOS(vsfg1.TextMatrix(vsfg1.Row, 1)) '������ ��������
       lngOrgLogCnt = vsfg1.ValueMatrix(vsfg1.Row, 2)    '������ �ŷ���ȣ
       '�ŷ���ȣ ���ϱ�
       If intChgDate = 1 Then '�������� �����̸�
          strChgDate = DTOS(vsfg1.TextMatrix(vsfg1.Row, 17)) '17.����ȸ�������
          strSQL = "spLogCounter '�������⳻��', '" & PB_regUserinfoU.UserBranchCode + DTOS(vsfg1.TextMatrix(vsfg1.Row, 17)) + "2" & "', " _
                            & "0, 0, '" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' "
          On Error GoTo ERROR_STORED_PROCEDURE
          P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
          lngLogCnt = P_adoRec(0)
          P_adoRec.Close
          '�������⳻�� �������� ����(���ݰ�꼭 ��ȣ�� ���� ���� ��)
          strSQL = "UPDATE �������⳻�� SET " _
                        & "��������� = '" & strChgDate & "', " _
                        & "�ŷ����� = '" & strChgDate & "', �ŷ���ȣ = " & lngLogCnt & " " _
                  & "WHERE ������ڵ� = '" & vsfg1.TextMatrix(vsfg1.Row, 0) & "' AND ������� = 2 AND ��뱸�� = 0 " _
                    & "AND �ŷ����� = '" & strOrgDate & "' " _
                    & "AND �ŷ���ȣ = " & lngOrgLogCnt & " "
          On Error GoTo ERROR_TABLE_UPDATE
          PB_adoCnnSQL.Execute strSQL
          '�׸��� ��,�� �ŷ����� �ٲٱ�
          '��
          With vsfg1
               .TextMatrix(.Row, 1) = .TextMatrix(.Row, 17)
               .TextMatrix(.Row, 2) = lngLogCnt
               .TextMatrix(.Row, 3) = .TextMatrix(.Row, 0) & "-" & Format(strChgDate, "0000/00/00") & "-" & CStr(lngLogCnt)
               .Cell(flexcpData, .Row, 3, lngR, 3) = Trim(.TextMatrix(.Row, 3)) 'FindRow ����� ����
          End With
          '��
          With vsfg2
               For lngRR = vsfg1.ValueMatrix(vsfg1.Row, 15) To vsfg1.ValueMatrix(vsfg1.Row, 16)
                   .TextMatrix(lngRR, 2) = vsfg1.TextMatrix(vsfg1.Row, 17) '�ŷ�(����)����
                   .TextMatrix(lngRR, 3) = lngLogCnt                       '�ŷ���ȣ
                   .TextMatrix(lngRR, 12) = .TextMatrix(lngRR, 1) & "-" & Format(strChgDate, "0000/00/00") _
                                     & "-" & CStr(lngLogCnt) & "-" & .TextMatrix(lngRR, 29) _
                                     & "-" & .TextMatrix(lngRR, 6) & "-" & .TextMatrix(lngRR, 10)
               Next lngRR
          End With
       End If
       If (intAddTax = 1 Or intAddTax = 3) Then '1.���� �Ǵ� 3.�������߰�
           strOldMakeYear = vsfg1.TextMatrix(vsfg1.Row, 9)
           lngOldLogCnt1 = vsfg1.ValueMatrix(vsfg1.Row, 10)
           lngOldLogCnt2 = vsfg1.ValueMatrix(vsfg1.Row, 11)
       End If
       If (intAddTax = 2 Or intAddTax = 3) Then '2.�߰� �Ǵ� 3.������ �߰�
          strNewMakeYear = Mid(PB_regUserinfoU.UserClientDate, 1, 4)
          'å��ȣ, �Ϸù�ȣ ���ϱ�
          strSQL = "spLogCounter '���ݰ�꼭', '" & PB_regUserinfoU.UserBranchCode + strNewMakeYear & "', " _
                               & "0, 0, '" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' "
          On Error GoTo ERROR_STORED_PROCEDURE
          P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
          lngNewLogCnt1 = P_adoRec(0)
          lngNewLogCnt2 = P_adoRec(1)
          P_adoRec.Close
       End If
       With vsfg2
            '�ŷ�����
            For lngRR = vsfg1.ValueMatrix(vsfg1.Row, 15) To vsfg1.ValueMatrix(vsfg1.Row, 16)
                If (.TextMatrix(lngRR, 28) = "I") Then 'And .ValueMatrix(lngRR, 16) <> 0) Then '�ŷ����� �߰�
                   strSQL = "INSERT INTO �������⳻��(������ڵ�, �з��ڵ�, �����ڵ�, �������, " _
                                                   & "���������, �����ð�, " _
                                                   & "�԰����, �԰�ܰ�, �԰�ΰ�, ������, ���ܰ�, ���ΰ�, " _
                                                   & "����ó�ڵ�, ����ó�ڵ�, �������������, ���۱���, " _
                                                   & "�߰�����, �߰߹�ȣ, �ŷ�����, �ŷ���ȣ, " _
                                                   & "��꼭���࿩��, ���ݱ���, ��������, ����, " _
                                                   & "�ۼ��⵵, å��ȣ, �Ϸù�ȣ, " _
                                                   & "��뱸��, ��������, ������ڵ�, ����̵�������ڵ�) VALUES( " _
                             & "'" & .TextMatrix(lngRR, 1) & "', '" & Mid(.TextMatrix(lngRR, 4), 1, 2) & "', " _
                             & "'" & Mid(.TextMatrix(lngRR, 4), 3) & "', 2, " _
                             & "'" & DTOS(.TextMatrix(lngRR, 2)) & "', '" & .TextMatrix(lngRR, 29) & "', " _
                             & "0, " & .ValueMatrix(lngRR, 19) & ", " _
                             & "" & .ValueMatrix(lngRR, 20) & ", " & .ValueMatrix(lngRR, 16) & ", " _
                             & "" & .ValueMatrix(lngRR, 23) & ", " & .ValueMatrix(lngRR, 24) & ", " _
                             & "'" & .TextMatrix(lngRR, 13) & "' , '" & .TextMatrix(lngRR, 8) & "', " _
                             & "'" & DTOS(.TextMatrix(lngRR, 2)) & "', 0, '', 0, " _
                             & "'" & DTOS(.TextMatrix(lngRR, 2)) & "', " & .ValueMatrix(lngRR, 3) & ", " _
                             & "" & IIf(intAddTax = 0, 0, 1) & ", " & intChkCash & ", 0, '" & Trim(.TextMatrix(lngRR, 27)) & "', " _
                             & "'" & IIf((intAddTax = 0 Or intAddTax = 1), strOldMakeYear, strNewMakeYear) & "', " _
                             & "" & IIf((intAddTax = 0 Or intAddTax = 1), lngOldLogCnt1, lngNewLogCnt1) & ", " _
                             & "" & IIf((intAddTax = 0 Or intAddTax = 1), lngOldLogCnt2, lngNewLogCnt2) & ", " _
                             & "0, '" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "', '' ) "
                   On Error GoTo ERROR_TABLE_INSERT
                   PB_adoCnnSQL.Execute strSQL
                   '������������ڰ���(������ڵ�, �з��ڵ�, �����ڵ�, �������)
                   strSQL = "sp����������������ڰ��� '" & PB_regUserinfoU.UserBranchCode & "', " _
                          & "'" & Mid(.TextMatrix(lngRR, 4), 1, 2) & "', '" & Mid(.TextMatrix(lngRR, 4), 3) & "', 2 "
                   On Error GoTo ERROR_STORED_PROCEDURE
                   PB_adoCnnSQL.Execute strSQL
                   '���������ܰ�����(������ڵ�, �з��ڵ�, �����ڵ�, �������, ��ü�ڵ�, �ܰ�, �ŷ�����)
                   If .ValueMatrix(lngRR, 23) > 0 And PB_intOAutoPriceGbn = 1 Then
                      strSQL = "sp���������ܰ����� '" & PB_regUserinfoU.UserBranchCode & "', " _
                             & "'" & Mid(.TextMatrix(lngRR, 4), 1, 2) & "', '" & Mid(.TextMatrix(lngRR, 4), 3) & "', 2, " _
                             & "'" & .TextMatrix(lngRR, 8) & "', " _
                             & "" & .ValueMatrix(lngRR, 23) & ", '" & DTOS(.TextMatrix(lngRR, 2)) & "' "
                      On Error GoTo ERROR_STORED_PROCEDURE
                      PB_adoCnnSQL.Execute strSQL
                   End If
                'ElseIf _
                '   (.TextMatrix(lngRR, 28) = "I" And .ValueMatrix(lngRR, 16) = 0) Then '�ŷ����� ���
                '   .TextMatrix(lngRR, 28) = "D"
                '   lngDelCnt = lngDelCnt + 1     '������ Row�� ���
                ElseIf _
                   (.TextMatrix(lngRR, 28) = "D") Then  '�ŷ����� ����
                   strSQL = "DELETE FROM �������⳻�� " _
                           & "WHERE ������ڵ� = '" & .TextMatrix(lngRR, 1) & "' " _
                             & "AND �з��ڵ� = '" & Mid(.TextMatrix(lngRR, 6), 1, 2) & "' " _
                             & "AND �����ڵ� = '" & Mid(.TextMatrix(lngRR, 6), 3) & "' " _
                             & "AND ������� = 2 " _
                             & "AND ��������� = '" & DTOS(.TextMatrix(lngRR, 2)) & "' " _
                             & "AND �����ð� = '" & .TextMatrix(lngRR, 29) & "' " _
                             & "AND ����ó�ڵ� = '" & .TextMatrix(lngRR, 10) & "' " _
                             & "AND �ŷ���ȣ = " & .ValueMatrix(lngRR, 3) & " "
                   lngDelCnt = lngDelCnt + 1     '������ Row�� ���
                   On Error GoTo ERROR_TABLE_DELETE
                   PB_adoCnnSQL.Execute strSQL
                   '������������ڰ���(������ڵ�, �з��ڵ�, �����ڵ�, �������)
                   strSQL = "sp����������������ڰ��� '" & PB_regUserinfoU.UserBranchCode & "', " _
                          & "'" & Mid(.TextMatrix(lngRR, 6), 1, 2) & "', '" & Mid(.TextMatrix(lngRR, 6), 3) & "', 2 " '������
                   On Error GoTo ERROR_STORED_PROCEDURE
                   PB_adoCnnSQL.Execute strSQL
                ElseIf _
                   (.TextMatrix(lngRR, 28) = "U") Then 'And .ValueMatrix(lngRR, 16) <> 0) Then '�ŷ����� ����
                   strSQL = "UPDATE �������⳻�� SET " _
                                 & "�з��ڵ� = '" & Mid(.TextMatrix(lngRR, 4), 1, 2) & "', " _
                                 & "�����ڵ� = '" & Mid(.TextMatrix(lngRR, 4), 3) & "', " _
                                 & "������ = " & .ValueMatrix(lngRR, 16) & ", " _
                                 & "��꼭���࿩�� = " & IIf(intAddTax = 0, 0, 1) & ", " _
                                 & "����ó�ڵ� = '" & .TextMatrix(lngRR, 8) & "', " _
                                 & "�԰�ܰ� = " & .ValueMatrix(lngRR, 19) & ", " _
                                 & "�԰�ΰ� = " & .ValueMatrix(lngRR, 20) & ", " _
                                 & "���ܰ� = " & .ValueMatrix(lngRR, 23) & "," _
                                 & "���ΰ� = " & .ValueMatrix(lngRR, 24) & ", " _
                                 & "���� = '" & .TextMatrix(lngRR, 27) & "', " _
                                 & "�ۼ��⵵ = '" & IIf((intAddTax = 0 Or intAddTax = 1), strOldMakeYear, strNewMakeYear) & "', " _
                                 & "å��ȣ = " & IIf((intAddTax = 0 Or intAddTax = 1), lngOldLogCnt1, lngNewLogCnt1) & " , " _
                                 & "�Ϸù�ȣ = " & IIf((intAddTax = 0 Or intAddTax = 1), lngOldLogCnt2, lngNewLogCnt2) & " , " _
                                 & "�������� = '" & PB_regUserinfoU.UserServerDate & "', " _
                                 & "������ڵ� = '" & PB_regUserinfoU.UserCode & "' " _
                           & "WHERE ������ڵ� = '" & .TextMatrix(lngRR, 1) & "' " _
                             & "AND �з��ڵ� = '" & Mid(.TextMatrix(lngRR, 6), 1, 2) & "' " _
                             & "AND �����ڵ� = '" & Mid(.TextMatrix(lngRR, 6), 3) & "' " _
                             & "AND ������� = 2 " _
                             & "AND ��������� = '" & DTOS(.TextMatrix(lngRR, 2)) & "' " _
                             & "AND �����ð� = '" & .TextMatrix(lngRR, 29) & "' " _
                             & "AND ����ó�ڵ� = '" & .TextMatrix(lngRR, 10) & "' " _
                             & "AND �ŷ���ȣ = " & .ValueMatrix(lngRR, 3) & " "
                   On Error GoTo ERROR_TABLE_UPDATE
                   PB_adoCnnSQL.Execute strSQL
                   '������������ڰ���(������ڵ�, �з��ڵ�, �����ڵ�, �������)
                   strSQL = "sp����������������ڰ��� '" & PB_regUserinfoU.UserBranchCode & "', " _
                          & "'" & Mid(.TextMatrix(lngRR, 4), 1, 2) & "', '" & Mid(.TextMatrix(lngRR, 4), 3) & "', 2 "    '������
                   On Error GoTo ERROR_STORED_PROCEDURE
                   PB_adoCnnSQL.Execute strSQL
                   If .TextMatrix(lngRR, 4) <> .TextMatrix(lngRR, 6) Then '�����ڵ� �����̸�
                      '������������ڰ���(������ڵ�, �з��ڵ�, �����ڵ�, �������)
                      strSQL = "sp����������������ڰ��� '" & PB_regUserinfoU.UserBranchCode & "', " _
                             & "'" & Mid(.TextMatrix(lngRR, 6), 1, 2) & "', '" & Mid(.TextMatrix(lngRR, 6), 3) & "', 2 " '������
                      On Error GoTo ERROR_STORED_PROCEDURE
                      PB_adoCnnSQL.Execute strSQL
                   End If
                   '���������ܰ�����(������ڵ�, �з��ڵ�, �����ڵ�, �������, ��ü�ڵ�, �ܰ�, �ŷ�����)
                   If .ValueMatrix(lngRR, 23) > 0 And PB_intOAutoPriceGbn = 1 Then
                      strSQL = "sp���������ܰ����� '" & PB_regUserinfoU.UserBranchCode & "', " _
                             & "'" & Mid(.TextMatrix(lngRR, 4), 1, 2) & "', '" & Mid(.TextMatrix(lngRR, 4), 3) & "', 2, " _
                             & "'" & .TextMatrix(lngRR, 8) & "', " _
                             & "" & .ValueMatrix(lngRR, 23) & ", '" & DTOS(.TextMatrix(lngRR, 2)) & "' "
                      On Error GoTo ERROR_STORED_PROCEDURE
                      PB_adoCnnSQL.Execute strSQL
                   End If
                ElseIf _
                   (.TextMatrix(lngRR, 28) = "") Then '�ƹ� ���������
                   strSQL = "UPDATE �������⳻�� SET " _
                                 & "��꼭���࿩�� = " & IIf(intAddTax = 0, 0, 1) & ", " _
                                 & "�ۼ��⵵ = '" & IIf((intAddTax = 0 Or intAddTax = 1), strOldMakeYear, strNewMakeYear) & "', " _
                                 & "å��ȣ = " & IIf((intAddTax = 0 Or intAddTax = 1), lngOldLogCnt1, lngNewLogCnt1) & " , " _
                                 & "�Ϸù�ȣ = " & IIf((intAddTax = 0 Or intAddTax = 1), lngOldLogCnt2, lngNewLogCnt2) & " , " _
                                 & "�������� = '" & PB_regUserinfoU.UserServerDate & "', " _
                                 & "������ڵ� = '" & PB_regUserinfoU.UserCode & "' " _
                           & "WHERE ������ڵ� = '" & .TextMatrix(lngRR, 1) & "' " _
                             & "AND �з��ڵ� = '" & Mid(.TextMatrix(lngRR, 6), 1, 2) & "' " _
                             & "AND �����ڵ� = '" & Mid(.TextMatrix(lngRR, 6), 3) & "' " _
                             & "AND ������� = 2 " _
                             & "AND ��������� = '" & DTOS(.TextMatrix(lngRR, 2)) & "' " _
                             & "AND �����ð� = '" & .TextMatrix(lngRR, 29) & "' " _
                             & "AND ����ó�ڵ� = '" & .TextMatrix(lngRR, 10) & "' " _
                             & "AND �ŷ���ȣ = " & .ValueMatrix(lngRR, 3) & " "
                   On Error GoTo ERROR_TABLE_UPDATE
                   PB_adoCnnSQL.Execute strSQL
                End If
                If ((.TextMatrix(lngRR, 28) = "I" Or .TextMatrix(lngRR, 28) = "U") And .ValueMatrix(lngRR, 16) <> 0) Then '�߰�, ����
                ElseIf _
                   (.TextMatrix(lngRR, 28) = "U" And .ValueMatrix(lngRR, 16) = 0) Then '�ŷ����� ����
                End If
            Next lngRR
            If intAddTax = 0 Then
            Else
               strSQL = "SELECT ISNULL(T1.�ۼ�����, '" & PB_regUserinfoU.UserClientDate & "') AS �ۼ�����, " _
                             & "ISNULL(T1.�ݾױ���, 3) AS �ݾױ���, ISNULL(T1.��û����, 1) AS ��û����, " _
                             & "ISNULL(T1.�ۼ�����, 0) AS �ۼ�����, ISNULL(T1.�̼�����, 1) AS �̼����� " _
                        & "FROM ���ݰ�꼭 T1 " _
                       & "WHERE T1.������ڵ� = '" & vsfg1.TextMatrix(vsfg1.Row, 0) & "' " _
                         & "AND T1.�ۼ��⵵ = '" & vsfg1.TextMatrix(vsfg1.Row, 9) & "' " _
                         & "AND T1.å��ȣ = " & vsfg1.ValueMatrix(vsfg1.Row, 10) & " " _
                         & "AND T1.�Ϸù�ȣ = " & vsfg1.ValueMatrix(vsfg1.Row, 11) & " "
               On Error GoTo ERROR_TABLE_SELECT
               P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
               If P_adoRec.RecordCount = 0 Then
                  P_adoRec.Close
                  strMakeDate = PB_regUserinfoU.UserClientDate
                  If intChkCash = 1 Then
                     intGbn1 = 0: intGbn2 = 0
                     intGbn3 = 0: intGbn4 = 1
                  Else
                     intGbn1 = 3: intGbn2 = 1
                     intGbn3 = 0: intGbn4 = 1
                  End If
               Else
                  If intAddTax = 1 Then     '����(���ݰ�꼭)
                     strMakeDate = P_adoRec("�ۼ�����")
                  Else
                     strMakeDate = PB_regUserinfoU.UserClientDate
                  End If
                  intGbn1 = IIf(intChkCash = 1, 0, P_adoRec("�ݾױ���"))
                  intGbn2 = IIf(intChkCash = 1, 0, P_adoRec("��û����"))
                  intGbn3 = P_adoRec("�ۼ�����")
                  intGbn4 = P_adoRec("�̼�����")
                  P_adoRec.Close
               End If
            End If
            If Len(vsfg1.TextMatrix(vsfg1.Row, 12)) <> 0 Then '���ݰ�꼭�� ������
               strSQL = "UPDATE �������⳻�� SET " _
                             & "��꼭���࿩�� = " & IIf(intAddTax = 0, 0, 1) & ", " _
                             & "�ۼ��⵵ = '" & IIf((intAddTax = 0 Or intAddTax = 1), strOldMakeYear, strNewMakeYear) & "', " _
                             & "å��ȣ = " & IIf((intAddTax = 0 Or intAddTax = 1), lngOldLogCnt1, lngNewLogCnt1) & " , " _
                             & "�Ϸù�ȣ = " & IIf((intAddTax = 0 Or intAddTax = 1), lngOldLogCnt2, lngNewLogCnt2) & " , " _
                             & "�������� = '" & PB_regUserinfoU.UserServerDate & "', " _
                             & "������ڵ� = '" & PB_regUserinfoU.UserCode & "' " _
                       & "WHERE ������ڵ� = '" & vsfg1.TextMatrix(vsfg1.Row, 0) & "' " _
                         & "AND �ۼ��⵵ = '" & vsfg1.TextMatrix(vsfg1.Row, 9) & "' " _
                         & "AND å��ȣ = " & vsfg1.ValueMatrix(vsfg1.Row, 10) & " " _
                         & "AND �Ϸù�ȣ = " & vsfg1.ValueMatrix(vsfg1.Row, 11) & " "
                       '& "WHERE ������ڵ� = '" & vsfg1.TextMatrix(vsfg1.Row, 0) & "' " _
                         & "AND �ŷ����� = '" & DTOS(vsfg1.TextMatrix(vsfg1.Row, 1)) & "' " _
                         & "AND �ŷ���ȣ = " & vsfg1.TextMatrix(vsfg1.Row, 2) & " "
               On Error GoTo ERROR_TABLE_UPDATE
               PB_adoCnnSQL.Execute strSQL
            End If
       End With
       '���ݸ���
       With vsfg1
            If blnASaveOK = True Then
               strSQL = "UPDATE �������⳻�� SET ���ݱ��� = " & intChkCash & " " _
                       & "WHERE ������ڵ� = '" & .TextMatrix(.Row, 0) & "' AND �ŷ����� = '" & DTOS(.TextMatrix(.Row, 1)) & "' " _
                         & "AND �ŷ���ȣ = " & .ValueMatrix(.Row, 2) & " "
               On Error GoTo ERROR_TABLE_UPDATE
               PB_adoCnnSQL.Execute strSQL
            End If
            '���ݸ����̰� ��꼭�����̸� �̼��ݳ��� �߰�
       End With
       With vsfg1
            If (.ValueMatrix(.Row, 16) - .ValueMatrix(.Row, 15) + 1) = lngDelCnt Then '�ŷ����� ��� ����
               strSQL = "UPDATE �������⳻�� SET " _
                             & "��뱸�� = 9, " _
                             & "�������� = '" & PB_regUserinfoU.UserServerDate & "', " _
                             & "������ڵ� = '" & PB_regUserinfoU.UserCode & "' " _
                       & "WHERE ������ڵ� = '" & .TextMatrix(.Row, 0) & "' AND �ŷ����� = '" & DTOS(.TextMatrix(.Row, 1)) & "' " _
                         & "AND �ŷ���ȣ = " & .ValueMatrix(.Row, 2) & " "
               On Error GoTo ERROR_TABLE_DELETE
               PB_adoCnnSQL.Execute strSQL
               If Len(vsfg1.TextMatrix(.Row, 12)) <> 0 Then '���ݰ�꼭�� ������
                  intRetVal = MsgBox("�̹� ����� ���ݰ�꼭�� �����Ͻðڽ��ϱ� ?", vbQuestion + vbYesNo + vbDefaultButton2, "���ݰ�꼭")
                  If intRetVal = vbYes Then '���ݰ�꼭 ����
                     strSQL = "UPDATE ���ݰ�꼭 SET " _
                                   & "��뱸�� = 9, " _
                                   & "�������� = '" & PB_regUserinfoU.UserServerDate & "', " _
                                   & "������ڵ� = '" & PB_regUserinfoU.UserCode & "' " _
                             & "WHERE ������ڵ� = '" & .TextMatrix(.Row, 0) & "' AND �ۼ��⵵ = '" & .TextMatrix(.Row, 9) & "' " _
                               & "AND å��ȣ = " & .ValueMatrix(.Row, 10) & " AND �Ϸù�ȣ = " & .ValueMatrix(.Row, 11) & " "
                      On Error GoTo ERROR_TABLE_DELETE
                      PB_adoCnnSQL.Execute strSQL
                   End If
               End If
               lngDelCntS = .ValueMatrix(.Row, 15): lngDelCntE = .ValueMatrix(.Row, 16)
               For lngRR = .ValueMatrix(.Row, 16) To .ValueMatrix(.Row, 15) Step -1
                   '������������ڰ���(������ڵ�, �з��ڵ�, �����ڵ�, �������)
                   strSQL = "sp����������������ڰ��� '" & PB_regUserinfoU.UserBranchCode & "', " _
                          & "'" & Mid(vsfg2.TextMatrix(lngRR, 6), 1, 2) & "', '" & Mid(vsfg2.TextMatrix(lngRR, 6), 3) & "', 2 "    '������
                   On Error GoTo ERROR_STORED_PROCEDURE
                   PB_adoCnnSQL.Execute strSQL
                   vsfg2.RemoveItem lngRR
               Next lngRR
               .RemoveItem .Row
               For lngRRR = 1 To .Rows - 1
                   If lngDelCntS < .ValueMatrix(lngRRR, 15) Then
                      .TextMatrix(lngRRR, 15) = .ValueMatrix(lngRRR, 15) - (lngDelCntE - lngDelCntS + 1)
                      .TextMatrix(lngRRR, 16) = .ValueMatrix(lngRRR, 16) - (lngDelCntE - lngDelCntS + 1)
                   End If
               Next lngRRR
               .Row = 0 '���缱�õ� ���� Row�� ����
            Else
               If (intAddTax = 2 Or intAddTax = 3) Then
                  .TextMatrix(.Row, 9) = IIf((intAddTax = 0 Or intAddTax = 1), strOldMakeYear, strNewMakeYear)
                  .TextMatrix(.Row, 10) = IIf((intAddTax = 0 Or intAddTax = 1), lngOldLogCnt1, lngNewLogCnt1)
                  .TextMatrix(.Row, 11) = IIf((intAddTax = 0 Or intAddTax = 1), lngOldLogCnt2, lngNewLogCnt2)
               End If
               lngDelCntS = .ValueMatrix(.Row, 15): lngDelCntE = .ValueMatrix(.Row, 16)
               For lngRR = .ValueMatrix(.Row, 16) To .ValueMatrix(.Row, 15) Step -1
                   If vsfg2.TextMatrix(lngRR, 28) = "D" Then
                      vsfg2.RemoveItem lngRR
                      '����
                      'For lngR = lngRR To vsfg2.Rows - 1
                      '    If vsfg2.RowHidden(lngR) = True Then Exit For
                      '    vsfg2.TextMatrix(lngR, 0) = vsfg2.ValueMatrix(lngR, 0) - 1
                      'Next lngR
                   End If
               Next lngRR
               vsfg1.TextMatrix(vsfg1.Row, 16) = vsfg1.ValueMatrix(vsfg1.Row, 16) - lngDelCnt
               For lngR = 1 To vsfg1.Rows - 1
                   If (lngR <> vsfg1.Row) Then
                      If (vsfg1.ValueMatrix(.Row, 16) < vsfg1.ValueMatrix(lngR, 15)) Then
                         vsfg1.TextMatrix(lngR, 15) = vsfg1.ValueMatrix(lngR, 15) - lngDelCnt
                         vsfg1.TextMatrix(lngR, 16) = vsfg1.ValueMatrix(lngR, 16) - lngDelCnt
                      End If
                   End If
               Next lngR
            End If
       End With
       With vsfg2
            '������(������)
            If vsfg1.Row > 0 Then
               For lngRR = vsfg1.ValueMatrix(vsfg1.Row, 15) To vsfg1.ValueMatrix(vsfg1.Row, 16)
                   If (.TextMatrix(lngRR, 28) = "I" Or .TextMatrix(lngRR, 28) = "U") Then 'And .ValueMatrix(lngRR, 16) > 0) Then   '�ŷ����� ����
                      .TextMatrix(lngRR, 6) = .TextMatrix(lngRR, 4)   '�����ڵ�(������->������)
                      .TextMatrix(lngRR, 7) = .TextMatrix(lngRR, 5)   '�����(������->������)
                      .TextMatrix(lngRR, 10) = .TextMatrix(lngRR, 8)  '����ó�ڵ�(������->������)
                      .TextMatrix(lngRR, 11) = .TextMatrix(lngRR, 9)  '����ó��(������->������)
                   End If
               Next lngRR
            End If
       End With
       With vsfg2
            '�ŷ�����(���� ����ġ)
            If vsfg1.Row > 0 Then '���缱�õ� �ŷ� Row�� ����
               vsfg1.Cell(flexcpBackColor, vsfg1.Row, vsfg1.FixedCols, vsfg1.Row, vsfg1.Cols - 1) = vbWhite
               vsfg1.Cell(flexcpForeColor, vsfg1.Row, vsfg1.FixedCols, vsfg1.Row, vsfg1.Cols - 1) = vbBlack
               .Cell(flexcpBackColor, vsfg1.ValueMatrix(vsfg1.Row, 15), 0, vsfg1.ValueMatrix(vsfg1.Row, 16), .FixedCols - 1) = _
               .Cell(flexcpBackColor, 0, 0, 0, 0)
               .Cell(flexcpForeColor, vsfg1.ValueMatrix(vsfg1.Row, 15), 0, vsfg1.ValueMatrix(vsfg1.Row, 16), .FixedCols - 1) = _
               .Cell(flexcpForeColor, 0, 0, 0, 0)
               .Cell(flexcpBackColor, vsfg1.ValueMatrix(vsfg1.Row, 15), .FixedCols, vsfg1.ValueMatrix(vsfg1.Row, 16), .Cols - 1) = vbWhite
               .Cell(flexcpForeColor, vsfg1.ValueMatrix(vsfg1.Row, 15), .FixedCols, vsfg1.ValueMatrix(vsfg1.Row, 16), .Cols - 1) = vbBlack
               .Cell(flexcpText, vsfg1.ValueMatrix(vsfg1.Row, 15), 28, vsfg1.ValueMatrix(vsfg1.Row, 16), 28) = "" 'SQL���� ����
            End If
       End With
       '������ ���� �κ� ����
       'With vsfg2
       '     If vsfg1.Row > 0 Then
       '        lngDelCntS = vsfg1.ValueMatrix(vsfg1.Row, 15): lngDelCntE = vsfg1.ValueMatrix(vsfg1.Row, 16)
       '        For lngRRR = 1 To vsfg1.Rows - 1
       '            If lngDelCntS < vsfg1.ValueMatrix(lngRRR, 15) Then
       '               vsfg1.TextMatrix(lngRRR, 15) = vsfg1.ValueMatrix(lngRRR, 15) - (lngDelCntE - lngDelCntS + 1)
       '               vsfg1.TextMatrix(lngRRR, 16) = vsfg1.ValueMatrix(lngRRR, 16) - (lngDelCntE - lngDelCntS + 1)
       '            End If
       '        Next lngRRR
       '     End If
       'End With
       
       'vsfg1.Row = 0: vsfg2.Row = 0 '(�������� �� ��)
       
       If (intAddTax = 1) Then '���ݰ�꼭 ����
          strSQL = "DELETE FROM ���ݰ�꼭 " _
                  & "WHERE ������ڵ� =  '" & vsfg1.TextMatrix(vsfg1.Row, 0) & "' " _
                    & "AND �ۼ��⵵ = '" & strOldMakeYear & "' " _
                    & "AND å��ȣ = " & lngOldLogCnt1 & " " _
                    & "AND �Ϸù�ȣ = " & lngOldLogCnt2 & " "
          On Error GoTo ERROR_TABLE_DELETE
          PB_adoCnnSQL.Execute strSQL
       End If
       If (intAddTax = 3) Then '���ݰ�꼭 �������߰�
          strSQL = "UPDATE ���ݰ�꼭 SET " _
                        & "��뱸�� = 9, " _
                        & "�������� = '" & PB_regUserinfoU.UserServerDate & "', " _
                        & "������ڵ� = '" & PB_regUserinfoU.UserCode & "' " _
                  & "WHERE ������ڵ� =  '" & vsfg1.TextMatrix(vsfg1.Row, 0) & "' " _
                    & "AND �ۼ��⵵ = '" & strOldMakeYear & "' " _
                    & "AND å��ȣ = " & lngOldLogCnt1 & " " _
                    & "AND �Ϸù�ȣ = " & lngOldLogCnt2 & " "
          On Error GoTo ERROR_TABLE_DELETE
          PB_adoCnnSQL.Execute strSQL
       End If
       '���ݰ�꼭 �ۼ�
       If (intAddTax > 0) Then
          '���ݰ�꼭
          strSQL = "INSERT INTO ���ݰ�꼭 " _
                 & "SELECT T1.������ڵ� AS ������ڵ�, T1.�ۼ��⵵ AS �ۼ��⵵, " _
                        & "T1.å��ȣ AS å��ȣ, T1.�Ϸù�ȣ AS �Ϸù�ȣ," _
                        & "T1.����ó�ڵ� AS ����ó�ڵ�, '" & strMakeDate & "' AS �ۼ�����, " _
                       & "(SELECT TOP 1 (S2.����� + SPACE(1) + S2.�԰�)  FROM �������⳻�� S1 " _
                          & "LEFT JOIN ���� S2 ON S1.�з��ڵ� = S2.�з��ڵ� AND S1.�����ڵ� = S2.�����ڵ� " _
                         & "WHERE S1.������ڵ� = T1.������ڵ� " _
                           & "AND S1.�ۼ��⵵ = T1.�ۼ��⵵ " _
                           & "AND S1.å��ȣ = T1.å��ȣ " _
                           & "AND S1.�Ϸù�ȣ = T1.�Ϸù�ȣ " _
                           & "AND S1.��뱸�� = 0 " _
                           & "AND S1.����ó�ڵ� = T1.����ó�ڵ� " _
                         & "ORDER BY S1.���������, S1.�����ð�) AS ǰ��ױ԰�, "
          strSQL = strSQL + _
                         "(SELECT (COUNT(S1.�����ڵ�) - 1) FROM �������⳻�� S1 " _
                         & "WHERE S1.������ڵ� = T1.������ڵ� " _
                           & "AND S1.�ۼ��⵵ = T1.�ۼ��⵵ " _
                           & "AND S1.å��ȣ = T1.å��ȣ " _
                           & "AND S1.�Ϸù�ȣ = T1.�Ϸù�ȣ " _
                           & "AND S1.��뱸�� = 0 " _
                           & "AND S1.����ó�ڵ� = T1.����ó�ڵ�) AS ����, " _
                        & "SUM(T1.������ * T1.���ܰ�) AS ���ް���, SUM(T1.������ * T1.���ΰ�) AS ����, " _
                        & "" & intGbn1 & " AS �ݾױ���, " & intGbn2 & " AS ��û����, " & intTaxPrt & " AS ���࿩��, " _
                        & "" & intGbn3 & " AS �ۼ�����, " & intGbn4 & " AS �̼�����, '' AS ����, 0 AS ��뱸��, " _
                        & "'" & PB_regUserinfoU.UserServerDate & "' AS ��������, '" & PB_regUserinfoU.UserCode & "' AS ������ڵ� " _
                   & "FROM �������⳻�� T1 " _
                   & "LEFT JOIN ���� T2 ON T2.�з��ڵ� = T1.�з��ڵ� AND T2.�����ڵ� = T1.�����ڵ� " _
                   & "LEFT JOIN ����ó T3 ON T3.����ó�ڵ� = T1.����ó�ڵ� " _
                  & "WHERE T1.������ڵ� = '" & vsfg1.TextMatrix(vsfg1.Row, 0) & "' " _
                    & "AND T1.�ۼ��⵵ = '" & IIf(intAddTax = 1, strOldMakeYear, strNewMakeYear) & "' " _
                    & "AND T1.å��ȣ = " & IIf(intAddTax = 1, lngOldLogCnt1, lngNewLogCnt1) & " " _
                    & "AND T1.�Ϸù�ȣ = " & IIf(intAddTax = 1, lngOldLogCnt2, lngNewLogCnt2) & " " _
                    & "AND T1.��뱸�� = 0 " _
                  & "GROUP BY T1.������ڵ�, T1.�ۼ��⵵, T1.å��ȣ, T1.�Ϸù�ȣ, T1.����ó�ڵ� " _
                  & "ORDER BY T1.������ڵ�, T1.�ۼ��⵵, T1.å��ȣ, T1.�Ϸù�ȣ "
          On Error GoTo ERROR_TABLE_INSERT
          PB_adoCnnSQL.Execute strSQL
          '���⼼�ݰ�꼭���(�ۼ��ð�:strTime)
          If (intAddTax = 2 Or intAddTax = 3) Then  '1.����, 2.�߰�, 3.�������߰�
             strSQL = "INSERT INTO ���⼼�ݰ�꼭��� " _
                    & "SELECT T1.������ڵ�, T1.�ۼ�����, '" & strTime & "', T1.����ó�ڵ�, " _
                           & "T1.ǰ��ױ԰�, T1.����, T1.���ް���, T1.����, " _
                           & "T1.�ݾױ���, T1.��û����, T1.���࿩��, T1.�ۼ�����, " _
                           & "T1.�̼�����, T1.����, T1.��뱸��, T1.��������, " _
                           & "T1.������ڵ�, T1.�ۼ��⵵, T1.å��ȣ, T1.�Ϸù�ȣ " _
                      & "FROM ���ݰ�꼭 T1 " _
                     & "WHERE T1.������ڵ� = '" & vsfg1.TextMatrix(vsfg1.Row, 0) & "' " _
                       & "AND T1.�ۼ��⵵ = '" & strNewMakeYear & "' " _
                       & "AND T1.å��ȣ = " & lngNewLogCnt1 & " " _
                       & "AND T1.�Ϸù�ȣ = " & lngNewLogCnt2 & " "
             On Error GoTo ERROR_TABLE_INSERT
             PB_adoCnnSQL.Execute strSQL
          End If
       End If
       'PB_adoCnnSQL.RollbackTrans
       PB_adoCnnSQL.CommitTrans
       If (intAddTax = 2 Or intAddTax = 3) Then '�߰� �Ǵ� ������ �߰�
          If intAddTax = 2 Then '�߰�
             vsfg1.TextMatrix(vsfg1.Row, 9) = strNewMakeYear
             vsfg1.TextMatrix(vsfg1.Row, 10) = lngNewLogCnt1
             vsfg1.TextMatrix(vsfg1.Row, 11) = lngNewLogCnt2
             vsfg1.TextMatrix(vsfg1.Row, 12) = vsfg1.TextMatrix(vsfg1.Row, 9) + "-" _
                                             + vsfg1.TextMatrix(vsfg1.Row, 10) + "-" + vsfg1.TextMatrix(vsfg1.Row, 11)
             If intTaxPrt = 0 Then
                vsfg1.Cell(flexcpChecked, vsfg1.Row, 13) = flexUnchecked
             Else
                vsfg1.Cell(flexcpChecked, vsfg1.Row, 13) = flexChecked
             End If
          ElseIf _
             intAddTax = 3 Then '������ �߰�
             For lngR = 1 To vsfg1.Rows - 1
                 If vsfg1.TextMatrix(lngR, 9) = strOldMakeYear And _
                    vsfg1.TextMatrix(lngR, 10) = lngOldLogCnt1 And _
                    vsfg1.TextMatrix(lngR, 11) = lngOldLogCnt2 And (lngR <> vsfg1.Row) Then
                    vsfg1.TextMatrix(lngR, 9) = strNewMakeYear
                    vsfg1.TextMatrix(lngR, 10) = lngNewLogCnt1
                    vsfg1.TextMatrix(lngR, 11) = lngNewLogCnt2
                    vsfg1.TextMatrix(lngR, 12) = vsfg1.TextMatrix(vsfg1.Row, 9) + "-" _
                                               + vsfg1.TextMatrix(vsfg1.Row, 10) + "-" + vsfg1.TextMatrix(vsfg1.Row, 11)
                    If intTaxPrt = 0 Then
                       vsfg1.Cell(flexcpChecked, vsfg1.Row, 13) = flexUnchecked
                    Else
                       vsfg1.Cell(flexcpChecked, vsfg1.Row, 13) = flexChecked
                    End If
                 End If
             Next lngR
          End If
       End If
       If (chkPrint.Value = 1) And (cboPrinter.ListIndex >= 0) Then
          SubPubPrint_DealBill p, PB_intPrtTypeGbn, DTOS(vsfg1.TextMatrix(vsfg1.Row, 1)), vsfg1.TextMatrix(vsfg1.Row, 2)  '�ŷ����� ���
       End If
       If (chkTaxBillPrint.Value = 1) And (cboPrinter.ListIndex >= 0) Then                        '���ݰ�꼭 ���
          SubPubPrint_TaxBill p, PB_intPrtTypeGbn, PB_regUserinfoU.UserBranchCode, strNewMakeYear, lngNewLogCnt1, lngNewLogCnt2, _
                              0, "", "", ""
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
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "���� ���� ����"
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
    If vsfg1.Row >= vsfg1.FixedRows Then
       If (vsfg1.ValueMatrix(vsfg1.Row, 18) = 1) And (Len(vsfg1.TextMatrix(vsfg1.Row, 12)) = 0) Then
          MsgBox "�̹� ��꼭�� �ۼ��Ǿ����� ��꼭��ȣ�� �˼� �����ϴ�. ���ݰ�꼭������ �ŷ������� �����ϼ���.!", vbCritical + vbOKOnly, "��꼭����"
          Screen.MousePointer = vbDefault
          Exit Sub
       End If
       intRetVal = MsgBox("����ó���� �ڷḦ �����Ͻðڽ��ϱ� ?", vbCritical + vbYesNo + vbDefaultButton2, "�ŷ����� ����")
       If intRetVal = vbNo Then
          vsfg1.SetFocus
          Exit Sub
       End If
       cmdDelete.Enabled = False
       Screen.MousePointer = vbHourglass
       PB_adoCnnSQL.BeginTrans
       P_adoRec.CursorLocation = adUseClient
       With vsfg1
            lngDelCntS = .ValueMatrix(.Row, 15): lngDelCntE = .ValueMatrix(.Row, 16)
            '�ŷ�����
            For lngRR = .ValueMatrix(.Row, 16) To .ValueMatrix(.Row, 15) Step -1
                strSQL = "UPDATE �������⳻�� SET " _
                              & "��뱸�� = 9, " _
                              & "�������� = '" & PB_regUserinfoU.UserServerDate & "', " _
                              & "������ڵ� = '" & PB_regUserinfoU.UserCode & "' " _
                        & "WHERE ������ڵ� = '" & vsfg2.TextMatrix(lngRR, 1) & "' " _
                          & "AND �з��ڵ� = '" & Mid(vsfg2.TextMatrix(lngRR, 6), 1, 2) & "' " _
                          & "AND �����ڵ� = '" & Mid(vsfg2.TextMatrix(lngRR, 6), 3) & "' " _
                          & "AND ������� = 2 " _
                          & "AND ��������� = '" & DTOS(vsfg2.TextMatrix(lngRR, 2)) & "' " _
                          & "AND �����ð� = '" & vsfg2.TextMatrix(lngRR, 29) & "' " _
                          & "AND ����ó�ڵ� = '" & vsfg2.TextMatrix(lngRR, 10) & "' " _
                          & "AND �ŷ���ȣ = " & vsfg2.ValueMatrix(lngRR, 3) & " "
                On Error GoTo ERROR_TABLE_UPDATE
                PB_adoCnnSQL.Execute strSQL
                '������������ڰ���(������ڵ�, �з��ڵ�, �����ڵ�, �������)
                strSQL = "sp����������������ڰ��� '" & PB_regUserinfoU.UserBranchCode & "', " _
                       & "'" & Mid(vsfg2.TextMatrix(lngRR, 6), 1, 2) & "', '" & Mid(vsfg2.TextMatrix(lngRR, 6), 3) & "', 2 "
                On Error GoTo ERROR_STORED_PROCEDURE
                PB_adoCnnSQL.Execute strSQL
                '���������ܰ�����(������ڵ�, �з��ڵ�, �����ڵ�, �������, ��ü�ڵ�, �ܰ�, �ŷ�����)
                'If vsfg2.ValueMatrix(lngRR, 23) > 0 And PB_intOAutoPriceGbn = 1 Then
                '   strSQL = "sp���������ܰ����� '" & PB_regUserinfoU.UserBranchCode & "', " _
                '          & "'" & Mid(vsfg2.TextMatrix(lngRR, 6), 1, 2) & "', '" & Mid(vsfg2.TextMatrix(lngRR, 6), 3) & "', 2, " _
                '          & "'" & vsfg2.TextMatrix(lngRR, 8) & "', " _
                '          & "" & vsfg2.ValueMatrix(lngRR, 23) & ", '" & vsfg2.ValueMatrix(lngRR, 2) & "' "
                '   On Error GoTo ERROR_STORED_PROCEDURE
                '   PB_adoCnnSQL.Execute strSQL
                'End If
                vsfg2.RemoveItem lngRR
            Next lngRR
            lblTotMny.Caption = Format(Vals(lblTotMny.Caption) - .ValueMatrix(.Row, 6), "#,#.00") '��ü�ݾ׿��� ����
            If Len(.TextMatrix(.Row, 12)) <> 0 Then '���ݰ�꼭�� ������
               intRetVal = MsgBox("�̹� ����� ���ݰ�꼭�� �����Ͻðڽ��ϱ� ?", vbQuestion + vbYesNo + vbDefaultButton2, "���ݰ�꼭")
               If intRetVal = vbYes Then '���ݰ�꼭 ����
                  strSQL = "UPDATE ���ݰ�꼭 SET " _
                                & "��뱸�� = 9, " _
                                & "�������� = '" & PB_regUserinfoU.UserServerDate & "', " _
                                & "������ڵ� = '" & PB_regUserinfoU.UserCode & "' " _
                          & "WHERE ������ڵ� = '" & .TextMatrix(.Row, 0) & "' AND �ۼ��⵵ = '" & .TextMatrix(.Row, 9) & "' " _
                            & "AND å��ȣ = " & .ValueMatrix(.Row, 10) & " AND �Ϸù�ȣ = " & .ValueMatrix(.Row, 11) & " "
                  On Error GoTo ERROR_TABLE_DELETE
                  PB_adoCnnSQL.Execute strSQL
                  strSQL = "UPDATE �������⳻�� SET " _
                                & "��꼭���࿩�� = 0, " _
                                & "�ۼ��⵵ = '', å��ȣ = 0, �Ϸù�ȣ = 0, " _
                                & "�������� = '" & PB_regUserinfoU.UserServerDate & "', " _
                                & "������ڵ� = '" & PB_regUserinfoU.UserCode & "' " _
                          & "WHERE ������ڵ� = '" & .TextMatrix(.Row, 0) & "' AND �ۼ��⵵ = '" & .TextMatrix(.Row, 9) & "' " _
                            & "AND å��ȣ = " & .ValueMatrix(.Row, 10) & " AND �Ϸù�ȣ = " & .ValueMatrix(.Row, 11) & " "
                  On Error GoTo ERROR_TABLE_UPDATE
                  PB_adoCnnSQL.Execute strSQL
                  For lngR = 1 To .Rows - 1
                      If (.TextMatrix(.Row, 0) = .TextMatrix(lngR, 0)) And (.TextMatrix(.Row, 9) = .TextMatrix(lngR, 9)) And _
                         (.ValueMatrix(.Row, 10) = .ValueMatrix(lngR, 10)) And (.ValueMatrix(.Row, 11) = .ValueMatrix(lngR, 11)) Then
                         .TextMatrix(lngR, 9) = "": .TextMatrix(lngR, 10) = "": .TextMatrix(lngR, 11) = ""
                         .TextMatrix(lngR, 12) = ""  '���ݰ�꼭��ȣ
                         .TextMatrix(lngR, 13) = "0" '��꼭���࿩��
                         .Cell(flexcpChecked, lngR, 13) = flexUnchecked  '2
                      End If
                  Next lngR
               End If
            End If
            .RemoveItem .Row
            For lngRRR = 1 To .Rows - 1
                If lngDelCntS < .ValueMatrix(lngRRR, 15) Then
                   .TextMatrix(lngRRR, 15) = .ValueMatrix(lngRRR, 15) - (lngDelCntE - lngDelCntS + 1)
                   .TextMatrix(lngRRR, 16) = .ValueMatrix(lngRRR, 16) - (lngDelCntE - lngDelCntS + 1)
                End If
            Next lngRRR
            .Row = 0
       End With
       PB_adoCnnSQL.CommitTrans
       cmdFind.SetFocus
       Screen.MousePointer = vbDefault
    End If
    cmdDelete.Enabled = True
    Exit Sub
ERROR_TABLE_SELECT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�ŷ����� �б� ����"
    Unload Me
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�ŷ����� ���� ����"
    Unload Me
    Exit Sub
ERROR_TABLE_DELETE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�ŷ����� ���� ����"
    Unload Me
    Exit Sub
ERROR_STORED_PROCEDURE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "���� ���� ����"
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
    Set frm������� = Nothing
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
    With vsfg1              'Rows 1, Cols 19, RowHeightMax(Min) 300
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
         .FixedCols = 3
         .Rows = 1             'Subvsfg1_Fill����ÿ� ����
         .Cols = 19
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1200   '������ڵ� 'H
         .ColWidth(1) = 1200   '�ŷ�����
         .ColWidth(2) = 1000   '�ŷ���ȣ
         .ColWidth(3) = 1730   '������ڵ�+�ŷ�����+�ŷ���ȣ(KEY) H
         .ColWidth(4) = 1500   '����ó�ڵ�
         .ColWidth(5) = 3300   '����ó��
         .ColWidth(6) = 2000   '�ݾ�(�ܰ���)
         .ColWidth(7) = 2000   '�ݾ�(�ΰ���) 'H
         .ColWidth(8) = 2000   '�ݾ�(�հ�) 'H
         .ColWidth(9) = 1000   '�ۼ��⵵   'H
         .ColWidth(10) = 1000  'å��ȣ     'H
         .ColWidth(11) = 1000  '�Ϸù�ȣ   'H
         .ColWidth(12) = 2000  '���ݰ�꼭��ȣ(�ۼ��⵵-å��ȣ-�Ϸù�ȣ)
         .ColWidth(13) = 1500  '��꼭����(���)����
         .ColWidth(14) = 1000  '���ⱸ��(���ݱ���)
         .ColWidth(15) = 1000  'ROW(vsfg2.Row)   Not Used 'H
         .ColWidth(16) = 1000  'COL(vsfg2.Row)   Not Used 'H
         .ColWidth(17) = 1200  '�������ں���
         .ColWidth(18) = 1000  '��꼭�ۼ�����
         
         .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "������ڵ�" 'H
         .TextMatrix(0, 1) = "�ŷ�����"
         .TextMatrix(0, 2) = "�ŷ���ȣ"
         .TextMatrix(0, 3) = "�ŷ�(KEY)"  'H (KEY)
         .TextMatrix(0, 4) = "����ó�ڵ�"
         .TextMatrix(0, 5) = "����ó��"
         .TextMatrix(0, 6) = "�ݾ�"
         .TextMatrix(0, 7) = "�ݾ�"       'H
         .TextMatrix(0, 8) = "�ݾ�"       'H
         .TextMatrix(0, 9) = "�ۼ��⵵"   'H
         .TextMatrix(0, 10) = "å��ȣ"    'H
         .TextMatrix(0, 11) = "�Ϸù�ȣ"  'H
         .TextMatrix(0, 12) = "���ݰ�꼭��ȣ"
         .TextMatrix(0, 13) = "��꼭���࿩��"
         .TextMatrix(0, 14) = "���ⱸ��"
         .TextMatrix(0, 15) = "Row"       'H
         .TextMatrix(0, 16) = "Col"       'H
         .TextMatrix(0, 17) = "��������"
         .TextMatrix(0, 18) = "�ۼ�����"  'H
         
         .ColHidden(0) = True: .ColHidden(3) = True:
         .ColHidden(7) = True: .ColHidden(8) = True
         .ColHidden(9) = True: .ColHidden(10) = True: .ColHidden(11) = True:
         .ColHidden(15) = True: .ColHidden(16) = True: .ColHidden(18) = True
         .ColFormat(6) = "#,#.00": .ColFormat(7) = "#,#.00": .ColFormat(8) = "#,#.00"
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 5
                         .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 0, 1, 2, 3, 4, 9, 10, 11, 12, 13, 14, 17, 18
                         .ColAlignment(lngC) = flexAlignCenterCenter
                    Case Else
                         .ColAlignment(lngC) = flexAlignRightCenter
             End Select
         Next lngC
         .MergeCells = flexMergeRestrictRows  'flexMergeFixedOnly
         .MergeRow(0) = True
         For lngC = 0 To 3
             .MergeCol(lngC) = True
         Next lngC
    End With
End Sub
Private Sub Subvsfg2_INIT()
Dim lngR    As Long
Dim lngC    As Long
    With vsfg2              'Rows 1, Cols 30, RowHeightMax(Min) 300
         .AllowBigSelection = False
         .AllowSelection = False
         .AllowUserResizing = flexResizeColumns
         .BackColorBkg = &H8000000F
         .BackColorSel = &H8000&
         .ExtendLastCol = True
         .FocusRect = flexFocusHeavy
         .ScrollBars = flexScrollBarBoth
         .ScrollTrack = True
         .SelectionMode = flexSelectionByRow
         .FixedRows = 1
         .FixedCols = 6
         .Rows = 1             'Subvsfg2_Fill����ÿ� ����
         .Cols = 30
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 400    'No
         .ColWidth(1) = 1000   '������ڵ�
         .ColWidth(2) = 1200   '�ŷ�����    '0000-00-00
         .ColWidth(3) = 1000   '�ŷ���ȣ
         .ColWidth(4) = 1900   'ǰ���ڵ�(������)
         .ColWidth(5) = 2600   'ǰ��(������)
         .ColWidth(6) = 1900   '�����ڵ�(������) 'H
         .ColWidth(7) = 2600   '�����(������) 'H
         .ColWidth(8) = 1000   '����ó�ڵ�(������) 'H
         .ColWidth(9) = 2000   '����ó��(������)   'H
         .ColWidth(10) = 1000  '����ó�ڵ�(������) 'H
         .ColWidth(11) = 2000  '����ó��(������) 'H
         .ColWidth(12) = 2000  '������ڵ�+��������+������ȣ+�����ð�+�����ڵ�+����ó�ڵ�+(KEY) 'H
         .ColWidth(13) = 1000  '����ó�ڵ� 'H
         .ColWidth(14) = 2500  '����ó��   'H
         .ColWidth(15) = 2200  '����԰�
         .ColWidth(16) = 1000  '����
         .ColWidth(17) = 800   '�������
         .ColWidth(18) = 800   '��꼭���࿩�� 'H
         .ColWidth(19) = 1600  '�԰�ܰ�   'H
         .ColWidth(20) = 1300  '�԰�ΰ�   'H
         .ColWidth(21) = 1600  '�԰���(�ܰ�+�ΰ�) 'H
         .ColWidth(22) = 1700  '�԰�ݾ�   'H
         .ColWidth(23) = 1600  '���ܰ�
         .ColWidth(24) = 1300  '���ΰ�   'H
         .ColWidth(25) = 1600  '�����(�ܰ�+�ΰ�) 'H
         .ColWidth(26) = 1700  '���ݾ�
         .ColWidth(27) = 5000  '����
         .ColWidth(28) = 800   'SQL����
         .ColWidth(29) = 1000  '�����ð� 'H
         .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "No"            '(����)
         .TextMatrix(0, 1) = "������ڵ�"   'H
         .TextMatrix(0, 2) = "�ŷ�����"     'H
         .TextMatrix(0, 3) = "�ŷ���ȣ"     'H
         .TextMatrix(0, 4) = "ǰ���ڵ�"     '������(Or ������)
         .TextMatrix(0, 5) = "ǰ��"         '������(Or ������)
         .TextMatrix(0, 6) = "ǰ���ڵ�"     'H, ������
         .TextMatrix(0, 7) = "ǰ��"         'H, ������
         .TextMatrix(0, 8) = "����ó�ڵ�"   'H, ������
         .TextMatrix(0, 9) = "����ó��"     'H, ������
         .TextMatrix(0, 10) = "����ó�ڵ�"  'H, ������
         .TextMatrix(0, 11) = "����ó��"    'H, ������
         .TextMatrix(0, 12) = "KEY"         'H
         .TextMatrix(0, 13) = "����ó�ڵ�"  'H
         .TextMatrix(0, 14) = "����ó��"    'H
         .TextMatrix(0, 15) = "�԰�"
         .TextMatrix(0, 16) = "����"
         .TextMatrix(0, 17) = "����"
         .TextMatrix(0, 18) = "���"        'H
         .TextMatrix(0, 19) = "���Դܰ�"    'H
         .TextMatrix(0, 20) = "���Ժΰ�"    'H
         .TextMatrix(0, 21) = "���԰���"    'H (�ܰ� + �ΰ�)
         .TextMatrix(0, 22) = "���Աݾ�"    'H
         .TextMatrix(0, 23) = "����ܰ�"
         .TextMatrix(0, 24) = "����ΰ�"    'H
         .TextMatrix(0, 25) = "���Ⱑ��"    'H (�ܰ� + �ΰ�)
         .TextMatrix(0, 26) = "����ݾ�"
         .TextMatrix(0, 27) = "����"
         .TextMatrix(0, 28) = "����"        'H(SQL����:I.Insert, U.Update, D.Delete)
         .TextMatrix(0, 29) = "�����ð�"  'H
         .ColHidden(1) = True: .ColHidden(2) = True: .ColHidden(3) = True
         .ColHidden(6) = True: .ColHidden(7) = True: .ColHidden(8) = True
         .ColHidden(9) = True: .ColHidden(10) = True: .ColHidden(11) = True: .ColHidden(12) = True
         .ColHidden(13) = True: .ColHidden(14) = True: .ColHidden(18) = True
         .ColHidden(19) = True: .ColHidden(20) = True: .ColHidden(21) = True: .ColHidden(22) = True
         .ColHidden(24) = True: .ColHidden(25) = True
         .ColHidden(28) = True: .ColHidden(29) = True
         .ColFormat(16) = "#,#"
         For lngC = 19 To 26
             .ColFormat(lngC) = "#,#.00"
         Next lngC
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 17, 27
                         .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 0, 1, 2, 3, 18, 28, 29
                         .ColAlignment(lngC) = flexAlignCenterCenter
                    Case Else
                         .ColAlignment(lngC) = flexAlignRightCenter
             End Select
         Next lngC
         .MergeCells = flexMergeRestrictRows 'flexMergeFixedOnly
         .MergeRow(0) = True
         For lngC = 0 To 5
             .MergeCol(lngC) = True
         Next lngC
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
    If Len(Text1(0).Text) = 0 Then
       strWhere = strWhere
    Else
       strWhere = "AND T1.����ó�ڵ� = '" & Trim(Text1(0).Text) & "' "
    End If
    strSQL = "SELECT T1.������ڵ� AS ������ڵ�, T1.�ŷ����� AS �ŷ�����, T1.�ŷ���ȣ AS �ŷ���ȣ, " _
                  & "T1.�ۼ��⵵ AS �ۼ��⵵, T1.å��ȣ AS å��ȣ, T1.�Ϸù�ȣ AS �Ϸù�ȣ, T3.���࿩�� AS ��꼭���࿩��, " _
                  & "T1.����ó�ڵ� AS ����ó�ڵ�, ISNULL(T2.����ó��, '') AS ����ó��, " _
                  & "SUM(T1.������ * T1.���ܰ�) AS �ܰ��ݾ�, SUM(T1.������ * T1.���ΰ�) AS �ΰ��ݾ�, " _
                  & "T1.���ݱ��� AS ���ݱ���, T1.��꼭���࿩�� AS ��꼭�ۼ����� " _
             & "FROM �������⳻�� T1 " _
             & "LEFT JOIN ����ó T2 ON T2.������ڵ� = T1.������ڵ� AND T2.����ó�ڵ� = T1.����ó�ڵ� " _
             & "LEFT JOIN ���ݰ�꼭 T3 ON T3.������ڵ� = T1.������ڵ� AND T3.å��ȣ = T1.å��ȣ AND T3.�Ϸù�ȣ = T1.�Ϸù�ȣ " _
            & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' AND T1.��뱸�� = 0 AND T1.������� = 2 " _
              & "" & strWhere & " " _
              & "AND T1.��������� BETWEEN '" & DTOS(dtpF_Date.Value) & "' AND '" & DTOS(dtpT_Date.Value) & "' " _
            & "GROUP BY T1.������ڵ�, T1.�ŷ�����, T1.�ŷ���ȣ, " _
                     & "T1.�ۼ��⵵, T1.å��ȣ, T1.�Ϸù�ȣ, T3.���࿩��, T1.����ó�ڵ�, ISNULL(T2.����ó��, ''), " _
                     & "T1.���ݱ���, T1.��꼭���࿩�� " _
            & "ORDER BY T1.������ڵ�, T1.�ŷ�����, T1.�ŷ���ȣ "
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
               .TextMatrix(lngR, 1) = Format(P_adoRec("�ŷ�����"), "0000-00-00")
               .TextMatrix(lngR, 2) = IIf(IsNull(P_adoRec("�ŷ���ȣ")), 0, P_adoRec("�ŷ���ȣ"))
               .TextMatrix(lngR, 3) = P_adoRec("������ڵ�") & "-" & Format(P_adoRec("�ŷ�����"), "0000/00/00") _
                                    & "-" & CStr(P_adoRec("�ŷ���ȣ"))
               .Cell(flexcpData, lngR, 3, lngR, 3) = Trim(.TextMatrix(lngR, 3)) 'FindRow ����� ����
               
               .TextMatrix(lngR, 4) = IIf(IsNull(P_adoRec("����ó�ڵ�")), "", P_adoRec("����ó�ڵ�"))
               .TextMatrix(lngR, 5) = IIf(IsNull(P_adoRec("����ó��")), "", P_adoRec("����ó��"))
               '.TextMatrix(lngR, 6) = IIf(IsNull(P_adoRec("�ܰ��ݾ�")), 0, P_adoRec("�ܰ��ݾ�"))
               .TextMatrix(lngR, 7) = IIf(IsNull(P_adoRec("�ΰ��ݾ�")), 0, P_adoRec("�ΰ��ݾ�"))
               .TextMatrix(lngR, 8) = .ValueMatrix(lngR, 6) + .ValueMatrix(lngR, 7)
               .TextMatrix(lngR, 9) = P_adoRec("�ۼ��⵵")
               .TextMatrix(lngR, 10) = P_adoRec("å��ȣ")
               .TextMatrix(lngR, 11) = P_adoRec("�Ϸù�ȣ")
               If Len(.TextMatrix(lngR, 9)) > 0 Then
                  .TextMatrix(lngR, 12) = .TextMatrix(lngR, 9) + "-" + .TextMatrix(lngR, 10) + "-" + .TextMatrix(lngR, 11)
               End If
               .TextMatrix(lngR, 13) = IIf(IsNull(P_adoRec("��꼭���࿩��")), 0, P_adoRec("��꼭���࿩��"))
               If P_adoRec("��꼭���࿩��") = 1 Then
                  .Cell(flexcpChecked, lngR, 13) = flexChecked    '1
               Else
                  .Cell(flexcpChecked, lngR, 13) = flexUnchecked  '2
               End If
               .Cell(flexcpText, lngR, 13) = "�� ��"
               .TextMatrix(lngR, 14) = IIf(IsNull(P_adoRec("���ݱ���")), 0, P_adoRec("���ݱ���"))
               If P_adoRec("���ݱ���") = 1 Then
                  .Cell(flexcpChecked, lngR, 14) = flexChecked    '1
               Else
                  .Cell(flexcpChecked, lngR, 14) = flexUnchecked  '2
               End If
               .Cell(flexcpText, lngR, 14) = "���ݸ���"
               .TextMatrix(lngR, 17) = Format(P_adoRec("�ŷ�����"), "0000-00-00") '��������
               'If .TextMatrix(lngR, 0) = PB_regUserinfoU.UserBranchCode Then
               '   lngRR = lngR
               'End If
               .Cell(flexcpText, lngR, 18) = IIf(IsNull(P_adoRec("��꼭�ۼ�����")), 0, P_adoRec("��꼭�ۼ�����"))
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
            vsfg1_EnterCell       'vsfg1_EnterCell �ڵ�����(���� �Ѱ� �϶��� ������ �ڵ�����)
            .SetFocus
       End With
    End If
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�ŷ� �б� ����"
    Unload Me
    Exit Sub
End Sub

Private Sub Subvsfg2_FILL()
Dim strSQL      As String
Dim strWhere    As String
Dim strOrderBy  As String
Dim lngR        As Long
Dim lngC        As Long
Dim lngRR       As Long
Dim lngRRR      As Long
Dim strCell     As String
Dim strSubTotal As String
    vsfg2.Rows = 1
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.������ڵ�, T1.�ŷ�����, T1.�ŷ���ȣ, T1.�����ð�, " _
                 & "(T1.�з��ڵ� + T1.�����ڵ�) AS �����ڵ�, ISNULL(T3.�����,'') AS �����, " _
                 & "ISNULL(T2.����ó�ڵ�, '') AS ����ó�ڵ�, ISNULL(T2.����ó��,'') AS ����ó��, " _
                 & "'' AS ����ó�ڵ�, '' AS ����ó��, T3.�԰� AS ����԰�, T1.������, " _
                 & "T3.���� AS �������, T1.��꼭���࿩��, T1.�԰�ܰ�, T1.�԰�ΰ�, " _
                 & "T1.���ܰ� , T1.���ΰ�, T1.���� AS ���� " _
            & "FROM �������⳻�� T1 " _
            & "LEFT JOIN ����ó T2 ON T2.������ڵ� = T1.������ڵ� AND T2.����ó�ڵ� = T1.����ó�ڵ� " _
            & "LEFT JOIN ���� T3 ON (T3.�з��ڵ� = T1.�з��ڵ� AND T3.�����ڵ� = T1.�����ڵ�) " _
           & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' AND T1.��뱸�� = 0 AND T1.������� = 2 " _
             & "AND T1.��������� BETWEEN '" & DTOS(dtpF_Date.Value) & "' AND '" & DTOS(dtpT_Date.Value) & "' " _
           & "ORDER BY T1.������ڵ�, T1.�ŷ�����, T1.�ŷ���ȣ, T1.�����ð� "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    vsfg2.Rows = P_adoRec.RecordCount + 1
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       cmdExit.SetFocus
       Exit Sub
    Else
       With vsfg2
            .Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlack
            If .Rows <= PC_intRowCnt2 Then
               '.ScrollBars = flexScrollBarHorizontal
            Else
               '.ScrollBars = flexScrollBarBoth
            End If
            Do Until P_adoRec.EOF
               lngR = lngR + 1
               '.TextMatrix(lngR, 0) = Format(P_adoRec("XX"), "0000-00-00")
               .TextMatrix(lngR, 1) = IIf(IsNull(P_adoRec("������ڵ�")), "", P_adoRec("������ڵ�"))
               .TextMatrix(lngR, 2) = Format(P_adoRec("�ŷ�����"), "0000-00-00")
               .TextMatrix(lngR, 3) = IIf(IsNull(P_adoRec("�ŷ���ȣ")), 0, P_adoRec("�ŷ���ȣ"))
               .TextMatrix(lngR, 4) = IIf(IsNull(P_adoRec("�����ڵ�")), "", P_adoRec("�����ڵ�"))
               .TextMatrix(lngR, 5) = IIf(IsNull(P_adoRec("�����")), "", P_adoRec("�����"))
               .TextMatrix(lngR, 6) = IIf(IsNull(P_adoRec("�����ڵ�")), "", P_adoRec("�����ڵ�"))
               .TextMatrix(lngR, 7) = IIf(IsNull(P_adoRec("�����")), "", P_adoRec("�����"))
               .TextMatrix(lngR, 8) = IIf(IsNull(P_adoRec("����ó�ڵ�")), "", P_adoRec("����ó�ڵ�"))
               .TextMatrix(lngR, 9) = IIf(IsNull(P_adoRec("����ó��")), "", P_adoRec("����ó��"))
               .TextMatrix(lngR, 10) = IIf(IsNull(P_adoRec("����ó�ڵ�")), "", P_adoRec("����ó�ڵ�"))
               .TextMatrix(lngR, 11) = IIf(IsNull(P_adoRec("����ó��")), "", P_adoRec("����ó��"))
               .TextMatrix(lngR, 12) = P_adoRec("������ڵ�") & "-" & Format(P_adoRec("�ŷ�����"), "0000/00/00") _
                                     & "-" & CStr(P_adoRec("�ŷ���ȣ")) & "-" & P_adoRec("�����ð�") _
                                     & "-" & P_adoRec("�����ڵ�") & "-" & P_adoRec("����ó�ڵ�")
               '.TextMatrix(lngR, 13) = IIf(IsNull(P_adoRec("����ó�ڵ�")), "", P_adoRec("����ó�ڵ�"))
               '.TextMatrix(lngR, 14) = IIf(IsNull(P_adoRec("����ó��")), "", P_adoRec("����ó��"))
               .TextMatrix(lngR, 15) = IIf(IsNull(P_adoRec("����԰�")), "", P_adoRec("����԰�"))
               .TextMatrix(lngR, 16) = Format(P_adoRec("������"), "#,#")
               .TextMatrix(lngR, 17) = IIf(IsNull(P_adoRec("�������")), "", P_adoRec("�������"))
               If P_adoRec("��꼭���࿩��") = 0 Then
                  .Cell(flexcpChecked, lngR, 18) = flexUnchecked
               Else
                  .Cell(flexcpChecked, lngR, 18) = flexChecked
               End If
               .Cell(flexcpText, lngR, 18) = "�� ��"
               .Cell(flexcpAlignment, lngR, 18, lngR, 18) = flexAlignLeftCenter
               .TextMatrix(lngR, 19) = Format(P_adoRec("�԰�ܰ�"), "#,#.00")
               .TextMatrix(lngR, 20) = Format(P_adoRec("�԰�ΰ�"), "#,#.00")
               .TextMatrix(lngR, 21) = .ValueMatrix(lngR, 19) + .ValueMatrix(lngR, 20)
               .TextMatrix(lngR, 22) = .ValueMatrix(lngR, 16) * .ValueMatrix(lngR, 21)
               .TextMatrix(lngR, 23) = Format(P_adoRec("���ܰ�"), "#,#.00")
               .TextMatrix(lngR, 24) = Format(P_adoRec("���ΰ�"), "#,#.00")
               .TextMatrix(lngR, 25) = .ValueMatrix(lngR, 23) + .ValueMatrix(lngR, 24) '���ܰ� + ���ΰ�
               .TextMatrix(lngR, 26) = .ValueMatrix(lngR, 16) * .ValueMatrix(lngR, 23) '���ݾ�(����*�ܰ�)
               .TextMatrix(lngR, 27) = IIf(IsNull(P_adoRec("����")), "", P_adoRec("����"))
               .TextMatrix(lngR, 29) = IIf(IsNull(P_adoRec("�����ð�")), "", P_adoRec("�����ð�"))
               'If .TextMatrix(lngR, 0) = PB_regUserinfoU.UserBranchCode Then
               '   lngRR = lngR
               'End If
               .RowHidden(lngR) = True
               P_adoRec.MoveNext
            Loop
            P_adoRec.Close
            If lngRR = 0 Then
               .Row = lngRR       'vsfg2_EnterCell �ڵ�����(���� �Ѱ� �϶��� �ڵ����� ����)
               If .Rows > PC_intRowCnt2 Then
                  '.TopRow = .Rows - PC_intRowCnt2 + .FixedRows
                  .TopRow = 1
               End If
            Else
               .Row = lngRR       'vsfg2_EnterCell �ڵ�����(���� �Ѱ� �϶��� �ڵ����� ����)
               If .Rows > PC_intRowCnt2 Then
                  .TopRow = .Row
               End If
            End If
            '.MultiTotals = True '(default value : true)
            '.Subtotal flexSTClear
            '.SubtotalPosition = flexSTBelow
            '.Subtotal flexSTCount, 6, 8, "#", vbRed, vbWhite, , "%s", , False
            '.Subtotal flexSTSum, 6, 10, , vbRed, vbWhite, , "%s", , False
            For lngR = 1 To .Rows - 1
                strCell = .TextMatrix(lngR, 1) & "-" & Format(DTOS(.TextMatrix(lngR, 2)), "0000/00/00") & "-" & .TextMatrix(lngR, 3)
                For lngRRR = 1 To vsfg1.Rows - 1
                    If strCell = vsfg1.TextMatrix(lngRRR, 3) Then
                       If vsfg1.ValueMatrix(lngRRR, 15) = 0 Then
                          vsfg1.TextMatrix(lngRRR, 15) = lngR
                       End If
                       vsfg1.TextMatrix(lngRRR, 16) = lngR
                       '�ŷ� �հ�ݾ� ���
                       vsfg1.TextMatrix(lngRRR, 6) = vsfg1.ValueMatrix(lngRRR, 6) + .ValueMatrix(lngR, 26)
                       lblTotMny.Caption = Format(Vals(lblTotMny.Caption) + .ValueMatrix(lngR, 26), "#,#.00")
                       Exit For
                    End If
                Next lngRRR
            Next lngR
            vsfg2_EnterCell       'vsfg1_EnterCell �ڵ�����(���� �Ѱ� �϶��� ������ �ڵ�����) 'Not Used
            .SetFocus                                                                         'Not Used
       End With
    End If
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�ŷ����� �б� ����"
    Unload Me
    Exit Sub
End Sub
