VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frm�����ü��˻� 
   Appearance      =   0  '���
   BackColor       =   &H00008000&
   BorderStyle     =   1  '���� ����
   Caption         =   "���� �ܰ� �˻�"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3720
   BeginProperty Font 
      Name            =   "����ü"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "�����ü��˻�.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   3720
   StartUpPosition =   2  'ȭ�� ���
   Begin VSFlex7Ctl.VSFlexGrid vsfg1 
      Height          =   2385
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   3495
      _cx             =   6165
      _cy             =   4207
      _ConvInfo       =   1
      Appearance      =   0
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483638
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483638
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
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
End
Attribute VB_Name = "frm�����ü��˻�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' ���α׷� �� �� : �����ü��˻�
' ���� Control :
' ������ Table   :
' ��  ��  ��  �� :
'
'+-------------------------------------------------------------------------------------------------------+
Private P_blnActived       As Boolean
Private P_intFindGbn       As Integer      '0.�����˻�, 1.�ڵ�, 2.�̸�(��), 3.�ڵ� + �̸�(��)
Private P_intButton        As Integer
Private P_strFindString2   As String
Private P_strFirstCode     As String
Private P_strAppDate       As String       '��������
Private P_blnSelect        As Boolean
Private P_adoRec           As New ADODB.Recordset
Private Const PC_intRowCnt As Integer = 10  '�׸��� �� ������ �� ���(FixedRows ����)

'+--------------------------------+
'/// LOAD FORM ( �ѹ��� ���� ) ///
'+--------------------------------+
Private Sub Form_Load()
    P_blnActived = False
    P_strFirstCode = PB_strMaterialsCode '���� ã������ �����ڵ�
    If (PB_strFMCCallFormName = "frm���ּ��ۼ�") Then
       With frm���ּ��ۼ�
            P_strAppDate = PB_regUserinfoU.UserClientDate
       End With
    End If
    If (PB_strFMCCallFormName = "frm���ּ�����") Then
       With frm���ּ�����
            P_strAppDate = .vsfg2.TextMatrix(.vsfg2.Row, 2)
       End With
    End If
    If (PB_strFMCCallFormName = "frm�����ۼ�1") Then
       With frm�����ۼ�1
            P_strAppDate = .vsfg2.TextMatrix(.vsfg2.Row, 2)
       End With
    End If
    If (PB_strFMCCallFormName = "frm�����ۼ�2") Then
       With frm�����ۼ�2
            P_strAppDate = DTOS(.dtpJ_Date.Value)
       End With
    End If
    If (PB_strFMCCallFormName = "frm���Լ���") Then
       With frm���Լ���
            P_strAppDate = .vsfg2.TextMatrix(.vsfg2.Row, 2)
       End With
    End If
    
    If (PB_strFMCCallFormName = "frm�������ۼ�") Then
       With frm�������ۼ�
            P_strAppDate = PB_regUserinfoU.UserClientDate
       End With
    End If
    If (PB_strFMCCallFormName = "frm����������") Then
       With frm����������
            P_strAppDate = .vsfg2.TextMatrix(.vsfg2.Row, 2)
       End With
    End If
    If (PB_strFMCCallFormName = "frm�����ۼ�1") Then
       With frm�����ۼ�1
            P_strAppDate = .vsfg2.TextMatrix(.vsfg2.Row, 2)
       End With
    End If
    If (PB_strFMCCallFormName = "frm�����ۼ�2") Then
       With frm�����ۼ�2
            P_strAppDate = DTOS(.dtpJ_Date.Value)
       End With
    End If
    If (PB_strFMCCallFormName = "frm�������") Then
       With frm�������
            P_strAppDate = .vsfg2.TextMatrix(.vsfg2.Row, 2)
       End With
    End If
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

'+------------+
'/// vsfg1 ///
'+------------+
Private Sub vsfg1_DblClick()
    vsfg1_KeyDown vbKeyReturn, 0
End Sub
Private Sub vsfg1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
       SubEscape
    End If
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
Private Sub vsfg1_EnterCell()
Dim lngR As Long
Dim lngC As Long
    With vsfg1
         If .Row >= .FixedRows Then
         End If
    End With
End Sub
'+---------------------------------------+
'/// �ü� �˻��� Return
'+---------------------------------------+
Private Sub vsfg1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim SQL         As String
Dim intRetVal   As Integer
Dim lngR        As Long
Dim blnDupOK    As Boolean
Dim varFormName As Variant
    With vsfg1
         If KeyCode = vbKeyReturn Then
            If .Row <= 0 Then
               PB_strMaterialsCode = ""
               PB_strMaterialsName = ""
            Else
               PB_strMaterialsCode = .TextMatrix(.Row, 0)
               PB_strMaterialsName = .TextMatrix(.Row, 1)
            End If
            varFormName = PB_strFMCCallFormName
            If .Row >= .FixedRows Then
               If (PB_strFMCCallFormName = "frm���ּ��ۼ�") Then
                  With frm���ּ��ۼ�
                       .vsfg1.TextMatrix(.vsfg1.Row, 8) = vsfg1.ValueMatrix(vsfg1.Row, 5)     '�ܰ�
                       .vsfg1.TextMatrix(.vsfg1.Row, 9) = vsfg1.ValueMatrix(vsfg1.Row, 6)     '�ΰ�
                       .vsfg1.TextMatrix(.vsfg1.Row, 10) = vsfg1.ValueMatrix(vsfg1.Row, 7)    '����
                       '�հ�ݾ� ���(�ݾ� �ٸ���)
                       If .vsfg1.ValueMatrix(.vsfg1.Row, 11) <> _
                          (.vsfg1.ValueMatrix(.vsfg1.Row, 3) * vsfg1.ValueMatrix(vsfg1.Row, 5)) Then
                          .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg1.ValueMatrix(.vsfg1.Row, 11) _
                                             + (.vsfg1.ValueMatrix(.vsfg1.Row, 3) * vsfg1.ValueMatrix(vsfg1.Row, 5)), "#,#.00")
                        End If
                        .vsfg1.TextMatrix(.vsfg1.Row, 11) = .vsfg1.ValueMatrix(.vsfg1.Row, 3) _
                                                          * vsfg1.ValueMatrix(vsfg1.Row, 5)   '�԰�ݾ�
                        P_blnSelect = True
                        PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm���ּ�����") Then
                  With frm���ּ�����
                       If (.vsfg2.ValueMatrix(.vsfg2.Row, 19) <> vsfg1.TextMatrix(vsfg1.Row, 5)) Then '�ܰ� �ٸ���
                           .vsfg2.Cell(flexcpBackColor, .vsfg2.Row, 19, .vsfg2.Row, 19) = vbRed
                           .vsfg2.Cell(flexcpForeColor, .vsfg2.Row, 19, .vsfg2.Row, 19) = vbWhite
                          If .vsfg2.TextMatrix(.vsfg2.Row, 28) = "" Then
                             .vsfg2.TextMatrix(.vsfg2.Row, 28) = "U"
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 19) = vsfg1.ValueMatrix(vsfg1.Row, 5)     '�ܰ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 20) = vsfg1.ValueMatrix(vsfg1.Row, 6)     '�ΰ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 21) = vsfg1.ValueMatrix(vsfg1.Row, 7)     '����
                          '�հ�ݾ� ���(�ݾ��� �ٸ���)
                          If .vsfg2.ValueMatrix(.vsfg2.Row, 22) <> _
                            (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5)) Then
                             .vsfg1.TextMatrix(.vsfg1.Row, 7) = .vsfg1.ValueMatrix(.vsfg1.Row, 7) - .vsfg2.ValueMatrix(.vsfg2.Row, 22) _
                                                             + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5))
                             .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg2.ValueMatrix(.vsfg2.Row, 22) _
                                                + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5)), "#,#.00")
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 22) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 5)   '�԰�ݾ�
                          P_blnSelect = True
                          PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       End If
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm�����ۼ�1") Then
                  With frm�����ۼ�1
                       If (.vsfg2.ValueMatrix(.vsfg2.Row, 19) <> vsfg1.TextMatrix(vsfg1.Row, 5)) Then '�ܰ� �ٸ���
                           .vsfg2.Cell(flexcpBackColor, .vsfg2.Row, 19, .vsfg2.Row, 19) = vbRed
                           .vsfg2.Cell(flexcpForeColor, .vsfg2.Row, 19, .vsfg2.Row, 19) = vbWhite
                          If .vsfg2.TextMatrix(.vsfg2.Row, 28) = "" Then
                             .vsfg2.TextMatrix(.vsfg2.Row, 28) = "U"
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 19) = vsfg1.ValueMatrix(vsfg1.Row, 5)     '�ܰ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 20) = vsfg1.ValueMatrix(vsfg1.Row, 6)     '�ΰ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 21) = vsfg1.ValueMatrix(vsfg1.Row, 7)     '����
                          '�հ�ݾ� ���(�ݾ��� �ٸ���)
                          If .vsfg2.ValueMatrix(.vsfg2.Row, 22) <> _
                            (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5)) Then
                             .vsfg1.TextMatrix(.vsfg1.Row, 7) = .vsfg1.ValueMatrix(.vsfg1.Row, 7) - .vsfg2.ValueMatrix(.vsfg2.Row, 22) _
                                                             + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5))
                             .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg2.ValueMatrix(.vsfg2.Row, 22) _
                                                + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5)), "#,#.00")
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 22) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 5)     '�ݾ�
                          P_blnSelect = True
                          PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       End If
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm�����ۼ�2") Then
                  With frm�����ۼ�2
                       If (.vsfg1.TextMatrix(.vsfg1.Row, 8) <> vsfg1.TextMatrix(vsfg1.Row, 5)) Then   '�ܰ� �ٸ���
                          .vsfg1.TextMatrix(.vsfg1.Row, 8) = vsfg1.ValueMatrix(vsfg1.Row, 5)          '�ܰ�
                          .vsfg1.TextMatrix(.vsfg1.Row, 9) = vsfg1.ValueMatrix(vsfg1.Row, 6)          '�ΰ�
                          .vsfg1.TextMatrix(.vsfg1.Row, 10) = vsfg1.ValueMatrix(vsfg1.Row, 7)         '����
                          '�հ�ݾ� ���(�ݾ��� �ٸ���)
                          If .vsfg1.ValueMatrix(.vsfg1.Row, 11) <> _
                            (.vsfg1.ValueMatrix(.vsfg1.Row, 3) * vsfg1.ValueMatrix(vsfg1.Row, 5)) Then
                             .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg1.ValueMatrix(.vsfg1.Row, 11) _
                                                + (.vsfg1.ValueMatrix(.vsfg1.Row, 3) * vsfg1.ValueMatrix(vsfg1.Row, 5)), "#,#.00")
                          End If
                          .vsfg1.TextMatrix(.vsfg1.Row, 11) = .vsfg1.ValueMatrix(.vsfg1.Row, 3) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 5)   '�԰�ݾ�
                          P_blnSelect = True
                          PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       End If
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm���Լ���") Then
                  With frm���Լ���
                       If (.vsfg2.TextMatrix(.vsfg2.Row, 19) <> vsfg1.TextMatrix(vsfg1.Row, 5)) Then '�ܰ� �ٸ���
                          .vsfg2.Cell(flexcpBackColor, .vsfg2.Row, 19, .vsfg2.Row, 19) = vbRed
                           .vsfg2.Cell(flexcpForeColor, .vsfg2.Row, 19, .vsfg2.Row, 19) = vbWhite
                          If .vsfg2.TextMatrix(.vsfg2.Row, 28) = "" Then
                             .vsfg2.TextMatrix(.vsfg2.Row, 28) = "U"
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 19) = vsfg1.ValueMatrix(vsfg1.Row, 5)   '�ܰ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 20) = vsfg1.ValueMatrix(vsfg1.Row, 6)   '�ΰ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 21) = vsfg1.ValueMatrix(vsfg1.Row, 7)   '����
                          '�հ�ݾ� ���(�ݾ��� �ٸ���)
                          If .vsfg2.ValueMatrix(.vsfg2.Row, 22) <> _
                            (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 10)) Then
                             .vsfg1.TextMatrix(.vsfg1.Row, 6) = .vsfg1.ValueMatrix(.vsfg1.Row, 6) - .vsfg2.ValueMatrix(.vsfg2.Row, 22) _
                                                             + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5))
                             .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg2.ValueMatrix(.vsfg2.Row, 22) _
                                                + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5)), "#,#.00")
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 22) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 5)   '�ݾ�
                          P_blnSelect = True
                          PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       End If
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm�������ۼ�") Then
                  With frm�������ۼ�
                       .vsfg1.TextMatrix(.vsfg1.Row, 12) = vsfg1.ValueMatrix(vsfg1.Row, 5)     '�ܰ�
                       .vsfg1.TextMatrix(.vsfg1.Row, 13) = vsfg1.ValueMatrix(vsfg1.Row, 6)     '�ΰ�
                       .vsfg1.TextMatrix(.vsfg1.Row, 14) = vsfg1.ValueMatrix(vsfg1.Row, 7)     '����
                       '�հ�ݾ� ���(�ݾ� �ٸ���)
                       If .vsfg1.ValueMatrix(.vsfg1.Row, 15) <> _
                          (.vsfg1.ValueMatrix(.vsfg1.Row, 3) * vsfg1.ValueMatrix(vsfg1.Row, 5)) Then
                          .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg1.ValueMatrix(.vsfg1.Row, 15) _
                                             + (.vsfg1.ValueMatrix(.vsfg1.Row, 3) * vsfg1.ValueMatrix(vsfg1.Row, 5)), "#,#.00")
                       End If
                       .vsfg1.TextMatrix(.vsfg1.Row, 15) = .vsfg1.ValueMatrix(.vsfg1.Row, 3) _
                                                         * vsfg1.ValueMatrix(vsfg1.Row, 5)
                       P_blnSelect = True
                       PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm����������") Then
                  With frm����������
                       If (.vsfg2.ValueMatrix(.vsfg2.Row, 23) <> vsfg1.TextMatrix(vsfg1.Row, 5)) Then '�ܰ� �ٸ���
                           .vsfg2.Cell(flexcpBackColor, .vsfg2.Row, 23, .vsfg2.Row, 23) = vbRed
                           .vsfg2.Cell(flexcpForeColor, .vsfg2.Row, 23, .vsfg2.Row, 23) = vbWhite
                          If .vsfg2.TextMatrix(.vsfg2.Row, 28) = "" Then
                             .vsfg2.TextMatrix(.vsfg2.Row, 28) = "U"
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 23) = vsfg1.ValueMatrix(vsfg1.Row, 5)     '�ܰ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 24) = vsfg1.ValueMatrix(vsfg1.Row, 6)     '�ΰ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 25) = vsfg1.ValueMatrix(vsfg1.Row, 7)     '����
                          '�հ�ݾ� ���(�ݾ��� �ٸ���)
                          If .vsfg2.ValueMatrix(.vsfg2.Row, 26) <> _
                            (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5)) Then
                             .vsfg1.TextMatrix(.vsfg1.Row, 7) = .vsfg1.ValueMatrix(.vsfg1.Row, 7) - .vsfg2.ValueMatrix(.vsfg2.Row, 26) _
                                                             + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5))
                             .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg2.ValueMatrix(.vsfg2.Row, 26) _
                                                + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5)), "#,#.00")
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 26) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 5)     '�ݾ�
                          P_blnSelect = True
                          PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       End If
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm�����ۼ�1") Then
                  With frm�����ۼ�1����
                       If (.vsfg2.ValueMatrix(.vsfg2.Row, 23) <> vsfg1.TextMatrix(vsfg1.Row, 5)) Then '�ܰ� �ٸ���
                           .vsfg2.Cell(flexcpBackColor, .vsfg2.Row, 23, .vsfg2.Row, 23) = vbRed
                           .vsfg2.Cell(flexcpForeColor, .vsfg2.Row, 23, .vsfg2.Row, 23) = vbWhite
                          If .vsfg2.TextMatrix(.vsfg2.Row, 28) = "" Then
                             .vsfg2.TextMatrix(.vsfg2.Row, 28) = "U"
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 23) = vsfg1.ValueMatrix(vsfg1.Row, 5)     '�ܰ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 24) = vsfg1.ValueMatrix(vsfg1.Row, 6)     '�ΰ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 25) = vsfg1.ValueMatrix(vsfg1.Row, 7)     '����
                          '�հ�ݾ� ���(�ݾ��� �ٸ���)
                          If .vsfg2.ValueMatrix(.vsfg2.Row, 26) <> _
                            (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5)) Then
                             .vsfg1.TextMatrix(.vsfg1.Row, 7) = .vsfg1.ValueMatrix(.vsfg1.Row, 7) - .vsfg2.ValueMatrix(.vsfg2.Row, 26) _
                                                             + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5))
                             .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg2.ValueMatrix(.vsfg2.Row, 26) _
                                                + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5)), "#,#.00")
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 26) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 5)     '�ݾ�
                          P_blnSelect = True
                          PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       End If
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm�����ۼ�2") Then
                  With frm�����ۼ�2
                       If (.vsfg1.TextMatrix(.vsfg1.Row, 12) <> vsfg1.TextMatrix(vsfg1.Row, 5)) Then
                          .vsfg1.TextMatrix(.vsfg1.Row, 12) = vsfg1.ValueMatrix(vsfg1.Row, 5)   '�ܰ�
                          .vsfg1.TextMatrix(.vsfg1.Row, 13) = vsfg1.ValueMatrix(vsfg1.Row, 6)   '�ΰ�
                          .vsfg1.TextMatrix(.vsfg1.Row, 14) = vsfg1.ValueMatrix(vsfg1.Row, 7)   '����
                          '�հ�ݾ� ���(�ݾ��� �ٸ���)
                          If .vsfg1.ValueMatrix(.vsfg1.Row, 15) <> _
                            (.vsfg1.ValueMatrix(.vsfg1.Row, 3) * vsfg1.ValueMatrix(vsfg1.Row, 5)) Then
                            .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg1.ValueMatrix(.vsfg1.Row, 15) _
                                               + (.vsfg1.ValueMatrix(.vsfg1.Row, 3) * vsfg1.ValueMatrix(vsfg1.Row, 5)), "#,#.00")
                          End If
                          .vsfg1.TextMatrix(.vsfg1.Row, 15) = .vsfg1.ValueMatrix(.vsfg1.Row, 3) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 5)   '�ݾ�
                          P_blnSelect = True
                          PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       End If
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm�������") Then
                  With frm�������
                       If (.vsfg2.ValueMatrix(.vsfg2.Row, 23) <> vsfg1.TextMatrix(vsfg1.Row, 5)) Then '�ܰ� �ٸ���
                          .vsfg2.Cell(flexcpBackColor, .vsfg2.Row, 23, .vsfg2.Row, 23) = vbRed
                           .vsfg2.Cell(flexcpForeColor, .vsfg2.Row, 23, .vsfg2.Row, 23) = vbWhite
                          If .vsfg2.TextMatrix(.vsfg2.Row, 28) = "" Then
                             .vsfg2.TextMatrix(.vsfg2.Row, 28) = "U"
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 23) = vsfg1.ValueMatrix(vsfg1.Row, 5)   '�ܰ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 24) = vsfg1.ValueMatrix(vsfg1.Row, 6)   '�ΰ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 25) = vsfg1.ValueMatrix(vsfg1.Row, 7)   '����
                          '�հ�ݾ� ���(�ݾ��� �ٸ���)
                          If .vsfg2.ValueMatrix(.vsfg2.Row, 26) <> _
                            (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5)) Then
                            .vsfg1.TextMatrix(.vsfg1.Row, 6) = .vsfg1.ValueMatrix(.vsfg1.Row, 6) - .vsfg2.ValueMatrix(.vsfg2.Row, 26) _
                                                             + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5))
                            .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg2.ValueMatrix(.vsfg2.Row, 26) _
                                               + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5)), "#,#.00")
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 26) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 5)   '�ݾ�
                          P_blnSelect = True
                          PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       End If
                  End With
               End If
            End If
            Unload Me
         End If
    End With
    Exit Sub
End Sub

'+-----------+
'/// ��ȸ ///
'+-----------+
Private Sub cmdFind_Click()
    Screen.MousePointer = vbHourglass
    cmdFind.Enabled = False
    'P_strFindString2 = Trim(Text1(1).Text)  '��ȸ�� ��� �˻��� ����� ����
    Subvsfg1_FILL
    cmdFind.Enabled = True
    Screen.MousePointer = vbDefault
End Sub

'+-----------+
'/// ���� ///
'+-----------+
Private Sub Form_Unload(Cancel As Integer)
    If P_blnSelect = False Then
       SubEscape
    End If
    Screen.MousePointer = vbDefault
    If P_adoRec.State <> adStateClosed Then
       P_adoRec.Close
    End If
    Set P_adoRec = Nothing
    Set frm�����ü��˻� = Nothing
End Sub

'+--------------------+
'/// Sub Procedure ///
'+--------------------+
'+----------------------------------+
'/// VsFlexGrid(vsfgGrid) �ʱ�ȭ ///
'+----------------------------------+
Private Sub Subvsfg1_INIT()
Dim lngR    As Long
Dim lngC    As Long
    With vsfg1           'Rows 1, Cols 8, RowHeightMax(Min) 300
         .AllowBigSelection = False
         .AllowSelection = False
         .AllowUserResizing = flexResizeColumns
         '.BackColorBkg = &H8000000F
         .BackColorSel = &H8000&
         .Ellipsis = flexEllipsisEnd
         '.ExplorerBar = flexExSortShow
         .ExtendLastCol = True
         .FocusRect = flexFocusHeavy
         .FontSize = 9
         .ScrollBars = flexScrollBarHorizontal
         .ScrollTrack = True
         .SelectionMode = flexSelectionByRow
         .FixedRows = 1
         .FixedCols = 0
         .Rows = 1             'SubvsfgUpGrid_Fill����ÿ� ����
         .Cols = 8
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1900   '�����ڵ�(�з��ڵ�+�����ڵ�) 'H
         .ColWidth(1) = 2500   '�����                      'H
         .ColWidth(2) = 2200   '�԰�                        'H
         .ColWidth(3) = 900    '����                        'H
         .ColWidth(4) = 1200   '����
         .ColWidth(5) = 1500   '�ܰ�
         .ColWidth(6) = 1200   '�ΰ�                        'H
         .ColWidth(7) = 1500   '����                        'H
         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "�ڵ�"
         .TextMatrix(0, 1) = "ǰ��"
         .TextMatrix(0, 2) = "�԰�"
         .TextMatrix(0, 3) = "����"
         
         If (PB_strFMCCallFormName = "frm���ּ��ۼ�") Or (PB_strFMCCallFormName = "frm���ּ�����") Then
            .TextMatrix(0, 4) = "��������"
            .TextMatrix(0, 5) = "���ִܰ�"
            .TextMatrix(0, 6) = "���ֺΰ�"
            .TextMatrix(0, 7) = "���ְ���"
         ElseIf _
            (PB_strFMCCallFormName = "frm�����ۼ�1") Or (PB_strFMCCallFormName = "frm�����ۼ�2") Or _
            (PB_strFMCCallFormName = "frm���Լ���") Then
            .TextMatrix(0, 4) = "��������"
            .TextMatrix(0, 5) = "���Դܰ�"
            .TextMatrix(0, 6) = "���Ժΰ�"
            .TextMatrix(0, 7) = "���԰���"
         ElseIf _
            (PB_strFMCCallFormName = "frm�������ۼ�") Or (PB_strFMCCallFormName = "frm����������") Then
            .TextMatrix(0, 4) = "��������"
            .TextMatrix(0, 5) = "�����ܰ�"
            .TextMatrix(0, 6) = "�����ΰ�"
            .TextMatrix(0, 7) = "��������"
         ElseIf _
            (PB_strFMCCallFormName = "frm�����ۼ�1") Or (PB_strFMCCallFormName = "frm�����ۼ�2") Or _
            (PB_strFMCCallFormName = "frm�������") Then
            .TextMatrix(0, 4) = "��������"
            .TextMatrix(0, 5) = "����ܰ�"
            .TextMatrix(0, 6) = "����ΰ�"
            .TextMatrix(0, 7) = "���Ⱑ��"
         End If
         .ColHidden(0) = True: .ColHidden(1) = True: .ColHidden(2) = True:: .ColHidden(3) = True
         .ColHidden(6) = True: .ColHidden(7) = True
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 5 To 7
                         .ColFormat(lngC) = "#,#.00"
             End Select
         Next lngC
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 0, 1, 2, 3
                        .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 4
                        .ColAlignment(lngC) = flexAlignCenterCenter
                    Case Else
                        .ColAlignment(lngC) = flexAlignRightCenter
             End Select
         Next lngC
         .MergeCells = flexMergeRestrictAll
         'For lngC = 0 To 5
         '    .MergeCol(lngC) = True
         'Next lngC
    End With
End Sub

'+---------------------------------+
'/// VsFlexGrid(vsfgGrid) ä���///
'+---------------------------------+
Private Sub Subvsfg1_FILL()
Dim strSQL     As String
Dim strWhere   As String
Dim strOrderBy As String
Dim lngR       As Long
Dim lngC       As Long
Dim lngRR      As Long
Dim lngRRR     As Long
Dim strAppDate As Long
    Screen.MousePointer = vbHourglass
    P_adoRec.CursorLocation = adUseClient
    If (PB_strFMCCallFormName = "frm���ּ��ۼ�") Or (PB_strFMCCallFormName = "frm���ּ�����") Then
       strSQL = "SELECT TOP 5 " _
                     & "T1.�����ڵ� AS �����ڵ�, J1.����� AS ǰ��, " _
                     & "J1.�԰� AS �԰�, J1.���� AS ����, T1.�������� AS ����, " _
                     & "T1.�԰�ܰ� AS �ܰ�, T1.�԰�ΰ� AS �ΰ� " _
                & "FROM ���ֳ��� T1 " _
                & "LEFT JOIN ���� J1 ON (J1.�з��ڵ� + J1.�����ڵ�) = T1.�����ڵ� " _
               & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                 & "AND T1.�������� <= '" & P_strAppDate & "' " _
                 & "AND T1.����ó�ڵ� = '" & PB_strSupplierCode & "' " _
                 & "AND T1.�����ڵ� = '" & PB_strMaterialsCode & "' " _
                 & "AND T1.��뱸�� = 0 AND T1.�԰�ܰ� > 0 " _
               & "GROUP BY T1.�����ڵ�, J1.�����, J1.�԰�, J1.����, T1.��������, T1.�԰�ܰ�, T1.�԰�ΰ� " _
               & "ORDER BY T1.�������� DESC "
    End If
    If (PB_strFMCCallFormName = "frm�����ۼ�1") Or (PB_strFMCCallFormName = "frm�����ۼ�2") Or _
       (PB_strFMCCallFormName = "frm���Լ���") Then
       strSQL = "SELECT TOP 5 " _
                     & "(T1.�з��ڵ� + T1.�����ڵ�) AS �����ڵ�, J1.����� AS ǰ��, " _
                     & "J1.�԰� AS �԰�, J1.���� AS ����, T1.��������� AS ����, " _
                     & "T1.�԰�ܰ� AS �ܰ�, T1.�԰�ΰ� AS �ΰ� " _
                & "FROM �������⳻�� T1 " _
                & "LEFT JOIN ���� J1 ON J1.�з��ڵ� = T1.�з��ڵ� AND J1.�����ڵ� = T1.�����ڵ� " _
               & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                 & "AND T1.��������� <= '" & P_strAppDate & "' " _
                 & "AND T1.����ó�ڵ� = '" & PB_strSupplierCode & "' " _
                 & "AND T1.�з��ڵ� = '" & Mid(PB_strMaterialsCode, 1, 2) & "' " _
                 & "AND T1.�����ڵ� = '" & Mid(PB_strMaterialsCode, 3) & "' " _
                 & "AND T1.��뱸�� = 0 AND T1.������� = 1 AND T1.�԰�ܰ� > 0 " _
               & "GROUP BY (T1.�з��ڵ� + T1.�����ڵ�), J1.�����, J1.�԰�, J1.����, T1.���������, T1.�԰�ܰ�, T1.�԰�ΰ� " _
               & "ORDER BY T1.��������� DESC "
    End If
    
    If (PB_strFMCCallFormName = "frm�������ۼ�") Or (PB_strFMCCallFormName = "frm����������") Then
       strSQL = "SELECT TOP 5 " _
                     & "T1.�����ڵ� AS �����ڵ�, J1.����� AS ǰ��, " _
                     & "J1.�԰� AS �԰�, J1.���� AS ����, T1.�������� AS ����, " _
                     & "T1.���ܰ� AS �ܰ�, T1.���ΰ� AS �ΰ� " _
                & "FROM �������� T1 " _
                & "LEFT JOIN ���� J1 ON (J1.�з��ڵ� + J1.�����ڵ�) = T1.�����ڵ� " _
               & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                 & "AND T1.�������� <= '" & P_strAppDate & "' " _
                 & "AND T1.����ó�ڵ� = '" & PB_strSupplierCode & "' " _
                 & "AND T1.�����ڵ� = '" & PB_strMaterialsCode & "' " _
                 & "AND T1.��뱸�� = 0 AND T1.���ܰ� > 0 " _
               & "GROUP BY T1.�����ڵ�, J1.�����, J1.�԰�, J1.����, T1.��������, T1.���ܰ�, T1.���ΰ� " _
               & "ORDER BY T1.�������� DESC "
    End If
    If (PB_strFMCCallFormName = "frm�����ۼ�1") Or (PB_strFMCCallFormName = "frm�����ۼ�2") Or _
       (PB_strFMCCallFormName = "frm�������") Then
       strSQL = "SELECT TOP 5 " _
                     & "(T1.�з��ڵ� + T1.�����ڵ�) AS �����ڵ�, J1.����� AS ǰ��, " _
                     & "J1.�԰� AS �԰�, J1.���� AS ����, T1.��������� AS ����, " _
                     & "T1.���ܰ� AS �ܰ�, T1.���ΰ� AS �ΰ� " _
                & "FROM �������⳻�� T1 " _
                & "LEFT JOIN ���� J1 ON J1.�з��ڵ� = T1.�з��ڵ� AND J1.�����ڵ� = T1.�����ڵ� " _
               & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                 & "AND T1.��������� <= '" & P_strAppDate & "' " _
                 & "AND T1.����ó�ڵ� = '" & PB_strSupplierCode & "' " _
                 & "AND T1.�з��ڵ� = '" & Mid(PB_strMaterialsCode, 1, 2) & "' " _
                 & "AND T1.�����ڵ� = '" & Mid(PB_strMaterialsCode, 3) & "' " _
                 & "AND T1.��뱸�� = 0 AND T1.������� = 2 AND T1.���ܰ� > 0 " _
               & "GROUP BY (T1.�з��ڵ� + T1.�����ڵ�), J1.�����, J1.�԰�, J1.����, T1.���������, T1.���ܰ�, T1.���ΰ� " _
               & "ORDER BY T1.��������� DESC "
    End If
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    vsfg1.Rows = P_adoRec.RecordCount + vsfg1.FixedRows
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       Screen.MousePointer = vbDefault
       Exit Sub
    Else
       With vsfg1
            .Cell(flexcpForeColor, 0, 0, .FixedRows - 1, .Cols - 1) = vbBlack
            .Rows = P_adoRec.RecordCount + 1
            If .Rows <= PC_intRowCnt Then
               .ScrollBars = flexScrollBarHorizontal
            Else
               .ScrollBars = flexScrollBarBoth
            End If
            Do Until P_adoRec.EOF
               lngR = lngR + 1
               .TextMatrix(lngR, 0) = P_adoRec("�����ڵ�")
               .TextMatrix(lngR, 1) = P_adoRec("ǰ��")
               .TextMatrix(lngR, 2) = P_adoRec("�԰�")
               .TextMatrix(lngR, 3) = P_adoRec("����")
               .TextMatrix(lngR, 4) = Format(P_adoRec("����"), "0000-00-00")
               .TextMatrix(lngR, 5) = P_adoRec("�ܰ�")
               .Cell(flexcpData, lngR, 4, lngR, 4) = .TextMatrix(lngR, 4) + "-" + .TextMatrix(lngR, 5)
               .TextMatrix(lngR, 6) = P_adoRec("�ΰ�")
               .TextMatrix(lngR, 7) = .TextMatrix(lngR, 5) + .TextMatrix(lngR, 6)
               'If PB_strMaterialsCode = Trim(.TextMatrix(lngR, 0)) Then
               '   lngRR = lngR
               'End If
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
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����ü� �б� ����"
    Unload Me
    Exit Sub
End Sub

'+----------+
'/// ESC ///
'+----------+
Private Sub SubEscape()
    PB_strFMCCallFormName = ""
    PB_strMaterialsCode = ""
    PB_strMaterialsName = ""
    PB_strSupplierCode = ""
    Unload Me
End Sub

