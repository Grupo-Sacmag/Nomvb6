VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form CapCst 
   Caption         =   "Modulo Captura Costo Directo"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9900
   FillColor       =   &H00E0E0E0&
   Icon            =   "CapCst.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   9900
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox EntXT 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   3015
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1680
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid CST 
      Height          =   4695
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   8281
      _Version        =   393216
      BackColor       =   16777215
      Rows            =   100
      FixedCols       =   0
      GridColor       =   -2147483627
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      HighLight       =   2
      GridLinesFixed  =   1
      GridLinesUnpopulated=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Poliza :"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Height          =   645
      Left            =   4800
      TabIndex        =   3
      Top             =   360
      Width           =   4095
   End
End
Attribute VB_Name = "CapCst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ccm As Long, Dcm As Long, Ecm As Long, Fcm As Long, Crgtro As Long, VaCio As Long
Dim Mx_Ren As Long, Po_ct As Integer, ACHCOS As String, Sal_do As Currency
Sub Apl_Cto()
 
 If Sal_do = 0 Then
  Close 11
  Open "C:\Archivos de programa\NOMINA1\perma.dno" For Random As #7 Len = Len(basico)
     fin_basico = LOF(7) / Len(basico)
     Get 7, 2, basico
     basico.datoarch = Trim(Direc_torio)
     Open basico.datoarch + "DATOS" For Random As 11 Len = Len(DATOS)
     Fcm = LOF(11) / Len(DATOS)
     Get 1, 1, DATOS
     Close 11
     
  Open Trim(basico.datoarch) + "DATOS" For Random As 11 Len = Len(oper)
  Fcm = LOF(11) / Len(oper)
  oper.cta = String(6 - Len(Str(Alarma.Otr)), " ") + Str(Alarma.Otr)
  If ultimo.TipoCap = 0 Then
        oper.descr = LTrim(RTrim(cheques.Text5.Text)) + _
                    " " + Left(cheques.Text1.Text, 20) + " " + _
                    Left(cheques.Text4.Text, 8)
        Else
        oper.descr = Trim(cheques.Text7.Text)
  End If
  oper.identi = "A"
  oper.fe = LTrim(RTrim(cheques.Text3.Text))
  oper.impte = ""
  oper.Real = ""
  Rem ***************************************************************************
  Fcm = Fcm + 1
  Put 11, Fcm, oper
  oper.cta = "  1105"
  oper.descr = ""
  oper.fe = ""
  impo_rte = Alarma.Imt
  oper.identi = "B"
  oper.impte = Str(impo_rte)
  oper.Real = Alarma.Clt
  Fcm = Fcm + 1
  Put 11, Fcm, oper
  Rem **************************************************************************
  For IW = 1 To CST.Rows - 1
     If IsNumeric(CST.TextMatrix(IW, 0)) Then
            oper.cta = String(6 - Len(Str(CST.TextMatrix(IW, 5))), " ") + _
                       Str(CST.TextMatrix(IW, 5))
            oper.descr = RTrim(CST.TextMatrix(IW, 4))
            oper.fe = ""
            impo_rte = CST.TextMatrix(IW, 2)
            oper.identi = "C"
            oper.impte = Str(impo_rte)
            oper.Real = CST.TextMatrix(IW, 0)
            Fcm = Fcm + 1
            Put 11, Fcm, oper
      End If
   Next IW
  oper.cta = "  5105"
  oper.descr = ""
  oper.fe = ""
  impo_rte = (Alarma.Imt * -1)
  oper.identi = "B"
  oper.impte = Str(impo_rte)
  oper.Real = Alarma.Gto
  Fcm = Fcm + 1
  Put 11, Fcm, oper
  CST.Clear
  Unload Me
        Else
        MsgBox "EL SALDO NO ES IGUAL A CEROS ", vbCritical
    
   End If
End Sub

Private Sub CST_EnterCell()
   Rem If CST.Row > 0 Then
       CST.CellBackColor = vbGreen
       EntXT.Text = CST.Text
       
   Rem End If
        
End Sub
Private Sub CST_KeyDown(KeyCode As Integer, Shift As Integer)
      Select Case KeyCode
            Case vbKeyDelete
                CST.Text = ""
                EntXT.Text = CST.Text
            Case vbKeyF2
                If CST.Text <> "" Then valcelant = CST.Text
                EntXT.Text = LTrim(RTrim(CST.Text))
                EntXT.SetFocus
               
       End Select
End Sub

Private Sub CST_KeyPress(KeyAscii As Integer)
         valcelant1 = CST.Text
         EntXT.Text = Chr$(KeyAscii)
         EntXT.SetFocus
End Sub

Private Sub CST_LeaveCell()
    Rem If CST.Row > 0 Then
       CST.CellBackColor = vbWhite
    Rem End If
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
     Select Case Button.Key
       Case Is = "Guardar"
            Apl_Cto
       Case Is = "CatScta"
            trcta.inicia = 811: trcta.termina = 1210
           Close 3: Open Dir_Costos + "CATAUX" For Random As 3 Len = Len(CATAUX)
           dm = LOF(3) / Len(CATAUX)
            CATSUB.Show 1
            CST.TextMatrix(CST.Row, 0) = trscta.num
            CST.TextMatrix(CST.Row, 1) = " " + trscta.nombre
            CST.TextMatrix(CST.Row, 5) = trscta.donde
            CST_LeaveCell
            CST.Col = 2
            CST_EnterCell
       Case Is = "Eliminar"
            CST.RemoveItem CST.Row
            Mx_Ren = Mx_Ren - 1
            Rem If CST.Row > 1 Then
               Sum_Sdo
            Rem End If
            CST_EnterCell
       Case Is = "CatCta"
            Rem
     End Select

End Sub

Sub Barr_Hta1()
   Dim imgX As ListImage
    Toolbar1.Style = tbrFlat
    Toolbar1.Align = 0
    ' Carga imágenes en el control ImageList.
    Set imgX = ImageList1.ListImages. _
    Add(, "Guardar", LoadPicture(Ruta_Acceso + "\save.bmp"))
        Set imgX = ImageList1.ListImages. _
    Add(, "CatScta", LoadPicture(Ruta_Acceso + "\tarjeta.bmp"))
        Set imgX = ImageList1.ListImages. _
    Add(, "Eliminar", LoadPicture(Ruta_Acceso + "\borrscta.bmp"))
        Set imgX = ImageList1.ListImages. _
    Add(, "CatCta", LoadPicture(Ruta_Acceso + "\CatCta.bmp"))
        Toolbar1.ImageList = ImageList1
    Dim btnX As Button
    Toolbar1.Buttons.Add , , , tbrPlaceholder
    Set btnX = Toolbar1.Buttons.Add(, "Guardar", , tbrDefault, "Guardar")
    btnX.ToolTipText = "Guardar Aplicación"
        Set btnX = Toolbar1.Buttons.Add(, "CatScta", , tbrDefault, "CatScta")
    btnX.ToolTipText = "Ver SubCuentas"
        btnX.Description = btnX.ToolTipText
        Set btnX = Toolbar1.Buttons.Add(, "Eliminar", , tbrDefault, "Eliminar")
    btnX.ToolTipText = "Eliminar Renglón"
        Set btnX = Toolbar1.Buttons.Add(, "CatCta", , tbrDefault, "CatCta")
    btnX.ToolTipText = "Catalogo de cuentas"
      

End Sub



Private Sub ENTXT_Change()
        
     CST.Text = EntXT.Text
End Sub
Private Sub ENTXT_GotFocus()
    SendKeys "{end}"
End Sub

Private Sub ENTXT_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         CST_LeaveCell
         Select Case CST.Col
            Case 4
               EntXT.Text = UCase(EntXT.Text)
               CST_LeaveCell
               CST.Row = CST.Row + 1
               CST_EnterCell
               CST.SetFocus
            Case 1
              Close 3: Open Dir_Costos + "CATAUX" For Random As 3 Len = Len(CATAUX)
              dm = LOF(3) / Len(CATAUX)
              If (IsNumeric(CST.TextMatrix(CST.Row, 0))) Then
                 Get 3, VaCio, CATAUX
                 EntXT.Text = UCase(EntXT.Text)
                 CATAUX.C1 = CST.TextMatrix(CST.Row, 0)
                 CATAUX.C2 = EntXT.Text
                 CST.TextMatrix(CST.Row, 5) = VaCio
                 Put 3, VaCio, CATAUX
                 Get 3, VaCio - 800, CATAUX
                 CATAUX.C1 = CST.TextMatrix(CST.Row, 0)
                 CATAUX.C2 = EntXT.Text
                 Put 3, VaCio - 800, CATAUX
                 Get 3, VaCio - 400, CATAUX
                 CATAUX.C1 = CST.TextMatrix(CST.Row, 0)
                 CATAUX.C2 = EntXT.Text
                 Put 3, VaCio - 400, CATAUX
                 CST.Text = EntXT.Text
                 CST_LeaveCell
                 CST.Col = 2: CST_EnterCell
                 CST.SetFocus
              End If
            Case 0
              If IsNumeric(EntXT.Text) Then
                 Close 3: Open Dir_Costos + "CATAUX" For Random As 3 Len = Len(CATAUX)
                 dm = LOF(3) / Len(CATAUX)
                 Bs_Vacio
                 exte = 0
                 Crgtro = EntXT.Text
                 For BKQ = 811 To 1210: Get 3, BKQ, CATAUX
                     If Crgtro = Val(CATAUX.C1) Then
                          exte = 1
                          VaCio = BKQ
                          CST.TextMatrix(CST.Row, 1) = " " + Trim(CATAUX.C2)
                          CST.TextMatrix(CST.Row, 5) = VaCio
                          CST.Col = 2
                          CST_EnterCell
                          CST.SetFocus
                          Exit For
                     End If
                 Next BKQ
                 If exte = 0 Then
                    CST.Col = 1
                    CST_EnterCell
                    CST.SetFocus
                 End If
                 Close
              End If
            Case 2
                If IsNumeric(EntXT) Then
                    CST.Text = Format(EntXT.Text, z1)
                    Sum_Sdo
                    CST.TextMatrix(CST.Row, 4) = Alarma.Rda
                    
                    CST.Row = CST.Row + 1: CST.Col = 0
                    CST_EnterCell
                    CST.SetFocus
                    Else
                    EntXT.Text = ""
                End If
         End Select
    End If
    
End Sub
Sub Sum_Sdo()
    Dim We As Long
    Sal_do = Alarma.Imt
    If CST.Row > Mx_Ren Then Mx_Ren = CST.Row
    
    For We = 1 To Mx_Ren
        If IsNumeric(CST.TextMatrix(We, 2)) Then
            Sal_do = Sal_do - CST.TextMatrix(We, 2)
        End If
        CST.TextMatrix(We, 3) = Format(Sal_do, z1)
    Next We
    
End Sub
Sub Bs_Vacio()
   
  For BKQ = 811 To 1210: Get 3, BKQ, CATAUX
         If Val(CATAUX.C1) = 0 Then
                 VaCio = BKQ
                 Exit For
         End If
  Next BKQ
End Sub
Private Sub Form_Load()
     Barr_Hta1
     CST.Cols = 6: CST.Row = 0
     CST.Col = 1: CST.ColWidth(0) = 1200:  CST.CellAlignment = 4: CST.Text = "Obra"
     CST.CellBackColor = vbYellow
     CST.Col = 2: CST.ColWidth(1) = 3200:  CST.CellAlignment = 4: CST.Text = "Nombre"
     CST.CellBackColor = vbYellow
     CST.Col = 2: CST.ColWidth(2) = 1200:  CST.CellAlignment = 4: CST.Text = "Importe"
     CST.CellBackColor = vbYellow
     CST.Col = 3: CST.ColWidth(3) = 1200:  CST.CellAlignment = 4: CST.Text = "Saldo"
     CST.TextMatrix(1, 3) = Format(Alarma.Imt, z1)
     CST.CellBackColor = vbYellow
     CST.Col = 4: CST.ColWidth(4) = 3200:  CST.CellAlignment = 4: CST.Text = "Redaccion"
     CST.CellBackColor = vbYellow
     CST.Col = 5: CST.ColWidth(5) = 100
     CST.CellBackColor = vbYellow
     CST.TextMatrix(1, 4) = Alarma.Rda
     CST.Row = 1: CST.Col = 0
     Ap_t
End Sub
Sub Ap_t()
   Dim Wr As Long
   Open Dir_Costos + "Datos" For Random As 8 Len = Len(DATOS)
   Ccm = LOF(8) / Len(DATOS)
   Get 8, 1, DATOS
   CapCst.Caption = CapCst.Caption + Trim(DATOS.D1) + Trim(DATOS.a_o)
   Open Dir_Costos + "Catmay" For Random As 9 Len = Len(CATMAY)
   Dcm = LOF(9) / Len(CATMAY)
   For Wr = 1 To Dcm: Get 9, Wr, CATMAY
          If Val(CATMAY.B1) = 1105 Then
              
              Alarma.Clt = Wr
              Exit For
          End If
   Next Wr
   For Wr = 1 To Dcm: Get 9, Wr, CATMAY
          If Val(CATMAY.B1) = 5105 Then
              
              Alarma.Gto = Wr
              Exit For
          End If
   Next Wr
   
   Open Dir_Costos + "Cataux" For Random As 10 Len = Len(CATAUX)
   Ecm = LOF(10) / Len(CATAUX)
   If m_m < 10 Then
         ACHCOS = "COS0" + Trim(Str(m_m))
         Else
         ACHCOS = "COS" + Trim(Str(m_m))
   End If
   Open Dir_Costos + ACHCOS For Random As 11 Len = Len(oper)
   Fcm = LOF(11) / Len(oper)
   If Fcm < 1 Then
        Alarma.Otr = 1
        Po_ct = 1
        Label1.Caption = " 1" + " - " + Alarma.Rda + " " + Trim(mm(m_m))
        Else
        For tw = Fcm To 1 Step -1: Get 11, tw, oper
            If oper.identi = "A" Then
              Po_ct = Val(oper.cta) + 1
              Alarma.Otr = Po_ct
              Label1.Caption = Trim(oper.cta) + " - " + Alarma.Rda + " " + Trim(mm(m_m))
              Exit For
            End If
        Next tw
   End If
   Close
End Sub

Private Sub Form_Resize()
   If CapCst.WindowState <> 1 Then
            Rem CST.Width = 11000
            CST.Width = CapCst.ScaleWidth - 400
            CST.Height = CapCst.ScaleHeight - 1200
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
        Close
End Sub
