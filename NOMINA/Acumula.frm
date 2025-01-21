VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   5475
   ClientLeft      =   1080
   ClientTop       =   1950
   ClientWidth     =   9360
   LinkTopic       =   "Form9"
   ScaleHeight     =   5475
   ScaleWidth      =   9360
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1920
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4335
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   7646
      _Version        =   327680
   End
   Begin VB.Menu Acu_arch 
      Caption         =   "&Archivos"
      Index           =   0
      Begin VB.Menu Acueleg 
         Caption         =   "&Elegir Archivo"
         Index           =   2
      End
      Begin VB.Menu AcuAcu 
         Caption         =   "&Actualizar"
         Index           =   3
         Begin VB.Menu Acttot 
            Caption         =   "&total"
            Index           =   1
         End
         Begin VB.Menu Actind 
            Caption         =   "Individual"
            Index           =   2
         End
      End
      Begin VB.Menu acuapl 
         Caption         =   "A&plicar Acumulado"
         Index           =   4
      End
      Begin VB.Menu Acu_sal 
         Caption         =   "&Salida"
         Index           =   1
      End
   End
   Begin VB.Menu Acuord 
      Caption         =   "&Ordenar"
      Index           =   0
      Begin VB.Menu Acuordalf 
         Caption         =   "&Alfabeticamente"
         Index           =   1
      End
      Begin VB.Menu AcuorNum 
         Caption         =   "&Numericamente"
         Index           =   2
      End
      Begin VB.Menu AcuorEli 
         Caption         =   "&Eliminar"
         Index           =   3
      End
   End
   Begin VB.Menu AcuVer 
      Caption         =   "&Verificar"
      Index           =   0
      Begin VB.Menu AcuVerVer 
         Caption         =   "&Verificar"
         Index           =   1
      End
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim limite As Long, Archivo_acu
Sub eliminacion()
           antecol = MSFlexGrid1.Col
           anteren = MSFlexGrid1.Row
           conta1 = 0
           re = 0
         Do Until re = (limite)
              re = re + 1
              dato_ent = 0: dato_sal = 0: valcelant = 0
              MSFlexGrid1.Row = re
              MSFlexGrid1.Col = 13
              If MSFlexGrid1.Text = "" Then
                    dato_ent = 0
                  Else
                    dato_ent = MSFlexGrid1.Text
              End If
              MSFlexGrid1.Col = 19
              If MSFlexGrid1.Text = "" Then
                    dato_sal = 0
                    Else
                    dato_sal = MSFlexGrid1.Text
               End If
             
              MSFlexGrid1.Col = 20
              If MSFlexGrid1.Text = "" Then
                    valcelant = 0
                    Else
                    valcelant = MSFlexGrid1.Text
              End If
         If dato_ent = 0 And dato_sal = 0 And valcelant = 0 Then
              MSFlexGrid1.RemoveItem re
              re = re - 1
              limite = limite - 1
              Else
                 dato_ent = 0
                 dato_sal = 0
                 valcelant = 0
         End If
    Loop
        If anteren <= limite Then MSFlexGrid1.Row = anteren Else MSFlexGrid1.Row = limite
        If anteren < 1 Then MSFlexGrid1.Row = 1
        MSFlexGrid1.Col = antecol
        
End Sub
Sub cargaper()
   For r = 1 To nq_3
       Get #3, r, personal
               aobaja = Val(Mid$(personal.fab, 7, 4))
               mesbaja = Val(Mid$(personal.fab, 4, 2))
               diabaja = Val(Mid$(personal.fab, 1, 2))
       If aobaja > 0 And aobaja < empresa.ao Then GoTo sigueL
               renglon = renglon + 1: MSFlexGrid1.Row = renglon
               MSFlexGrid1.Col = 0: MSFlexGrid1.Text = Format(r, "#####")
               MSFlexGrid1.Col = 1
               MSFlexGrid1.Text = RTrim(personal.ape1) + " " + RTrim(personal.ape2) + " " + RTrim(personal.nom)
               limite = limite + 1
sigueL:
   Next r
End Sub

Sub iniciar()
   
   MSFlexGrid1.Font = "Arial"
   MSFlexGrid1.Font.Size = 8
   MSFlexGrid1.Font.Bold = True
   MSFlexGrid1.Row = 0: MSFlexGrid1.Col = 0
   MSFlexGrid1.ColWidth(0) = 400: MSFlexGrid1.CellAlignment = 4
   MSFlexGrid1.Text = "No."
   MSFlexGrid1.Col = 1: MSFlexGrid1.CellAlignment = 4: MSFlexGrid1.ColWidth(1) = 3500: MSFlexGrid1.Text = "Nombre"
   MSFlexGrid1.Col = 2: MSFlexGrid1.CellAlignment = 4: MSFlexGrid1.ColWidth(2) = 1200: MSFlexGrid1.Text = "dias T."
   MSFlexGrid1.Col = 3: MSFlexGrid1.CellAlignment = 4: MSFlexGrid1.ColWidth(3) = 1200: MSFlexGrid1.Text = "Sueldo"
   MSFlexGrid1.Col = 4: MSFlexGrid1.CellAlignment = 4: MSFlexGrid1.ColWidth(4) = 1200: MSFlexGrid1.Text = "hs.Extras"
   MSFlexGrid1.Col = 5: MSFlexGrid1.CellAlignment = 4: MSFlexGrid1.ColWidth(5) = 1200: MSFlexGrid1.Text = "Viaticos"
   MSFlexGrid1.Col = 6: MSFlexGrid1.CellAlignment = 4: MSFlexGrid1.ColWidth(6) = 1200: MSFlexGrid1.Text = "P.Vacac."
   MSFlexGrid1.Col = 7: MSFlexGrid1.CellAlignment = 4: MSFlexGrid1.ColWidth(7) = 1200: MSFlexGrid1.Text = "Otras"
   MSFlexGrid1.Col = 8: MSFlexGrid1.CellAlignment = 4: MSFlexGrid1.ColWidth(8) = 1200: MSFlexGrid1.Text = "Exentos"
   MSFlexGrid1.Col = 9: MSFlexGrid1.CellAlignment = 4: MSFlexGrid1.ColWidth(9) = 1200: MSFlexGrid1.Text = "Aguinaldo"
   MSFlexGrid1.Col = 10: MSFlexGrid1.CellAlignment = 4: MSFlexGrid1.ColWidth(10) = 1200: MSFlexGrid1.Text = "Part.de.Util."
   MSFlexGrid1.Col = 11: MSFlexGrid1.CellAlignment = 4: MSFlexGrid1.ColWidth(11) = 1200: MSFlexGrid1.Text = "Tot.Ingr."
   MSFlexGrid1.Col = 12: MSFlexGrid1.CellAlignment = 4: MSFlexGrid1.ColWidth(12) = 1200: MSFlexGrid1.Text = "Ispt"
   MSFlexGrid1.Col = 13: MSFlexGrid1.CellAlignment = 4: MSFlexGrid1.ColWidth(13) = 1200: MSFlexGrid1.Text = "Cr.Salario"
   MSFlexGrid1.Col = 14: MSFlexGrid1.CellAlignment = 4: MSFlexGrid1.ColWidth(14) = 1200: MSFlexGrid1.Text = "Imss"
   MSFlexGrid1.Col = 15: MSFlexGrid1.CellAlignment = 4: MSFlexGrid1.ColWidth(15) = 1200: MSFlexGrid1.Text = "Prestamos"
   MSFlexGrid1.Col = 16: MSFlexGrid1.CellAlignment = 4: MSFlexGrid1.ColWidth(16) = 1200: MSFlexGrid1.Text = "Fonacot"
   MSFlexGrid1.Col = 17: MSFlexGrid1.CellAlignment = 4: MSFlexGrid1.ColWidth(17) = 1200: MSFlexGrid1.Text = "Telefonos"
   MSFlexGrid1.Col = 18: MSFlexGrid1.CellAlignment = 4: MSFlexGrid1.ColWidth(18) = 1200: MSFlexGrid1.Text = "Infonavit"
   MSFlexGrid1.Col = 19: MSFlexGrid1.CellAlignment = 4: MSFlexGrid1.ColWidth(19) = 1200: MSFlexGrid1.Text = "Tot.Deduc"
   MSFlexGrid1.Col = 20: MSFlexGrid1.CellAlignment = 4: MSFlexGrid1.ColWidth(20) = 1200: MSFlexGrid1.Text = "Neto"
   MSFlexGrid1.Col = 21: MSFlexGrid1.ColWidth(21) = 120: Rem  MSFlexGrid1.Text = "Neto"
   MSFlexGrid1.Col = 22: MSFlexGrid1.ColWidth(22) = 120: Rem  MSFlexGrid1.Text = "Neto"
   MSFlexGrid1.Col = 23: MSFlexGrid1.ColWidth(23) = 120: Rem  MSFlexGrid1.Text = "Neto"
End Sub

Private Sub Actind_Click(Index As Integer)
    inter_clave = 0
    Close 1: Close 2
    Open "Acumul.nom" For Random As 2 Len = Len(acumulado)
    nq_2 = LOF(2) / Len(acumulado)
    If nq_2 = 0 Then blanc0
    If Archivo_acu <> "" Then
    Open Archivo_acu For Random As 1 Len = Len(nomina)
    nq_1 = LOF(1) / Len(nomina)
    Load Form10
    Form10.Show
     Else
     MsgBox "Primero debe elegir el archivo que actualiza", vbCritical + vbDefaultButton1, "Acumulado de sueldos"
    End If

End Sub

Private Sub Acttot_Click(Index As Integer)
  inter_clave = 2
  Load Form10
  Form10.Show
  
End Sub

Private Sub Acu_sal_Click(Index As Integer)
   Unload Form9
   Load Form1
   Form1.Show
End Sub

Private Sub acuapl_Click(Index As Integer)
    Close 2
    Open "Acumul.nom" For Random As 2 Len = Len(acumulado)
    nq_2 = LOF(2) / Len(acumulado)
    If nq_2 = 0 Then
       MsgBox "No hay archivo de acumulados"
       Else
        For r = 1 To limite:
           MSFlexGrid1.Row = r
           Get 2, MSFlexGrid1.TextMatrix(r, 0), acumulado
           acucar
        Next r
    End If
End Sub
Sub acucar()
   Rem MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = RTrim$(personal.ape1) + " " + RTrim$(personal.ape2) + " " + RTrim$(personal.nom)
   ingresos = 0: deducciones = 0: neto = 0
   li = MSFlexGrid1.Row: ii = 2
   If acumulado.dias <> 0 Then MSFlexGrid1.TextMatrix(li, 2) = Format(acumulado.dias, "##0.00") Else MSFlexGrid1.TextMatrix(li, ii) = ""
   ii = ii + 1: Rem 3
   If acumulado.sueldo <> 0 Then MSFlexGrid1.TextMatrix(li, 3) = Format(acumulado.sueldo, z1$) Else MSFlexGrid1.TextMatrix(li, ii) = ""
   ii = ii + 1: Rem 4
   If acumulado.hs_ext <> 0 Then MSFlexGrid1.TextMatrix(li, 4) = Format(acumulado.hs_ext, z1$) Else MSFlexGrid1.TextMatrix(li, ii) = ""
   ii = ii + 1: Rem 5
   If acumulado.viaticos <> 0 Then MSFlexGrid1.TextMatrix(li, 5) = Format(acumulado.viaticos, z1$) Else MSFlexGrid1.TextMatrix(li, ii) = ""
   ii = ii + 1: Rem 6
   If acumulado.pvac <> 0 Then MSFlexGrid1.TextMatrix(li, 6) = Format(acumulado.pvac, z1$) Else MSFlexGrid1.TextMatrix(li, ii) = ""
   ii = ii + 1: Rem 7
   If acumulado.otras <> 0 Then MSFlexGrid1.TextMatrix(li, 7) = Format(acumulado.otras, z1$) Else MSFlexGrid1.TextMatrix(li, ii) = ""
   ii = ii + 1: Rem 8
   If acumulado.exentos <> 0 Then MSFlexGrid1.TextMatrix(li, 8) = Format(acumulado.exentos, z1$) Else MSFlexGrid1.TextMatrix(li, ii) = ""
   ii = ii + 1: Rem 9
   If acumulado.aguin <> 0 Then MSFlexGrid1.TextMatrix(li, 9) = Format(acumulado.aguin, z1$) Else MSFlexGrid1.TextMatrix(li, ii) = ""
   ii = ii + 1: Rem 10
   If acumulado.ptu <> 0 Then MSFlexGrid1.TextMatrix(li, 10) = Format(acumulado.ptu, z1$) Else MSFlexGrid1.TextMatrix(li, ii) = ""
   Rem cuotahr = personal.ingr / 8
   ingresos = acumulado.sueldo + acumulado.hs_ext + acumulado.viaticos + acumulado.pvac + acumulado.otras + acumulado.exentos + acumulado.aguin + acumulado.ptu
   ii = ii + 1: Rem 11
   If ingresos <> 0 Then MSFlexGrid1.TextMatrix(li, 11) = Format(ingresos, z1$) Else MSFlexGrid1.TextMatrix(li, ii) = ""
   ii = ii + 1: Rem 12
   If acumulado.ispt <> 0 Then MSFlexGrid1.TextMatrix(li, 12) = Format(acumulado.ispt, z1$) Else MSFlexGrid1.TextMatrix(li, ii) = ""
   ii = ii + 1: Rem 13
   If acumulado.crdsal <> 0 Then MSFlexGrid1.TextMatrix(li, 13) = Format(acumulado.crdsal, z1$) Else MSFlexGrid1.TextMatrix(li, ii) = ""
   ii = ii + 1: Rem 14
   If acumulado.imss <> 0 Then MSFlexGrid1.TextMatrix(li, 14) = Format(acumulado.imss, z1$) Else MSFlexGrid1.TextMatrix(li, ii) = ""
   ii = ii + 1: Rem 15
   If acumulado.prestamos <> 0 Then MSFlexGrid1.TextMatrix(li, 15) = Format(acumulado.prestamos, z1$) Else MSFlexGrid1.TextMatrix(li, ii) = ""
   ii = ii + 1: Rem 16
   If acumulado.fonacot <> 0 Then MSFlexGrid1.TextMatrix(li, 16) = Format(acumulado.fonacot, z1$) Else MSFlexGrid1.TextMatrix(li, ii) = ""
   ii = ii + 1: Rem 17
   If acumulado.telefono <> 0 Then MSFlexGrid1.TextMatrix(li, 17) = Format(acumulado.telefono, z1$) Else MSFlexGrid1.TextMatrix(li, ii) = ""
   ii = ii + 1: Rem 18
   If acumulado.otraded <> 0 Then MSFlexGrid1.TextMatrix(li, 18) = Format(acumulado.otraded, z1$) Else MSFlexGrid1.TextMatrix(li, ii) = ""
   ii = ii + 1: Rem 19
   deducciones = acumulado.crdsal + acumulado.ispt + acumulado.imss + acumulado.prestamos + acumulado.fonacot + acumulado.telefono + acumulado.otraded
   If deducciones <> 0 Then MSFlexGrid1.TextMatrix(li, 19) = Format(deducciones, z1$) Else MSFlexGrid1.TextMatrix(li, ii) = ""
   neto = ingresos - deducciones
   ii = ii + 1: Rem 20
   If neto <> 0 Then MSFlexGrid1.TextMatrix(li, 20) = Format(neto, z1$) Else MSFlexGrid1.TextMatrix(li, ii) = ""
   MSFlexGrid1.TextMatrix(li, 21) = 0
   MSFlexGrid1.TextMatrix(li, 22) = 0
   MSFlexGrid1.TextMatrix(li, 23) = 0

End Sub
Private Sub Acueleg_Click(Index As Integer)
    CommonDialog1.CancelError = True
     On Error GoTo ErrHandler
     CommonDialog1.filename = "*.NOM"
     CommonDialog1.ShowOpen
     Archivo_acu = CommonDialog1.FileTitle

Exit Sub
    
ErrHandler:
    ' El usuario ha hecho clic en el botón Cancelar
    Exit Sub

End Sub

Private Sub Acuordalf_Click(Index As Integer)
    colanti = MSFlexGrid1.Col
    renati = MSFlexGrid1.Row
    MSFlexGrid1.Row = 1
    MSFlexGrid1.Col = 1
    MSFlexGrid1.RowSel = limite
    MSFlexGrid1.Sort = 1
    MSFlexGrid1.Col = colanti
    MSFlexGrid1.Row = renati
    MSFlexGrid1.SetFocus

End Sub

Private Sub AcuorEli_Click(Index As Integer)
   eliminacion
End Sub

Private Sub AcuorNum_Click(Index As Integer)
    colanti = MSFlexGrid1.Col
    renati = MSFlexGrid1.Row
    MSFlexGrid1.Row = 1
    MSFlexGrid1.Col = 0
    MSFlexGrid1.RowSel = limite
    MSFlexGrid1.Sort = 3
    MSFlexGrid1.Col = colanti
    MSFlexGrid1.Row = renati
    MSFlexGrid1.SetFocus

End Sub

Private Sub AcuVerVer_Click(Index As Integer)
     nq_2 = 0
     inter_clave = 1
     Load Form10
     Form10.Show
     
End Sub

Private Sub Form_Load()
     nq_1 = 0: nq_2 = 0: nq_3 = 0
     Archivo_acu = ""
     meselegido = 1
On Error GoTo Manejo9
    Close 3
     Open "personal.nom" For Random As 3 Len = Len(personal)
     nq_3 = LOF(3) / Len(personal)
     Get #3, 1, personal
     If nq_3 > 0 Then
      MSFlexGrid1.Cols = 24
      MSFlexGrid1.Rows = nq_3 + 2
      iniciar
      cargaper
     Else
       Close 3
       MsgBox "No existe archivo de personal no es posible capturar la nomina"
    End If
    GoTo sale9
Manejo9:
      Acu_sal_Click 1
sale9:

End Sub

Private Sub MSFlexGrid1_EnterCell()
    If MSFlexGrid1.Col > 1 And MSFlexGrid1.Row > 0 Then
        MSFlexGrid1.CellBackColor = &H80FF80
    End If
End Sub


Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
      Select Case KeyCode
            Case vbKeyDelete
                MSFlexGrid1.Text = ""
                
            Case vbKeyF2
                Text2.Text = MSFlexGrid1.Text
                Text2.SetFocus
       End Select
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
         MSFlexGrid1.SelectionMode = flexSelectionFree
         If MSFlexGrid1.Text <> "" Then valcelant = MSFlexGrid1.Text
         Text2.Text = Chr$(KeyAscii)
         Text2.SetFocus
End Sub

Private Sub MSFlexGrid1_LeaveCell()
   If MSFlexGrid1.Col > 1 And MSFlexGrid1.Row > 0 Then
         MSFlexGrid1.CellBackColor = vbWhite
   End If
End Sub

Private Sub MSFlexGrid1_RowColChange()
       Rem Text2.Text = MSFlexGrid1.Text
End Sub

