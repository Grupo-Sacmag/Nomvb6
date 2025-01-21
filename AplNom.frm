VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form AplNom 
   Caption         =   "Poliza de aplicacion nomina"
   ClientHeight    =   6195
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12315
   Icon            =   "AplNom.frx":0000
   LinkTopic       =   "Form9"
   ScaleHeight     =   6195
   ScaleWidth      =   12315
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid NomPol 
      Height          =   4815
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8493
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      BackColorBkg    =   -2147483633
      Appearance      =   0
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   5655
   End
   Begin VB.Menu Ar 
      Caption         =   "&Archivo"
      Begin VB.Menu ArGr 
         Caption         =   "&Guardar"
      End
      Begin VB.Menu ArSep1 
         Caption         =   "-"
      End
      Begin VB.Menu ArImpr 
         Caption         =   "&Imprimir"
      End
   End
   Begin VB.Menu Edt 
      Caption         =   "&Edicion"
      Begin VB.Menu EdSelT 
         Caption         =   "&Seleccionar Todo"
      End
      Begin VB.Menu EdSep1 
         Caption         =   "-"
      End
      Begin VB.Menu EdCop 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "AplNom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim F_Aum As Long, MesNom As Integer, Direc_t As String
Private Sub EdCop_Click()
 Dim Temporal1
 Clipboard.Clear
   difer = NomPol.RowSel - NomPol.Row
   For i = NomPol.Row To NomPol.RowSel
      For f = NomPol.Col To NomPol.ColSel
            Temporal1 = Temporal1 + NomPol.TextMatrix(i, f)
            If f < NomPol.ColSel Then
                Temporal1 = Temporal1 & Chr(9)
            End If
      Next f
      Temporal1 = Temporal1 & Chr(13) & Chr(10)
   Next i
    Clipboard.SetText Temporal1
    
    
End Sub

Private Sub EdSelT_Click()
   Clipboard.Clear
    NomPol.Col = 0: NomPol.Row = 0
    NomPol.RowSel = NomPol.Rows - 1
    NomPol.ColSel = NomPol.Cols - 1
    
End Sub

Private Sub Form_Load()
   If (Len(Form8.Label7.Caption) > 30) Then
            Label1.Caption = UCase(Mid(Form8.Label7.Caption, 13))
            Else
            Label1.Caption = UCase(Form8.Label7.Caption)
   End If
   NomPol.Clear
   ColDefine
    z1 = "##,###,##0.00": z2 = "####0"
    Carga_nomina
    
    
    Rem NomPol.Row = 0: NomPol.Col = 10
    Rem NomPol.ColWidth(10) = 3200: NomPol.CellFontBold = True
    
End Sub
Sub ColDefine()
    NomPol.Row = 0
    NomPol.Col = 0: NomPol.ColWidth(0) = 600:  NomPol.CellAlignment = 4: NomPol.Text = "Cuenta"
    NomPol.Col = 1: NomPol.ColWidth(1) = 600:  NomPol.CellAlignment = 4: NomPol.Text = "SubCta"
    NomPol.Col = 2: NomPol.ColWidth(2) = 3350: NomPol.CellFontBold = True: NomPol.CellAlignment = 4: NomPol.Text = "Nombre"
    NomPol.Col = 3: NomPol.ColWidth(3) = 1200: NomPol.CellFontBold = True: NomPol.CellAlignment = 4: NomPol.Text = "Parcial"
    NomPol.Col = 4: NomPol.ColWidth(4) = 1100: NomPol.CellFontBold = True: NomPol.CellAlignment = 4: NomPol.Text = "Debe"
    NomPol.Col = 5: NomPol.ColWidth(5) = 1100: NomPol.CellFontBold = True: NomPol.CellAlignment = 4: NomPol.Text = "Haber"
    NomPol.Col = 6: NomPol.ColWidth(6) = 3200: NomPol.CellFontBold = True: NomPol.CellAlignment = 3: NomPol.Text = "Redaccion"
    NomPol.Col = 7: NomPol.ColWidth(7) = 480
       
End Sub
Sub Carga_nomina()
     Close
     Open "C:\GconTA\perma.dno" For Random As #7 Len = Len(basico)
     fin_basico = LOF(7) / Len(basico)
     Get 7, fin_basico, basico
     Direc_torio = Trim(basico.datoarch)
     Open Direc_torio + "DATOS" For Random As 10 Len = Len(Datos)
     FCM = LOF(10) / Len(Datos)
     Get 10, 1, Datos
     Rem If Direc_torio < "" Then
     Direc_t = InputBox("El directorio de costos es " + Trim(Direc_torio) + " es correcto ??? ", _
               "Aplicacion nomina", Direc_torio)
     Direc_torio = Trim(Direc_t)

     Rem End If
     MesNom = InputBox("Esta poliza se va a registrar en " + Trim(mm(ReferOper.Mes)) + " es correcto ??? ", _
               "Aplicacion nomina", ReferOper.Mes)
     ReferOper.Mes = MesNom
     
     If ReferOper.Mes < 10 Then
        ArcOper = "COS" + "0" + Trim(Str(ReferOper.Mes))
        Else
        ArcOper = "COS" + Trim(Str(ReferOper.Mes))
     End If
     
     Open Direc_torio + "\" + ArcOper For Random As 11 Len = Len(oper)
     Dm = LOF(11) / Len(oper)
     
     Rem *******************************************************
     RECPOL
    Exit Sub
ErrHandler:

    Exit Sub
End Sub
Sub RECPOL()
Dim L As Long, W As Long, Rren As Long, Adicion, Obron As Integer
Dim UB As Integer, NomUb As String, Monto As Currency, SdoCta As Currency
   
   NomPol.Rows = 1
   Rren = Rren + 1: UB = 0: Obron = 1104
   VerCcta Obron, UB, NomUb
   Adicion = Obron & Chr(9) & "" & Chr(9) & _
                            Trim(NomUb) & _
                            Chr(9) & "" & Chr(9) & "" & Chr(9) & _
                            "" & Chr(9) & "" & Chr(9) & UB
                    NomPol.AddItem Adicion
                    Adicion = ""
    SdoCta = 0
   For L = 1 To (Poliza.Apl1.Rows - 2)
        For W = 3 To (Poliza.Apl1.Cols - 1)
              If IsNumeric(Poliza.Apl1.TextMatrix(L, W)) Then
                Obron = Poliza.Apl1.TextMatrix(0, W)
                Monto = Poliza.Apl1.TextMatrix(L, W)
                Select Case Obron
                Case 9000
                    Rem ES GASTOS NO CUENTA
                Case Else
                    Rren = Rren + 1
                    VerScta Obron, UB, NomUb
                    SdoCta = SdoCta + Monto
                    Adicion = "" & Chr(9) & Poliza.Apl1.TextMatrix(0, W) & Chr(9) & _
                            NomUb & _
                            Chr(9) & Poliza.Apl1.TextMatrix(L, W) & Chr(9) & "" & Chr(9) & _
                            "" & Chr(9) & Poliza.Apl1.TextMatrix(L, 1) & Chr(9) & UB
                NomPol.AddItem Adicion
                Adicion = ""
                End Select
             End If
        Next W
   Next L
   NomPol.TextMatrix(1, 4) = Format(SdoCta, z1)
   Rren = Rren + 1: SdoCta = SdoCta * -1
   Rren = Rren + 1: UB = 0: Obron = 5104
   VerCcta Obron, UB, NomUb

   Adicion = "5104" & Chr(9) & "" & Chr(9) & _
                            NomUb & _
                            Chr(9) & "" & Chr(9) & "" & Chr(9) & _
                            Format(SdoCta, z1) & Chr(9) & "" & Chr(9) & UB
                    NomPol.AddItem Adicion
                    Adicion = ""
   SdoCta = 0
End Sub
Sub VerScta(Obron, UB, NomUb)
    Dim FCATX As Long, WW As Long
    Close 12: UB = 0
    Open Direc_torio + "\CATAUX" For Random As 12 Len = Len(CATAUX)
     FCATX = LOF(12) / Len(CATAUX)
     
     For WW = 411 To 810: Get 12, WW, CATAUX
          If Val(CATAUX.C1) = Obron Then
             UB = WW: NomUb = Trim(CATAUX.C2)
             Exit For
          End If
     Next WW
     If UB = 0 Then
         NomUb = "NOMBRE NO LOCALIZADO"
     End If
End Sub
Sub VerCcta(Obron, UB, NomUb)
    Dim FCATM As Long, WW1 As Long
    Close 12: UB = 0
    Open Direc_torio + "\CATMAY" For Random As 13 Len = Len(CATMAY)
     FCATM = LOF(13) / Len(CATMAY)
     
     For WW1 = 1 To FCATM: Get 13, WW1, CATMAY
          If Val(CATMAY.B1) = Obron Then
             UB = WW1: NomUb = Trim(CATMAY.B2)
             Exit For
          End If
     Next WW1
     If UB = 0 Then
         NomUb = "CUENTA NO LOCALIZADA"
     End If
     Close 13
End Sub


Private Sub Form_Resize()
If AplNom.WindowState <> 1 Then
      NomPol.Height = ScaleHeight - 1200
      NomPol.Width = ScaleWidth - 400
      F_Aum = (NomPol.Width - 400) / 9200
      ColDfn
   End If
End Sub
Sub ColDfn()
    NomPol.FontWidth = 3 * F_Aum
    NomPol.ColWidth(0) = 600 * F_Aum
    NomPol.ColWidth(1) = 600 * F_Aum
    NomPol.ColWidth(2) = 3350 * F_Aum
    NomPol.ColWidth(3) = 1200 * F_Aum
    NomPol.ColWidth(4) = 1100 * F_Aum
    NomPol.ColWidth(5) = 1100 * F_Aum
    NomPol.ColWidth(6) = 3200 * F_Aum
    NomPol.ColWidth(7) = 480 * F_Aum
    
    
End Sub
