VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Poliza1 
   BackColor       =   &H80000014&
   BorderStyle     =   0  'None
   Caption         =   "Polizas de Diario"
   ClientHeight    =   7875
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   11115
   Icon            =   "Poliza1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   11115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2400
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      Begin MSComDlg.CommonDialog FijarDir1 
         Left            =   3720
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5640
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   720
      Width           =   5055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10080
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid Pol2 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   11245
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      BackColor       =   -2147483644
      BackColorSel    =   -2147483639
      BackColorBkg    =   -2147483644
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   5535
   End
   Begin VB.Menu ArPol 
      Caption         =   "&Archivo"
      Begin VB.Menu ArchCamb 
         Caption         =   "&Cambio Directorio"
      End
      Begin VB.Menu ArPolAbre 
         Caption         =   "&Abrir"
      End
      Begin VB.Menu Polsep1 
         Caption         =   "-"
      End
      Begin VB.Menu ArPolSale 
         Caption         =   "&Salida"
      End
   End
   Begin VB.Menu PolEdi 
      Caption         =   "&Edición"
      Begin VB.Menu PolEdiCopiar 
         Caption         =   "&Copiar"
      End
   End
   Begin VB.Menu PolImpr 
      Caption         =   "&Imprimir"
      Begin VB.Menu PolRango 
         Caption         =   "&Rango"
      End
      Begin VB.Menu sepimp 
         Caption         =   "-"
      End
      Begin VB.Menu PolTdas 
         Caption         =   "&Todas"
      End
   End
End
Attribute VB_Name = "Poliza1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim num_pol As Integer, des_pol
Dim TopeBorrego As Long, z2 As String, letrero1 As String, Fe_dia As Integer
Sub Carga_nomina()
     Rem CommonDialog1.CancelError = True
     Rem ON ERROR GoTo ErrHandler
     Rem CommonDialog1.FileName = "CATAUX*.*"
     Rem CommonDialog1.ShowOpen
     Rem Direc_torio = Mid(CommonDialog1.FileName, 1, Len(CommonDialog1.FileName) - Len(CommonDialog1.FileTitle))
     Close
     Open "C:\Archivos de programa\NOMINA1\perma.dno" For Random As #7 Len = Len(basico)
     fin_basico = LOF(7) / Len(basico)
     Get 7, 2, basico
     Direc_torio = Trim(basico.datoarch)
     Open Direc_torio + "DATOS" For Random As 10 Len = Len(Datos)
     FCM = LOF(10) / Len(Datos)
     Get 10, 1, Datos
     If ReferOper.Mes < 10 Then
        ArcOper = Trim(Datos.No_arch) + "0" + Trim(Str(ReferOper.Mes))
        Else
        ArcOper = Trim(Datos.No_arch) + Trim(Str(ReferOper.Mes))
     End If
     Open Direc_torio + "\" + ArcOper For Random As 11 Len = Len(oper)
     Dm = LOF(11) / Len(oper)
     
     Rem *******************************************************
     RECPOL
     
    Exit Sub
Errhandler:
    ' El usuario ha hecho clic en el botón Cancelar
    Exit Sub
End Sub
Sub RECPOL()
Dim L As Long, W As Long, Rren As Long, Adicion, Obron As Integer
Dim UB As Integer, NomUb As String, Monto As Currency, SdoCta As Currency
   
   Pol2.Rows = 1
   Rren = Rren + 1: UB = 3
   Adicion = "1104" & Chr(9) & "" & Chr(9) & _
                            "COSTO DIRECTO" & _
                            Chr(9) & "" & Chr(9) & "" & Chr(9) & _
                            "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" _
                            & Chr(9) & UB & Chr(9) & _
                            ""
                    Pol2.AddItem Adicion
                    Adicion = ""
    SdoCta = 0
   For L = 1 To (poliza.Apl1.Rows - 2)
        For W = 3 To (poliza.Apl1.Cols - 1)
              If IsNumeric(poliza.Apl1.TextMatrix(L, W)) Then
                Obron = poliza.Apl1.TextMatrix(0, W)
                Monto = poliza.Apl1.TextMatrix(L, W)
                Select Case Obron
                Case 9000
                    Rem ES GASTOS NO CUENTA
                Case Else
                    Rren = Rren + 1
                    
                    VerScta Obron, UB, NomUb
                    SdoCta = SdoCta + Monto
                    Adicion = "" & Chr(9) & poliza.Apl1.TextMatrix(0, W) & Chr(9) & _
                            NomUb & _
                            Chr(9) & poliza.Apl1.TextMatrix(L, W) & Chr(9) & "" & Chr(9) & _
                            "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" _
                            & Chr(9) & UB & Chr(9) & _
                            poliza.Apl1.TextMatrix(L, 1)
                Pol2.AddItem Adicion
                Adicion = ""
                End Select
             End If
        Next W
   Next L
   Pol2.TextMatrix(1, 4) = Format(SdoCta, z1)
   Rren = Rren + 1: UB = 22: SdoCta = SdoCta * -1
   Adicion = "5104" & Chr(9) & "" & Chr(9) & _
                            "CTRA. CTA. COSTO DIRECTO" & _
                            Chr(9) & "" & Chr(9) & "" & Chr(9) & _
                            Format(SdoCta, z1) & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" _
                            & Chr(9) & UB & Chr(9) & _
                            ""
                    Pol2.AddItem Adicion
                    Adicion = ""
   SdoCta = 0
End Sub
Sub VerScta(Obron, UB, NomUb)
    Dim FCATX As Long, WW As Long
    Close 12: UB = 0
    Open Direc_torio + "CATAUX" For Random As 12 Len = Len(CATAUX)
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
Sub titulopoliza()
     
    centrar ancho2, (Printer.ScaleWidth), RTrim(Datos.D1)
    Printer.CurrentX = ancho2
    Printer.Print RTrim(Datos.D1)
    letrero1 = "P  O   L  I  Z  A    D  E    D  I  A  R  I  O"
    centrar ancho2, (Printer.ScaleWidth), letrero1
    Printer.CurrentX = ancho2
    Printer.Print letrero1
    Printer.Print
    Printer.Print String(230, "-")
    Printer.Print
    Printer.Print Tab(3); "Cuenta-SubCta";
    Printer.Print Tab(23); "N  O  M   B   R  E";
    Printer.Print Tab(55); "PARCIAL";
    Printer.Print Tab(70); "D E B E";
    Printer.Print Tab(85); "H A B E R";
    Printer.Print Tab(100); "R e d a c c i o n   A u x i l i a r"
    Printer.Print
    Printer.Print String(230, "-")
    Printer.Print
End Sub



Private Sub Form_Unload(Cancel As Integer)
   CONS_ULTA = 0
   Close 2, 3, 4
   Unload Poliza1
End Sub

Private Sub PolEdiCopiar_Click()
  Clipboard.Clear
   Clipboard.SetText Clipboard.GetText + Poliza1.Caption & Chr(13)
   Clipboard.SetText Clipboard.GetText + Label1.Caption & Chr(13)
   Pol2.RowSel = Pol2.Rows - 1
   Pol2.ColSel = Pol2.Cols - 1
   For i = 0 To Pol2.RowSel
      For f = 0 To Pol2.ColSel
            Clipboard.SetText Clipboard.GetText + Pol2.TextMatrix(i, f) & Chr(9)
      Next f
      Clipboard.SetText Clipboard.GetText + Chr(13)
   Next i
   difer = Pol2.RowSel - Pol2.Row

End Sub

Private Sub PolRango_Click()
On Error GoTo Zacatecas1
If Pol2.TextMatrix(1, 0) <> "" Then
  ultimo.TipoCap = Printer.FontSize
  Printer.FontSize = 8
  For ImprPoliza = Rango_Inic To Rango_Final
     If ultimo.num = 0 Then
        Rem ultimo.num = LocPol.LPol1.TextMatrix(ImprPoliza, 2)
     End If
     verpoliza
    If ultimo.num > 0 Then
       ImpresionPoliza
       ultimo.num = 0
    End If
  Next ImprPoliza
Else
    MsgBox "No existe poliza que imprimir"
End If
Zacatecas1:
End Sub

Sub ImpresionPoliza()
       titulopoliza
        hojita = 1
    For C = 1 To Pol2.Rows - 2
        cuenta = cuenta + 1
        For r = 0 To Pol2.Cols - 1
        Select Case r
            Case 0
            If Pol2.TextMatrix(C, r) <> "" Then
                Printer.Print Tab(6);
                colocar ancho2, Pol2.TextMatrix(C, r), z2
                Printer.CurrentX = Printer.CurrentX + ancho2
                Printer.FontBold = True
                Printer.FontUnderline = True
                Printer.Print Format(Pol2.TextMatrix(C, r), z2); "  ";
                
            End If
            Case 1
            If Pol2.TextMatrix(C, r) <> "" Then
                Printer.Print Tab(10);
                colocar ancho2, Pol2.TextMatrix(C, r), z2
                Printer.CurrentX = Printer.CurrentX + ancho2
                Printer.Print Format(Pol2.TextMatrix(C, r), z2); "  ";
                
            End If
            Case 2
            If Pol2.TextMatrix(C, r) <> "" Then
                
                Printer.Print RTrim(Pol2.TextMatrix(C, r));
                Printer.FontBold = False
                Printer.FontUnderline = False
            End If
            Case 3
            If Pol2.TextMatrix(C, r) <> "" Then
                Printer.CurrentX = 4725
                colocar ancho2, Pol2.TextMatrix(C, r), z2
                Printer.CurrentX = Printer.CurrentX + ancho2
                Printer.Print Format(Pol2.TextMatrix(C, r), z1);
            End If
            Case 4
            If Pol2.TextMatrix(C, r) <> "" Then
                Printer.CurrentX = 6110
                colocar ancho2, Pol2.TextMatrix(C, r), z2
                Printer.CurrentX = Printer.CurrentX + ancho2
                Printer.Print Format(Pol2.TextMatrix(C, r), z1);
            End If
            Case 5
            If Pol2.TextMatrix(C, r) <> "" Then
                Printer.CurrentX = 7520
                colocar ancho2, Pol2.TextMatrix(C, r), z2
                Printer.CurrentX = Printer.CurrentX + ancho2
                Printer.Print Format(Pol2.TextMatrix(C, r), z1);
            End If
            Case 10
            If Pol2.TextMatrix(C, r) <> "" Then
                Printer.CurrentX = 8500
                Printer.Print RTrim(Pol2.TextMatrix(C, r))
            End If

             
          End Select
          Next r
          If cuenta > 52 Then
                hojita = hojita + 1
                Printer.Print Tab(50); "Continua hoja......."; hojita
                
                Printer.NewPage
                cuenta = 0
                titulopoliza
          End If

    Next C
    
    Pol2.Row = Pol2.Rows - 1
    
    Printer.CurrentY = 12000
        If Pol2.TextMatrix(Pol2.Row, 4) <> "" Then
                Printer.CurrentY = Printer.CurrentY + 50
                Printer.CurrentX = 5310
                antes = Printer.DrawWidth
                Printer.DrawWidth = 3
                Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX + 3000, Printer.CurrentY), , B
                Rem Printer.Print Tab(70); String(18, "-");
                Rem Printer.Print Tab(86); String(18, "-")
                Printer.Print Tab(35); "Sumas Iguales  ....";: Rem  Tab(75);
                Printer.CurrentX = 6110
                colocar ancho2, Pol2.TextMatrix(Pol2.Row, 4), z2
                Printer.CurrentX = Printer.CurrentX + ancho2
                Printer.Print Format(Pol2.TextMatrix(Pol2.Row, 4), z1);
        End If
        If Pol2.TextMatrix(Pol2.Row, 5) <> "" Then
                Printer.CurrentX = 7520
                colocar ancho2, Pol2.TextMatrix(Pol2.Row, 5), z2
                Printer.CurrentX = Printer.CurrentX + ancho2
                Printer.Print Format(Pol2.TextMatrix(Pol2.Row, 5), z1)
                Printer.CurrentY = Printer.CurrentY + 50
                Printer.CurrentX = 5310
                Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX + 3000, Printer.CurrentY), , B
                Printer.CurrentX = 5310
                Printer.Line (Printer.CurrentX, Printer.CurrentY + 20)-(Printer.CurrentX + 3000, Printer.CurrentY + 20), , B
                Printer.DrawWidth = antes
                Printer.Print
                Rem Printer.Print ; Tab(70); String(10, "="); Tab(86); String(10, "=")
        End If
        Printer.CurrentX = 1700
        Printer.Print "C O N C E P T O";
        Printer.CurrentX = 5250
        Printer.Print "F E C H A";
        Printer.CurrentX = 8000
        Printer.Print "POLIZA No"
        Printer.Print
        regresa = Printer.CurrentY
        Printer.Print
        ubica_pol = Printer.CurrentX + 800
        Printer.Line (ubica_pol, regresa + 50)-(ubica_pol + 9000, Printer.CurrentY + 300), , B
        Printer.CurrentY = regresa
        Printer.CurrentY = Printer.CurrentY + 200
        Printer.CurrentX = 900
        Printer.Print des_pol;
        Printer.CurrentX = 4700
        Printer.Print Fe_dia; " "; RTrim(mm(Mes_Act)); " de "; Datos.a_o;
        Printer.CurrentX = 8500
        Printer.Print num_pol;
        antes = Printer.DrawWidth
        Printer.DrawWidth = 3
        
        Rem Printer.Line (ubica_pol, Printer.CurrentY + 500)-(ubica_pol + 9000, Printer.CurrentY + 500), , B
        
        Printer.DrawWidth = antes
    Printer.EndDoc
    Printer.FontSize = ultimo.TipoCap

End Sub

Private Sub PolTdas_Click()

   If (Pol2.TextMatrix(1, 0) <> "") And (Pol2.TextMatrix(1, 0) <> "") Then
  
  ultimo.TipoCap = Printer.FontSize
  Printer.FontSize = 8
  
  For ImprPoliza = 1 To LocPol.LPol1.Rows - 1
  
     ultimo.num = LocPol.LPol1.TextMatrix(ImprPoliza, 2)
     verpoliza
    If ultimo.num > 0 Then
       ImpresionPoliza
    End If
  Next ImprPoliza
Else
    MsgBox "No existe poliza que imprimir"
End If
 
      
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
     Select Case Button.Key
       Case Is = "Cambiar"
            ArPolAbre_Click
       Case Is = "Imprimir"
            PolRango_Click
       Case Is = "CatCta"
            LocPol.Show
            Rem Combo1.Visible = True
            Rem Label1.Visible = False
     End Select

End Sub

Sub barra_htas()

    Dim imgX As ListImage
    Ruta_Acceso = "\2010\PROGVB\NOMINA2"
    Toolbar1.Style = tbrFlat
    Toolbar1.Align = 4
    ' Carga imágenes en el control ImageList.
    Set imgX = ImageList1.ListImages. _
    Add(, "Cambiar", LoadPicture(Ruta_Acceso + "\open.bmp"))
        Set imgX = ImageList1.ListImages. _
    Add(, "Polizas", LoadPicture(Ruta_Acceso + "\note12.ico"))
        Set imgX = ImageList1.ListImages. _
    Add(, "Imprimir", LoadPicture(Ruta_Acceso + "\Print.bmp"))
        Set imgX = ImageList1.ListImages. _
    Add(, "CatCta", LoadPicture(Ruta_Acceso + "\CatCta.bmp"))
        Toolbar1.ImageList = ImageList1
    Dim btnX As Button
    Toolbar1.Buttons.Add , , , tbrPlaceholder
    Set btnX = Toolbar1.Buttons.Add(, "Cambiar", , tbrDefault, "Cambiar")
    btnX.ToolTipText = "Abrir archivo de operaciones"
        Set btnX = Toolbar1.Buttons.Add(, "Imprimir", , tbrDefault, "Imprimir")
    btnX.ToolTipText = "Enviar poliza a impresora"
        btnX.Description = btnX.ToolTipText
        Set btnX = Toolbar1.Buttons.Add(, "CatCta", , tbrDefault, "CatCta")
    btnX.ToolTipText = "Volver a la lista de polizas en el archivo"
      
End Sub
Private Sub ArchCamb_Click()
  Close #1
    
    mientras = ""
Rem On Err GoTo Errhandler
    MIDIR = CurDir
    MIDIR = RTrim(MIDIR)
    If Right(MIDIR, 1) = "\" Then
        MIDIR = Left(MIDIR, Len(MIDIR) - 1)
    End If
    FijarDir1.InitDir = MIDIR
    FijarDir1.Flags = cdlOFNHideReadOnly
    FijarDir1.Filter = "Archivos de Contabilidad(Dat*.*)|Dat*.*"
    FijarDir1.ShowOpen
    If FijarDir1.FileName <> "" Then
        For i = 1 To Len(FijarDir1.FileName)
                If Mid(FijarDir1.FileName, i, 1) = "\" Then tope = i
        Next i
        mientras = Mid(FijarDir1.FileName, 1, tope)
        ChDir mientras
        Close 3
        Open Ruta_Acceso_Contr + "\Gcont.Arr" For Random As 3 Len = Len(SCont)
        SCont.guarda = mientras
        Put 3, 1, SCont
        Close 3
        cm = 0
        inicio
        Poliza1.Caption = RTrim(Datos.D1) + " " + Datos.a_o
    End If
Errhandler:
End Sub

Private Sub ArPolAbre_Click()
On Error GoTo Errhandler
    CommonDialog1.CancelError = True
    MIDIR = CurDir
    MIDIR = RTrim(MIDIR)
    If Right(MIDIR, 1) = "\" Then
        MIDIR = Left(MIDIR, Len(MIDIR) - 1)
    End If
    CommonDialog1.InitDir = MIDIR
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNReadOnly
    Datos.No_arch = RTrim(Datos.No_arch)
    CommonDialog1.FileName = MIDIR + "\" + RTrim(Datos.No_arch) + "*.*"
    
    anuncio = ""
    Arch_act = ""
    anuncio = "Archivos de Operaciones (" + RTrim(Datos.No_arch) + "*.*)|" + RTrim(Datos.No_arch) + "*.*|Todos los archivos (*.*)|*.*"
    CommonDialog1.Filter = anuncio
    CommonDialog1.ShowOpen
    donde = Len(RTrim(Datos.No_arch)) + 1
    Arch_act = RTrim(CommonDialog1.FileTitle)
    Mes_Act = 0
    For i = 1 To Len(Arch_act)
       If (Mid(Arch_act, i, 1) >= Chr(48)) And (Mid(Arch_act, i, 1)) <= Chr(57) Then
           Mes_Act = Val(Right(RTrim(Arch_act), 2))
            If (Mid(Arch_act, i + 2, 1)) = "." Then
                 MsgBox "Archivo no Valido ", vbCritical
                 Exit Sub
            End If
            Exit For
       End If
    Next i
    recorrer
    ultimo.num = 0
    GoTo SaleBien
Errhandler:
  Arch_act = ""
  Mes_Act = 0
  Combo1.Visible = False
  Poliza1.Caption = RTrim(Datos.D1) + " " + Datos.a_o + " No se Eligio Archivo"
  Label1.Visible = False
  Combo1.Visible = False
  Pol2.Clear
  Exit Sub
SaleBien:
  Pol2.Clear
  Poliza1.Caption = RTrim(Datos.D1) + " Polizas de " + RTrim(mm(Mes_Act)) + " " + Datos.a_o
End Sub
Sub verpoliza()
   Dim Mvtodebe As Currency, MvtoHaber As Currency, Mientrassalgo As Integer
   If ultimo1.num > 0 Then Mientrassalgo = ultimo.num: ultimo.num = ultimo1.num: ultimo1.num = Mientrassalgo
    If ultimo.num > 0 Then
        Pol2.Clear
        titulocolumnas
        Pol2.Rows = 1
        Pol2.Row = 0
        Combo1.Visible = False
        Label1.Visible = True
        Label1.BackColor = vbYellow
        Close 3
        Open Arch_act For Random As 3 Len = Len(oper)
        cm = LOF(3) / Len(oper)
        Get 3, ultimo.num, oper
        Label1.Caption = "Poliza: " + oper.CTA + "  " + oper.descr
        num_pol = oper.CTA: Fe_dia = Val(oper.fe): Rem mm(Mes_Act)
        des_pol = oper.descr
        Open "Catmay" For Random As 2 Len = Len(CATMAY)
        Open "Cataux" For Random As 4 Len = Len(CATAUX)
        Mvtodebe = 0: MvtoHaber = 0
        For r = (ultimo.num + 1) To cm
                Get 3, r, oper
                If oper.identi = "B" Then
                        Get 2, Val(oper.real), CATMAY
                        If Val(oper.impte) > 0 Then
                            Pol2.AddItem Format(oper.CTA, "###0") & Chr(9) & "" & Chr(9) & _
                                     CATMAY.B2 & Chr(9) & "" & Chr(9) & Format(oper.impte, "###,###,##0.00") _
                                     & Chr(9) & "" & Chr(9) & "" & Chr(9) & oper.descr
                            Mvtodebe = Mvtodebe + oper.impte
                            Else
                            Pol2.AddItem Format(oper.CTA, "###0") & Chr(9) & "" & Chr(9) & _
                                     CATMAY.B2 & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format(oper.impte, "###,###,##0.00") _
                                     & Chr(9) & "" & Chr(9) & "" & Chr(9) & oper.descr
                            MvtoHaber = MvtoHaber + oper.impte
                        End If
                                     
                End If
                If oper.identi = "C" Then
                        Get 4, Val(oper.CTA), CATAUX
                          
                            Pol2.AddItem "" & Chr(9) & Format(CATAUX.C1, "###0") & Chr(9) & _
                                     (" " + CATAUX.C2) & Chr(9) & Format(oper.impte, "###,###,##0.00") _
                                     & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & _
                                     "" & Chr(9) & "" & Chr(9) & oper.descr
                End If
                If oper.identi = "A" Then Exit For
        Next r
        Close 2, 4
        Pol2.AddItem "" & Chr(9) & "" & Chr(9) & _
                                     "Sumas Iguales" & Chr(9) & "" & Chr(9) & Format(Mvtodebe, "###,###,##0.00") & Chr(9) & Format(MvtoHaber, "###,###,##0.00") _
                                     & Chr(9) & "" & Chr(9) & "" & Chr(9) & ""
                
        
        Pol2.Height = (Pol2.Rows + 4) * (Pol2.CellHeight)
        If Pol2.Height > TopeBorrego Then Pol2.Height = TopeBorrego
            Form_Resize
        End If

       Close 3
End Sub
Sub recorrer()
  If ultimo.num > 0 Then Unload LocPol
  Load LocPol
  LocPol.Show
  
   'Close 3
   'Open Arch_act For Random As 3 Len = Len(oper)
   'cm = LOF(3) / Len(oper)
   'Combo1.Visible = True
   'Combo1.Clear
   'For R = 1 To cm: Get 3, R, oper
        'If oper.identi = "A" Then
             'numero = Val(oper.cta)
             'num1 = Format(numero, "#####0")
             
             'num1 = String(6 - Len(num1), " ") + num1
             
             'muestra = num1 + " " + oper.descr + Str(R)
            'Combo1.AddItem muestra
        'End If
        
   'Next R
   'Combo1.Text = Combo1.List(0)
End Sub
Sub inicio()
On Error GoTo saltalo
    Open (Ruta_Acceso_Contr + "\Gcont.Arr") For Random As 3 Len = Len(SCont)
    Get 3, 1, SCont
    If SCont.guarda <= " " Then
        ChDir Ruta_Acceso_Contr
        Else
        
        If Left(SCont.guarda, 1) <> "C" Then
                    ChDrive Left(SCont.guarda, 1)
        End If
        ChDir SCont.guarda
    End If
saltalo:
    Close 3
     
    Close #1
    Archivo = "DATOS"
    Open Archivo For Random As #1 Len = Len(Datos)
    cm = LOF(1) / Len(Datos)
    Get 1, 1, Datos
End Sub

Private Sub ArPolSale_Click()
   CONS_ULTA = 0
   
   Close 2, 3, 4
   Unload Poliza1
End Sub

Private Sub Combo1_dblClick()
    Combo1_KeyPress 13
End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        Dim Mvtodebe As Currency, MvtoHaber As Currency
        ultimo.num = Val(Mid(Combo1.List(Combo1.ListIndex), 39))
        If ultimo.num > 0 Then
        Pol2.Clear
        titulocolumnas
        Pol2.Rows = 1
        Pol2.Row = 0
        Combo1.Visible = False
        Label1.Visible = True
        Label1.BackColor = vbYellow
        Get 3, ultimo.num, oper
        Label1.Caption = "Poliza: " + oper.CTA + "  " + oper.descr
        num_pol = oper.CTA: Fe_dia = oper.fe: Rem mm(Mes_Act)
        des_pol = oper.descr
        Rem Open "Catmay" For Random As 2 Len = Len(CATMAY)
        Open "Cataux" For Random As 4 Len = Len(CATAUX)
        Mvtodebe = 0: MvtoHaber = 0
        For r = (ultimo.num + 1) To cm
                Get 3, r, oper
                If oper.identi = "B" Then
                        Get 1, Val(oper.real), CATMAY
                        If Val(oper.impte) > 0 Then
                            Pol2.AddItem Format(oper.CTA, "###0") & Chr(9) & "" & Chr(9) & _
                                     CATMAY.B2 & Chr(9) & "" & Chr(9) & Format(oper.impte, "###,###,##0.00") _
                                     & Chr(9) & "" & Chr(9) & "" & Chr(9) & oper.descr
                            Mvtodebe = Mvtodebe + oper.impte
                            Else
                            Pol2.AddItem Format(oper.CTA, "###0") & Chr(9) & "" & Chr(9) & _
                                     CATMAY.B2 & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format(oper.impte, "###,###,##0.00") _
                                     & Chr(9) & "" & Chr(9) & "" & Chr(9) & oper.descr
                            MvtoHaber = MvtoHaber + oper.impte
                        End If
                                     
                End If
                If oper.identi = "C" Then
                        Get 4, Val(oper.CTA), CATAUX
                          
                            Pol2.AddItem "" & Chr(9) & Format(CATAUX.C1, "###0") & Chr(9) & _
                                     (" " + CATAUX.C2) & Chr(9) & Format(oper.impte, "###,###,##0.00") _
                                     & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & _
                                     "" & Chr(9) & "" & Chr(9) & oper.descr
                End If
                If oper.identi = "A" Then Exit For
        Next r
        Close 4
        Pol2.AddItem "" & Chr(9) & "" & Chr(9) & _
                                     "Sumas Iguales" & Chr(9) & "" & Chr(9) & Format(Mvtodebe, "###,###,##0.00") & Chr(9) & Format(MvtoHaber, "###,###,##0.00") _
                                     & Chr(9) & "" & Chr(9) & "" & Chr(9) & ""
                
        
        Pol2.Height = (Pol2.Rows + 4) * (Pol2.CellHeight)
        If Pol2.Height > TopeBorrego Then Pol2.Height = TopeBorrego
        Form_Resize
        End If
   End If
End Sub
Sub titulocolumnas()
    Pol2.Row = 0
    Pol2.Col = 0: Pol2.ColWidth(0) = 600:  Pol2.CellAlignment = 4: Pol2.Text = "Cuenta"
    Pol2.Col = 1: Pol2.ColWidth(1) = 600:  Pol2.CellAlignment = 4: Pol2.Text = "SubCta"
    Pol2.Col = 2: Pol2.ColWidth(2) = 3350: Pol2.CellFontBold = True: Pol2.CellAlignment = 4: Pol2.Text = "Nombre"
    Pol2.Col = 3: Pol2.ColWidth(3) = 1200: Pol2.CellFontBold = True: Pol2.CellAlignment = 4: Pol2.Text = "Parcial"
    Pol2.Col = 4: Pol2.ColWidth(4) = 1100: Pol2.CellFontBold = True: Pol2.CellAlignment = 4: Pol2.Text = "Debe"
    Pol2.Col = 5: Pol2.ColWidth(5) = 1100: Pol2.CellFontBold = True: Pol2.CellAlignment = 4: Pol2.Text = "Haber"
    Pol2.Col = 6: Pol2.ColWidth(6) = 280
    Pol2.Col = 7: Pol2.ColWidth(7) = 280
    Pol2.Col = 8: Pol2.ColWidth(8) = 280
    Pol2.Col = 9: Pol2.ColWidth(9) = 680
    Pol2.Col = 10: Pol2.ColWidth(10) = 3200: Pol2.CellFontBold = True: Pol2.CellAlignment = 3: Pol2.Text = "Redaccion"
   
End Sub
Private Sub Form_Load()
    barra_htas
    Rem inicio
    Combo1.Visible = False
    Rem Poliza1.Caption = RTrim(Datos.D1) + " " + Datos.a_o
    z1 = "##,###,##0.00": z2 = "####0"
    Pol2.Cols = 11
    Label1.Visible = False
    Pol2.Clear
    titulocolumnas
    Pol2.Row = 0: Pol2.Col = 10
    Pol2.ColWidth(10) = 3200: Pol2.CellFontBold = True
    Pol2.CellAlignment = 3: Pol2.Text = "REDACCION"
    For i = 0 To 10
        ancho = ancho + Pol2.ColWidth(i)
    Next i
    Poliza1.Width = ancho + 800
    TopeBorrego = Pol2.Height
    Pol2.Width = ancho + 57.5
    Carga_nomina
    Rem Pol2.ColIsVisible(5) = False
    Rem Pol2.Rows = 20
End Sub

Private Sub Form_Resize()
     Poliza1.Height = 1600 + Pol2.Height
     Poliza1.Width = Pol2.Width * 1.05
     Poliza1.Refresh
End Sub

Private Sub Pol2_KeyPress(KeyAscii As Integer)
   Rem On Error GoTo SalePoliza
     If KeyAscii = 27 Then
         Rem Combo1.Visible = True
         Rem Label1.Visible = False
         Clipboard.Clear
         If CONS_ULTA = 1 Then LocPol.Show Else ArPolSale_Click
     End If
SalePoliza:
End Sub

