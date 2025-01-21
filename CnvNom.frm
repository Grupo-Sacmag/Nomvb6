VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form CnvNom 
   Caption         =   "Conversion Nomina"
   ClientHeight    =   4980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   0
      Width           =   6015
   End
   Begin MSFlexGridLib.MSFlexGrid CxNom 
      Height          =   4215
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   8080
      _ExtentX        =   14261
      _ExtentY        =   7435
      _Version        =   393216
      Cols            =   5
      FixedCols       =   2
      FocusRect       =   2
   End
   Begin MSComDlg.CommonDialog DialogoAbrir 
      Left            =   3960
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu Ar 
      Caption         =   "&Archivo"
      Begin VB.Menu ArAbr 
         Caption         =   "&Abrir"
      End
      Begin VB.Menu ArSep1 
         Caption         =   "-"
      End
      Begin VB.Menu ArTrNom 
         Caption         =   "&Transferir Nomina"
      End
      Begin VB.Menu ArcSep2 
         Caption         =   "-"
      End
      Begin VB.Menu ArcSalida 
         Caption         =   "&Salida"
      End
   End
End
Attribute VB_Name = "CnvNom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Directorio As String
Private Sub ArAbr_Click()
     Dim Arch As String, g As Integer
     
  On Err GoTo ErrHandler
            DialogoAbrir.Flags = cdlOFNHideReadOnly
            If Option1 = True Then
                    DialogoAbrir.FileName = "*.*"
                    DialogoAbrir.Filter = "Archivos de captura (*.GNO)|*.GNO"
                    Else
                    DialogoAbrir.FileName = "*.GNO"
                    DialogoAbrir.Filter = "Archivos de catalogos(*.GNO)|*.GNO"
            End If
            DialogoAbrir.ShowOpen
            If DialogoAbrir.FileName <> "" Then
               'CARGA
               Arch = Trim(DialogoAbrir.FileName)
               For r = 1 To Len(Arch)
                   If (Mid(Arch, r, 1) = "\") Or (Mid(Arch, r, 1) = "/") Then
                      g = r
                   End If
               
               Next r
               Directorio = Left(Arch, g)
               carga
               Exit Sub
                    Else
                    nombrearchivo = ""
                    MsgBox "No se elegio archivo ", vbExclamation
            End If
  
ErrHandler:
   Close
   Exit Sub

 
End Sub
Sub carga()
      Dim fecha1 As Date, Noma

      For r = 1 To Dm: Get 2, r, personal
           If Trim(personal.fab) <> "" Then
              fecha1 = personal.fab
              If fecha1 < Date Then GoTo Sale
              
              Else
              If (Trim(personal.ape1) <> "") Then
                    Noma = Format(r, "####0") & Chr(9) & Trim(personal.ape1) + " " + Trim(personal.ape2) _
                        + " " + Trim(personal.nom)
                    CxNom.AddItem Noma
              End If
           End If
Sale:
      
      Next r
      ColocarDias
End Sub
Sub ColocarDias()
    Dim DiasT As Integer, Salario As Currency
    Close 7
    Open DialogoAbrir.FileName For Random As 7 Len = Len(NxF)
    Wm = LOF(7) / Len(NxF)
    For r = 1 To Wm: Get 7, r, NxF
           For g = 1 To CxNom.Rows - 1
               DiasT = 0
               
               If CxNom.TextMatrix(g, 0) = NxF.Empleado Then
                  Get 2, NxF.Empleado, personal
                  DiasT = Val(CxNom.TextMatrix(g, 2)) + NxF.DiasTrab
                  
                  If DiasT > 15 Then
                     CxNom.Col = 2: CxNom.Row = g
                     CxNom.CellBackColor = vbRed
                     DiasT = 15
                  End If
                  CxNom.TextMatrix(g, 2) = Format(DiasT, "####0")
                  CxNom.TextMatrix(g, 3) = Format((personal.ingr * DiasT), "###,##0.00")
                  Salario = (personal.ingr * DiasT)
                  Exit For
               End If
           Next g
    Next r
End Sub

Private Sub ArDirImp_Click()
 Dir_imptos = InputBox("Directorio de tablas de impuestos ", _
                "Captura rapida nomina ", Dir_imptos)
 CnvNom.Caption = "CONVERSION NOMINA --> Directorio : " + Dir_imptos
End Sub

Private Sub ArcSalida_Click()
  Unload CnvNom
End Sub

Private Sub ArTrNom_Click()
  Dim Wb As Integer, Wc As Integer
    For Wb = 1 To CxNom.Rows - 1
        If IsNumeric(CxNom.TextMatrix(Wb, 2)) Then
             For Wc = 1 To Form8.ConNom1.Rows - 1
                   If CxNom.TextMatrix(Wb, 0) = Form8.ConNom1.TextMatrix(Wc, 0) Then
                       Form8.ConNom1.TextMatrix(Wc, 2) = CxNom.TextMatrix(Wb, 2)
                       Form8.ConNom1.Row = Wc: Form8.ConNom1.Col = 2
                       Form8.Text2.Text = CxNom.TextMatrix(Wb, 2)
                       Form8.checar
                   End If
             Next Wc
        
        End If
    Next Wb
End Sub

Private Sub Form_Load()
    Limpieza
    empresa.ao = 2010
    Dir_imptos = "C:\TARIFA10\"
    z1 = "#,##0.00;(#,##0.00)"
    CnvNom.Caption = "CONVERSION NOMINA --> directorio : " + Dir_imptos
    CxNom.Cols = 6: CxNom.Rows = 1
    CxNom.Row = 0
    CxNom.Col = 0: CxNom.ColWidth(0) = 800: CxNom.CellAlignment = 3: CxNom.Text = "Num.": CxNom.CellFontBold = True
    CxNom.Col = 1: CxNom.ColWidth(1) = 3800: CxNom.CellAlignment = 3: CxNom.Text = "Nombre": CxNom.CellFontBold = True
    CxNom.Col = 2: CxNom.ColWidth(2) = 800: CxNom.CellAlignment = 3: CxNom.Text = "Dias": CxNom.CellFontBold = True
    CxNom.Col = 3: CxNom.ColWidth(3) = 1500: CxNom.CellAlignment = 3: CxNom.Text = "Salario": CxNom.CellFontBold = True
End Sub
Sub Limpieza()
   Dim WA As Integer, We As Integer
   For WA = 1 To Form8.ConNom1.Rows - 1
      For We = 2 To Form8.ConNom1.Cols - 1
           Form8.ConNom1.TextMatrix(WA, We) = ""
      Next We
   Next WA
End Sub
