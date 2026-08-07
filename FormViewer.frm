VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormViewer 
   Caption         =   "Visualizador de Nomina Histórica"
   ClientHeight    =   8200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   LinkTopic       =   "FormViewer"
   ScaleHeight     =   8200
   ScaleWidth      =   12000
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdCopyAll 
      Caption         =   "Copiar Todo"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton CmdCopySel 
      Caption         =   "Copiar Selección"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Regresar"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label LblFile 
      Caption         =   "Archivo:"
      Height          =   255
      Left            =   5640
      TabIndex        =   4
      Top             =   180
      Width           =   6200
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid ConNom1 
      Height          =   7200
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   12700
      _Version        =   393216
      FixedRows       =   1
      AllowUserResizing=   3
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   7850
      Visible         =   0   'False
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "FormViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim N_ormal As Integer
Dim diat As Double
Dim Arch As String
Dim Arch1 As String
Dim Dm As Long
Dim nm As Long
Dim fi_nm As Long
Dim limite As Long
Dim renglon As Long

Private Sub Form_Load()
    ' Asegurar que z1$ esté definido
    If z1$ = "" Then z1$ = "###,###,##0.00"
End Sub

Public Sub LoadPayroll(ByVal fName As String)
    On Error GoTo ErrorHandler
    
    Dim nameOnly As String
    nameOnly = fName
    
    Dim pos As Integer
    pos = InStrRev(fName, "\")
    If pos > 0 Then nameOnly = Mid$(fName, pos + 1)
    
    LblFile.Caption = "Archivo: " & nameOnly
    
    ' Abrir archivo de personal para obtener el total de empleados (Dm)
    Close 2
    Open "personal.dno" For Random As 2 Len = Len(personal)
    Dm = LOF(2) / Len(personal)
    
    If Dm <= 0 Then
        MsgBox "No existe archivo de personal (personal.dno). No se puede abrir la nomina.", vbCritical, "Error"
        Close 2
        Exit Sub
    End If
    
    If IsNormalPayrollName(nameOnly) Then
        N_ormal = 0
        diat = 15 ' Valor por defecto para carga
    Else
        N_ormal = 1
        diat = 0
    End If
    
    Arch = nameOnly
    Arch1 = Left$(nameOnly, Len(nameOnly) - 4) & ".cmp"
    
    Close 6
    Close 12
    Close 14
    Close 8
    
    Open Arch For Random As 6 Len = Len(nomina)
    nm = LOF(6) / Len(nomina)
    
    Open Arch1 For Random As 14 Len = Len(nom_com)
    Open "bnxcla.dno" For Random As 12 Len = Len(Clbnx)
    Open "maestro.dno" For Random As 8 Len = Len(maestro)
    
    If nm > 0 Then
        If nm < Dm Then
            fi_nm = Dm
        Else
            fi_nm = nm
        End If
        
        ConNom1.Clear
        define
        
        ConNom1.Rows = fi_nm + 2
        limite = 0
        renglon = 0
        
        ProgressBar1.Visible = True
        ProgressBar1.Min = 0
        ProgressBar1.Max = fi_nm
        ProgressBar1.Value = 0
        
        Dim r As Long
        Dim yavas As Integer
        
        For r = 1 To fi_nm
            ProgressBar1.Value = r
            Get #6, r, nomina
            Get #2, r, personal
            Get #8, r, maestro
            Get #14, r, nom_com
            
            yavas = 0
            verifica yavas
            
            If yavas > 0 Then
                renglon = renglon + 1
                ConNom1.Row = renglon
                ConNom1.Col = 0
                ConNom1.Text = Format(r, "#####")
                limite = limite + 1
                regtro = r
                carganom
            End If
        Next r
        
        eliminacion
        sumavert
    Else
        MsgBox "El archivo de nomina está vacío o no contiene registros.", vbInformation, "Visor de Nomina"
    End If
    
    ProgressBar1.Visible = False
    Close 6
    Close 14
    Close 12
    Close 8
    Close 2
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al cargar la nomina: " & Err.Description, vbCritical, "Error"
    Close 6
    Close 14
    Close 12
    Close 8
    Close 2
End Sub

Private Function IsNormalPayrollName(ByVal fName As String) As Boolean
    Dim namePart As String
    If Len(fName) = 12 Then
        namePart = UCase(Left$(fName, 3))
        If namePart = "ENE" Or namePart = "FEB" Or namePart = "MAR" Or namePart = "ABR" Or _
           namePart = "MAY" Or namePart = "JUN" Or namePart = "JUL" Or namePart = "AGO" Or _
           namePart = "SEP" Or namePart = "OCT" Or namePart = "NOV" Or namePart = "DIC" Then
            
            Dim q As String
            q = Mid$(fName, 4, 1)
            If q = "1" Or q = "2" Then
                If IsNumeric(Mid$(fName, 5, 4)) Then
                    IsNormalPayrollName = True
                    Exit Function
                End If
            End If
        End If
    End If
    IsNormalPayrollName = False
End Function

Sub carganom()
    Dim ingresos As Currency
    Dim deducciones As Currency
    Dim neto As Currency
    Dim ii As Integer
    Dim li As Long
    
    ConNom1.TextMatrix(ConNom1.Row, 1) = RTrim$(personal.ape1) + " " + RTrim$(personal.ape2) + " " + RTrim$(personal.nom)
    li = ConNom1.Row
    ii = 2
    
    If nomina.dias <> 0 Then ConNom1.TextMatrix(li, 2) = Format(nomina.dias, "##0.00") Else ConNom1.TextMatrix(li, ii) = ""
    ii = ii + 1
    
    If nomina.sueldo <> 0 Then ConNom1.TextMatrix(li, 3) = Format(nomina.sueldo, z1$) Else ConNom1.TextMatrix(li, ii) = ""
    ii = ii + 1
    
    If nomina.hs_nor <> 0 Then ConNom1.TextMatrix(li, 4) = Format(nomina.hs_nor, z1$) Else ConNom1.TextMatrix(li, ii) = ""
    ii = ii + 1
    
    If N_ormal = 1 Then
        If nomina.aguin <> 0 Then ConNom1.TextMatrix(li, 5) = Format(nomina.aguin, z1$) Else ConNom1.TextMatrix(li, ii) = ""
    Else
        If nomina.hs_dbl <> 0 Then ConNom1.TextMatrix(li, 5) = Format(nomina.hs_dbl, z1$) Else ConNom1.TextMatrix(li, ii) = ""
    End If
    ii = ii + 1
   
    If N_ormal = 1 Then
       If nomina.ptu <> 0 Then ConNom1.TextMatrix(li, 6) = Format(nomina.ptu, z1$) Else ConNom1.TextMatrix(li, ii) = ""
    Else
       If nomina.hs_tri <> 0 Then ConNom1.TextMatrix(li, 6) = Format(nomina.hs_tri, z1$) Else ConNom1.TextMatrix(li, ii) = ""
    End If
    ii = ii + 1
    
    If nomina.viaticos <> 0 Then ConNom1.TextMatrix(li, 7) = Format(nomina.viaticos, z1$) Else ConNom1.TextMatrix(li, ii) = ""
    ii = ii + 1
     
    If nomina.pvac <> 0 Then ConNom1.TextMatrix(li, 8) = Format(nomina.pvac, z1$) Else ConNom1.TextMatrix(li, ii) = ""
    ii = ii + 1
    
    If nomina.otras <> 0 Then ConNom1.TextMatrix(li, 9) = Format(nomina.otras, z1$) Else ConNom1.TextMatrix(li, ii) = ""
    ii = ii + 1
    
    If nomina.exentos <> 0 Then ConNom1.TextMatrix(li, 10) = Format(nomina.exentos, z1$) Else ConNom1.TextMatrix(li, ii) = ""
    ii = ii + 1
    
    ingresos = nomina.sueldo + nomina.hs_nor + nomina.hs_dbl + nomina.hs_tri + nomina.aguin + nomina.ptu + nomina.viaticos + nomina.pvac + nomina.otras + nomina.exentos
    If ingresos <> 0 Then ConNom1.TextMatrix(li, 11) = Format(ingresos, z1$) Else ConNom1.TextMatrix(li, ii) = ""
    ii = ii + 1
    
    If nomina.ispt <> 0 Then ConNom1.TextMatrix(li, 12) = Format(nomina.ispt, z1$) Else ConNom1.TextMatrix(li, ii) = ""
    ii = ii + 1
    
    If nomina.crdsal <> 0 Then ConNom1.TextMatrix(li, 13) = Format(nomina.crdsal, z1$) Else ConNom1.TextMatrix(li, ii) = ""
    ii = ii + 1
    
    If nomina.imss <> 0 Then ConNom1.TextMatrix(li, 14) = Format(nomina.imss, z1$) Else ConNom1.TextMatrix(li, ii) = ""
    ii = ii + 1
    
    If nomina.prestamos <> 0 Then ConNom1.TextMatrix(li, 15) = Format(nomina.prestamos, z1$) Else ConNom1.TextMatrix(li, ii) = ""
    ii = ii + 1
    
    If nomina.fonacot <> 0 Then ConNom1.TextMatrix(li, 16) = Format(nomina.fonacot, z1$) Else ConNom1.TextMatrix(li, ii) = ""
    ii = ii + 1
    
    If nomina.telefono <> 0 Then ConNom1.TextMatrix(li, 17) = Format(nomina.telefono, z1$) Else ConNom1.TextMatrix(li, ii) = ""
    ii = ii + 1
    
    If nomina.otraded <> 0 Then ConNom1.TextMatrix(li, 18) = Format(nomina.otraded, z1$) Else ConNom1.TextMatrix(li, ii) = ""
    ii = ii + 1
    
    deducciones = nomina.crdsal + nomina.ispt + nomina.imss + nomina.prestamos + nomina.fonacot + nomina.telefono + nomina.otraded
    If deducciones <> 0 Then ConNom1.TextMatrix(li, 19) = Format(deducciones, z1$) Else ConNom1.TextMatrix(li, ii) = ""
    ii = ii + 1
     
    neto = ingresos - deducciones
    If neto <> 0 Then ConNom1.TextMatrix(li, 20) = Format(neto, z1$) Else ConNom1.TextMatrix(li, ii) = ""
    
    ConNom1.TextMatrix(li, 21) = 0
    ConNom1.TextMatrix(li, 22) = 0
    ConNom1.TextMatrix(li, 23) = 0
    If N_ormal = 1 Then ConNom1.TextMatrix(ConNom1.Row, 23) = "1"
    
    Get 12, regtro, Clbnx
    ConNom1.TextMatrix(li, 24) = (" " + Clbnx.Q1)
End Sub

Sub define()
     ConNom1.Cols = 25
     ConNom1.Font = "Arial": ConNom1.Font.Size = 8: ConNom1.Font.Bold = True
     
     ConNom1.Row = 0
     ConNom1.Col = 0: ConNom1.CellAlignment = 4: ConNom1.ColWidth(0) = 600: ConNom1.Text = "No."
     ConNom1.Col = 1: ConNom1.CellAlignment = 4: ConNom1.ColWidth(1) = 3500: ConNom1.Text = "Nombre"
     ConNom1.Col = 2: ConNom1.CellAlignment = 4: ConNom1.ColWidth(2) = 1200: ConNom1.Text = "dias T."
     ConNom1.Col = 3: ConNom1.CellAlignment = 4: ConNom1.ColWidth(3) = 1200: ConNom1.Text = "Sueldo"
     ConNom1.Col = 4: ConNom1.CellAlignment = 4: ConNom1.ColWidth(4) = 1200: ConNom1.Text = "hs.Norm."
     ConNom1.Col = 5: ConNom1.CellAlignment = 4: ConNom1.ColWidth(5) = 1200
      
      If N_ormal = 0 Then
          ConNom1.Text = "hs.Dobles"
      Else
          ConNom1.Text = "Aguinaldo"
      End If
      
      ConNom1.Col = 6: ConNom1.CellAlignment = 4: ConNom1.ColWidth(6) = 1200
      
      If N_ormal = 0 Then
          ConNom1.Text = "hs.Triples"
      Else
          ConNom1.Text = "Ptu"
      End If
      
      If N_ormal = 0 Then
          ConNom1.Col = 7: ConNom1.CellAlignment = 4: ConNom1.ColWidth(7) = 1200: ConNom1.Text = "O.F."
      Else
          ConNom1.Col = 7: ConNom1.CellAlignment = 4: ConNom1.ColWidth(7) = 1200: ConNom1.Text = "Premio punt."
      End If
     
      ConNom1.Col = 8: ConNom1.CellAlignment = 4: ConNom1.ColWidth(8) = 1200: ConNom1.Text = "P.Vacac."
      ConNom1.Col = 9: ConNom1.CellAlignment = 4: ConNom1.ColWidth(9) = 1200: ConNom1.Text = "Otras"
      ConNom1.Col = 10: ConNom1.CellAlignment = 4: ConNom1.ColWidth(10) = 1200: ConNom1.Text = "Perc.exenta."
      ConNom1.Col = 11: ConNom1.CellAlignment = 4: ConNom1.ColWidth(11) = 1200: ConNom1.Text = "Tot.Ingr."
      ConNom1.Col = 12: ConNom1.CellAlignment = 4: ConNom1.ColWidth(12) = 1200: ConNom1.Text = "Ispt"
      ConNom1.Col = 13: ConNom1.CellAlignment = 4: ConNom1.ColWidth(13) = 1200: ConNom1.Text = "Sub.P/Empl."
      ConNom1.Col = 14: ConNom1.CellAlignment = 4: ConNom1.ColWidth(14) = 1200: ConNom1.Text = "Imss"
      ConNom1.Col = 15: ConNom1.CellAlignment = 4: ConNom1.ColWidth(15) = 1200: ConNom1.Text = "Prestamos"
      ConNom1.Col = 16: ConNom1.CellAlignment = 4: ConNom1.ColWidth(16) = 1200: ConNom1.Text = "Fonacot"
      ConNom1.Col = 17: ConNom1.CellAlignment = 4: ConNom1.ColWidth(17) = 1200: ConNom1.Text = "Pension Alimenticia"
      ConNom1.Col = 18: ConNom1.CellAlignment = 4: ConNom1.ColWidth(18) = 1200: ConNom1.Text = "Infonavit"
      ConNom1.Col = 19: ConNom1.CellAlignment = 4: ConNom1.ColWidth(19) = 1200: ConNom1.Text = "Tot.Deduc"
      ConNom1.Col = 20: ConNom1.CellAlignment = 4: ConNom1.ColWidth(20) = 1200: ConNom1.Text = "Neto"
      ConNom1.Col = 21: ConNom1.ColWidth(21) = 0
      ConNom1.Col = 22: ConNom1.ColWidth(22) = 0
      ConNom1.Col = 23: ConNom1.ColWidth(23) = 0
      ConNom1.Col = 24: ConNom1.CellAlignment = 4: ConNom1.ColWidth(24) = 2400: ConNom1.Text = "Banamex"
End Sub

Sub verifica(yavas)
    If nomina.dias <> 0 Then yavas = 1
    If nomina.hs_nor <> 0 Then yavas = 1
    If nomina.hs_dbl <> 0 Then yavas = 1
    If nomina.hs_tri <> 0 Then yavas = 1
    If nomina.ispt <> 0 Then yavas = 1
    If nomina.crdsal <> 0 Then yavas = 1
    If nomina.imss <> 0 Then yavas = 1
    If nomina.sueldo <> 0 Then yavas = 1
    If nomina.viaticos <> 0 Then yavas = 1
    If nomina.pvac <> 0 Then yavas = 1
    If nomina.otras <> 0 Then yavas = 1
    If nomina.aguin <> 0 Then yavas = 1
    If nomina.ptu <> 0 Then yavas = 1
    If nomina.exentos <> 0 Then yavas = 1
    If nomina.prestamos <> 0 Then yavas = 1
    If nomina.fonacot <> 0 Then yavas = 1
    If nomina.telefono <> 0 Then yavas = 1
    If nomina.otraded <> 0 Then yavas = 1
End Sub

Sub eliminacion()
    Dim re As Long
    re = 0
    Do Until re = (limite)
        re = re + 1
        If (Trim(ConNom1.TextMatrix(re, 11)) = "") Then
            ConNom1.RemoveItem re
            re = re - 1
            limite = limite - 1
        End If
    Loop
End Sub

Sub sumavert()
    Dim Vw As Integer
    Dim late As Integer
    Dim li As Long
    Dim ii As Integer
    Dim sumv(20) As Currency
    
    ProgressBar1.Min = 0
    ProgressBar1.Max = ConNom1.Rows
    ProgressBar1.Value = 0
    ProgressBar1.Visible = True
    
    Vw = 0
    For late = 3 To 20
        sumv(late) = 0
    Next late
        
    For li = 1 To limite
        Vw = Vw + 1
        ProgressBar1.Value = li
        For late = 3 To 20
            ii = late
            If ConNom1.TextMatrix(li, ii) <> "" Then
                sumv(late) = sumv(late) + CCur(ConNom1.TextMatrix(li, ii))
            End If
        Next late
    Next li
    
    li = limite + 1
    ConNom1.TextMatrix(li, 1) = "Empleados... " + Str(Vw) + " S u m a s ....."
    
    For late = 3 To 20
        ii = late
        If sumv(late) <> 0 Then
            ConNom1.TextMatrix(li, ii) = Format(sumv(late), z1$)
        Else
            ConNom1.TextMatrix(li, ii) = ""
        End If
    Next late
    
    ProgressBar1.Visible = False
End Sub

Private Sub CmdCopyAll_Click()
    On Error GoTo CopyAllError
    Dim r As Long, c As Long
    Dim txt As String
    
    For r = 0 To ConNom1.Rows - 1
        Dim rowTxt As String
        rowTxt = ""
        For c = 0 To ConNom1.Cols - 1
            If c > 0 Then rowTxt = rowTxt & vbTab
            rowTxt = rowTxt & ConNom1.TextMatrix(r, c)
        Next c
        If r > 0 Then txt = txt & vbCrLf
        txt = txt & rowTxt
    Next r
    
    Clipboard.Clear
    Clipboard.SetText txt
    MsgBox "Se copió toda la información de la nómina al portapapeles. Ya puedes pegarla en Excel con Ctrl+V.", vbInformation, "Copiar Todo"
    Exit Sub
CopyAllError:
    MsgBox "Error al copiar: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub CmdCopySel_Click()
    On Error GoTo CopySelError
    Dim rStart As Long, rEnd As Long
    Dim cStart As Long, cEnd As Long
    Dim r As Long, c As Long
    Dim txt As String
    
    If ConNom1.Row < ConNom1.RowSel Then
        rStart = ConNom1.Row
        rEnd = ConNom1.RowSel
    Else
        rStart = ConNom1.RowSel
        rEnd = ConNom1.Row
    End If
    
    If ConNom1.Col < ConNom1.ColSel Then
        cStart = ConNom1.Col
        cEnd = ConNom1.ColSel
    Else
        cStart = ConNom1.ColSel
        cEnd = ConNom1.Col
    End If
    
    For r = rStart To rEnd
        Dim rowTxt As String
        rowTxt = ""
        For c = cStart To cEnd
            If c > cStart Then rowTxt = rowTxt & vbTab
            rowTxt = rowTxt & ConNom1.TextMatrix(r, c)
        Next c
        If r > rStart Then txt = txt & vbCrLf
        txt = txt & rowTxt
    Next r
    
    Clipboard.Clear
    Clipboard.SetText txt
    MsgBox "Se copió el rango seleccionado al portapapeles. Ya puedes pegarlo en Excel con Ctrl+V.", vbInformation, "Copiar Selección"
    Exit Sub
CopySelError:
    MsgBox "Error al copiar selección: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub CmdExit_Click()
    Unload Me
    Form1.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState <> vbMinimized Then
        If Me.ScaleWidth > 240 Then
            ConNom1.Width = Me.ScaleWidth - 240
            ProgressBar1.Width = Me.ScaleWidth - 240
        End If
        If Me.ScaleHeight > 1000 Then
            ConNom1.Height = Me.ScaleHeight - 1000
        End If
        If Me.ScaleHeight > 350 Then
            ProgressBar1.Top = Me.ScaleHeight - 350
        End If
    End If
End Sub
