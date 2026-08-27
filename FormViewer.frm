VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormViewer 
   Caption         =   "Visualizador de Nomina Histórica"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   17490
   LinkTopic       =   "FormViewer"
   ScaleHeight     =   8220
   ScaleWidth      =   17490
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdLimpiarFiltro 
      Caption         =   "Limpiar"
      Height          =   375
      Left            =   16320
      TabIndex        =   8
      Top             =   120
      Width           =   1000
   End
   Begin VB.TextBox TxtBuscar 
      Height          =   315
      Left            =   13200
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   120
      Width           =   3000
   End
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
   Begin MSFlexGridLib.MSFlexGrid ConNom1 
      Height          =   7200
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   17265
      _ExtentX        =   30454
      _ExtentY        =   12700
      _Version        =   393216
      AllowUserResizing=   3
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   7845
      Visible         =   0   'False
      Width           =   17265
      _ExtentX        =   30454
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Buscar:"
      Height          =   255
      Left            =   12480
      TabIndex        =   7
      Top             =   240
      Width           =   495
   End
   Begin VB.Label LblFile 
      Caption         =   "Archivo:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   4
      Top             =   180
      Width           =   6195
   End
   Begin VB.Menu mnuOrdenar 
      Caption         =   "Ordenar"
      Begin VB.Menu mnuOrdenarID 
         Caption         =   "Por ID"
         Begin VB.Menu mnuIDDesc 
            Caption         =   "Mayor a Menor"
         End
         Begin VB.Menu mnuIDAsc 
            Caption         =   "Menor a Mayor"
         End
      End
      Begin VB.Menu mnuOrdenarNombre 
         Caption         =   "Por Nombre"
         Begin VB.Menu mnuNombreAsc 
            Caption         =   "A - Z"
         End
         Begin VB.Menu mnuNombreDesc 
            Caption         =   "Z - A"
         End
      End
   End
   Begin VB.Menu mnuExcel 
      Caption         =   "Excel"
      Begin VB.Menu mnuExportarExcel 
         Caption         =   "Exportar a Excel"
      End
   End
   Begin VB.Menu mnuPortapapeles 
      Caption         =   "Portapapeles"
      Begin VB.Menu mnuCopiarTodo 
         Caption         =   "Copiar todo"
      End
      Begin VB.Menu mnuCopiarSeleccion 
         Caption         =   "Copiar selección"
      End
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

Private cargandoNomina As Boolean

Public CargaCorrecta As Boolean

Private Const FILA_INICIO_DATOS As Long = 1
Private Const COL_BUSQUEDA As Long = 1

Private ordenAscendente As Boolean
Private textoFiltro As String

Private Const COL_ID As Long = 0
Private Const COL_NOMBRE As Long = 1

Private Sub CmdLimpiarFiltro_Click()

    Dim i As Long
    Dim ultimaFila As Long

    If cargandoNomina Then Exit Sub

    TxtBuscar.Text = ""
    textoFiltro = ""

    If ConNom1.Rows <= 1 Then Exit Sub

    ultimaFila = limite

    If ultimaFila > ConNom1.Rows - 2 Then
        ultimaFila = ConNom1.Rows - 2
    End If

    If ultimaFila >= 1 Then

        For i = FILA_INICIO_DATOS To ultimaFila

            If Trim$(ConNom1.TextMatrix(i, COL_NOMBRE)) <> "" Then

                ConNom1.RowHeight(i) = -1

            Else

                ConNom1.RowHeight(i) = 0

            End If

        Next i

    End If

    If limite + 1 >= 0 And limite + 1 < ConNom1.Rows Then
        ConNom1.RowHeight(limite + 1) = -1
    End If

    If ConNom1.Rows > 1 Then
        ConNom1.Row = 1
        ConNom1.Col = 0
    End If

End Sub

Private Sub Form_Load()

    cargandoNomina = True
    CargaCorrecta = False

    If z1$ = "" Then
        z1$ = "###,###,##0.00"
    End If

    limite = 0
    renglon = 0
    nm = 0
    fi_nm = 0
    Dm = 0

    textoFiltro = ""

    TxtBuscar.Text = ""

    ConNom1.SelectionMode = flexSelectionByRange
    ConNom1.AllowUserResizing = flexResizeColumns

    cargandoNomina = False

End Sub

Public Sub LoadPayroll(ByVal fName As String)

    On Error GoTo ErrorHandler

    Dim nameOnly As String
    Dim pos As Long
    Dim rutaBase As String

    Dim r As Long
    Dim yavas As Integer

    Dim totalMaestro As Long
    Dim totalCmp As Long
    Dim totalBnx As Long

    Dim numeroError As Long
    Dim descripcionError As String


    '========================================================
    ' INICIO DE CARGA
    '========================================================

    cargandoNomina = True
    CargaCorrecta = False

    CerrarArchivosVisor


    '========================================================
    ' REINICIAR ESTADO DE LA NOMINA ANTERIOR
    '========================================================

    limite = 0
    renglon = 0

    nm = 0
    fi_nm = 0
    Dm = 0

    textoFiltro = ""

    TxtBuscar.Text = ""


    '========================================================
    ' OBTENER NOMBRE Y RUTA
    '========================================================

    nameOnly = fName

    pos = InStrRev(fName, "\")

    If pos > 0 Then

        nameOnly = Mid$(fName, pos + 1)
        rutaBase = Left$(fName, pos)

    Else

        rutaBase = CurDir$

        If Right$(rutaBase, 1) <> "\" Then
            rutaBase = rutaBase & "\"
        End If

    End If


    LblFile.Caption = "Archivo: " & nameOnly


    '========================================================
    ' VALIDAR NOMBRE
    '========================================================

    If Len(nameOnly) < 5 Then

        MsgBox "El nombre del archivo de nómina no es válido.", _
               vbExclamation, "Visor de nómina"

        GoTo Salida

    End If


    '========================================================
    ' ABRIR PERSONAL.DNO
    '========================================================

    Open rutaBase & "personal.dno" _
        For Random As #2 _
        Len = Len(personal)

    Dm = LOF(2) / Len(personal)


    If Dm <= 0 Then

        MsgBox "No existe información válida en personal.dno." & _
               vbCrLf & _
               "No se puede abrir la nómina.", _
               vbCritical, "Error"

        GoTo Salida

    End If


    '========================================================
    ' DETERMINAR NOMINA NORMAL / ESPECIAL
    '========================================================

    If IsNormalPayrollName(nameOnly) Then

        N_ormal = 0
        diat = 15

    Else

        N_ormal = 1
        diat = 0

    End If


    '========================================================
    ' ARCHIVOS
    '========================================================

    Arch = fName

    Arch1 = rutaBase & _
            Left$(nameOnly, Len(nameOnly) - 4) & ".cmp"


    '========================================================
    ' ABRIR NOMINA
    '========================================================

    Open Arch _
        For Random As #6 _
        Len = Len(nomina)

    nm = LOF(6) / Len(nomina)


    If nm <= 0 Then

        MsgBox "El archivo de nómina está vacío o no contiene registros.", _
               vbInformation, "Visor de Nómina"

        GoTo Salida

    End If


    '========================================================
    ' ABRIR ARCHIVOS AUXILIARES
    '========================================================

    Open Arch1 _
        For Random As #14 _
        Len = Len(nom_com)

    Open rutaBase & "bnxcla.dno" _
        For Random As #12 _
        Len = Len(Clbnx)

    Open rutaBase & "maestro.dno" _
        For Random As #8 _
        Len = Len(maestro)


    totalCmp = LOF(14) / Len(nom_com)
    totalBnx = LOF(12) / Len(Clbnx)
    totalMaestro = LOF(8) / Len(maestro)


    '========================================================
    ' DETERMINAR NUMERO SEGURO DE REGISTROS
    '========================================================
    '
    ' Solamente NOM y personal.dno son indispensables
    ' para determinar hasta dónde se puede recorrer.
    '
    ' maestro y cmp se protegen individualmente.
    '
    '========================================================

    If nm < Dm Then

        fi_nm = nm

    Else

        fi_nm = Dm

    End If


    If fi_nm <= 0 Then

        MsgBox "No existen registros compatibles entre la nómina " & _
               "y personal.dno.", _
               vbInformation, "Visor de Nómina"

        GoTo Salida

    End If


    '========================================================
    ' PREPARAR GRID COMPLETAMENTE DESDE CERO
    '========================================================

    ConNom1.Redraw = False

    ConNom1.Clear

    'Primero las columnas
    ConNom1.Cols = 25

    'Después las filas
    ConNom1.Rows = fi_nm + 2

    'Restablecer selección ANTES de trabajar
    ConNom1.Row = 0
    ConNom1.RowSel = 0

    ConNom1.Col = 0
    ConNom1.ColSel = 0

    define

    limite = 0
    renglon = 0


    '========================================================
    ' PROGRESSBAR
    '========================================================

    ProgressBar1.Visible = True

    ProgressBar1.Min = 0
    ProgressBar1.Max = fi_nm
    ProgressBar1.Value = 0


    '========================================================
    ' LEER REGISTROS
    '========================================================

    For r = 1 To fi_nm

        ProgressBar1.Value = r


        '--------------------------------------------
        ' NOMINA
        '--------------------------------------------

        Get #6, r, nomina


        '--------------------------------------------
        ' PERSONAL
        '--------------------------------------------

        Get #2, r, personal


        '--------------------------------------------
        ' MAESTRO
        '--------------------------------------------

        If r <= totalMaestro Then
            Get #8, r, maestro
        End If


        '--------------------------------------------
        ' CMP
        '--------------------------------------------

        If r <= totalCmp Then
            Get #14, r, nom_com
        End If


        '--------------------------------------------
        ' VERIFICAR SI TIENE INFORMACION
        '--------------------------------------------

        yavas = 0

        verifica yavas


        If yavas > 0 Then

            renglon = renglon + 1

            'Garantizar que la fila exista
            If renglon >= ConNom1.Rows Then
                ConNom1.Rows = renglon + 2
            End If

            ConNom1.Row = renglon
            ConNom1.Col = 0

            ConNom1.Text = Format$(r, "#####")

            limite = limite + 1

            regtro = r

            carganom

        End If

    Next r


    '========================================================
    ' ELIMINAR REGISTROS SIN INFORMACION
    '========================================================

    eliminacion


    '========================================================
    ' MUY IMPORTANTE:
    '
    ' Antes de REDUCIR el número de filas,
    ' regresar Row / RowSel a cero.
    '
    ' Esto evita uno de los escenarios del ERROR 381.
    '========================================================

    ConNom1.Row = 0
    ConNom1.RowSel = 0

    ConNom1.Col = 0
    ConNom1.ColSel = 0


    '========================================================
    ' AJUSTAR AL NUMERO REAL DE REGISTROS
    '========================================================

    If limite < 0 Then limite = 0

    ConNom1.Rows = limite + 2


    '========================================================
    ' SUMAS
    '========================================================

    sumavert


    '========================================================
    ' DEJAR SELECCION EN UNA FILA VALIDA
    '========================================================

    If limite > 0 Then

        ConNom1.Row = 1
        ConNom1.RowSel = 1

    Else

        ConNom1.Row = 0
        ConNom1.RowSel = 0

    End If

    ConNom1.Col = 0
    ConNom1.ColSel = 0


    '========================================================
    ' FINALIZAR CORRECTAMENTE
    '========================================================

    ConNom1.Redraw = True

    ProgressBar1.Visible = False

    CargaCorrecta = True

    GoTo Salida


'============================================================
' ERROR
'============================================================

ErrorHandler:

    numeroError = Err.Number
    descripcionError = Err.Description

    CargaCorrecta = False

    On Error Resume Next

    ConNom1.Redraw = True
    ProgressBar1.Visible = False

    On Error GoTo 0

    MsgBox "Error al cargar la nómina." & vbCrLf & _
           "Número: " & numeroError & vbCrLf & _
           "Descripción: " & descripcionError, _
           vbCritical, "Error"


'============================================================
' SALIDA COMUN
'============================================================

Salida:

    CerrarArchivosVisor

    cargandoNomina = False

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

    ConNom1.TextMatrix(ConNom1.Row, 1) = _
        RTrim$(personal.ape1) & " " & _
        RTrim$(personal.ape2) & " " & _
        RTrim$(personal.nom)

    li = ConNom1.Row
    ii = 2


    If nomina.dias <> 0 Then
        ConNom1.TextMatrix(li, 2) = Format$(nomina.dias, "##0.00")
    Else
        ConNom1.TextMatrix(li, ii) = ""
    End If

    ii = ii + 1


    If nomina.sueldo <> 0 Then
        ConNom1.TextMatrix(li, 3) = Format$(nomina.sueldo, z1$)
    Else
        ConNom1.TextMatrix(li, ii) = ""
    End If

    ii = ii + 1


    If nomina.hs_nor <> 0 Then
        ConNom1.TextMatrix(li, 4) = Format$(nomina.hs_nor, z1$)
    Else
        ConNom1.TextMatrix(li, ii) = ""
    End If

    ii = ii + 1


    If N_ormal = 1 Then

        If nomina.aguin <> 0 Then
            ConNom1.TextMatrix(li, 5) = Format$(nomina.aguin, z1$)
        Else
            ConNom1.TextMatrix(li, ii) = ""
        End If

    Else

        If nomina.hs_dbl <> 0 Then
            ConNom1.TextMatrix(li, 5) = Format$(nomina.hs_dbl, z1$)
        Else
            ConNom1.TextMatrix(li, ii) = ""
        End If

    End If

    ii = ii + 1


    If N_ormal = 1 Then

        If nomina.ptu <> 0 Then
            ConNom1.TextMatrix(li, 6) = Format$(nomina.ptu, z1$)
        Else
            ConNom1.TextMatrix(li, ii) = ""
        End If

    Else

        If nomina.hs_tri <> 0 Then
            ConNom1.TextMatrix(li, 6) = Format$(nomina.hs_tri, z1$)
        Else
            ConNom1.TextMatrix(li, ii) = ""
        End If

    End If

    ii = ii + 1


    If nomina.viaticos <> 0 Then
        ConNom1.TextMatrix(li, 7) = Format$(nomina.viaticos, z1$)
    Else
        ConNom1.TextMatrix(li, ii) = ""
    End If

    ii = ii + 1


    If nomina.pvac <> 0 Then
        ConNom1.TextMatrix(li, 8) = Format$(nomina.pvac, z1$)
    Else
        ConNom1.TextMatrix(li, ii) = ""
    End If

    ii = ii + 1


    If nomina.otras <> 0 Then
        ConNom1.TextMatrix(li, 9) = Format$(nomina.otras, z1$)
    Else
        ConNom1.TextMatrix(li, ii) = ""
    End If

    ii = ii + 1


    If nomina.exentos <> 0 Then
        ConNom1.TextMatrix(li, 10) = Format$(nomina.exentos, z1$)
    Else
        ConNom1.TextMatrix(li, ii) = ""
    End If

    ii = ii + 1


    ingresos = nomina.sueldo + _
               nomina.hs_nor + _
               nomina.hs_dbl + _
               nomina.hs_tri + _
               nomina.aguin + _
               nomina.ptu + _
               nomina.viaticos + _
               nomina.pvac + _
               nomina.otras + _
               nomina.exentos


    If ingresos <> 0 Then
        ConNom1.TextMatrix(li, 11) = Format$(ingresos, z1$)
    Else
        ConNom1.TextMatrix(li, ii) = ""
    End If

    ii = ii + 1


    If nomina.ispt <> 0 Then
        ConNom1.TextMatrix(li, 12) = Format$(nomina.ispt, z1$)
    Else
        ConNom1.TextMatrix(li, ii) = ""
    End If

    ii = ii + 1


    If nomina.crdsal <> 0 Then
        ConNom1.TextMatrix(li, 13) = Format$(nomina.crdsal, z1$)
    Else
        ConNom1.TextMatrix(li, ii) = ""
    End If

    ii = ii + 1


    If nomina.imss <> 0 Then
        ConNom1.TextMatrix(li, 14) = Format$(nomina.imss, z1$)
    Else
        ConNom1.TextMatrix(li, ii) = ""
    End If

    ii = ii + 1


    If nomina.prestamos <> 0 Then
        ConNom1.TextMatrix(li, 15) = Format$(nomina.prestamos, z1$)
    Else
        ConNom1.TextMatrix(li, ii) = ""
    End If

    ii = ii + 1


    If nomina.fonacot <> 0 Then
        ConNom1.TextMatrix(li, 16) = Format$(nomina.fonacot, z1$)
    Else
        ConNom1.TextMatrix(li, ii) = ""
    End If

    ii = ii + 1


    If nomina.telefono <> 0 Then
        ConNom1.TextMatrix(li, 17) = Format$(nomina.telefono, z1$)
    Else
        ConNom1.TextMatrix(li, ii) = ""
    End If

    ii = ii + 1


    If nomina.otraded <> 0 Then
        ConNom1.TextMatrix(li, 18) = Format$(nomina.otraded, z1$)
    Else
        ConNom1.TextMatrix(li, ii) = ""
    End If

    ii = ii + 1


    deducciones = nomina.crdsal + _
                  nomina.ispt + _
                  nomina.imss + _
                  nomina.prestamos + _
                  nomina.fonacot + _
                  nomina.telefono + _
                  nomina.otraded


    If deducciones <> 0 Then
        ConNom1.TextMatrix(li, 19) = Format$(deducciones, z1$)
    Else
        ConNom1.TextMatrix(li, ii) = ""
    End If

    ii = ii + 1


    neto = ingresos - deducciones


    If neto <> 0 Then
        ConNom1.TextMatrix(li, 20) = Format$(neto, z1$)
    Else
        ConNom1.TextMatrix(li, ii) = ""
    End If


    ConNom1.TextMatrix(li, 21) = "0"
    ConNom1.TextMatrix(li, 22) = "0"
    ConNom1.TextMatrix(li, 23) = "0"


    If N_ormal = 1 Then
        ConNom1.TextMatrix(li, 23) = "1"
    End If


    '========================================================
    ' BANAMEX
    '========================================================

    If regtro > 0 Then

        If regtro <= LOF(12) / Len(Clbnx) Then

            Get #12, regtro, Clbnx

            ConNom1.TextMatrix(li, 24) = _
                " " & Clbnx.Q1

        Else

            ConNom1.TextMatrix(li, 24) = ""

        End If

    Else

        ConNom1.TextMatrix(li, 24) = ""

    End If

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

    re = 1

    Do While re <= limite

        ' Si no tiene total de ingresos, se considera
        ' un registro sin información de nómina.
        If Trim$(ConNom1.TextMatrix(re, 11)) = "" Then

            ConNom1.RemoveItem re

            limite = limite - 1

        Else

            re = re + 1

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

    Dim r As Long
    Dim c As Long
    Dim txt As String
    Dim filaTexto As String

    txt = ""

    For r = 0 To limite + 1

        filaTexto = ""

        For c = 0 To ConNom1.Cols - 1

            If c > 0 Then
                filaTexto = filaTexto & vbTab
            End If

            filaTexto = filaTexto & ConNom1.TextMatrix(r, c)

        Next c

        If r > 0 Then
            txt = txt & vbCrLf
        End If

        txt = txt & filaTexto

    Next r

    Clipboard.Clear
    Clipboard.SetText txt, vbCFText

    Exit Sub

CopyAllError:

    MsgBox "Error al copiar: " & Err.Description, _
           vbCritical, "Copiar información"
End Sub

Private Sub CmdCopySel_Click()
 On Error GoTo CopySelError

    Dim rStart As Long
    Dim rEnd As Long
    Dim cStart As Long
    Dim cEnd As Long
    Dim r As Long
    Dim c As Long

    Dim txt As String
    Dim filaTexto As String

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

    ' Nunca copiar la fila de totales si se seleccionó junto con ella
    If rEnd > limite Then
        rEnd = limite
    End If

    If rStart < 1 Then
        rStart = 1
    End If

    If rStart > rEnd Then Exit Sub

    txt = ""

    For r = rStart To rEnd

        filaTexto = ""

        For c = cStart To cEnd

            If c > cStart Then
                filaTexto = filaTexto & vbTab
            End If

            filaTexto = filaTexto & ConNom1.TextMatrix(r, c)

        Next c

        If r > rStart Then
            txt = txt & vbCrLf
        End If

        txt = txt & filaTexto

    Next r

    Clipboard.Clear
    Clipboard.SetText txt, vbCFText

    Exit Sub

CopySelError:

    MsgBox "Error al copiar selección: " & Err.Description, _
           vbCritical, "Copiar selección"
End Sub

Private Sub CmdExit_Click()

    Unload Me

End Sub
Private Sub Form_Unload(Cancel As Integer)

    cargandoNomina = True
    CargaCorrecta = False

    CerrarArchivosVisor

End Sub

Private Sub Form_Resize()

    On Error Resume Next

    If Me.WindowState = vbMinimized Then Exit Sub

    If Me.ScaleWidth > 240 Then

        ConNom1.Width = Me.ScaleWidth - 240
        ProgressBar1.Width = Me.ScaleWidth - 240

    End If

    If Me.ScaleHeight > 1500 Then

        ConNom1.Height = Me.ScaleHeight - 1500

    End If

    If Me.ScaleHeight > 350 Then

        ProgressBar1.Top = Me.ScaleHeight - 350

    End If

End Sub

Private Sub mnuCopiarSeleccion_Click()
    CmdCopySel_Click
End Sub

Private Sub mnuCopiarTodo_Click()
    CmdCopyAll_Click
End Sub

Private Sub mnuExportarExcel_Click()
    ExportarNominaExcel
End Sub

Private Sub ExportarNominaExcel()

    On Error GoTo ErrorExcel

    Dim objXL As Object
    Dim wbXL As Object
    Dim wsXL As Object

    Dim r As Long
    Dim c As Long

    Set objXL = CreateObject("Excel.Application")

    If objXL Is Nothing Then

        MsgBox "Necesitas tener instalado Microsoft Excel.", _
               vbExclamation, "Exportar a Excel"

        Exit Sub

    End If

    Set wbXL = objXL.Workbooks.Add
    Set wsXL = wbXL.Worksheets(1)

    objXL.Visible = True

    '------------------------------------------
    ' ENCABEZADOS
    '------------------------------------------

    For c = 0 To ConNom1.Cols - 1

        If ConNom1.ColWidth(c) > 0 Then

            wsXL.Cells(1, c + 1).Value = _
                ConNom1.TextMatrix(0, c)

        End If

    Next c

    '------------------------------------------
    ' DATOS
    '------------------------------------------

    For r = 1 To limite

        For c = 0 To ConNom1.Cols - 1

            If ConNom1.ColWidth(c) > 0 Then

                wsXL.Cells(r + 1, c + 1).Value = _
                    ConNom1.TextMatrix(r, c)

            End If

        Next c

    Next r

    '------------------------------------------
    ' FILA DE TOTALES
    '------------------------------------------

    For c = 0 To ConNom1.Cols - 1

        If ConNom1.ColWidth(c) > 0 Then

            wsXL.Cells(limite + 2, c + 1).Value = _
                ConNom1.TextMatrix(limite + 1, c)

        End If

    Next c

    '------------------------------------------
    ' FORMATO
    '------------------------------------------

    With wsXL.Range( _
        wsXL.Cells(1, 1), _
        wsXL.Cells(1, ConNom1.Cols))

        .Font.Bold = True

    End With

    With wsXL.Range( _
        wsXL.Cells(limite + 2, 1), _
        wsXL.Cells(limite + 2, ConNom1.Cols))

        .Font.Bold = True

    End With

    wsXL.Columns.AutoFit

    wsXL.Activate

    Set wsXL = Nothing
    Set wbXL = Nothing
    Set objXL = Nothing

    Exit Sub

ErrorExcel:

    On Error Resume Next

    If Not objXL Is Nothing Then
        objXL.Visible = True
    End If

    MsgBox "No fue posible exportar la nómina a Excel." & vbCrLf & _
           "Error: " & Err.Description, _
           vbCritical, "Exportar a Excel"

End Sub

Private Sub mnuIDAsc_Click()

    Dim filaActual As Long
    Dim colActual As Long

    If limite <= 1 Then Exit Sub

    filaActual = ConNom1.Row
    colActual = ConNom1.Col

    ConNom1.Redraw = False

    ConNom1.Row = FILA_INICIO_DATOS
    ConNom1.RowSel = limite

    ConNom1.Col = COL_ID
    ConNom1.ColSel = COL_ID

    ' 3 = Numeric Ascending
    ConNom1.Sort = 3

    ConNom1.Row = filaActual
    ConNom1.Col = colActual

    ConNom1.Redraw = True
End Sub

Private Sub mnuIDDesc_Click()
    Dim filaActual As Long
    Dim colActual As Long

    If limite <= 1 Then Exit Sub

    filaActual = ConNom1.Row
    colActual = ConNom1.Col

    ConNom1.Redraw = False

    ConNom1.Row = FILA_INICIO_DATOS
    ConNom1.RowSel = limite

    ConNom1.Col = COL_ID
    ConNom1.ColSel = COL_ID

    ' 4 = Numeric Descending
    ConNom1.Sort = 4

    ConNom1.Row = filaActual
    ConNom1.Col = colActual

    ConNom1.Redraw = True
End Sub

Private Sub mnuNombreAsc_Click()
    Dim filaActual As Long
    Dim colActual As Long

    If limite <= 1 Then Exit Sub

    filaActual = ConNom1.Row
    colActual = ConNom1.Col

    ConNom1.Redraw = False

    ConNom1.Row = FILA_INICIO_DATOS
    ConNom1.RowSel = limite

    ConNom1.Col = COL_NOMBRE
    ConNom1.ColSel = COL_NOMBRE

    ' 5 = String Ascending
    ConNom1.Sort = 5

    ConNom1.Row = filaActual
    ConNom1.Col = colActual

    ConNom1.Redraw = True
End Sub

Private Sub mnuNombreDesc_Click()
    Dim filaActual As Long
    Dim colActual As Long

    If limite <= 1 Then Exit Sub

    filaActual = ConNom1.Row
    colActual = ConNom1.Col

    ConNom1.Redraw = False

    ConNom1.Row = FILA_INICIO_DATOS
    ConNom1.RowSel = limite

    ConNom1.Col = COL_NOMBRE
    ConNom1.ColSel = COL_NOMBRE

    ' 6 = String Descending
    ConNom1.Sort = 6

    ConNom1.Row = filaActual
    ConNom1.Col = colActual

    ConNom1.Redraw = True
End Sub

Private Sub TxtBuscar_Change()

    Dim i As Long
    Dim filtro As String
    Dim nombre As String
    Dim ultimaFila As Long

    'No filtrar mientras se está cargando otra nómina
    If cargandoNomina Then Exit Sub

    'El grid debe tener al menos encabezado + una fila
    If ConNom1.Rows <= 1 Then Exit Sub

    filtro = UCase$(Trim$(TxtBuscar.Text))
    textoFiltro = filtro

    'Nunca permitir recorrer más filas de las que existen
    ultimaFila = limite

    If ultimaFila > ConNom1.Rows - 2 Then
        ultimaFila = ConNom1.Rows - 2
    End If

    If ultimaFila < 1 Then Exit Sub

    For i = FILA_INICIO_DATOS To ultimaFila

        nombre = Trim$(ConNom1.TextMatrix(i, COL_NOMBRE))

        If nombre = "" Then

            ConNom1.RowHeight(i) = 0

        ElseIf filtro = "" Then

            ConNom1.RowHeight(i) = -1

        ElseIf InStr(1, UCase$(nombre), filtro, vbTextCompare) > 0 Then

            ConNom1.RowHeight(i) = -1

        Else

            ConNom1.RowHeight(i) = 0

        End If

    Next i

    'Fila de totales
    If limite + 1 >= 0 And limite + 1 < ConNom1.Rows Then
        ConNom1.RowHeight(limite + 1) = -1
    End If

End Sub

Private Sub ConNom1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyC And (Shift And vbCtrlMask) <> 0 Then

        CmdCopySel_Click
        KeyCode = 0

    End If

End Sub

Private Sub CerrarArchivosVisor()

    On Error Resume Next

    Close #2
    Close #6
    Close #8
    Close #12
    Close #14

    On Error GoTo 0

End Sub

' comentario
