VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Form1 
   Caption         =   "Nomina:"
   ClientHeight    =   7320
   ClientLeft      =   6630
   ClientTop       =   3570
   ClientWidth     =   9255
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Nomina.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   9255
   Begin VB.ListBox lstNominas 
      Height          =   3375
      Left            =   5160
      TabIndex        =   10
      Top             =   3840
      Width           =   3975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8640
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.DriveListBox Drive1 
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   240
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   2640
      TabIndex        =   7
      Top             =   6840
      Width           =   2295
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2955
      Left            =   5160
      TabIndex        =   5
      Top             =   4200
      Width           =   3975
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2565
      Left            =   5160
      TabIndex        =   2
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Nombre directorio nuevo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   6840
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FFFF&
      Caption         =   "--------------------------------- ARCHIVOS ---------------------------------"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   6
      Top             =   3480
      Width           =   3975
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   4815
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Directorio:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   4815
   End
   Begin VB.Menu per_sonal 
      Caption         =   "&Personal"
      Index           =   0
      Begin VB.Menu percap 
         Caption         =   "&Captura"
         Index           =   1
         Shortcut        =   {F2}
      End
      Begin VB.Menu PerCfdi 
         Caption         =   "C&fdi Complemento"
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu peredic 
         Caption         =   "&Edicion"
         Index           =   2
         Shortcut        =   {F1}
      End
      Begin VB.Menu perimpr 
         Caption         =   "&Impresion"
         Index           =   3
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu persal 
         Caption         =   "&Salida"
         Index           =   4
      End
   End
   Begin VB.Menu nom_na 
      Caption         =   "&Nomina"
      Index           =   0
      Begin VB.Menu nomcap 
         Caption         =   "&Captura"
         Index           =   1
         Shortcut        =   {F5}
      End
      Begin VB.Menu nomimpr 
         Caption         =   "&Impresion"
         Index           =   2
      End
   End
   Begin VB.Menu emp_sa 
      Caption         =   "&Empresa"
      Index           =   0
      Begin VB.Menu Emp_dat 
         Caption         =   "&Datos"
         Index           =   1
      End
      Begin VB.Menu ArSep2 
         Caption         =   "-"
      End
      Begin VB.Menu empcom 
         Caption         =   "&Complementarios"
      End
      Begin VB.Menu arSep3 
         Caption         =   "-"
      End
      Begin VB.Menu arDirTar 
         Caption         =   "&Directorio de Tarifas"
      End
   End
   Begin VB.Menu crear 
      Caption         =   "&Crear subdirectorio"
      Index           =   0
      Begin VB.Menu creasub 
         Caption         =   "&Subdirectorio"
         Index           =   1
      End
   End
   Begin VB.Menu tarimp 
      Caption         =   "&Tarifas impuestos"
      Index           =   0
      Begin VB.Menu tarart 
         Caption         =   "&Articulo 113"
         Index           =   1
      End
      Begin VB.Menu tarsub 
         Caption         =   "&Subsidio 114"
         Index           =   2
      End
      Begin VB.Menu tarcre 
         Caption         =   "&Credito al salario"
         Index           =   3
      End
      Begin VB.Menu taranu 
         Caption         =   "A&nual"
         Index           =   4
         Begin VB.Menu taranisr 
            Caption         =   "&ISR Anual 177"
            Index           =   5
         End
         Begin VB.Menu taransub 
            Caption         =   "&Subsidio anual 178"
            Index           =   6
         End
         Begin VB.Menu tarancre 
            Caption         =   "&Credito salario"
            Index           =   7
         End
      End
   End
   Begin VB.Menu Ayuda 
      Caption         =   "&Ayuda"
      Begin VB.Menu Version 
         Caption         =   "&Version"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private NavegandoDirectorio As Boolean

Private Sub Form_Load()
    Dim rutaInicio As String
    On Error GoTo ManejoError
    ElegirEntornoNomina
    Apertura

    If gModoTesting Then
        rutaInicio = RutaInicialNomina()

        NavegandoDirectorio = True
    
        If mID$(rutaInicio, 2, 1) = ":" Then
            ChDrive Left$(rutaInicio, 2)
            Drive1.Drive = Left$(rutaInicio, 2)
        End If
    
        ChDir rutaInicio
        Dir1.Path = rutaInicio
        NavegandoDirectorio = False
    
    Else
        NavegandoDirectorio = True
        veridir
        NavegandoDirectorio = False
    End If

    ActualizarDirectorio

    Kincenal = 1

    InicializarMeses
    InicializarDiasMes

    z1$ = "#,###,##0.00"
    z2$ = "##,##0.00"

    empresarial
    AplicarPoliticaConexion
    
    File1.Pattern = "*.NOM"

    File1.Visible = False
    lstNominas.Visible = True
    lstNominas.ZOrder 0

    Label4.Caption = Dir1.Path

    ActualizarListaNominasAgrupadas

    Me.Caption = "Nomina [" & EtiquetaEntornoNomina() & "]"

    If empresa.ao > 2007 Then
        a_opago = empresa.ao
        tarcre(3).Visible = False
    End If

    Exit Sub

ManejoError:
    MsgBox "No fue posible iniciar el entorno." & vbCrLf & Err.Description, vbCritical, "Nomina"
    gEntornoElegido = False
    gRutaTrabajoAprobada = vbNullString
    Close
    End
End Sub

Private Sub Form_Activate()

    On Error GoTo ErrorActualizarLista
    If NavegandoDirectorio Then Exit Sub
    If StrComp(File1.Path, Dir1.Path, vbTextCompare) <> 0 Then
        Exit Sub
    End If

    File1.Refresh
    ActualizarListaNominasAgrupadas
    Exit Sub
ErrorActualizarLista:
    Debug.Print "Error actualizando listado: " & Err.Number & " - " & Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Close
    End
End Sub

Private Sub InicializarMeses()
    mm(1) = "ENERO"
    mm(2) = "FEBRERO"
    mm(3) = "MARZO"
    mm(4) = "ABRIL"
    mm(5) = "MAYO"
    mm(6) = "JUNIO"
    mm(7) = "JULIO"
    mm(8) = "AGOSTO"
    mm(9) = "SEPTIEMBRE"
    mm(10) = "OCTUBRE"
    mm(11) = "NOVIEMBRE"
    mm(12) = "DICIEMBRE"
End Sub

Private Sub InicializarDiasMes()
    dd(1) = 31
    dd(2) = 28
    dd(3) = 31
    dd(4) = 30
    dd(5) = 31
    dd(6) = 30
    dd(7) = 31
    dd(8) = 31
    dd(9) = 30
    dd(10) = 31
    dd(11) = 30
    dd(12) = 31
End Sub

Sub Apertura()
    On Error GoTo mens1
    
    DIR_CALC = "C:\GconTA\Rutabl.gcs"
    
    miarchivo = ""
    miarchivo = Dir(DIR_CALC, vbDirectory)

    If miarchivo <> "" Then
        Open DIR_CALC For Input As #9
        Mi_limite = LOF(9)
    Else
        Mi_limite = 0
    End If

        If Mi_limite = 0 Then
        Close #9
        DirecT_arifas = "C:\TARIFA26\"
        CambioDirectorio
    Else
        Line Input #9, DirecT_arifas
        Close #9
    End If

    DirecT_arifas = Trim$(DirecT_arifas)
    If DirecT_arifas <> "" And Right$(DirecT_arifas, 1) <> "\" Then
        DirecT_arifas = DirecT_arifas & "\"
    End If
    
    SubT_Mes = DirecT_arifas
    Dir_imptos = DirecT_arifas

    Exit Sub

mens1:
    MsgBox "Error en Apertura: " & Err.Description, vbCritical
End Sub
Sub CambioDirectorio()
    On Error GoTo mens2

    Close #9

    Tot_Dat = 0
    CapturaDatos.Caption = Form1.Caption
    CapturaDatos.Label1(0).Caption = "Subdirectorio Tarifas de retención mensual  : "
    CapturaDatos.Text1(0).Text = DirecT_arifas
    CapturaDatos.Examinar(0).Caption = "Archivar"
    CapturaDatos.Examinar(1).Caption = "Salir"
    CapturaDatos.Show 1

    If DirecT_arifas <= "" Then
        DirecT_arifas = "C:\TARIFA25\"
    End If

    Open DIR_CALC For Output As 9
    Print #9, DirecT_arifas
    Close #9

    Exit Sub

mens2:
    MsgBox "No puedo crear el Rutabl.gcs", vbCritical
End Sub

Sub empresarial()
    On Error GoTo Empre2
    g_EmpresaCargada = False

    Close #1

    AbreEmpr = "empresa.dno"

    If Dir$(AbreEmpr) = "" Then
        MostrarEmpresaNoEncontrada
        Exit Sub
    End If

    Open AbreEmpr For Random As #1 Len = Len(empresa)

    cm = LOF(1) / Len(empresa)

    If cm > 0 Then
        Get #1, cm, empresa
        g_EmpresaCargada = True
        emp = empresa.name

        Label1.Caption = "Empresa : " & empresa.name
        Label2.Caption = "Año : " & Str$(empresa.ao) & Chr(13) & " Salario Minimo $" & Str$(empresa.psub) & Chr(13) & "UMA x dia :$ " & Format(empresa.sm, z1$)

        Close #1
    Else
        Close #1
        MostrarEmpresaNoEncontrada
    End If

    Exit Sub

Empre2:
    On Error Resume Next

    Close #1
    MostrarEmpresaNoEncontrada
End Sub

Private Sub MostrarEmpresaNoEncontrada()
    Label1.Caption = "No existe Empresa Cambie subdirectorio o Capture datos"
    Label2.Caption = " "
End Sub

Private Sub ConectarBaseDatos()
    On Error Resume Next
    If Not SincronizacionRemotaPermitida() Then Exit Sub
    If con.State <> 0 Then
        con.Close
    End If
    Err.Clear
    strconnect = "Provider=SQLOLEDB;Data Source=SQL6012.site4now.net;Initial Catalog=db_a4091b_sacmag"
    con.Open strconnect, "db_a4091b_sacmag_admin", "Sacmag2020"
    Err.Clear

    On Error GoTo 0
End Sub

Private Sub Drive1_Change()
    On Error GoTo ErrorUnidad
    If NavegandoDirectorio Then Exit Sub
    ChDrive Left$(Drive1.Drive, 2)
    Dir1.Path = Left$(Drive1.Drive, 2) & "\"

    Exit Sub

ErrorUnidad:
    MsgBox "No fue posible cambiar de unidad." & vbCrLf & Err.Description, vbExclamation, "Cambio de unidad"
End Sub

Private Sub Dir1_DblClick()
    AbrirDirectorioSeleccionado
End Sub

Private Sub Dir1_Change()
    If NavegandoDirectorio Then Exit Sub
    ActualizarDirectorio
End Sub


Private Sub Dir1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        AbrirDirectorioSeleccionado
    End If
End Sub


Private Sub ActualizarDirectorio()

    Dim RutaActual As String
    On Error GoTo ErrorActualizar

    RutaActual = Trim$(Dir1.Path)

    If RutaActual = "" Then Exit Sub
    If Not RutaPermitidaPorEntorno(RutaActual) Then
        MsgBox "La carpeta seleccionada no existe o no se puede acceder a ella.", vbExclamation, "Directorio"
        Exit Sub
    End If

    If mID$(RutaActual, 2, 1) = ":" Then
        ChDrive Left$(RutaActual, 2)
    End If

    ChDir RutaActual
    gRutaTrabajoAprobada = CurDir$

    File1.Pattern = "*.NOM"
    File1.Path = RutaActual
    File1.Refresh
    
    ActualizarListaNominasAgrupadas

    Direc_torio = RutaActual
    Label4.Caption = RutaActual

    GuardarUltimoDirectorio RutaActual

    empresarial
    AplicarPoliticaConexion

    Exit Sub

ErrorActualizar:
    NavegandoDirectorio = False
    MsgBox "No fue posible cambiar de carpeta." & vbCrLf & Err.Description, vbExclamation, "Directorio"
End Sub

Private Sub GuardarUltimoDirectorio(ByVal RutaActual As String)
    Dim h As Integer
    On Error GoTo ErrorGuardar

    If Not RutaTrabajoSegura(False) Then Exit Sub

    If Not RutaPermitidaPorEntorno(RutaActual) Then
        Exit Sub
    End If

    h = FreeFile
    Open RutaArchivoEntornoNomina() For Random Access Read Write As #h Len = Len(basico)
    basico.datoarch = RutaActual
    Put #h, 1, basico

Salida:
    On Error Resume Next

    If h <> 0 Then Close #h

    On Error GoTo 0
    Exit Sub

ErrorGuardar:
    Resume Salida
End Sub

Private Sub AbrirDirectorioSeleccionado()
    Dim RutaSeleccionada As String

    On Error GoTo ErrorDirectorio

    If NavegandoDirectorio Then Exit Sub
    If Dir1.ListIndex = -1 Then Exit Sub

    RutaSeleccionada = Dir1.List(Dir1.ListIndex)

    If Trim$(RutaSeleccionada) = "" Then Exit Sub

    If StrComp(RutaSeleccionada, Dir1.Path, vbTextCompare) = 0 Then Exit Sub

    If Not RutaPermitidaPorEntorno(RutaSeleccionada) Then
        MsgBox "Carpeta no permitida en este entorno.", _
               vbExclamation, "Directorio"
        Exit Sub
    End If

    NavegandoDirectorio = True
    Dir1.Path = RutaSeleccionada
    NavegandoDirectorio = False

    ActualizarDirectorio

    Exit Sub

ErrorDirectorio:
    NavegandoDirectorio = False

    MsgBox "No fue posible acceder al directorio." & vbCrLf & "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Directorio"
End Sub

Private Sub Dir1_KeyDown(KeyCode As Integer, Shift As Integer)
    'Retroceso o Supr: subir un nivel.
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        KeyCode = 0
        SubirDirectorio
    End If
End Sub

Private Sub SubirDirectorio()
    Dim RutaActual As String
    Dim RutaPadre As String
    Dim p As Long

    On Error GoTo ErrorSubir

    If NavegandoDirectorio Then Exit Sub

    RutaActual = Dir1.Path

    'Si termina en "\" ya estamos en la raiz de la unidad.
    If Right$(RutaActual, 1) = "\" Then Exit Sub

    p = InStrRev(RutaActual, "\")
    If p = 0 Then Exit Sub

    RutaPadre = Left$(RutaActual, p - 1)

    'C:\Carpeta -> "C:" -> "C:\"
    If Right$(RutaPadre, 1) = ":" Then RutaPadre = RutaPadre & "\"

    If Not RutaPermitidaPorEntorno(RutaPadre) Then
        Beep
        Exit Sub
    End If

    NavegandoDirectorio = True
    Dir1.Path = RutaPadre
    NavegandoDirectorio = False

    ActualizarDirectorio

    Exit Sub

ErrorSubir:
    NavegandoDirectorio = False

    MsgBox "No fue posible subir de carpeta." & vbCrLf & "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Directorio"
End Sub

Private Sub File1_KeyPress(KeyAscii As Integer)

    Dim archivoDirectorio As String
    Dim rutaCompleta As String

    Dim numeroError As Long
    Dim descripcionError As String

    If KeyAscii <> vbKeyReturn Then Exit Sub
    KeyAscii = 0
    
    archivoDirectorio = Trim$(File1.FileName)

    If Len(archivoDirectorio) = 0 Then Exit Sub
    If Right$(File1.Path, 1) = "\" Then
        rutaCompleta = File1.Path & archivoDirectorio
    Else
        rutaCompleta = File1.Path & "\" & archivoDirectorio
    End If

    On Error GoTo ErrorAbrir

    If UCase$(Right$(rutaCompleta, 4)) <> ".NOM" Then
        MsgBox "Seleccione un archivo de nomina NOM.", vbExclamation, "Abrir nomina"
        Exit Sub
    End If

    If Len(Dir$(rutaCompleta)) = 0 Then
        MsgBox "El archivo seleccionado no existe.", vbExclamation, "Abrir nomina"
        Exit Sub
    End If

    Load Form8

    If Not Form8.ProgramarEdicionDesdeVisor(rutaCompleta) Then
        Exit Sub
    End If

    Form8.Show vbModal, Me

    Unload Form8
    File1.Refresh
    ActualizarListaNominasAgrupadas
    If Me.Visible And lstNominas.Visible Then
        lstNominas.SetFocus
    End If
    Exit Sub

ErrorAbrir:

    numeroError = Err.Number
    descripcionError = Err.Description

    On Error Resume Next
        File1.Refresh
    ActualizarListaNominasAgrupadas

    If Me.Visible And lstNominas.Visible Then
        lstNominas.SetFocus
    End If

    On Error GoTo 0

    MsgBox "No fue posible abrir la nomina." & vbCrLf & "Error " & CStr(numeroError) & ": " & descripcionError, vbCritical, "Abrir nomina"

End Sub

Private Sub arDirTar_Click()
    CambioDirectorio
    SubT_Mes = DirecT_arifas
    Dir_imptos = DirecT_arifas
    Close 9
End Sub

Private Sub capturarCédulaFiscal_Click()
    FormCedulaFiscal.Show
End Sub

Private Sub configNom_Click()
    EstadisticasPersonal.Show
End Sub

Private Sub creasub_Click(Index As Integer)
    Text1.SetFocus
End Sub

Private Sub Emp_dat_Click(Index As Integer)
    Load Form3
    Form3.Show
End Sub

Private Sub empcom_Click()
    Load Repre
    Repre.Show 1
End Sub

Private Sub NomAcu_Click(Index As Integer)
    Unload Form1
    Load Form9
    Form9.Show
End Sub

Private Sub nomcap_Click(Index As Integer)

    Close #2
    Open "personal.dno" For Random As #2 Len = Len(personal)
    Dm = LOF(2) / Len(personal)

    If Dm > 0 Then
        Load Form8
        Form8.PrepararCapturaNuevaDesdeForm1
        Form1.Hide
        Form8.Show
    Else
        Close #2
        MsgBox "No existe archivo de personal no es posible capturar nomina", vbCritical + vbDefaultButton1, "Nomina"
    End If
End Sub

Private Sub percap_Click(Index As Integer)
    Load Form2
    Form2.IniciarNuevo
    Form2.Show
End Sub

Private Sub PerCfdi_Click()
    PeCfdi.Show
End Sub

Private Sub peredic_Click(Index As Integer)
    Load Form4
    Form4.Show
End Sub

Private Sub perimpr_Click(Index As Integer)
    Load Form4
    Form4.Show
End Sub

Private Sub persal_Click(Index As Integer)
    Unload Form1
    Close
    End
End Sub

Private Sub tarancre_Click(Index As Integer)
    ta_r = 3
    arch_tr = (Dir_imptos + "CRE116.03")
    Load Form7
    Form7.Show
End Sub

Private Sub taranisr_Click(Index As Integer)
    ta_r = 1
    arch_tr = (Dir_imptos + "ISR177.03")
    Load Form5
    Form5.Show
End Sub

Private Sub taransub_Click(Index As Integer)
    ta_r = 2
    arch_tr = (Dir_imptos + "SUB178.03")
    Load Form5
    Form5.Show
End Sub

Private Sub tarart_Click(Index As Integer)
    ta_r = 0
    Load Form5
    Form5.Show
End Sub

Private Sub tarcre_Click(Index As Integer)
    ta_r = 0
    Load Form7
    Form7.Show
End Sub

Private Sub tarsub_Click(Index As Integer)
    ta_r = 0
    Load Form6
    Form6.Show
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MkDir Text1.Text
        MsgBox ("Carpeta creada con éxito!")
        Dir1.Refresh
        Text1.Text = Empty
    End If
End Sub

Private Sub Version_Click()
    Load Form19
    Form19.Show
End Sub

Private Sub AplicarPoliticaConexion()
    On Error Resume Next

    If SincronizacionRemotaPermitida() Then
        If con.State = 0 Then ConectarBaseDatos
    Else
        If con.State <> 0 Then con.Close
        Err.Clear
    End If
End Sub

Private Function MesNombreArchivo(ByVal abreviatura As String) As Integer
    Select Case UCase$(abreviatura)

        Case "ENE": MesNombreArchivo = 1
        Case "FEB": MesNombreArchivo = 2
        Case "MAR": MesNombreArchivo = 3
        Case "ABR": MesNombreArchivo = 4
        Case "MAY": MesNombreArchivo = 5
        Case "JUN": MesNombreArchivo = 6
        Case "JUL": MesNombreArchivo = 7
        Case "AGO": MesNombreArchivo = 8
        Case "SEP": MesNombreArchivo = 9
        Case "OCT": MesNombreArchivo = 10
        Case "NOV": MesNombreArchivo = 11
        Case "DIC": MesNombreArchivo = 12

        Case Else
            MesNombreArchivo = 0
    End Select
End Function

Private Function ClaveOrdenArchivoNomina(ByVal nombreArchivo As String) As String

    Dim nombre As String
    Dim mes As Integer
    Dim quincena As String

    Dim posicion As Long
    Dim numero As String
    Dim tipo As String

    nombre = UCase$(Trim$(nombreArchivo))

    If Right$(nombre, 4) = ".NOM" Then
        nombre = Left$(nombre, Len(nombre) - 4)
    End If
    If Len(nombre) = 8 Then
        mes = MesNombreArchivo(Left$(nombre, 3))
        quincena = mID$(nombre, 4, 1)

        If mes > 0 And _
           (quincena = "1" Or quincena = "2") And Right$(nombre, 4) Like "####" Then
            ClaveOrdenArchivoNomina = "0" & Right$(nombre, 4) & Right$("0" & CStr(mes), 2) & quincena & nombre
            Exit Function
        End If
    End If

    posicion = 1

    Do While posicion <= Len(nombre)
        If mID$(nombre, posicion, 1) < "0" Or mID$(nombre, posicion, 1) > "9" Then
            Exit Do
        End If
        posicion = posicion + 1
    Loop

    If posicion > 1 Then
        numero = Left$(nombre, posicion - 1)
        tipo = mID$(nombre, posicion)

        If Len(numero) <= 9 Then
            ClaveOrdenArchivoNomina = "1" & Right$(String$(9, "0") & numero, 9) & tipo
        Else
            ClaveOrdenArchivoNomina = "1" & numero & tipo
        End If
    Else
        ClaveOrdenArchivoNomina = "1" & String$(9, "9") & nombre
    End If
End Function

Private Function ClaveEsNominaNormal(ByVal clave As String) As Boolean
    ClaveEsNominaNormal = (Left$(clave, 1) = "0")
End Function

Private Sub ActualizarListaNominasAgrupadas()

    Dim nombres() As String
    Dim claves() As String
    Dim indices() As Long

    Dim cantidad As Long
    Dim i As Long
    Dim j As Long

    Dim temporalTexto As String
    Dim temporalIndice As Long

    Dim indiceVisual As Long
    Dim grupoActual As Integer
    Dim grupoArchivo As Integer

    On Error GoTo ErrorLista
    
    lstNominas.Clear
    cantidad = File1.ListCount

    If cantidad <= 0 Then
        lstNominas.AddItem "(No hay nominas guardadas)"
        lstNominas.ItemData(0) = -1
        Exit Sub
    End If

    ReDim nombres(0 To cantidad - 1)
    ReDim claves(0 To cantidad - 1)
    ReDim indices(0 To cantidad - 1)

    For i = 0 To cantidad - 1
        nombres(i) = File1.List(i)
        claves(i) = ClaveOrdenArchivoNomina(nombres(i))
        indices(i) = i
    Next i
    
    For i = 0 To cantidad - 2
        For j = i + 1 To cantidad - 1
            If StrComp(claves(i), claves(j), vbTextCompare) > 0 Then
                       
                temporalTexto = claves(i)
                claves(i) = claves(j)
                claves(j) = temporalTexto

                temporalTexto = nombres(i)
                nombres(i) = nombres(j)
                nombres(j) = temporalTexto

                temporalIndice = indices(i)
                indices(i) = indices(j)
                indices(j) = temporalIndice
            End If
        Next j
    Next i

    grupoActual = -1
    
    For i = 0 To cantidad - 1
        If ClaveEsNominaNormal(claves(i)) Then
            grupoArchivo = 0
        Else
            grupoArchivo = 1
        End If
        
        If grupoArchivo <> grupoActual Then
            If grupoActual <> -1 Then
                lstNominas.AddItem ""
                lstNominas.ItemData(lstNominas.NewIndex) = -1
            End If
            If grupoArchivo = 0 Then
                lstNominas.AddItem _
                    "---- NOMINA NORMAL ----"
            Else
                lstNominas.AddItem _
                    "---- NOMINA ESPECIAL ----"
            End If
            lstNominas.ItemData(lstNominas.NewIndex) = -1
            grupoActual = grupoArchivo
        End If
        
        lstNominas.AddItem nombres(i)
        indiceVisual = lstNominas.NewIndex
        lstNominas.ItemData(indiceVisual) = indices(i)
    Next i
    Exit Sub

ErrorLista:
    Debug.Print "Error agrupando nominas: " & Err.Number & " - " & Err.Description
End Sub

Private Sub AbrirNominaListaAgrupada()

    Dim indiceArchivo As Long
    Dim tecla As Integer

    If lstNominas.ListIndex < 0 Then Exit Sub
    indiceArchivo = lstNominas.ItemData(lstNominas.ListIndex)

    If indiceArchivo < 0 Then Exit Sub
    If indiceArchivo >= File1.ListCount Then
        MsgBox "La lista de nominas ha cambiado." & vbCrLf & "Actualice la seleccion.", vbExclamation, "Nomina"

        File1.Refresh
        ActualizarListaNominasAgrupadas
        Exit Sub
    End If

    File1.ListIndex = indiceArchivo

    tecla = vbKeyReturn
    File1_KeyPress tecla

End Sub

Private Sub lstNominas_DblClick()
    AbrirNominaListaAgrupada
End Sub

Private Sub lstNominas_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        AbrirNominaListaAgrupada
    End If
End Sub

