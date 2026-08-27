VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Form1 
   Caption         =   "Nomina:"
   ClientHeight    =   7530
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
   ScaleHeight     =   7530
   ScaleWidth      =   9255
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8640
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.DriveListBox Drive1 
      Height          =   375
      Left            =   5280
      TabIndex        =   9
      Top             =   240
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   2640
      TabIndex        =   7
      Top             =   6960
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
      Height          =   3345
      Left            =   5160
      TabIndex        =   5
      Top             =   4080
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
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   6960
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FFFF&
      Caption         =   "Archivos:"
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
      Top             =   3600
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
      Height          =   2295
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
   Begin VB.Menu configNom 
      Caption         =   "&Configuración"
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

Private Sub capturarCédulaFiscal_Click()
    FormCedulaFiscal.Show
End Sub

Private Sub configNom_Click()
    Rem nominaConfiguracion.Show
    EstadisticasPersonal.Show
End Sub

Private Sub Form_Load()

    On Error GoTo ManejoError

    Apertura
    veridir
    ActualizarDirectorio

    Kincenal = 1

    mm(1) = "ENERO": mm(2) = "FEBRERO": mm(3) = "MARZO": mm(4) = "ABRIL"
    mm(5) = "MAYO": mm(6) = "JUNIO": mm(7) = "JULIO": mm(8) = "AGOSTO"
    mm(9) = "SEPTIEMBRE": mm(10) = "OCTUBRE": mm(11) = "NOVIEMBRE": mm(12) = "DICIEMBRE"

    dd(1) = 31: dd(2) = 28: dd(3) = 31: dd(4) = 30
    dd(5) = 31: dd(6) = 30: dd(7) = 31: dd(8) = 31
    dd(9) = 30: dd(10) = 31: dd(11) = 30: dd(12) = 31

inicio:

    z1$ = "#,###,##0.00"
    z2$ = "##,##0.00"

    empresarial

    ' Intentar conexión, pero NO impedir el arranque
    ConectarBaseDatos

    File1.Pattern = "*.NOM"
    Label4.Caption = Dir1.Path

    If empresa.ao > 2007 Then
        a_opago = empresa.ao
        tarcre(3).Visible = False
    End If

    GoTo Sale

ManejoError:
    ChDir "C:\"
    Dir1 = "C:\"
    File1 = Dir1
    GoTo inicio

Sale:
    Exit Sub

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
        DirecT_arifas = "C:\TARIFA25\"
        Dir_imptos = DirecT_arifas
        CambioDirectorio
        SubT_Mes = DirecT_arifas
        Dir_imptos = DirecT_arifas
    Else
        Line Input #9, DirecT_arifas
        SubT_Mes = DirecT_arifas
        Close #9
        Exit Sub
   End If

mens1:
    MsgBox ("No se encuentra Microsoft Excel")
    MkDir "C:\GconTA"
End Sub

Sub CambioDirectorio()
    On Error GoTo mens2
    Close
    
    Tot_Dat = 0
    CapturaDatos.Caption = Form1.Caption
    CapturaDatos.Label1(0).Caption = "Subdirectorio Tarifas de retención mensual  : "
    CapturaDatos.Text1(0).Text = DirecT_arifas
    CapturaDatos.Examinar(0).Caption = "Archivar"
    CapturaDatos.Examinar(1).Caption = "Salir"
    CapturaDatos.Show 1
          
    If DirecT_arifas <= "" Then
        DirecT_arifas = "C:\TARIFA25\"
        Open DIR_CALC For Output As 9
        Print #9, DirecT_arifas
        Close #9
    Else
        Open DIR_CALC For Output As 9
        Print #9, DirecT_arifas
        Close #9
    End If
    
Exit Sub

mens2:
    MsgBox "No puedo crear el Rutabl.gcs", vbCritical
End Sub

Sub empresarial()

    On Error GoTo Empre2

    Close #1

    AbreEmpr = "empresa.dno"

    '-----------------------------------------
    ' VERIFICAR SI EL ARCHIVO YA EXISTE
    '-----------------------------------------
    If Dir$(AbreEmpr) = "" Then

        Label1.Caption = _
            "No existe Empresa Cambie subdirectorio o Capture datos"

        Label2.Caption = " "

        Exit Sub

    End If

    '-----------------------------------------
    ' ABRIR SOLO SI YA EXISTE
    '-----------------------------------------
    Open AbreEmpr For Random As #1 Len = Len(empresa)

    cm = LOF(1) / Len(empresa)

    If cm > 0 Then

        Get #1, cm, empresa

        emp = empresa.name

        Label1.Caption = "Empresa : " & empresa.name

        Label2.Caption = _
            "Año : " & Str$(empresa.ao) & Chr(13) & _
            " Salario Minimo $" & Str$(empresa.psub) & Chr(13) & _
            "UMA x dia :$ " & Format(empresa.sm, z1$)

        Close #1

    Else

        Close #1

        Label1.Caption = _
            "No existe Empresa Cambie subdirectorio o Capture datos"

        Label2.Caption = " "

    End If

    Exit Sub

Empre2:

    On Error Resume Next

    Close #1

    Label1.Caption = _
        "No existe Empresa Cambie subdirectorio o Capture datos"

    Label2.Caption = " "

End Sub

Private Sub arDirTar_Click()
   CambioDirectorio
   SubT_Mes = DirecT_arifas
   Dir_imptos = DirecT_arifas
   Close 9
End Sub

Private Sub creasub_Click(Index As Integer)
   Text1.SetFocus
End Sub

Private Sub Dir1_Click()
    
    ' Un solo clic únicamente selecciona la carpeta.
    ' No se navega ni se actualiza información.

End Sub

Private Sub Dir1_DblClick()

    AbrirDirectorioSeleccionado

End Sub
Private Sub Dir1_Change()

    ActualizarDirectorio

End Sub
Private Sub Dir1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then

        KeyAscii = 0

        AbrirDirectorioSeleccionado

    End If

End Sub
Private Sub Drive1_Change()

    On Error GoTo manejadorError

    ChDrive Drive1.Drive

    Dir1.Path = Drive1.Drive

    Exit Sub

manejadorError:

    MsgBox Err.Description, vbExclamation, "Cambio de unidad"

End Sub

Private Sub Emp_dat_Click(Index As Integer)
    Load Form3
    Form3.Show
End Sub


Private Sub empcom_Click()
    Load Repre
    Repre.Show 1
End Sub

Private Sub File1_KeyPress(KeyAscii As Integer)

    Dim archivoDirectorio As String
    Dim rutaCompleta As String
    Dim numeroError As Long
    Dim descripcionError As String

    If KeyAscii <> 13 Then Exit Sub

    KeyAscii = 0

    archivoDirectorio = Trim$(File1.FileName)

    If archivoDirectorio = "" Then Exit Sub

    If Right$(File1.Path, 1) = "\" Then
        rutaCompleta = File1.Path & archivoDirectorio
    Else
        rutaCompleta = File1.Path & "\" & archivoDirectorio
    End If

    On Error GoTo ErrorAbrir

    'Cargar el formulario
    Load FormViewer

    'Cargar la información ANTES de mostrarlo
    FormViewer.LoadPayroll rutaCompleta

    'Si ocurrió un problema durante la carga,
    'no mostramos un visor incompleto
    If Not FormViewer.CargaCorrecta Then

        Unload FormViewer

        File1.Refresh
        File1.SetFocus

        Exit Sub

    End If

    'Ocultar principal solamente cuando
    'la nómina ya fue cargada correctamente
    Me.Hide

    'IMPORTANTE:
    'El visor será MODAL.
    FormViewer.Show vbModal

    'Cuando FormViewer se cierre,
    'el código continúa exactamente aquí.
    Me.Show

    File1.Refresh
    File1.SetFocus

    Exit Sub


ErrorAbrir:

    numeroError = Err.Number
    descripcionError = Err.Description

    On Error Resume Next

    Unload FormViewer

    Me.Show
    File1.Refresh
    File1.SetFocus

    On Error GoTo 0

    MsgBox "No fue posible abrir la nómina." & vbCrLf & _
           "Error " & numeroError & ": " & descripcionError, _
           vbCritical, "Visor de nómina"

End Sub
Private Sub NomAcu_Click(Index As Integer)
  Unload Form1
  Load Form9
  Form9.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Close: End
End Sub

Private Sub nomcap_Click(Index As Integer)
  Close 2: Open "personal.dno" For Random As 2 Len = Len(personal)
  Dm = LOF(2) / Len(personal)

    If Dm > 0 Then
        Form1.Hide
        Load Form8
        Form8.Show
        Form1.Hide
    Else
        Close 2
        MsgBox "No existe archivo de personal no es posible capturar nomina", vbCritical + vbDefaultButton1, "Nomina"
    End If
    
End Sub

Private Sub percap_Click(Index As Integer)
  Load Form2
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
Close: End
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

Private Sub ConectarBaseDatos()

    On Error Resume Next

    If con.State <> 0 Then
        con.Close
    End If

    Err.Clear

    strconnect = "Provider=SQLOLEDB;Data Source=SQL6012.site4now.net;Initial Catalog=db_a4091b_sacmag"

    con.Open strconnect, _
             "db_a4091b_sacmag_admin", _
             "Sacmag2020"

    ' Si falla SQL, simplemente continuar
    Err.Clear

    On Error GoTo 0

End Sub
Private Sub ActualizarDirectorio()

    Dim RutaActual As String

    On Error GoTo ErrorActualizar

    RutaActual = Trim$(Dir1.Path)

    If RutaActual = "" Then Exit Sub

    '-----------------------------------------
    ' Cambiar el directorio actual del programa
    '-----------------------------------------
    ChDir RutaActual

    '-----------------------------------------
    ' Actualizar FileListBox
    '-----------------------------------------
    File1.Pattern = "*.NOM"
    File1.Path = RutaActual
    File1.Refresh

    '-----------------------------------------
    ' Guardar directorio actual
    '-----------------------------------------
    Direc_torio = RutaActual

    '-----------------------------------------
    ' Mostrar ruta
    '-----------------------------------------
    Label4.Caption = RutaActual

    '-----------------------------------------
    ' Guardar último directorio utilizado
    '-----------------------------------------
    On Error Resume Next

    Close #7

    Open "C:\GconTA\perma.dno" _
        For Random As #7 Len = Len(basico)

    If Err.Number = 0 Then

        fin_basico = LOF(7) / Len(basico)

        basico.datoarch = RutaActual

        Put #7, 1, basico

    End If

    Close #7

    Err.Clear

    On Error GoTo ErrorActualizar

    '-----------------------------------------
    ' Actualizar información de la empresa
    '-----------------------------------------
    Close #1

    empresarial

    Exit Sub

ErrorActualizar:

    On Error Resume Next

    Close #7

End Sub
Private Sub AbrirDirectorioSeleccionado()

    Dim RutaSeleccionada As String

    On Error GoTo ErrorDirectorio

    If NavegandoDirectorio Then Exit Sub

    If Dir1.ListIndex < 0 Then Exit Sub

    RutaSeleccionada = Dir1.List(Dir1.ListIndex)

    If Trim$(RutaSeleccionada) = "" Then Exit Sub

    NavegandoDirectorio = True

    ' Cambiar realmente al directorio seleccionado
    If UCase$(Dir1.Path) <> UCase$(RutaSeleccionada) Then
        Dir1.Path = RutaSeleccionada
    End If

    NavegandoDirectorio = False

    Exit Sub

ErrorDirectorio:

    NavegandoDirectorio = False

    MsgBox "No fue posible acceder al directorio." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbExclamation, "Directorio"

End Sub

' comentario
