VERSION 5.00
Begin VB.Form Form2 
   ClientHeight    =   8385
   ClientLeft      =   7560
   ClientTop       =   3570
   ClientWidth     =   10770
   ClipControls    =   0   'False
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8385
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnCancelarCambios 
      Caption         =   "Cancelar Cambios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   41
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton BtnHabilitarEdicion 
      Caption         =   "Habilitar Edición"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   40
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Actualizar"
      Height          =   495
      Left            =   8520
      TabIndex        =   39
      Top             =   7080
      Width           =   1935
   End
   Begin VB.TextBox Text16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   1560
      TabIndex        =   38
      Top             =   120
      Width           =   4695
   End
   Begin VB.CommandButton Command2 
      Height          =   390
      Left            =   9960
      Picture         =   "Form2.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Archivar"
      Height          =   510
      Left            =   6480
      TabIndex        =   32
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Frame Frame4 
      Height          =   1215
      Left            =   6480
      TabIndex        =   23
      Top             =   5640
      Width           =   4095
      Begin VB.TextBox Text14 
         Height          =   390
         Left            =   2160
         TabIndex        =   31
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text13 
         Height          =   390
         Left            =   120
         TabIndex        =   30
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label14 
         Caption         =   "Porcentaje %:"
         Height          =   255
         Left            =   2160
         TabIndex        =   29
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label13 
         Caption         =   "Num Obra:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos generales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   6135
      Begin VB.TextBox Text8 
         Height          =   390
         Left            =   120
         TabIndex        =   18
         Top             =   4920
         Width           =   5775
      End
      Begin VB.TextBox Text7 
         Height          =   390
         Left            =   120
         TabIndex        =   17
         Top             =   4080
         Width           =   5775
      End
      Begin VB.TextBox Text6 
         Height          =   390
         Left            =   120
         TabIndex        =   16
         Top             =   3240
         Width           =   5775
      End
      Begin VB.TextBox Text5 
         Height          =   390
         Left            =   120
         TabIndex        =   15
         Top             =   2400
         Width           =   5775
      End
      Begin VB.TextBox Text4 
         Height          =   390
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   5775
      End
      Begin VB.TextBox Text3 
         Height          =   390
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   5775
      End
      Begin VB.Label Label8 
         Caption         =   "Numero de seguridad social:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   4560
         Width           =   5655
      End
      Begin VB.Label Label3 
         Caption         =   "RFC:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "CURP:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Apellido peterno:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Apellido materno:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   3720
         Width           =   2775
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   6135
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   3240
         TabIndex        =   11
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Height          =   390
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha baja:"
         Height          =   375
         Left            =   3240
         TabIndex        =   9
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha ingreso:"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ingresos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   6480
      TabIndex        =   0
      Top             =   720
      Width           =   4095
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
         Height          =   390
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   4080
         Width           =   3735
      End
      Begin VB.TextBox Text12 
         Alignment       =   1  'Right Justify
         Height          =   390
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   3240
         Width           =   3735
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         Height          =   390
         Left            =   120
         TabIndex        =   26
         Text            =   "0"
         Top             =   2400
         Width           =   3735
      End
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         Height          =   390
         Left            =   120
         TabIndex        =   25
         Text            =   "0"
         Top             =   1560
         Width           =   3735
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         Height          =   405
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label Label15 
         Caption         =   "Impuesto mensual ( ISR ):"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   3720
         Width           =   3615
      End
      Begin VB.Label Label12 
         Caption         =   "Salario diario integrado:"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   2880
         Width           =   2775
      End
      Begin VB.Label Label11 
         Caption         =   "Otros diarios fijos:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Label Label10 
         Caption         =   "Viaticos diarios:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label9 
         Caption         =   "Ingreso diario normal:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Label LabelId 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Caption         =   "ID actual:"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   35
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label LabelAnterior 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6360
      TabIndex        =   34
      Top             =   120
      Width           =   3495
   End
   Begin VB.Menu cedulaFiscal 
      Caption         =   "Cedula fiscal"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim largoPersonal As Integer
Dim largoOtros As Integer
Dim largoMaestro As Integer

Dim camposValidosGlobal As Boolean
Dim empleadoDuplicado As Boolean
Dim algunCampoVacioGlobal As Boolean
Dim ModoEdicion As Boolean
Private EsNuevo As Boolean
Private mIDFijo As Long
Private mRutaFija As String
Private Const PAT_LETRAS As String = "[A-Za-zÁÉÍÓÚÜÑáéíóúüñ ]"
Private Const PAT_ALFANUM As String = "[A-Za-z0-9Ññ&]"

' ============================================================
' CARGA / MODOS DE TRABAJO
' ============================================================

Public Sub Form_Load()
    Text1.MaxLength = 10
    Text2.MaxLength = 10
    Text3.MaxLength = 13
    Text4.MaxLength = 18
    Text8.MaxLength = 11
    Text5.MaxLength = Len(personal.nom)
    Text6.MaxLength = Len(personal.ape1)
    Text7.MaxLength = Len(personal.ape2)
End Sub

Public Sub IniciarNuevo()
    EsNuevo = True
    ModoEdicion = True          'En captura nueva no hay que habilitar edición.

    AbrirArchivosPersistentes
    LimpiarCampos
    
    

    Text16.Text = (largoPersonal + 1)
    mIDFijo = CLng(Text16.Text)
    mRutaFija = CurDir$

    If largoPersonal <= 0 Then
        MsgBox ("Achis, no hay personal. Intenta con otra carpeta!")
        Get 2, 1, personal
    Else
        Get 2, largoPersonal, personal
        LabelAnterior.Caption = ("ID anterior:  " & largoPersonal)
    End If

    LabelAnterior.Visible = True
    ActualizarBotones
End Sub

Public Sub IniciarConsulta(ByVal id As Integer)
    EsNuevo = False
    ModoEdicion = False

    AbrirArchivosPersistentes
    cargarEmpleado id
    mIDFijo = CLng(id)
    mRutaFija = CurDir$

    LabelAnterior.Visible = False
    ActualizarBotones
End Sub

' ============================================================
' ARCHIVOS / RESPALDOS
' ============================================================

Private Sub AbrirArchivosPersistentes()
    Close 2, 3, 8

    Open "personal.dno" For Random As 2 Len = Len(personal)
    largoPersonal = LOF(2) / Len(personal)

    Open "PerOtre.dno" For Random As 3 Len = Len(Otros_Rgtros)
    largoOtros = LOF(3) \ Len(Otros_Rgtros)

    Open "maestro.dno" For Random As 8 Len = Len(maestro)
    largoMaestro = LOF(8) / Len(maestro)
End Sub

Private Sub RespaldarArchivos()
    Close 2, 3, 4, 8

    FileCopy "personal.dno", "personal.bak"
    FileCopy "PerOtre.dno", "PerOtre.bak"

    If Dir("maestro.dno") <> "" Then FileCopy "maestro.dno", "maestro.bak"
    If Dir("Bnxcla.dno") <> "" Then FileCopy "Bnxcla.dno", "Bnxcla.bak"
End Sub

Private Sub RestaurarArchivos()
    Close 2, 3, 4, 8

    FileCopy "personal.bak", "personal.dno"
    FileCopy "PerOtre.bak", "PerOtre.dno"

    If Dir("maestro.bak") <> "" Then FileCopy "maestro.bak", "maestro.dno"
    If Dir("Bnxcla.bak") <> "" Then FileCopy "Bnxcla.bak", "Bnxcla.dno"
End Sub

' ============================================================
' ESTADO DE BOTONES / CAMPOS
' ============================================================

Private Sub ActualizarBotones()
    If EsNuevo Then
        Command1.Visible = True
        Command3.Visible = True
        Command2.Visible = False

        BtnHabilitarEdicion.Visible = False
        BtnCancelarCambios.Visible = False

        Command1.Enabled = True
        Command3.Enabled = True
    Else
        Command1.Visible = False
        Command2.Visible = True
        Command3.Visible = True

        BtnHabilitarEdicion.Visible = True
        BtnCancelarCambios.Visible = True

        BtnHabilitarEdicion.Enabled = Not ModoEdicion
        BtnCancelarCambios.Enabled = ModoEdicion
        Command2.Enabled = ModoEdicion
        Command3.Enabled = ModoEdicion
    End If

    BloquearCampos Not ModoEdicion
End Sub

Private Sub BloquearCampos(ByVal bloquear As Boolean)
    Dim i As Integer

    For i = 1 To 15
        Me.Controls("Text" & i).Locked = bloquear
    Next i

    'Text16 (ID) se deja libre: con Enter sirve para navegar a otro empleado.
End Sub

Private Sub BtnHabilitarEdicion_Click()
    RespaldarArchivos
    AbrirArchivosPersistentes

    ModoEdicion = True
    ActualizarBotones
End Sub


Private Sub BtnCancelarCambios_Click()

    If MsgBox("Desea descartar los cambios sin guardar?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Cancelar edicion") <> vbYes Then

        Exit Sub

    End If

    If EsNuevo Then
        Unload Me
        Exit Sub
    End If

    If Not RutaTrabajoSegura() Then Exit Sub

    If mIDFijo < 1 Or mIDFijo > 32767 Then
        MsgBox "El ID seleccionado no es valido.", vbCritical, "Cancelar cambios"
        Exit Sub
    End If

    If StrComp(CurDir$, mRutaFija, vbTextCompare) <> 0 Then
        MsgBox "La carpeta de trabajo cambio.", vbCritical, "Cancelar cambios"
        Exit Sub
    End If

    ModoEdicion = False

    'Recuperar los valores actuales del mismo empleado.
    'Nunca restaurar archivos .bak completos.
    cargarEmpleado CInt(mIDFijo)

    ActualizarBotones

End Sub


' ============================================================
' EVENTOS DE TEXTBOX
' ============================================================

Private Sub Text1_Change()
    calcularSalarioDiario
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text3.SetFocus
        Exit Sub
    End If
    FiltrarTecla KeyAscii, "FECHA"
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then FiltrarTecla KeyAscii, "FECHA"
    If EsNuevo Then Exit Sub

    If KeyAscii = 13 And Text2.Text <> "" Then
        EliminarNumeroTarjeta
    End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    FiltrarTecla KeyAscii, "ALFANUM"
End Sub

Private Sub Text3_Change()

    If Not EsNuevo Then Exit Sub

    If Len(Text3.Text) = 13 And Len(Text4.Text) <= 10 Then
        Text4.Text = Left$(Text3.Text, 10)
    End If

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    FiltrarTecla KeyAscii, "ALFANUM"
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    FiltrarTecla KeyAscii, "LETRAS"
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    FiltrarTecla KeyAscii, "LETRAS"
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
    FiltrarTecla KeyAscii, "LETRAS"
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
    FiltrarTecla KeyAscii, "NUM"
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
    FiltrarTecla KeyAscii, "DEC", Text9.Text
End Sub

Private Sub Text9_Change()
    calcularSalarioDiario
    calcularImpuestoMensual
End Sub

Private Sub Text10_Change()
    calcularSalarioDiario
    calcularImpuestoMensual
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
    FiltrarTecla KeyAscii, "DEC", Text10.Text
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
    FiltrarTecla KeyAscii, "DEC", Text11.Text
End Sub


Private Sub Text11_Change()
    calcularSalarioDiario
    calcularImpuestoMensual
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
    FiltrarTecla KeyAscii, "NUM"
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
    FiltrarTecla KeyAscii, "DEC", Text14.Text
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Text16.Text <> "" Then
        cargarEmpleado Val(Text16.Text)
    End If
End Sub

' ============================================================
' EVENTOS DE BOTONES
' ============================================================



Private Sub cedulaFiscal_Click()

    Dim sID As String
    Dim idEmpleado As Long

    sID = Trim$(Me.Text16.Text)

    If sID = "" Or sID Like "*[!0-9]*" Then
        MsgBox "Seleccione un empleado con ID valido.", _
               vbExclamation, "Cedula Fiscal"
        Exit Sub
    End If

    If Len(sID) > 5 Then
        MsgBox "ID fuera de rango.", _
               vbExclamation, "Cedula Fiscal"
        Exit Sub
    End If

    idEmpleado = CLng(sID)

    If idEmpleado < 1 Or idEmpleado > 32767 Then
        MsgBox "ID fuera de rango (1 a 32767).", _
               vbExclamation, "Cedula Fiscal"
        Exit Sub
    End If

    Load FormCedulaFiscal

    If Not FormCedulaFiscal.PrepararEmpleado(idEmpleado) Then
        Unload FormCedulaFiscal
        Exit Sub
    End If

    FormCedulaFiscal.Show vbModal, Me

End Sub


Private Sub Command1_Click()
    Dim sinCamposVacios As Boolean
    Dim camposValidos As Boolean
    Dim noDuplicado As Boolean

    If Not ModoEdicion Then
        MsgBox "Debes habilitar edición primero.", vbExclamation
        Exit Sub
    End If

    verificarCamposVacios
    sinCamposVacios = Not algunCampoVacioGlobal

    ValidarCampos
    camposValidos = camposValidosGlobal

    verificarEmpleadoDuplicado
    noDuplicado = Not empleadoDuplicado

    
    If camposValidos And noDuplicado And sinCamposVacios Then
    
        If Not RutaTrabajoSegura() Then Exit Sub
    
        If MsgBox( _
            "Confirma guardar al empleado con ID " & _
            Trim$(Text16.Text) & "?" & vbCrLf & _
            "Entorno: " & EtiquetaEntornoNomina(), _
            vbYesNo Or vbQuestion Or vbDefaultButton2, _
            "Confirmar alta de empleado") = vbYes Then
    
            guardarRegistro
    
        End If

    ElseIf Not sinCamposVacios Then
        MsgBox "Faltan campos por llenar (marcados en rojo).", vbExclamation
    End If
End Sub

Private Sub Command2_Click()
    eliminarRegistro
End Sub

Private Sub Command3_Click()
    If EsNuevo Then
        Command1_Click          'Mismas validaciones: vacíos, RFC/CURP y duplicado.
    Else
        actualizarRegistro
    End If
End Sub

' ============================================================
' CARGA / LIMPIEZA DE DATOS EN PANTALLA
' ============================================================

Public Sub cargarEmpleado(id As Integer)


    'Validar ID y longitud fisica de los archivos.
    If id < 1 Then
        Err.Raise vbObjectError + 4301, , "ID de empleado no valido."
    End If
    
    If LOF(2) Mod Len(personal) <> 0 Or LOF(3) Mod Len(Otros_Rgtros) <> 0 Or LOF(8) Mod Len(maestro) <> 0 Then
    
        Err.Raise vbObjectError + 4302, , "Hay archivos de empleados desalineados."
    End If
    
    If id > LOF(2) \ Len(personal) Or id > LOF(3) \ Len(Otros_Rgtros) Or id > LOF(8) \ Len(maestro) Then
    
        Err.Raise vbObjectError + 4303, , "El ID no existe en todos los archivos requeridos."
    End If

    Get 2, id, personal
    Get 3, id, Otros_Rgtros
    Get 8, id, maestro

    Text16.Text = id

    Text1.Text = Trim(personal.fal)                    'Fecha alta.
    Text2.Text = Trim(personal.fab)                    'Fecha baja.
    Text3.Text = Trim(personal.RFC)                    'RFC.
    Text4.Text = Trim(Otros_Rgtros.curp)               'CURP.
    Text5.Text = Trim(personal.nom)                    'Nombre.
    Text6.Text = Trim(personal.ape1)                   'Apellido paterno.
    Text7.Text = Trim(personal.ape2)                   'Apellido materno.
    Text8.Text = Replace(Trim(personal.imss), "-", "") 'Número SS.
    Text9.Text = personal.ingr                         'Ingreso diario normal.
    Text10.Text = personal.viat                        'O.F. diarios.
    Text11.Text = personal.otras                       'Otros ingresos.
    Text12.Text = personal.integrado                   'Salario diario integrado.
    Text13.Text = maestro.O_1                          'Número obra.
    Text14.Text = maestro.por_1                        'Porcentaje.
    
    mIDFijo = CLng(id)
    mRutaFija = CurDir$
    
End Sub

Private Sub LimpiarCampos()
    Text1.Text = Empty
    Text2.Text = Empty
    Text3.Text = Empty
    Text4.Text = Empty
    Text5.Text = Empty
    Text6.Text = Empty
    Text7.Text = Empty
    Text8.Text = Empty
    Text9.Text = Empty
    Text10.Text = 0
    Text11.Text = 0
    Text12.Text = Empty
    Text13.Text = Empty
    Text14.Text = Empty
    Text15.Text = Empty
End Sub

' ============================================================
' CALCULOS
' ============================================================

Sub calcularSalarioDiario()
    Dim antiguedad As Integer
    Dim facto As Double
    Dim totalIngr As Currency

    antiguedad = CalcularAntiguedad(Text1.Text)
    facto = 0
    totalIngr = 0

    If IsNumeric(Text9.Text) Then totalIngr = totalIngr + Text9.Text
    If IsNumeric(Text10.Text) Then totalIngr = totalIngr + Text10.Text
    If IsNumeric(Text11.Text) Then totalIngr = totalIngr + Text11.Text

    factor antiguedad, facto

    Text12.Text = Format((totalIngr * facto), z1$)

    If IsNumeric(Text12.Text) Then
        If Text12.Text > (25 * empresa.sm) Then
            Text12.Text = Format((25 * empresa.sm), z1$)
        End If
    End If
End Sub

Private Sub calcularImpuestoMensual()
    Dim ingresoTotal As Currency
    Dim largoArticulo As Double
    Dim salarioDiario As Currency
    Dim viaticosDiarios As Currency
    Dim otrosDiarios As Currency
    Dim impuesto As Currency
    Dim i As Integer

    On Error GoTo ErrorCalculo

    salarioDiario = 0
    viaticosDiarios = 0
    otrosDiarios = 0

    If IsNumeric(Text9.Text) Then salarioDiario = CCur(Text9.Text)
    If IsNumeric(Text10.Text) Then viaticosDiarios = CCur(Text10.Text)
    If IsNumeric(Text11.Text) Then otrosDiarios = CCur(Text11.Text)

    ingresoTotal = salarioDiario + viaticosDiarios + otrosDiarios
    ingresoTotal = ingresoTotal * 30

    Close 99
    Open (Trim(Dir_imptos) + "Tab08Mes.ISR") For Random As 99 Len = Len(articulo)

    largoArticulo = LOF(99) / Len(articulo)

    For i = 1 To largoArticulo
        Get 99, i, articulo

        impuesto = ingresoTotal

        If ingresoTotal >= articulo.liminf And ingresoTotal <= articulo.limsup Then
            impuesto = (((impuesto - articulo.liminf) * articulo.porcsl / 100) + articulo.cuotaf)
            Text15.Text = Format(impuesto, z1$)
        End If
    Next i

ErrorCalculo:
End Sub

' ============================================================
' VALIDACIONES
' ============================================================

Private Sub verificarCamposVacios()
    Dim i As Integer
    Dim txt As TextBox
    Dim lbl As Label
    Dim algunCampoVacio As Boolean

    algunCampoVacio = False

    For i = 1 To 14
        If i = 2 Then GoTo SiguienteIteracion

        On Error Resume Next
        Set txt = Me.Controls("Text" & i)
        Set lbl = Me.Controls("Label" & i)
        On Error GoTo 0

        If Not txt Is Nothing And Not lbl Is Nothing Then
            If txt.Text = "" Then
                lbl.ForeColor = vbRed
                algunCampoVacio = True
            Else
                lbl.ForeColor = vbBlack
            End If
        Else
            MsgBox "Control " & "Text" & i & " o " & "Label" & i & " no existe."
        End If

        Set txt = Nothing
        Set lbl = Nothing

SiguienteIteracion:
    Next i

    algunCampoVacioGlobal = algunCampoVacio
End Sub

Private Sub ValidarCampos()
    Dim errores As String

    If Not IsValidDate(Text1.Text) Then errores = errores & "- Fecha de alta no válida (dd/mm/aaaa)." & vbCrLf

    If Trim(Text2.Text) <> "" Then
        If Not IsValidDate(Text2.Text) Then errores = errores & "- Fecha de baja no válida (dd/mm/aaaa)." & vbCrLf
    End If

   
    '============================================================
    ' VALIDACION SIMPLE DE RFC
    '============================================================
    
    If Len(Trim$(Text3.Text)) <> 13 Then
        errores = errores & "- El RFC debe contener exactamente 13 caracteres." & vbCrLf
    End If
    
    '============================================================
    ' VALIDACION SIMPLE DE CURP
    '============================================================
    
    If Len(Trim$(Text4.Text)) <> 18 Then
        errores = errores & "- La CURP debe contener exactamente 18 caracteres." & vbCrLf
    End If


    If Not SoloCaracteres(Text5.Text, PAT_LETRAS) Then errores = errores & "- El nombre solo puede llevar letras." & vbCrLf
    If Not SoloCaracteres(Text6.Text, PAT_LETRAS) Then errores = errores & "- El apellido paterno solo puede llevar letras." & vbCrLf
    If Not SoloCaracteres(Text7.Text, PAT_LETRAS) Then errores = errores & "- El apellido materno solo puede llevar letras." & vbCrLf

    If Not SoloCaracteres(Text8.Text, "[0-9]") Then errores = errores & "- El número de seguro social solo puede llevar dígitos." & vbCrLf

    If Not IsNumeric(Text9.Text) Then errores = errores & "- Ingreso diario normal no es un número." & vbCrLf
    If Not IsNumeric(Text10.Text) Then errores = errores & "- O.F. diarios no es un número." & vbCrLf
    If Not IsNumeric(Text11.Text) Then errores = errores & "- Otros ingresos no es un número." & vbCrLf

    camposValidosGlobal = (errores = "")

    If errores <> "" Then
        MsgBox "Corrige lo siguiente:" & vbCrLf & vbCrLf & errores, vbExclamation
    End If
End Sub

Private Sub verificarEmpleadoDuplicado()
    Dim i As Integer
    Dim empleadoExistente As Boolean

    empleadoExistente = False

    For i = 1 To largoPersonal
        Get 2, i, personal

        If Text3.Text = Trim(personal.RFC) Then
            empleadoExistente = True
            MsgBox "Este empleado ya existe."
            Exit For
        End If
    Next i

    empleadoDuplicado = empleadoExistente
End Sub


Private Function IsValidDate(ByVal fecha As String) As Boolean

    Dim dia As Integer
    Dim mes As Integer
    Dim anio As Integer
    Dim fechaValidada As Date

    On Error GoTo FechaIncorrecta

    IsValidDate = False

    fecha = Trim$(fecha)

    'Formato exacto: dd/mm/aaaa
    If Len(fecha) <> 10 Then Exit Function

    If mID$(fecha, 3, 1) <> "/" Then Exit Function
    If mID$(fecha, 6, 1) <> "/" Then Exit Function

    'Verificar los componentes numericos.
    If Left$(fecha, 2) Like "*[!0-9]*" Then Exit Function
    If mID$(fecha, 4, 2) Like "*[!0-9]*" Then Exit Function
    If Right$(fecha, 4) Like "*[!0-9]*" Then Exit Function

    dia = CInt(Left$(fecha, 2))
    mes = CInt(mID$(fecha, 4, 2))
    anio = CInt(Right$(fecha, 4))

    If dia < 1 Or dia > 31 Then Exit Function
    If mes < 1 Or mes > 12 Then Exit Function
    If anio < 100 Or anio > 9999 Then Exit Function

    fechaValidada = DateSerial(anio, mes, dia)

    'DateSerial puede ajustar fechas inexistentes.
    'Comprobamos que no haya modificado dia, mes o anio.
    If Day(fechaValidada) <> dia Then Exit Function
    If Month(fechaValidada) <> mes Then Exit Function
    If Year(fechaValidada) <> anio Then Exit Function

    IsValidDate = True
    Exit Function

FechaIncorrecta:

    IsValidDate = False

End Function


Private Function IsValidCURP(ByVal curp As String) As Boolean
    Dim regex As Object

    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^[A-Z]{4}[0-9]{6}[HM][A-Z]{5}[A-Z0-9][0-9]$"

    IsValidCURP = regex.Test(UCase$(Trim$(curp)))
End Function

' ============================================================
' GUARDAR / ACTUALIZAR / ELIMINAR
' ============================================================

Private Sub actualizarRegistro()
    Dim respuesta As VbMsgBoxResult

    respuesta = MsgBox("¿Desea actualizar los registros?", vbYesNo + vbQuestion, "Confirmar actualización")

    If respuesta = vbNo Then
        MsgBox "Operación cancelada.", vbInformation
        Exit Sub
    End If

    guardarRegistro
End Sub


Private Sub guardarRegistro()

    Dim id As Long
    Dim totalPersonal As Long
    Dim totalOtros As Long
    Dim totalMaestro As Long

    Dim personalVacio As per
    Dim otrosVacios As OtrasCh
    Dim maestroVacio As ob

    If Not RutaTrabajoSegura() Then Exit Sub

    'No utilizar directamente un ID editable.
    If mIDFijo < 1 Or mIDFijo > 32767 Then
        MsgBox "No hay un ID de empleado valido.", vbCritical, "Guardar personal"
        Exit Sub
    End If

    If Trim$(Text16.Text) <> CStr(mIDFijo) Then
        MsgBox "El ID en pantalla no coincide con el empleado seleccionado." & vbCrLf & "No se guardaron cambios.", vbCritical, "Guardar personal"
        Exit Sub
    End If

    If StrComp(CurDir$, mRutaFija, vbTextCompare) <> 0 Then
        MsgBox "La carpeta de trabajo cambio." & vbCrLf & "No se guardaron datos.", vbCritical, "Guardar personal"
        Exit Sub
    End If

    id = mIDFijo

    On Error GoTo ErrorValidacion

    'Comprobar que los manejadores estan abiertos
    'y que las estructuras son compatibles.
    If LOF(2) Mod Len(personal) <> 0 Then
        Err.Raise vbObjectError + 4101, , "personal.dno esta desalineado."
    End If

    If LOF(3) Mod Len(Otros_Rgtros) <> 0 Then
        Err.Raise vbObjectError + 4102, , "PerOtre.dno esta desalineado."
    End If

    If LOF(8) Mod Len(maestro) <> 0 Then
        Err.Raise vbObjectError + 4103, , "maestro.dno esta desalineado."
    End If

    totalPersonal = LOF(2) \ Len(personal)
    totalOtros = LOF(3) \ Len(Otros_Rgtros)
    totalMaestro = LOF(8) \ Len(maestro)

    If EsNuevo Then

        'La nueva alta solo puede agregarse despues
        'del ultimo registro de personal.
        If id <> totalPersonal + 1 Then
            Err.Raise vbObjectError + 4104, , "La posicion del nuevo empleado cambio."
        End If

        'No sobrescribir posiciones auxiliares existentes.
        If id <= totalOtros Or id <= totalMaestro Then
            Err.Raise vbObjectError + 4105, , "La posicion nueva ya contiene un registro en un archivo auxiliar. Requiere revision."
        End If

        'Para un empleado nuevo, iniciar registros vacios.
        'Nunca arrastrar buffers del empleado anterior.
        personal = personalVacio
        Otros_Rgtros = otrosVacios
        maestro = maestroVacio

    Else

        'Un empleado existente debe encontrarse en
        'los tres archivos antes de cualquier escritura.
        If id > totalPersonal Or _
           id > totalOtros Or _
           id > totalMaestro Then

            Err.Raise vbObjectError + 4106, , "El empleado no tiene registros alineados en los tres archivos."
        End If

        Get #2, id, personal
        Get #3, id, Otros_Rgtros
        Get #8, id, maestro

    End If

    GoTo DatosValidados

ErrorValidacion:

    MsgBox "Guardado bloqueado:" & vbCrLf & Err.Description, vbCritical, "Proteccion de archivos"
    Exit Sub

DatosValidados:

    personal.fal = Text1.Text
    personal.fab = Text2.Text
    personal.RFC = Text3.Text
    personal.nom = Text5.Text
    personal.ape1 = Text6.Text
    personal.ape2 = Text7.Text
    personal.imss = Text8.Text

    If IsNumeric(Text9.Text) Then personal.ingr = CCur(Text9.Text)
    If IsNumeric(Text10.Text) Then personal.viat = CCur(Text10.Text)
    If IsNumeric(Text11.Text) Then personal.otras = CCur(Text11.Text)
    If IsNumeric(Text12.Text) Then personal.integrado = CCur(Text12.Text)

    Otros_Rgtros.curp = Text4.Text

    maestro.O_1 = Text13.Text

    If IsNumeric(Text14.Text) Then maestro.por_1 = CCur(Text14.Text)

    Put 2, id, personal
    Put 3, id, Otros_Rgtros
    Put 8, id, maestro

    MsgBox "Se guardó con éxito.", vbInformation

    Unload Me
End Sub


Private Sub eliminarRegistro()

    MsgBox "La eliminacion fisica de empleados " & "esta deshabilitada para proteger sus IDs." & vbCrLf & vbCrLf & "Para registrar una baja, utilice " & "la fecha de baja del empleado.", vbExclamation, "Proteccion de personal"

End Sub


Private Sub EliminarNumeroTarjeta()
    Dim respuesta As Integer
    Dim respuestaDos As Integer
    Dim idNomina As String

    If Not ModoEdicion Then
        MsgBox "Debes habilitar edición primero.", vbExclamation
        Exit Sub
    End If

    Close 4
    Open "Bnxcla.dno" For Random As #4 Len = Len(Clbnx)

    idNomina = Text16.Text

    Get #4, idNomina, Clbnx

    respuesta = MsgBox("¿Se eliminará el número de tarjeta: " & Clbnx.Q1 & "?" & vbCrLf & vbCrLf & "¿Deseas continuar?", vbOKCancel + vbQuestion)

    If respuesta = vbOK Then
        respuestaDos = MsgBox("Esta acción es irreversible. " & "Asegúrate de tener una copia del número de tarjeta.", vbOKCancel + vbExclamation)

        If respuestaDos = vbOK Then
            Clbnx.Q1 = "0"
            Put #4, idNomina, Clbnx
            MsgBox "Se eliminó el número de tarjeta.", vbInformation
        Else
            MsgBox "Operación cancelada.", vbInformation
        End If
    Else
        MsgBox "Operación cancelada.", vbInformation
    End If

    Close #4
End Sub

Private Function SoloCaracteres(ByVal txt As String, ByVal patron As String) As Boolean
    Dim i As Integer

    SoloCaracteres = True
    For i = 1 To Len(txt)
        If Not (mID$(txt, i, 1) Like patron) Then
            SoloCaracteres = False
            Exit Function
        End If
    Next i
End Function

Private Sub FiltrarTecla(ByRef KeyAscii As Integer, ByVal tipo As String, Optional ByVal textoActual As String = "")
    Dim c As String

    If KeyAscii < 32 Then Exit Sub          'Backspace y teclas de control.

    c = Chr$(KeyAscii)

    Select Case tipo
        Case "LETRAS"
            If Not (c Like PAT_LETRAS) Then KeyAscii = 0
        Case "ALFANUM"
            If Not (c Like PAT_ALFANUM) Then KeyAscii = 0
        Case "NUM"
            If Not (c Like "[0-9]") Then KeyAscii = 0
        Case "DEC"
            If c = "." Then
                If InStr(textoActual, ".") > 0 Then KeyAscii = 0
            ElseIf Not (c Like "[0-9]") Then
                KeyAscii = 0
            End If
        Case "FECHA"
            If Not (c Like "[0-9/]") Then KeyAscii = 0
    End Select

    If KeyAscii <> 0 And (tipo = "LETRAS" Or tipo = "ALFANUM") Then
        KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    End If
End Sub
