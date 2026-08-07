VERSION 5.00
Begin VB.Form Form2 
   ClientHeight    =   7890
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
   ScaleHeight     =   7890
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Actualizar"
      Height          =   495
      Left            =   8640
      TabIndex        =   39
      Top             =   7200
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
      Left            =   10080
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
      Top             =   7200
      Width           =   1935
   End
   Begin VB.Frame Frame4 
      Height          =   1215
      Left            =   6480
      TabIndex        =   23
      Top             =   5760
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
      Left            =   6480
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

Public Sub Form_Load()
    Close 2, 3, 8, 15
    
    Open "personal.dno" For Random As 2 Len = Len(personal): largoPersonal = LOF(2) / Len(personal)
    Open "PerOtre.dno" For Random As 3 Len = Len(Otros_Rgtros): largoOtros = LOF(3) / Len(personal)
    Open "maestro.dno" For Random As 8 Len = Len(maestro): largoMaestro = LOF(8) / Len(maestro)
    
    If largoPersonal <= 0 Then
        MsgBox ("Achis, no hay personal. Intenta con otra carpeta!")
        Get 2, 1, personal
    Else
        Get 2, largoPersonal, personal
    
        Text16.Text = (largoPersonal + 1)
        LabelAnterior.Caption = ("ID anterior:  " & largoPersonal)
        
    End If
End Sub

Private Sub Text1_Change()
    calcularSalarioDiario
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text3.SetFocus
    End If
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Text16.Text <> "" Then
        cargarEmpleado Val(Text16.Text)
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text2.Text <> "" Then
    Dim respuesta As Integer
    Dim respuestaDos As Integer
    Dim idNomina As String
    
    Close 4
    Open "Bnxcla.dno" For Random As 4 Len = Len(Clbnx)
       
    idNomina = Text16.Text
    
    Get 4, idNomina, Clbnx
    
    respuesta = MsgBox("Se eliminar? el n?mero de tarjeta: " + Clbnx.Q1 + " , ?Deseas continuar?", vbOKCancel)
    
    If respuesta = vbOK Then
        ' C?digo para eliminar los datos bancarios
        respuestaDos = MsgBox("Esta acci?n es irreversible asegurate de tener un copia del n?mero de tarjeta.", vbOKCancel)
        
        If respuestaDos = vbOK Then

            Clbnx.Q1 = 0
            
            Put 4, idNomina, Clbnx
            
            MsgBox ("Se elimino el n?mero de tarjeta.")
            
        ElseIf respuestaDos = vbCancel Then
            MsgBox "Operaci?n cancelada"
        End If

    ElseIf respuesta = vbCancel Then

        MsgBox "Operaci?n cancelada."
    End If

End If

End Sub

Private Sub Text3_Change()
    ' Obtener el texto de Text1
    Dim texto As String
    texto = Text3.Text
    
    ' Verificar que el texto tiene m?s de 3 caracteres
    If Len(texto) > 3 Then
        ' Copiar el texto en Text2 menos los ?ltimos 3 caracteres
        Text4.Text = Left(texto, Len(texto) - 3)
    Else
        ' Si el texto tiene 3 o menos caracteres, Text2 ser? vac?o
        Text4.Text = ""
    End If
    
End Sub

Private Sub Text9_Change()
        calcularSalarioDiario
        calcularImpuestoMensual
End Sub

Private Sub Text10_Change()
        calcularSalarioDiario
        calcularImpuestoMensual
End Sub

Private Sub Text11_Change()
        calcularSalarioDiario
        calcularImpuestoMensual
End Sub


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

Private Sub cedulaFiscal_Click()
    FormCedulaFiscal.Show 1
End Sub

Private Sub Command2_Click()
    eliminarRegistro
End Sub


Private Sub eliminarRegistro()
    Dim respuesta As VbMsgBoxResult
    Dim registro As Integer
    Dim i As Integer
    Dim personalList() As per
    Dim otrosList() As OtrasCh
    Dim maestroList() As ob
    Dim countPersonal As Integer
    Dim countOtros As Integer
    Dim countMaestro As Integer
    
    ' Pregunta al usuario si desea eliminar el registro
    respuesta = MsgBox("?Desea eliminar el registro?", vbYesNo + vbQuestion, "Confirmar eliminaci?n")

    If respuesta = vbNo Then
        MsgBox "Operaci?n cancelada.", vbInformation
        Exit Sub
    End If

    registro = Val(Text16.Text)
    
    countPersonal = LOF(2) / Len(personal)
    countOtros = LOF(3) / Len(Otros_Rgtros)
    countMaestro = LOF(8) / Len(maestro)
    
    ReDim personalList(1 To countPersonal)
    ReDim otrosList(1 To countOtros)
    ReDim maestroList(1 To countMaestro)
    
    For i = 1 To countPersonal
        Get 2, i, personalList(i)
    Next i
    For i = 1 To countOtros
        Get 3, i, otrosList(i)
    Next i
    For i = 1 To countMaestro
        Get 8, i, maestroList(i)
    Next i
    
    ' Cerrar archivos
    Close 2
    Close 3
    Close 8
    
    ' Eliminar el registro de las listas
    If registro > 0 And registro <= countPersonal Then
        For i = registro To countPersonal - 1
            personalList(i) = personalList(i + 1)
        Next i
        ReDim Preserve personalList(1 To countPersonal - 1)
    End If
    
    If registro > 0 And registro <= countOtros Then
        For i = registro To countOtros - 1
            otrosList(i) = otrosList(i + 1)
        Next i
        ReDim Preserve otrosList(1 To countOtros - 1)
    End If
    
    If registro > 0 And registro <= countMaestro Then
        For i = registro To countMaestro - 1
            maestroList(i) = maestroList(i + 1)
        Next i
        ReDim Preserve maestroList(1 To countMaestro - 1)
    End If
    
    ' Reescribir los registros en los archivos
    Open "personal.dno" For Output As #2
    Open "PerOtre.dno" For Output As #3
    Open "maestro.dno" For Output As #8
    
    Close #2
    Close #3
    Close #8
    
    Open "personal.dno" For Random As 2 Len = Len(personal)
    Open "PerOtre.dno" For Random As 3 Len = Len(Otros_Rgtros)
    Open "maestro.dno" For Random As 8 Len = Len(maestro)
    
    For i = 1 To UBound(personalList)
        Put 2, i, personalList(i)
    Next i
    For i = 1 To UBound(otrosList)
        Put 3, i, otrosList(i)
    Next i
    For i = 1 To UBound(maestroList)
        Put 8, i, maestroList(i)
    Next i
    
    ' Cerrar archivos
    Close 2
    Close 3
    Close 8
    
    MsgBox "Registro eliminado con ?xito.", vbInformation
    limpiarCampos
End Sub
Private Sub limpiarCampos()
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
    
    Me.Form_Load
End Sub

Private Sub calcularImpuestoMensual()
Dim ingresoTotal As Currency
Dim largoArticulo As Double

On Error GoTo Error:

    Dim salarioDiario As Currency: salarioDiario = 0
    Dim viaticosDiarios As Currency: viaticosDiarios = 0
    Dim otrosDiarios As Currency: otrosDiarios = 0
    
    If IsNumeric(Text9.Text) Then salarioDiario = CCur(Text9.Text)
    If IsNumeric(Text10.Text) Then viaticosDiarios = CCur(Text10.Text)
    If IsNumeric(Text11.Text) Then otrosDiarios = CCur(Text11.Text)

    ingresoTotal = salarioDiario + viaticosDiarios + otrosDiarios
    ingresoTotal = ingresoTotal * 30

    Close 99: Open (Trim(Dir_imptos) + "Tab08Mes.ISR") For Random As 99 Len = Len(articulo): largoArticulo = LOF(99) / Len(articulo)
    
    Dim i As Integer
    
    For i = 1 To largoArticulo
        Get 99, i, articulo
        Dim impuesto As Currency
            
            impuesto = ingresoTotal
        If ingresoTotal >= articulo.liminf And ingresoTotal <= articulo.limsup Then
                
            impuesto = (((impuesto - articulo.liminf) * articulo.porcsl / 100) + articulo.cuotaf)
            Text15.Text = Format(impuesto, z1$)
            
        End If
    Next i
Error:

End Sub

Public Sub cargarEmpleado(id As Integer)
    Get 2, id, personal
    Get 3, id, Otros_Rgtros
    Get 8, id, maestro
    
    Text16.Text = id
    
    ' FECHA ALTA
    Text1.Text = Trim(personal.fal)
    ' FECHA BAJA
    Text2.Text = Trim(personal.fab)
    ' RFC
    Text3.Text = Trim(personal.RFC)
    ' CURP
    Text4.Text = Trim(Otros_Rgtros.curp)
    ' NOMBRE
    Text5.Text = Trim(personal.nom)
    ' APELLIDO P
    Text6.Text = Trim(personal.ape1)
    ' APELLIDO M
    Text7.Text = Trim(personal.ape2)
    ' NUMERO SS
    Text8.Text = Replace(Trim(personal.imss), "-", "")
    ' INGRESO DIARIO NORMAL
    Text9.Text = personal.ingr
    ' O.F DIARIOS
    Text10.Text = personal.viat
    ' OTROS INGRESOS
    Text11.Text = personal.otras
    ' SALARIO DIARIO INTEGRADO
    Text12.Text = personal.integrado
    ' NUMERO OBRA
    Text13.Text = maestro.O_1
    ' PORCENTAJE
    Text14.Text = maestro.por_1
    
End Sub

Private Sub Command1_Click()
    Dim sinCamposVacios As Boolean
    Dim camposValidos As Boolean
    Dim noDuplicado As Boolean
    
    ' Verificar campos vac?os y establecer la bandera correspondiente
    verificarCamposVacios
    sinCamposVacios = Not algunCampoVacioGlobal
    
    ' Validar campos y establecer la bandera correspondiente
    ValidarCampos
    camposValidos = camposValidosGlobal
    
    ' Verificar empleado duplicado y establecer la bandera correspondiente
    verificarEmpleadoDuplicado
    noDuplicado = Not empleadoDuplicado
    
    ' Comprobar todas las banderas antes de guardar
    If camposValidos And noDuplicado And sinCamposVacios Then
        guardarRegistro
    Else
        MsgBox "Hay errores en los campos o el empleado ya existe. No se puede guardar."
    End If
End Sub


Private Sub verificarCamposVacios()
    Dim i As Integer
    Dim txt As TextBox
    Dim lbl As Label
    Dim algunCampoVacio As Boolean
    
    algunCampoVacio = False
    
    For i = 1 To 14
        If i = 2 Then
            GoTo SiguienteIteracion
        End If
        
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
    
    ' Asignar valor a la variable global
    algunCampoVacioGlobal = algunCampoVacio
End Sub

Private Sub ValidarCampos()
    Dim fechaValida As Boolean
    Dim fechaOpcionalValida As Boolean
    Dim rfcValido As Boolean
    Dim curpValido As Boolean
    
    fechaValida = IsValidDate(Text1.Text)
    
    If Trim(Text2.Text) = "" Then
        fechaOpcionalValida = True
    Else
        fechaOpcionalValida = IsValidDate(Text2.Text)
    End If
    
    rfcValido = (Len(Text3.Text) = 13)
    curpValido = IsValidCURP(Text4.Text)
    
    ' Asignar valor a la variable global
    camposValidosGlobal = fechaValida And fechaOpcionalValida And rfcValido And curpValido
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
    
    ' Asignar valor a la variable global
    empleadoDuplicado = empleadoExistente
End Sub


Private Sub Command3_Click()
    actualizarRegistro
End Sub

Private Sub actualizarRegistro()
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox("?Desea actualizar los registros?", vbYesNo + vbQuestion, "Confirmar actualizaci?n")

    If respuesta = vbNo Then
        MsgBox "Operaci?n cancelada.", vbInformation
        Exit Sub
    End If

    guardarRegistro

End Sub

Private Sub guardarRegistro()

Close 4
On Error GoTo manejador
    Dim registro As Integer
    registro = Val(Text16.Text)
       
    personal.fal = Text1.Text
    personal.fab = Text2.Text
    personal.RFC = Text3.Text
    Otros_Rgtros.curp = Text4.Text
    personal.nom = Text5.Text
    personal.ape1 = Text6.Text
    personal.ape2 = Text7.Text
    personal.imss = Replace(Text8.Text, "-", "")
    If IsNumeric(Text9.Text) Then personal.ingr = Text9.Text Else personal.ingr = 0
    If IsNumeric(Text10.Text) Then personal.viat = Text10.Text Else personal.viat = 0
    If IsNumeric(Text11.Text) Then personal.otras = Text11.Text Else personal.otras = 0
    If IsNumeric(Text12.Text) Then personal.integrado = Text12.Text Else personal.integrado = 0
    If IsNumeric(Text13.Text) Then maestro.O_1 = Text13.Text Else maestro.O_1 = 0
    If IsNumeric(Text14.Text) Then maestro.por_1 = Text14.Text Else maestro.por_1 = 0
     
    Put 2, registro, personal
    Put 3, registro, Otros_Rgtros
    Put 8, registro, maestro
     
    MsgBox "Se guardó con éxito."
    
    limpiarCampos
    Exit Sub
    
manejador:
    If Err.Number = 13 Then
        MsgBox "Error 13: Error de tipo." & vbCrLf & _
               "Por favor, asegúrese de que los campos numéricos (Sueldo Diario, Viáticos, Otras Percepciones, Número de Obra, Porcentaje) no estén vacíos y contengan únicamente números.", vbCritical, "Error de Validación de Datos"
    Else
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error al Guardar"
    End If
End Sub

Private Function IsValidDate(ByVal fecha As String) As Boolean
    Dim testDate As Date
    IsValidDate = False
    On Error Resume Next
    testDate = DateValue(fecha)
    On Error GoTo 0
    If Err.Number = 0 Then
        IsValidDate = True
    End If
End Function

Private Function IsValidCURP(ByVal curp As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Patr?n de expresi?n regular para validar CURP
    regex.Pattern = "^[A-Z]{4}[0-9]{6}[HM][A-Z]{5}[0-9]{2}$"
    
    ' Aplicar la expresi?n regular a la cadena CURP
    IsValidCURP = regex.Test(curp)
End Function

