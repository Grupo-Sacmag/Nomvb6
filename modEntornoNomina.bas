Attribute VB_Name = "modEntornoNomina"
Option Explicit

'============================================================
' CONFIGURACION EXCLUSIVA DEL DESARROLLADOR
' Cambiar esta unica linea ANTES de compilar:
'   True  = TESTING (rutas y perma_testing.dno)
'   False = PRODUCCION (perma.dno habitual)
' NO aparece ninguna pregunta al usuario final.
'============================================================
Public Const MODO_TESTING_DESARROLLADOR As Boolean = True
' False para PRODUCCIÓN, True para TESTING

'Contiene COPIAS independientes de las carpetas de empresa.
Public Const RAIZ_NOMINA_TESTING As String = "C:\NominaTesting"

'Conservar estas variables: las utilizan los cambios en Form1.
Public gModoTesting As Boolean
Public gEntornoElegido As Boolean
Public gRutaTrabajoAprobada As String

'Conserva el nombre utilizado en Form1, pero ya NO pregunta nada.
Public Sub ElegirEntornoNomina()
    gModoTesting = MODO_TESTING_DESARROLLADOR
    gEntornoElegido = True
    gRutaTrabajoAprobada = vbNullString
End Sub

'La configuracion de PRODUCCION permanece intacta.
'El archivo TESTING usa el mismo tipo basico (basini).
Public Function RutaArchivoEntornoNomina() As String
    If gModoTesting Then
        RutaArchivoEntornoNomina = "C:\GconTA\perma_testing.dno"
    Else
        RutaArchivoEntornoNomina = "C:\GconTA\perma.dno"
    End If
End Function


Private Function RutaSimple(ByVal ruta As String) As String

    ruta = UCase$(Trim$(Replace(ruta, vbNullChar, vbNullString)))

    ruta = Replace(ruta, "/", "\")

    Do While Len(ruta) > 3 And Right$(ruta, 1) = "\"

        ruta = Left$(ruta, Len(ruta) - 1)

    Loop

    RutaSimple = ruta

End Function


Public Function RutaBajoTesting(ByVal ruta As String) As Boolean
    Dim base As String
    Dim actual As String

    base = RutaSimple(RAIZ_NOMINA_TESTING)
    actual = RutaSimple(ruta)
    If Len(actual) = 0 Then Exit Function

    RutaBajoTesting = (actual = base Or _
                       Left$(actual, Len(base) + 1) = base & "\")
End Function


Public Function RutaPermitidaPorEntorno(ByVal ruta As String) As Boolean

    Dim atributos As Long

    On Error GoTo ErrorRuta

    RutaPermitidaPorEntorno = False

    If Not gEntornoElegido Then Exit Function

    ruta = Trim$(Replace(ruta, vbNullChar, ""))

    If Len(ruta) = 0 Then Exit Function

    'Aceptar cualquier directorio existente y accesible.
    'No restringir Testing a una carpeta determinada.

    atributos = GetAttr(ruta)

    If (atributos And vbDirectory) = vbDirectory Then
        RutaPermitidaPorEntorno = True
    End If

    Exit Function

ErrorRuta:
    RutaPermitidaPorEntorno = False

End Function


'Obtener ruta inicial de TESTING sin tocar perma.dno.
'Si el archivo aun no existe o la ruta es invalida, usar la raiz.

Public Function RutaInicialNomina() As String

    Dim h As Integer
    Dim candidata As String

    On Error GoTo ErrorLectura

    'Si no hay ruta guardada, iniciar desde
    'el directorio actual de la aplicacion.
    RutaInicialNomina = CurDir$

    If Not gModoTesting Then Exit Function

    'Si el archivo no existe, continuar
    'normalmente desde el directorio actual.
    If Dir$(RutaArchivoEntornoNomina()) = "" Then
        Exit Function
    End If

    h = FreeFile

    Open RutaArchivoEntornoNomina() For Random Access Read As #h Len = Len(basico)

    If LOF(h) >= Len(basico) Then

        Get #h, 1, basico

        candidata = Trim$(Replace(basico.datoarch, vbNullChar, ""))

        'Aceptar cualquier carpeta valida.
        If RutaPermitidaPorEntorno(candidata) Then

            RutaInicialNomina = candidata

        End If

    End If

Salida:

    On Error Resume Next

    If h <> 0 Then Close #h

    On Error GoTo 0
    Exit Function

ErrorLectura:

    Debug.Print "No se pudo recuperar la ruta Testing: " & Err.Description

    RutaInicialNomina = CurDir$

    Resume Salida

End Function

'Comprobar el directorio efectivo antes de escribir archivos de empresa.
Public Function RutaTrabajoSegura( _
    Optional ByVal mostrar As Boolean = True) As Boolean

    Dim actual As String
    actual = CurDir$
    RutaTrabajoSegura = RutaPermitidaPorEntorno(actual)

    If Len(gRutaTrabajoAprobada) = 0 Then
        RutaTrabajoSegura = False
    ElseIf StrComp(RutaSimple(actual), _
                   RutaSimple(gRutaTrabajoAprobada), vbBinaryCompare) <> 0 Then
        RutaTrabajoSegura = False
    End If

    If Not RutaTrabajoSegura And mostrar Then
        MsgBox "Operacion bloqueada: la carpeta no corresponde al entorno " & _
               "activo." & vbCrLf & "Ruta: " & actual, _
               vbCritical, "Seguridad de archivos"
    End If
End Function

Public Function EtiquetaEntornoNomina() As String
    If gModoTesting Then
        EtiquetaEntornoNomina = "TESTING"
    Else
        EtiquetaEntornoNomina = "PRODUCCION"
    End If
End Function
