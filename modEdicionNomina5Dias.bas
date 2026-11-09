Attribute VB_Name = "modEdicionNomina5Dias"
Option Explicit

Private Function FechaPrimeraCreacionNomina(ByVal rutaNom As String) As Date

    Dim sistema As Object
    Dim archivoNom As Object
    Dim archivoCmp As Object
    Dim rutaCmp As String
    Dim fechaNom As Date
    Dim fechaCmp As Date

    If Len(rutaNom) < 5 Then
        Err.Raise vbObjectError + 2810, , "Ruta de nomina invalida."
    End If

    If UCase$(Right$(rutaNom, 4)) <> ".NOM" Then
        Err.Raise vbObjectError + 2811, , "El archivo debe tener extension NOM."
    End If

    rutaCmp = Left$(rutaNom, Len(rutaNom) - 4) & ".cmp"

    Set sistema = CreateObject("Scripting.FileSystemObject")

    If Not sistema.FileExists(rutaNom) Then
        Err.Raise vbObjectError + 2812, , "No existe el archivo NOM."
    End If

    If Not sistema.FileExists(rutaCmp) Then
        Err.Raise vbObjectError + 2813, , "No existe el complemento CMP."
    End If

    Set archivoNom = sistema.GetFile(rutaNom)
    Set archivoCmp = sistema.GetFile(rutaCmp)

    fechaNom = archivoNom.DateCreated
    fechaCmp = archivoCmp.DateCreated

    'No aceptar fechas posteriores al reloj actual.
    If fechaNom > Now Or fechaCmp > Now Then
        Err.Raise vbObjectError + 2814, , "Los archivos presentan fechas futuras."
    End If

    'La fecha mas antigua es el origen conservador. 'Una modificacion posterior no renueva el plazo.
    If fechaNom <= fechaCmp Then
        FechaPrimeraCreacionNomina = fechaNom
    Else
        FechaPrimeraCreacionNomina = fechaCmp
    End If

End Function

Public Function PuedeModificarNomina5Dias(ByVal rutaNom As String, ByRef detalle As String) As Boolean

    Dim origen As Date
    Dim vencimiento As Date

    On Error GoTo ErrorFecha

    detalle = ""
    PuedeModificarNomina5Dias = False

    origen = FechaPrimeraCreacionNomina(rutaNom)
    vencimiento = DateAdd("d", 5, origen)

    If Now >= vencimiento Then
        detalle = "El plazo de cinco dias ha vencido." & vbCrLf & "Generada: " & Format$(origen, "dd/mm/yyyy hh:nn:ss") & vbCrLf & "Limite: " & Format$(vencimiento, "dd/mm/yyyy hh:nn:ss")
        Exit Function
    End If

    PuedeModificarNomina5Dias = True
    Exit Function

ErrorFecha:
    detalle = "No se puede autorizar la edicion: " & Err.Description
    PuedeModificarNomina5Dias = False
End Function


