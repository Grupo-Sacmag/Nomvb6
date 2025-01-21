Attribute VB_Name = "COMCFDI"
Type EMP_1
    Direccion As String * 100
    Colonia As String * 100
    Ciudad As String * 100
    Estado As String * 64
    Delegacion As String * 64
    Cpostal As String * 6
    correo As String * 64
End Type

Type EmpC
     Folio As Long
     serie As String * 10
     Consecutivo As Integer
     RegPatr As String * 32
     RiesgoImss As String * 20
     Otra1 As String * 30
     Otra2 As Long
End Type

Type EMP_2
    NumConc As String * 10
    Concepto As String * 100
End Type

Public EmpCFDI As EmpC
Public DATcf As EMP_2
Public Empleado_1 As EMP_1
Public MiFecha
