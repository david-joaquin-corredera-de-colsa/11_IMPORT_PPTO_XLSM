Attribute VB_Name = "Modulo_SmartView_05_SUB_PPAL"
Option Explicit

Public Function M002_SmartView_Paso_00_Crear_Conexion(ByVal vConnection_Username As String, ByVal vConnection_Password As String, ByVal vConnection_Provider As String, ByVal vConnection_URL As String, ByVal vConnection_Server As String, ByVal vConnection_Application As String, ByVal vConnection_Database As String, ByVal vConnection_Name As String, ByVal vConnection_Description As String, ByVal vConnection_Create_MostrarMensajes As Boolean, ByVal vConnection_Create_MostrarMensajeFinal As Boolean) As Boolean
    
    M002_SmartView_Paso_00_Crear_Conexion = SmartView_CreateConnection(vConnection_Username, vConnection_Password, vConnection_Provider, vConnection_URL, vConnection_Server, vConnection_Application, vConnection_Database, vConnection_Name, vConnection_Description, vConnection_Create_MostrarMensajes, vConnection_Create_MostrarMensajeFinal)
        
End Function

Public Function M002_SmartView_Paso_01_Establecer_Conexion_Activa_para_Hoja(ByVal vConnection_Username As String, ByVal vConnection_Password As String, ByVal vConnection_Provider As String, ByVal vConnection_URL As String, ByVal vConnection_Server As String, ByVal vConnection_Application As String, ByVal vConnection_Database As String, ByVal vConnection_Name As String, ByVal vConnection_Description As String, ByVal vConnection_Create_MostrarMensajes As Boolean, ByVal vConnection_Create_MostrarMensajeFinal As Boolean, ByVal vNombreHojaConexion As Variant) As Boolean
    
    M002_SmartView_Paso_01_Establecer_Conexion_Activa_para_Hoja = SmartView_SetActiveConnection_x_Sheet(vConnection_Username, vConnection_Password, vConnection_Provider, vConnection_URL, vConnection_Server, vConnection_Application, vConnection_Database, vConnection_Name, vConnection_Description, vConnection_Create_MostrarMensajes, vConnection_Create_MostrarMensajeFinal, vNombreHojaConexion)
        
End Function


Sub M002_SmartView_Paso_02_EstablecerOpciones(vNombreDeLaHoja As String)

    'Dim x As Boolean
    'x = SmartView_CreateConnection
        
    'Dim vNombreDeLaHoja As String
    'vNombreDeLaHoja = ThisWorkbook.ActiveSheet.Name
    
    Dim vReturn_SmartView_Options_DataOptions As Integer
    vReturn_SmartView_Options_DataOptions = SmartView_Options_DataOptions_Estandar(vNombreDeLaHoja)
    
End Sub
Sub M002_SmartView_Paso_02_Hacer_Refresh(vNombreDeLaHoja As String)

    'Dim x As Boolean
    'x = SmartView_CreateConnection
    
    'Dim vNombreDeLaHoja As String
    'vNombreDeLaHoja = ThisWorkbook.ActiveSheet.Name
    
    Dim vReturn_SmartView_Retrieve As Integer
    vReturn_SmartView_Retrieve = SmartView_Retrieve(vNombreDeLaHoja)
    
End Sub

Sub M002_SmartView_Paso_02_EstablecerOpciones_CrearAdHoc(vNombreDeLaHoja As String)

    'Dim x As Boolean
    'x = SmartView_CreateConnection
    
    
    'Dim vNombreDeLaHoja As String
    'vNombreDeLaHoja = ThisWorkbook.ActiveSheet.Name
    
'    Dim vReturn_SmartView_Options_DataOptions As Integer
'    vReturn_SmartView_Options_DataOptions = SmartView_Options_DataOptions_Estandar(vNombreDeLaHoja)
    
'    Dim vReturn_SmartView_Retrieve As Integer
'    vReturn_SmartView_Retrieve = SmartView_Retrieve(vNombreDeLaHoja)
    
    Call M002_SmartView_Paso_02_EstablecerOpciones(vNombreDeLaHoja)
    Call M002_SmartView_Paso_02_Hacer_Refresh(vNombreDeLaHoja)

    
End Sub

Sub M003_SmartView_Paso_02_Submit(vNombreDeLaHoja As String)

'    Dim x As Boolean
'    x = SmartView_CreateConnection
    
    
    'Dim vNombreDeLaHoja As String
    'vNombreDeLaHoja = ThisWorkbook.ActiveSheet.Name
    
'    Dim vReturn_SmartView_Options_DataOptions As Integer
'    vReturn_SmartView_Options_DataOptions = SmartView_Options_DataOptions_Estandar(vNombreDeLaHoja)
    
    Dim vReturn_SmartView_Submit As Integer
    vReturn_SmartView_Submit = SmartView_Submit(vNombreDeLaHoja)
    
End Sub

Sub M003_SmartView_Paso_02_Submit_without_Refresh(vNombreDeLaHoja As String)

'    Dim x As Boolean
'    x = SmartView_CreateConnection
    
    
    'Dim vNombreDeLaHoja As String
    'vNombreDeLaHoja = ThisWorkbook.ActiveSheet.Name
    
'    Dim vReturn_SmartView_Options_DataOptions As Integer
'    vReturn_SmartView_Options_DataOptions = SmartView_Options_DataOptions_Estandar(vNombreDeLaHoja)
    
    Dim vReturn_SmartView_Submit_without_Refresh As Integer
    vReturn_SmartView_Submit_without_Refresh = SmartView_Submit_without_Refresh(vNombreDeLaHoja)
    
End Sub

Public Sub xx_Stand_Alone_M003_SmartView_Paso_02_Submit_without_Refresh()

'    Dim x As Boolean
'    x = SmartView_CreateConnection
    
    
    Dim vNombreDeLaHoja As String
    vNombreDeLaHoja = ThisWorkbook.ActiveSheet.Name
    MsgBox "vNombreDeLaHoja=" & vNombreDeLaHoja
    
'    Dim vReturn_SmartView_Options_DataOptions As Integer
'    vReturn_SmartView_Options_DataOptions = SmartView_Options_DataOptions_Estandar(vNombreDeLaHoja)
    
    Dim vReturn_SmartView_Submit_without_Refresh As Integer
    vReturn_SmartView_Submit_without_Refresh = SmartView_Submit_without_Refresh(vNombreDeLaHoja)
    
End Sub
Sub xx_Stand_Alone_M003_SmartView_Paso_02_Submit()

'    Dim x As Boolean
'    x = SmartView_CreateConnection
    
    
    Dim vNombreDeLaHoja As String
    vNombreDeLaHoja = ThisWorkbook.ActiveSheet.Name
    MsgBox "vNombreDeLaHoja=" & vNombreDeLaHoja
    
'    Dim vReturn_SmartView_Options_DataOptions As Integer
'    vReturn_SmartView_Options_DataOptions = SmartView_Options_DataOptions_Estandar(vNombreDeLaHoja)
    
    Dim vReturn_SmartView_Submit As Integer
    vReturn_SmartView_Submit = SmartView_Submit(vNombreDeLaHoja)
    
End Sub

Public Sub xx_Editar_Celdas()
    Dim r As Integer
    Dim c As Integer
    Dim vValor As Variant
    
    
    For r = 8 To 10
        For c = 13 To 24
            vValor = Cells(r, c).Value
            Cells(r, c).Value = vValor
        Next c
    Next r
End Sub
