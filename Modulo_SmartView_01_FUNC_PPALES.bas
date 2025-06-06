Attribute VB_Name = "Modulo_SmartView_01_FUNC_PPALES"
Option Explicit

Public Function Pedir_Credenciales(ByRef vUsername As String, ByRef vPassword As String) As Boolean
    'Declaracion de Variables para conectar
    'Dim vUsername As String
    'Dim vPassword As String
    
    ' Variables para control de errores
    Dim strFuncion As String
    ' Inicialización
    strFuncion = "Pedir_Credenciales"

    ' Validar/Crear hoja UserName
    If Not fun802_SheetExists(CONST_HOJA_USERNAME) Then
        If Not F002_Crear_Hoja(CONST_HOJA_USERNAME) Then
            Err.Raise ERROR_BASE_IMPORT + 3, strFuncion, _
                "Error al crear la hoja " & gstrHoja_UserName
        End If
    End If
    ' Verificar si debemos ocultar la hoja UserName (comprobando la constante global CONST_OCULTAR_HOJA_USERNAME)
    If CONST_OCULTAR_HOJA_USERNAME = True Then
        ' Ocultar la hoja de delimitadores
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(CONST_HOJA_USERNAME)
        If Not fun809_OcultarHojaDelimitadores(ws) Then
            Debug.Print "ADVERTENCIA: Error al ocultar la hoja " & gstrHoja_UserName & " - Función: F000_Comprobaciones_Iniciales - " & Now()
            ' Nota: No es un error crítico, el proceso puede continuar
        End If
    End If
    
    'Tomamos el vUsername de la hoja Username, de la celda referenciada en la constante global (con CONST_CELDA_USERNAME = "C2" por ejemplo)
    vUsername = ws.Range(CONST_CELDA_USERNAME).Value
    ThisWorkbook.Save
    
    'Para la celda en la que almacenamos el "Username", vamos a crearle un Header
    '   (normalmente a su izquierda con CONST_CELDA_HEADER_USERNAME = "B2" por ejemplo, y con CONST_VALOR_HEADER_USERNAME = "Username:")
    ws.Range(CONST_CELDA_HEADER_USERNAME).Value = CONST_VALOR_HEADER_USERNAME

    'Connection Inputs
    'Inicializamos con valor en blanco los 2 TextBox del UserForm (el de Username y el de Password)
    UserForm_Username_Password.TextBox_Username.Value = vUsername
    UserForm_Username_Password.TextBox_Password.Value = ""
    'Mostramos el User Form, para pedir el Username y la Password
    UserForm_Username_Password.Show vbModal
    
    'Ahora a las variables de vUsername y vPassword le pasamos el valor de los TextBox de Username y Password
    vUsername = UserForm_Username_Password.TextBox_Username.Value
    vPassword = UserForm_Username_Password.TextBox_Password.Value
    'Y guardamos su valor en la hoja de Username/Password (CONST_HOJA_USERNAME)
    ws.Range(CONST_CELDA_USERNAME).Value = vUsername
    ThisWorkbook.Save

End Function
              

Public Function SmartView_CreateConnection(ByVal vConnection_Username As String, ByVal vConnection_Password As String, ByVal vConnection_Provider As String, _
    ByVal vConnection_URL As String, ByVal vConnection_Server As String, ByVal vConnection_Application As String, ByVal vConnection_Database As String, _
    ByVal vConnection_Name As String, ByVal vConnection_Description As String, ByVal vConnection_Create_MostrarMensajes As Boolean, _
    ByVal vConnection_Create_MostrarMensajeFinal As Boolean) As Boolean

    'Declaracion de Variables para recoger el código de error retornado por SmartView
    Dim vExisteConexion_Return As Boolean
    Dim vEliminarConexion_Return As Integer
    Dim vCrearConexion_Return As Integer
    Dim vDesconectarTodo_Return As Integer
    Dim vConexionActiva_Return As Integer
    Dim vConectar_Return As Variant 'Integer
    
    'Inicializo la variable vEliminarConexion_Return
    vEliminarConexion_Return = 0
    'Inicializo la variable vConexionActiva_Return
    vConexionActiva_Return = 9999
    
    'Verificar si la conexion existe
    vExisteConexion_Return = HypConnectionExists(vConnection_Name)
    'Acciones a realizar si existe la conexion / acciones a realizar si no existe
    If vExisteConexion_Return = True Then
        If vConnection_Create_MostrarMensajes Then MsgBox ("La conexion ya existe. Vamos a actualizarla.")
        'Desconectar de todas las aplicaciones
        vDesconectarTodo_Return = HypDisconnectAll()
        'No tiene sentido verificar si se desconecto de todas las aplicaciones
        
        'Eliminar la conexion vConnection_Name
        vEliminarConexion_Return = HypRemoveConnection(vConnection_Name)
        'Verificar si la conexion se ha eliminado
        If vEliminarConexion_Return = 0 Then
           If vConnection_Create_MostrarMensajes Then MsgBox "Se eliminó la conexion " & Chr(34) & vConnection_Name & Chr(34) & vbCrLf & "para volver a crearla con valores actualizados"
        Else
          If vConnection_Create_MostrarMensajes Then MsgBox "Hubo un error al intentar eliminar la conexión " & Chr(34) & vConnection_Name & Chr(34) & vbCrLf & _
              "con el proposito de volver a crearla con valores actualizados." & vbCrLf & "Error Number = " & vEliminarConexion_Return
        End If
        
    Else
        If vConnection_Create_MostrarMensajes Then MsgBox "La conexion NO existe. Vamos a crearla." & vbCrLf & "Error Number = " & vExisteConexion_Return
    End If

    
    'Crear una conexion via SmartView a la aplicacion deseada
    '   Solo si el valor de la variable vEliminarConexion_Return sigue siendo 0
    
    If vEliminarConexion_Return = 0 Then
        vCrearConexion_Return = HypCreateConnection(Empty, vConnection_Username, vConnection_Password, vConnection_Provider, vConnection_URL, vConnection_Server, _
        vConnection_Application, vConnection_Database, vConnection_Name, vConnection_Description)
        'Verificar si la conexion se ha creado
        If vCrearConexion_Return = 0 Then
            If vConnection_Create_MostrarMensajes Then MsgBox "Se creo correctamente la conexion" & Chr(34) & vConnection_Name & Chr(34) & vbCrLf & "con valores actualizados"
            'Conectar a SmartView usando la nueva conexion > en la siguiente funcion (para establecerla como activa - SmartView_SetActiveConnection_x_Sheet)
            'Verificar si nos hemos conectado > en la siguiente funcion (para establecerla como activa - SmartView_SetActiveConnection_x_Sheet)
            'Fijar como conexion activa la nueva conexion > en la siguiente funcion (para establecerla como activa - SmartView_SetActiveConnection_x_Sheet)
            'Verificar si hemos puesto como conexion activa la recien creada > en la siguiente funcion (para establecerla como activa - SmartView_SetActiveConnection_x_Sheet)
        Else
            If vConnection_Create_MostrarMensajes Then MsgBox "Fallo la creacion de la conexion" & Chr(34) & vConnection_Name & Chr(34) & vbCrLf & _
                "Error Number = " & vCrearConexion_Return
        End If
    End If
      
    If vCrearConexion_Return = 0 Then
        SmartView_CreateConnection = True
        If vConnection_Create_MostrarMensajeFinal Then MsgBox "Se creo la conexion " & Chr(34) & vConnection_Name & Chr(34)
    Else
        SmartView_CreateConnection = False
        If vConnection_Create_MostrarMensajeFinal Then MsgBox "NO se consiguio crear la conexion " & Chr(34) & vConnection_Name & Chr(34)
    End If

End Function


Public Function SmartView_SetActiveConnection_x_Sheet(ByVal vConnection_Username As String, ByVal vConnection_Password As String, ByVal vConnection_Provider As String, _
    ByVal vConnection_URL As String, ByVal vConnection_Server As String, ByVal vConnection_Application As String, ByVal vConnection_Database As String, _
    ByVal vConnection_Name As String, ByVal vConnection_Description As String, ByVal vConnection_Create_MostrarMensajes As Boolean, _
    ByVal vConnection_Create_MostrarMensajeFinal As Boolean, ByVal vNombreHojaConexion As Variant) As Boolean

    'Declaracion de Variables para recoger el código de error retornado por SmartView
    Dim vExisteConexion_Return As Boolean
    Dim vCrearConexion_Return As Integer
    Dim vDesconectarTodo_Return As Integer
    Dim vConexionActiva_Return As Integer
    Dim vConectar_Return As Variant 'Integer
    
    'Inicializo la variable vConexionActiva_Return
    vConexionActiva_Return = 9999
    
    'Seleccionar y activar la hoja sobre la que queremos activar la conexion
    If Not IsEmpty(vNombreHojaConexion) And Not IsNull(vNombreHojaConexion) And Trim(CStr(vNombreHojaConexion)) <> "" Then
        vNombreHojaConexion = Trim(CStr(vNombreHojaConexion))
    Else
        vNombreHojaConexion = ThisWorkbook.ActiveSheet.Name
    End If
    
    ThisWorkbook.Worksheets(vNombreHojaConexion).Select
    ThisWorkbook.Worksheets(vNombreHojaConexion).Activate
    ActiveWindow.Zoom = 70

    
    'Verificar si la conexion existe
    vExisteConexion_Return = HypConnectionExists(vConnection_Name)
    'Acciones a realizar si existe la conexion / acciones a realizar si no existe
    If vExisteConexion_Return = True Then
        If vConnection_Create_MostrarMensajes Then MsgBox "Ya existe la conexion " & Chr(34) & vConnection_Name & Chr(34) & vbCrLf & "Vamos a actualizarla."
        'Desconectar de todas las aplicaciones
        vDesconectarTodo_Return = HypDisconnectAll()
        'No tiene sentido verificar si se desconecto de todas las aplicaciones
    Else
        If vConnection_Create_MostrarMensajes Then MsgBox "NO existe la conexion " & Chr(34) & vConnection_Name & Chr(34) & vbCrLf & "Vamos a crearla." & _
            vbCrLf & "Error Number = " & vExisteConexion_Return
    End If

    
    'Si la conexion no existe, la creamos.
    
    If vExisteConexion_Return = False Then
        vCrearConexion_Return = HypCreateConnection(Empty, vConnection_Username, vConnection_Password, vConnection_Provider, vConnection_URL, vConnection_Server, _
        vConnection_Application, vConnection_Database, vConnection_Name, vConnection_Description)
        'Verificar si la conexion se ha creado
        If vCrearConexion_Return = 0 Then
            If vConnection_Create_MostrarMensajes Then MsgBox ("Se creo correctamente la conexion" & Chr(34) & vConnection_Name & Chr(34) & vbCrLf & _
                "con valores actualizados")
            'Conectar a SmartView usando la nueva conexion
        Else
            If vConnection_Create_MostrarMensajes Then MsgBox "Fallo al crear la conexion" & Chr(34) & vConnection_Name & Chr(34) & vbCrLf & "Error Number = " & vCrearConexion_Return
        End If
    End If
    
    'Una vez que ya exite (existía o la acabamos de crear),
    '   nos conectamos a ella, y la fijamos como activa
    
    If (vExisteConexion_Return = True) Or (vCrearConexion_Return = 0) Then
        'Conectar a SmartView usando la nueva conexion
        vConectar_Return = HypConnect(Empty, vConnection_Username, vConnection_Password, vConnection_Name)
        'Verificar si nos hemos conectado
        If vConectar_Return = 0 Then
            If vConnection_Create_MostrarMensajes Then MsgBox ("Nos hemos conectado correctamente a " & vConnection_Name)
            'Fijar como conexion activa la nueva conexion
            vConexionActiva_Return = HypSetActiveConnection(vConnection_Name)
            'Verificar si hemos puesto como conexion activa la recien creada
            If vConexionActiva_Return = 0 Then
                If vConnection_Create_MostrarMensajes Then MsgBox "Hemos establecido " & Chr(34) & vConnection_Name & Chr(34) & vbCrLf & "como conexion 'activa'" & vbCrLf & _
                    "Contexto > en la hoja " & Chr(34) & vNombreHojaConexion & Chr(34)
                
            Else
                If vConnection_Create_MostrarMensajes Then MsgBox "Fallo al establecer " & Chr(34) & vConnection_Name & Chr(34) & vbCrLf _
                & "como conexion 'activa'." & vbCrLf & "Error Number = " & vConexionActiva_Return & vbCrLf & _
                "Contexto > en la hoja " & Chr(34) & vNombreHojaConexion & Chr(34)
            End If
        Else
            If vConnection_Create_MostrarMensajes Then MsgBox "Fallo al conectarnos." & vbCrLf & "Error Number = " & vConectar_Return & vbCrLf & _
                "Contexto > en la hoja " & Chr(34) & vNombreHojaConexion & Chr(34)
        End If
    End If
      
    'Y aqui fijamos el valor de retorno de la función actual
    If vConexionActiva_Return = 0 Then
        SmartView_SetActiveConnection_x_Sheet = True
        If vConnection_Create_MostrarMensajeFinal Then MsgBox "Se fijó como 'activa' la conexion " & Chr(34) & vConnection_Name & Chr(34) & vbCrLf & _
            "Contexto > en la hoja " & Chr(34) & vNombreHojaConexion & Chr(34)
    Else
        SmartView_SetActiveConnection_x_Sheet = False
        If vConnection_Create_MostrarMensajeFinal Then MsgBox "NO se pudo fijar como 'activa' la conexion " & Chr(34) & vConnection_Name & Chr(34) & vbCrLf & _
            "Contexto > en la hoja " & Chr(34) & vNombreHojaConexion & Chr(34)
    End If

End Function

Public Function SmartView_Options_DataOptions_Estandar(vHoja As Variant) As Boolean

    Dim vInteger01, vInteger02, vInteger03, vInteger04, vInteger05, vInteger06, vInteger07, vInteger08 As Double
    
    vInteger01 = SmartView_Options_MemberOptions_Indent_None(vHoja)
    
    vInteger02 = SmartView_Options_MemberOptions_DisplayNameOnly(vHoja)
    
    vInteger03 = SmartView_Options_DataOptions_CellDisplay(vHoja)
    
    vInteger04 = SmartView_Options_DataOptions_Supress_Missing(vHoja, False)
    
    vInteger05 = SmartView_Options_DataOptions_Supress_Zero(vHoja, False)
    
    vInteger06 = SmartView_Options_DataOptions_Supress_Repeated(vHoja, False)
    
    vInteger07 = SmartView_Options_DataOptions_Supress_Invalid(vHoja, False)
    
    vInteger08 = SmartView_Options_DataOptions_Supress_NoAccess(vHoja, False)
    
    If (vInteger01 <> 0) Or (vInteger02 <> 0) Or (vInteger03 <> 0) Or (vInteger04 <> 0) Or (vInteger05 <> 0) Or (vInteger06 <> 0) Or _
        (vInteger07 <> 0) Or (vInteger08 <> 0) Then
        SmartView_Options_DataOptions_Estandar = False
    Else
        SmartView_Options_DataOptions_Estandar = True
    End If

    
End Function

Public Function SmartView_Retrieve(vHoja As Variant) As Boolean
    Dim vErrorNumber As Long
    Dim vEnabled_Parts As Boolean
    
    ThisWorkbook.Worksheets(vHoja).Activate
    'vNombreDeLaHoja = ThisWorkbook.ActiveSheet.Name
    vHoja = ThisWorkbook.ActiveSheet.Name
    vErrorNumber = HypRetrieve(vHoja)
    'MsgBox "vLong = " & vLong
    
    vEnabled_Parts = False
    If vEnabled_Parts Then
        If vErrorNumber = 0 Then
            MsgBox "Refrescado con exito."
        Else
            MsgBox "Hubo un error al refrescar. Error Number = " & SmartView_Retrieve
        End If
    End If 'vEnabled_Parts Then
    
    If vErrorNumber = 0 Then
        SmartView_Retrieve = True
    Else
        SmartView_Retrieve = False
    End If
    
End Function

Public Function SmartView_Submit(vHoja As Variant) As Boolean
    Dim vLong As Long
    
    ThisWorkbook.Worksheets(vHoja).Activate
    'vNombreDeLaHoja = ThisWorkbook.ActiveSheet.Name
    vHoja = ThisWorkbook.ActiveSheet.Name
    vLong = HypSubmitData(vHoja)
    'MsgBox "vLong = " & vLong
    
    If vLong = 0 Then
        MsgBox "Submit ejecutado con exito."
    Else
        MsgBox "Hubo un error al ejecutar Submit. Error Number = " & SmartView_Submit
    End If
    If vLong = 0 Then
        SmartView_Submit = True
    Else
        SmartView_Submit = False
    End If
    
End Function

Public Function SmartView_Submit_without_Refresh(vHoja As Variant) As Boolean
    Dim vLong As Long
    
    ThisWorkbook.Worksheets(vHoja).Activate
    'vNombreDeLaHoja = ThisWorkbook.ActiveSheet.Name
    vHoja = ThisWorkbook.ActiveSheet.Name
    vLong = HypSubmitSelectedRangeWithoutRefresh(vHoja, False, False, False)
    'MsgBox "vLong = " & vLong
    
    If vLong = 0 Then
        MsgBox "Submitted without Refresh - ejecutado con exito."
    Else
        MsgBox "Hubo un error al ejecutar - Submit without Refresh. Error Number = " & SmartView_Submit_without_Refresh
    End If
    If vLong = 0 Then
        SmartView_Submit_without_Refresh = True
    Else
        SmartView_Submit_without_Refresh = False
    End If
    
End Function

