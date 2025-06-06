Attribute VB_Name = "Modulo_501_SUB_principal"
Option Explicit

'******************************************************************************
' Módulo: M001_Ejecutar_Proceso_Principal
' Fecha y Hora de Creación: 2025-05-26 05:39:34 UTC
' Autor: david-joaquin-corredera-de-colsa
'
' Descripción:
' Este módulo contiene el procedimiento principal que coordina la ejecución
' de los procesos de importación y gestión de datos.
'******************************************************************************

Public Sub M001_Ejecutar_Proceso_Principal()

    '--------------------------------------------------------------------------
    ' Variables para control de errores y seguimiento
    '--------------------------------------------------------------------------
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    Dim blnResult As Boolean
    Dim x As Boolean
    
    'Variable para habilitar/deshabilitar partes de esta SUB
    Dim vEnabled_Parts As Boolean
    
    ' Inicialización
    strFuncion = "M001_Ejecutar_Proceso_Principal"
    lngLineaError = 0
    
    
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 0. Configuración inicial del entorno
    '--------------------------------------------------------------------------
    lngLineaError = 44
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Inicializar variables globales
    Call InitializeGlobalVariables
    
    fun801_LogMessage "Iniciando proceso principal..."
    
    '--------------------------------------------------------------------------
    ' 1. Ejecución de comprobaciones iniciales (F000)
    '--------------------------------------------------------------------------
    lngLineaError = 54
    fun801_LogMessage "Ejecutando comprobaciones iniciales..."
    
    blnResult = F000_Comprobaciones_Iniciales()
    
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1001, strFuncion, _
            "Las comprobaciones iniciales no se completaron correctamente"
    End If
    
    
    '--------------------------------------------------------------------------
    ' 2. Creacion de hojas de importacion (F001)
    '--------------------------------------------------------------------------
    lngLineaError = 55
    fun801_LogMessage "Creando hojas de importacion..."

    blnResult = F001_Crear_hojas_de_Importacion()

    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1001, strFuncion, _
            "Las hojas de importacion no se crearon correctamente"
    End If
    
    '--------------------------------------------------------------------------
    ' 2a. Detectar delimitadores Originales del sistema |||||||
    '--------------------------------------------------------------------------
    lngLineaError = 61
    Call fun801_LogMessage("Detectando delimitadores del sistema", False)

    blnResult = F004_Detectar_Delimitadores_en_Excel()
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 6, strFuncion, _
            "Error en la detección de delimitadores"
    End If
    
    ThisWorkbook.Save
    
    '--------------------------------------------------------------------------
    ' 2b. Forzar delimitadores Especificos en el sistema  |||||||
    '--------------------------------------------------------------------------
    lngLineaError = 62
    Call fun801_LogMessage("Forzando delimitadores Especificos en el sistema", False)

    blnResult = F004_Forzar_Delimitadores_en_Excel()
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 6, strFuncion, _
            "Error en la detección de delimitadores"
    End If
    
    '--------------------------------------------------------------------------
    ' 3. Mostrar información de las hojas creadas
    '--------------------------------------------------------------------------
    lngLineaError = 66
    
    ' Mostrar nombre de la hoja de importación
    If CONST_MOSTRAR_MENSAJES_HOJAS_CREADAS Then MsgBox "Hoja de Importación:" & vbCrLf & vbCrLf & _
           gstrNuevaHojaImportacion & vbCrLf & vbCrLf & _
           "Esta hoja contendrá los datos importados.", _
           vbInformation, _
           "Hoja de Importación - " & strFuncion
    
    ' Mostrar nombre de la hoja de trabajo
    If CONST_MOSTRAR_MENSAJES_HOJAS_CREADAS Then MsgBox "Hoja de Trabajo (Working):" & vbCrLf & vbCrLf & _
           gstrNuevaHojaImportacion_Working & vbCrLf & vbCrLf & _
           "Esta hoja se utilizará para procesamiento temporal.", _
           vbInformation, _
           "Hoja de Trabajo - " & strFuncion
    
    ' Mostrar nombre de la hoja de envío
    If CONST_MOSTRAR_MENSAJES_HOJAS_CREADAS Then MsgBox "Hoja de Envío:" & vbCrLf & vbCrLf & _
           gstrNuevaHojaImportacion_Envio & vbCrLf & vbCrLf & _
           "Esta hoja contendrá los datos listos para envío.", _
           vbInformation, _
           "Hoja de Envío - " & strFuncion
           
    ' Mostrar nombre de la hoja de comprobación
    If CONST_MOSTRAR_MENSAJES_HOJAS_CREADAS Then MsgBox "Hoja de Comprobación:" & vbCrLf & vbCrLf & _
           gstrNuevaHojaImportacion_Comprobacion & vbCrLf & vbCrLf & _
           "Esta hoja se utilizará para verificación y control de calidad.", _
           vbInformation, _
           "Hoja de Comprobación - " & strFuncion
           
    '--------------------------------------------------------------------------
    ' 3a. Localizar hoja de envío anterior
    '--------------------------------------------------------------------------
    lngLineaError = 89
    fun801_LogMessage "Iniciando localización de hoja de envío anterior..."
    
    blnResult = F009_Localizar_Hoja_Envio_Anterior()
    
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1011, strFuncion, _
            "La localización de hoja de envío anterior no se completó correctamente"
    End If
    
    Dim vPrefijo_Del_Prev_Envio, vSufijo_Del_Prev_Envio As String
    vPrefijo_Del_Prev_Envio = "Del_Prev_Envio_"
    vSufijo_Del_Prev_Envio = Right(gstrPreviaHojaImportacion_Envio, 15)
    
    gstrPrevDelHojaImportacion_Envio = vPrefijo_Del_Prev_Envio & vSufijo_Del_Prev_Envio
    
    '--------------------------------------------------------------------------
    ' 3b. Copiar hoja de envío anterior
    '--------------------------------------------------------------------------
    lngLineaError = 90
    fun801_LogMessage "Iniciando copia de hoja de envío anterior..."
    
    blnResult = F010_Copiar_Hoja_Envio_Anterior()
    
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1012, strFuncion, _
            "La copia de hoja de envío anterior no se completó correctamente"
    End If
    
    '--------------------------------------------------------------------------
    ' 4. Ejecutar proceso de importación (F002)
    '--------------------------------------------------------------------------
    lngLineaError = 91
    fun801_LogMessage "Iniciando proceso de importación..."
    
    blnResult = F002_Importar_Fichero(gstrNuevaHojaImportacion, _
                                     gstrNuevaHojaImportacion_Working, _
                                     gstrNuevaHojaImportacion_Envio, _
                                     gstrNuevaHojaImportacion_Comprobacion)
    
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1002, strFuncion, _
            "El proceso de importación no se completó correctamente"
    End If
        
    '--------------------------------------------------------------------------
    ' 5. Procesar hoja de envío
    '--------------------------------------------------------------------------
    lngLineaError = 95
    fun801_LogMessage "Iniciando procesamiento de hoja de envío..."
    
    blnResult = F003_Procesar_Hoja_Envio(gstrNuevaHojaImportacion_Working, _
                                        gstrNuevaHojaImportacion_Envio)
    
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1003, strFuncion, _
            "El procesamiento de la hoja de envío no se completó correctamente"
    End If
        
    '--------------------------------------------------------------------------
    ' 6. Procesar hoja de comprobación
    '--------------------------------------------------------------------------
    lngLineaError = 97
    fun801_LogMessage "Iniciando procesamiento de hoja de comprobación..."
    
    blnResult = F005_Procesar_Hoja_Comprobacion()
    
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1004, strFuncion, _
            "El procesamiento de la hoja de comprobación NO se completó correctamente"
    End If
    
    '--------------------------------------------------------------------------
    ' 6.8.1.    SmartView: Para la hoja gstrPrevDelHojaImportacion_Envio:
    '           Crear Conexiones, Actualizar Opciones, Crear AdHoc
    '--------------------------------------------------------------------------
    
    ThisWorkbook.Worksheets(gstrPrevDelHojaImportacion_Envio).Select
    ThisWorkbook.Worksheets(gstrPrevDelHojaImportacion_Envio).Activate
    ActiveWindow.Zoom = 70
    
    lngLineaError = 106
    fun801_LogMessage "Creando conexiones para la hoja " & gstrPrevDelHojaImportacion_Envio & "..."
    
    'Pedimos las credenciales
    x = Pedir_Credenciales(vUsername, vPassword)
    
    'MsgBox "Username=" & vUsername & "|Password=" & vPassword
    Dim vConnection_Username As String: vConnection_Username = vUsername
    Dim vConnection_Password As String: vConnection_Password = vPassword
    Dim vConnection_Provider As String: vConnection_Provider = CONST_PROVIDER
    Dim vConnection_URL As String: vConnection_URL = CONST_PROVIDER_URL
    Dim vConnection_Server As String: vConnection_Server = CONST_SERVER_NAME
    Dim vConnection_Application As String: vConnection_Application = CONST_APPLICATION_NAME
    Dim vConnection_Database As String: vConnection_Database = CONST_DATABASE_NAME
    Dim vConnection_Name As String: vConnection_Name = CONST_CONNECTION_FRIENDLY_NAME
    Dim vConnection_Description As String: vConnection_Description = CONST_DESCRIPTION
    Dim vConnection_Create_MostrarMensajes As Boolean: vConnection_Create_MostrarMensajes = CONST_MOSTRAR_MENSAJES_SMARTVIEW_CREAR_CONEXION
    Dim vConnection_Create_MostrarMensajeFinal As Boolean: vConnection_Create_MostrarMensajeFinal = CONST_MOSTRAR_MENSAJE_FINAL_SMARTVIEW_CREAR_CONEXION
    
    'Creamos la conexion
'    x = M002_SmartView_Paso_00_Crear_Conexion(vConnection_Username, vConnection_Password, vConnection_Provider, vConnection_URL, vConnection_Server, _
'        vConnection_Application, vConnection_Database, vConnection_Name, vConnection_Description, _
'        vConnection_Create_MostrarMensajes, vConnection_Create_MostrarMensajeFinal)
    x = SmartView_CreateConnection(vConnection_Username, vConnection_Password, vConnection_Provider, vConnection_URL, vConnection_Server, _
        vConnection_Application, vConnection_Database, vConnection_Name, vConnection_Description, _
        vConnection_Create_MostrarMensajes, vConnection_Create_MostrarMensajeFinal)
    
    
    Dim vConnection_FijarActiva_MostrarMensajes As Boolean: vConnection_Create_MostrarMensajes = CONST_MOSTRAR_MENSAJES_SMARTVIEW_FIJAR_CONEXION_ACTIVA
    Dim vConnection_FijarActiva_MostrarMensajeFinal As Boolean: vConnection_Create_MostrarMensajeFinal = CONST_MOSTRAR_MENSAJE_FINAL_SMARTVIEW_FIJAR_CONEXION_ACTIVA
    
    'La establecemos como activa para la hoja gstrPrevDelHojaImportacion_Envio
'    x = M002_SmartView_Paso_01_Establecer_Conexion_Activa_para_Hoja(vConnection_Username, vConnection_Password, vConnection_Provider, vConnection_URL, vConnection_Server, _
'        vConnection_Application, vConnection_Database, vConnection_Name, vConnection_Description, _
'        vConnection_Create_MostrarMensajes, vConnection_Create_MostrarMensajeFinal, gstrPrevDelHojaImportacion_Envio)
    x = SmartView_SetActiveConnection_x_Sheet(vConnection_Username, vConnection_Password, vConnection_Provider, vConnection_URL, vConnection_Server, _
        vConnection_Application, vConnection_Database, vConnection_Name, vConnection_Description, _
        vConnection_Create_MostrarMensajes, vConnection_Create_MostrarMensajeFinal, gstrPrevDelHojaImportacion_Envio)
    
    Call M002_SmartView_Paso_02_EstablecerOpciones_CrearAdHoc(gstrPrevDelHojaImportacion_Envio)
    
    '--------------------------------------------------------------------------
    ' 6.8.2. SmartView: localizo final inicial y final
    '                   (como en F007_Copiar_Datos_de_Comprobacion_a_Envio)
    '                   y voy editando cada celda de datos (con valor en blanco)
    '--------------------------------------------------------------------------
    
    vEnabled_Parts = False
    If vEnabled_Parts Then
    
        ThisWorkbook.Worksheets(gstrPrevDelHojaImportacion_Envio).Select
        ThisWorkbook.Worksheets(gstrPrevDelHojaImportacion_Envio).Activate
        ActiveWindow.Zoom = 70
    
        lngLineaError = 107
        fun801_LogMessage "Copiando datos de hoja de comprobación a hoja de envío..."
    
        'blnResult = F007_Preparar_Datos_para_Borrado(gstrPrevDelHojaImportacion_Envio) 'Crear esta nueva funcion 20250604
    
        If Not blnResult Then
            Err.Raise ERROR_BASE_IMPORT + 1009, strFuncion, _
                "La copia de datos de comprobación a envío NO se completó correctamente"
        End If
    
    End If 'vEnabled_Parts Then
    
    '--------------------------------------------------------------------------
    ' 7. SmartView: Crear Conexiones, Actualizar Opciones, Crear AdHoc
    '--------------------------------------------------------------------------
    
    ThisWorkbook.Worksheets(gstrNuevaHojaImportacion_Envio).Select
    ThisWorkbook.Worksheets(gstrNuevaHojaImportacion_Envio).Activate
    ActiveWindow.Zoom = 70
    
    'Call M002_SmartView_Paso_02_EstablecerOpciones_CrearAdHoc(gstrNuevaHojaImportacion_Envio)
    'x = M002_SmartView_Paso_00_Crear_Conexion(vConnection_Username, vConnection_Password, vConnection_Provider, vConnection_URL, vConnection_Server, vConnection_Application, vConnection_Database, vConnection_Name, vConnection_Description, vConnection_Create_MostrarMensajes, vConnection_Create_MostrarMensajeFinal)
    'Call M002_SmartView_Paso_02_EstablecerOpciones_CrearAdHoc(gstrNuevaHojaImportacion_Envio)
    
    'La establecemos como activa para la hoja gstrPrevDelHojaImportacion_Envio
'    x = M002_SmartView_Paso_01_Establecer_Conexion_Activa_para_Hoja(vConnection_Username, vConnection_Password, vConnection_Provider, vConnection_URL, vConnection_Server, _
'        vConnection_Application, vConnection_Database, vConnection_Name, vConnection_Description, _
'        vConnection_Create_MostrarMensajes, vConnection_Create_MostrarMensajeFinal, gstrNuevaHojaImportacion_Envio)
    x = SmartView_SetActiveConnection_x_Sheet(vConnection_Username, vConnection_Password, vConnection_Provider, vConnection_URL, vConnection_Server, _
        vConnection_Application, vConnection_Database, vConnection_Name, vConnection_Description, _
        vConnection_Create_MostrarMensajes, vConnection_Create_MostrarMensajeFinal, gstrNuevaHojaImportacion_Envio)
    
    Call M002_SmartView_Paso_02_EstablecerOpciones_CrearAdHoc(gstrNuevaHojaImportacion_Envio)
    
    
    '--------------------------------------------------------------------------
    ' 8.1. Copiar datos de comprobación a envío
    '--------------------------------------------------------------------------
    lngLineaError = 107
    fun801_LogMessage "Copiando datos de hoja de comprobación a hoja de envío..."
    
    blnResult = F007_Copiar_Datos_de_Comprobacion_a_Envio(gstrNuevaHojaImportacion_Comprobacion, _
                                                          gstrNuevaHojaImportacion_Envio)
    
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1009, strFuncion, _
            "La copia de datos de comprobación a envío NO se completó correctamente"
    End If
    
    '--------------------------------------------------------------------------
    ' 8.2. SmartView: Enviar Datos / Submit
    '--------------------------------------------------------------------------
    
    ThisWorkbook.Worksheets(gstrNuevaHojaImportacion_Envio).Select
    ThisWorkbook.Worksheets(gstrNuevaHojaImportacion_Envio).Activate
    ActiveWindow.Zoom = 70
    
    Call M003_SmartView_Paso_02_Submit(gstrNuevaHojaImportacion_Envio)
    'Call M003_SmartView_Paso_02_Submit_without_Refresh(gstrNuevaHojaImportacion_Envio)
    
    '--------------------------------------------------------------------------
    ' 8.2. SmartView: Comprobar gstrNuevaHojaImportacion_Envio contra gstrNuevaHojaImportacion_Comprobacion > En CURSO
    '--------------------------------------------------------------------------
    
    'Primero hacer refresh de gstrNuevaHojaImportacion_Envio
    Call M002_SmartView_Paso_02_Hacer_Refresh(gstrNuevaHojaImportacion_Envio)
    
    'Segundo comparar celda por celda, y señalar en gstrNuevaHojaImportacion_Envio la linea con problemas
    ' basandome en lo siguiente:
    
'    lngLineaError = 107
'    fun801_LogMessage "Copiando datos de hoja de comprobación a hoja de envío..."
'
'    blnResult = F008_Comparar_Datos_HojaComprobacion_vs_HojaEnvio(gstrNuevaHojaImportacion_Comprobacion, _
'                                                          gstrNuevaHojaImportacion_Envio)
'
'    If Not blnResult Then
'        Err.Raise ERROR_BASE_IMPORT + 1009, strFuncion, _
'            "La copia de datos de comprobación a envío NO se completó correctamente"
'    End If
    
    
    '--------------------------------------------------------------------------
    ' 9. Restaurar delimitadores Originales del sistema |||||||
    '--------------------------------------------------------------------------
    lngLineaError = 103
    Call fun801_LogMessage("Detectando delimitadores del sistema", False)

    blnResult = F004_Restaurar_Delimitadores_en_Excel()
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1007, strFuncion, _
            "Error en la detección de delimitadores"
    End If
    
    ThisWorkbook.Save
    
    '--------------------------------------------------------------------------
    ' 9a. Limpieza de Hojas Historicas
    '--------------------------------------------------------------------------
    
    vEnabled_Parts = True
    If vEnabled_Parts Then
    
        lngLineaError = 109
        fun801_LogMessage "Iniciando proceso de limpieza de hojas historicas..."
    
        blnResult = Function_Return_Integer_to_Boolean(F011_Limpieza_Hojas_Historicas())
        If Not blnResult Then
            Err.Raise ERROR_BASE_IMPORT + 1010, strFuncion, _
                "El proceso de limpieza de hojas historicas NO se completó correctamente"
        End If
        
    End If 'vEnabled_Parts Then
    
    '--------------------------------------------------------------------------
    ' 9b. Inventariar Hojas existentes en el libro
    '--------------------------------------------------------------------------
    lngLineaError = 110
    fun801_LogMessage "Iniciando proceso de inventariar hojas..."
    
    blnResult = Function_Return_Integer_to_Boolean(F012_Inventariar_Hojas())
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1010, strFuncion, _
            "El proceso de inventariar NO se completó correctamente"
    End If
    
    '--------------------------------------------------------------------------
    ' 10. Proceso completado
    '--------------------------------------------------------------------------
    lngLineaError = 111
    fun801_LogMessage "Proceso principal completado correctamente"
    
    MsgBox "El proceso se ha completado correctamente." & vbCrLf & vbCrLf & _
           "- Hojas creadas: 4" & vbCrLf & _
           "- Datos importados en: " & gstrNuevaHojaImportacion & vbCrLf & _
           "- Rango de datos: " & glngLineaInicial_HojaImportacion & " a " & _
           glngLineaFinal_HojaImportacion & vbCrLf & _
           "- Hoja de comprobación preparada: " & gstrNuevaHojaImportacion_Comprobacion, _
           vbInformation, _
           "Éxito - " & strFuncion
               
    ThisWorkbook.Worksheets("00_Ejecutar_Procesos").Select
    
    ThisWorkbook.Save
           
CleanExit:
    '--------------------------------------------------------------------------
    ' 7. Restauración del entorno
    '--------------------------------------------------------------------------
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

GestorErrores:
    ' Construcción del mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "Línea: " & lngLineaError & vbCrLf & _
                      "Número de Error: " & Err.Number & vbCrLf & _
                      "Origen: " & Err.Source & vbCrLf & _
                      "Descripción: " & Err.Description
    
    ' Registro del error
    fun801_LogMessage strMensajeError, True
    
    ' Mostrar mensaje al usuario
    MsgBox strMensajeError, vbCritical, "Error en " & strFuncion
    
    ' Asegurar que se restaura la configuración de Excel
    Resume CleanExit
End Sub


