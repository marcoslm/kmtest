Attribute VB_Name = "mod_kmtest"
Option Explicit



Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const conSwNormal = 1

'Constantes para pasarle a la función Api SetWindowPos
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
    
' Función Api SetWindowPos
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                                ByVal hWndInsertAfter As Long, _
                                                ByVal X As Long, ByVal Y As Long, _
                                                ByVal cX As Long, _
                                                ByVal cY As Long, _
                                                ByVal wFlags As Long) As Long
                            
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CLEANBOOT& = 67                    ' 0=normal, 1=safemode, 2=safemode+network
Private Const SM_CMONITORS = 80                     ' Cantidad de monitores activos
Private Const SM_CXSCREEN = 0                       ' Ancho en pixels de la pantalla en monitor principal
Private Const SM_CYSCREEN = 1                       ' Alto en pixels de la pantalla en monitor principal
Private Const SM_CMOUSEBUTTONS = 43                 ' Número de botones del mouse (0 si no hay mouse)
Private Const SM_MOUSEHORIZONTALWHEELPRESENT = 91   ' Devuelve mayor de cero si el mouse tiene una rueda de scroll horizontal
Private Const SM_MOUSEWHEELPRESENT = 75             ' Devuelve mayor de cero si el mouse tiene una rueda de scroll vertical

' Para recoger información de estado de los "LEDs"
'Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
'Const VK_CAPITAL = &H14
'Const VK_NUMLOCK = &H90
            
' Enum para distinguir entre evento de mouse o de teclado (para función "NuevoEvento")
Public Enum eTipoEvento
    EVENTO_TECLADO = 0
    EVENTO_MOUSE = 1
End Enum

' Tipo de datos para almacenar información/características del Teclado y Mouse
Public Type tInfoPerifericos
	Mouse_Presente As Boolean
    Mouse_Botones As Integer
    Mouse_RuedaVertical As Boolean
    Mouse_RuedaHorizontal As Boolean
End Type

' Variable para almacenar información/características del Teclado y Mouse
Public InfoPerifericos As tInfoPerifericos




Public Sub RecogeInfoPerifericos()
    If GetSystemMetrics(SM_CMOUSEBUTTONS) > 0 Then
        InfoPerifericos.Mouse_Presente = True
        InfoPerifericos.Mouse_Botones = GetSystemMetrics(SM_CMOUSEBUTTONS)
    End If
    If GetSystemMetrics(SM_MOUSEWHEELPRESENT) > 0 Then
    	InfoPerifericos.Mouse_RuedaVertical = True
    End If
    If GetSystemMetrics(SM_MOUSEHORIZONTALWHEELPRESENT) > 0 Then
    	InfoPerifericos.Mouse_RuedaHorizontal = True
    End If
End Sub


            
