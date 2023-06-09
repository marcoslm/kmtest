Attribute VB_Name = "HookTecladoMouse"
Option Explicit

'--------------------------------------------------------------
' Hooks para el teclado y mouse
'--------------------------------------------------------------
' Por Marcos López Merayo (11-2009) Para Bluebit
' Portado a twinBasic 10-2022
'
'
'


' Para guardar el gancho creado con SetWindowsHookEx
Private mHook As Long ' Este para el teclado
Private mmHook As Long ' Y este para el ratón
' Para indicar a SetWindowsHookEx que tipo de hook queremos instalar
Private Const WH_KEYBOARD_LL As Long = 13&
' Lo mimso pero para el ratón
Private Const WH_MOUSE_LL As Long = 14&
'
Private Type tagKBDLLHOOKSTRUCT
    vkCode      As Long
    scanCode    As Long
    flags       As Long
    time        As Long
    dwExtraInfo As Long
End Type
Private Type tagPOINT
    X As Long
    Y As Long
End Type
Private Type MSLLHOOKSTRUCT
  pt            As tagPOINT 'POINTAPI
  mouseData     As Long
  flags         As Long
  time          As Long
  dwExtraInfo   As Long
End Type

' Constantes para eventos del teclado
'--------------------------------------------------------------
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105

' Constantes de teclado
'--------------------------------------------------------------
Public Const VK_LBUTTON = &H1
Public Const VK_RBUTTON = &H2
Public Const VK_CANCEL = &H3
Public Const VK_MBUTTON = &H4
Public Const VK_BACK = &H8
Public Const VK_TAB = &H9
Public Const VK_CLEAR = &HC
Public Const VK_RETURN = &HD
Public Const VK_SHIFT = &H10
Public Const VK_CONTROL = &H11
Public Const VK_MENU = &H12 ' Alt
Public Const VK_PAUSE = &H13
Public Const VK_CAPITAL = &H14
Public Const VK_ESCAPE = &H1B
Public Const VK_SPACE = &H20
Public Const VK_PRIOR = &H21
Public Const VK_NEXT = &H22
Public Const VK_END = &H23
Public Const VK_HOME = &H24
Public Const VK_LEFT = &H25
Public Const VK_UP = &H26
Public Const VK_RIGHT = &H27
Public Const VK_DOWN = &H28
Public Const VK_Select = &H29
Public Const VK_PRINT = &H2A
Public Const VK_EXECUTE = &H2B
Public Const VK_SNAPSHOT = &H2C
Public Const VK_Insert = &H2D
Public Const VK_Delete = &H2E
Public Const VK_HELP = &H2F
Public Const VK_0 = &H30
Public Const VK_1 = &H31
Public Const VK_2 = &H32
Public Const VK_3 = &H33
Public Const VK_4 = &H34
Public Const VK_5 = &H35
Public Const VK_6 = &H36
Public Const VK_7 = &H37
Public Const VK_8 = &H38
Public Const VK_9 = &H39
Public Const VK_A = &H41
Public Const VK_B = &H42
Public Const VK_C = &H43
Public Const VK_D = &H44
Public Const VK_E = &H45
Public Const VK_F = &H46
Public Const VK_G = &H47
Public Const VK_H = &H48
Public Const VK_I = &H49
Public Const VK_J = &H4A
Public Const VK_K = &H4B
Public Const VK_L = &H4C
Public Const VK_M = &H4D
Public Const VK_N = &H4E
Public Const VK_O = &H4F
Public Const VK_P = &H50
Public Const VK_Q = &H51
Public Const VK_R = &H52
Public Const VK_S = &H53
Public Const VK_T = &H54
Public Const VK_U = &H55
Public Const VK_V = &H56
Public Const VK_W = &H57
Public Const VK_X = &H58
Public Const VK_Y = &H59
Public Const VK_Z = &H5A
Public Const VK_STARTKEY = &H5B
Public Const VK_CONTEXTKEY = &H5D
Public Const VK_NUMPAD0 = &H60
Public Const VK_NUMPAD1 = &H61
Public Const VK_NUMPAD2 = &H62
Public Const VK_NUMPAD3 = &H63
Public Const VK_NUMPAD4 = &H64
Public Const VK_NUMPAD5 = &H65
Public Const VK_NUMPAD6 = &H66
Public Const VK_NUMPAD7 = &H67
Public Const VK_NUMPAD8 = &H68
Public Const VK_NUMPAD9 = &H69
Public Const VK_MULTIPLY = &H6A
Public Const VK_ADD = &H6B
Public Const VK_SEPARATOR = &H6C
Public Const VK_SUBTRACT = &H6D
Public Const VK_DECIMAL = &H6E
Public Const VK_DIVIDE = &H6F
Public Const VK_F1 = &H70
Public Const VK_F2 = &H71
Public Const VK_F3 = &H72
Public Const VK_F4 = &H73
Public Const VK_F5 = &H74
Public Const VK_F6 = &H75
Public Const VK_F7 = &H76
Public Const VK_F8 = &H77
Public Const VK_F9 = &H78
Public Const VK_F10 = &H79
Public Const VK_F11 = &H7A
Public Const VK_F12 = &H7B
Public Const VK_F13 = &H7C
Public Const VK_F14 = &H7D
Public Const VK_F15 = &H7E
Public Const VK_F16 = &H7F
Public Const VK_F17 = &H80
Public Const VK_F18 = &H81
Public Const VK_F19 = &H82
Public Const VK_F20 = &H83
Public Const VK_F21 = &H84
Public Const VK_F22 = &H85
Public Const VK_F23 = &H86
Public Const VK_F24 = &H87
Public Const VK_NUMLOCK = &H90
Public Const VK_OEM_SCROLL = &H91
Public Const VK_OEM_1 = &HBA
Public Const VK_OEM_PLUS = &HBB
Public Const VK_OEM_COMMA = &HBC
Public Const VK_OEM_MINUS = &HBD
Public Const VK_OEM_PERIOD = &HBE
Public Const VK_OEM_2 = &HBF
Public Const VK_OEM_3 = &HC0
Public Const VK_OEM_4 = &HDB
Public Const VK_OEM_5 = &HDC
Public Const VK_OEM_6 = &HDD
Public Const VK_OEM_7 = &HDE
Public Const VK_OEM_8 = &HDF
Public Const VK_ICO_F17 = &HE0
Public Const VK_ICO_F18 = &HE1
Public Const VK_OEM102 = &HE2
Public Const VK_ICO_HELP = &HE3
Public Const VK_ICO_00 = &HE4
Public Const VK_ICO_CLEAR = &HE6
Public Const VK_OEM_RESET = &HE9
Public Const VK_OEM_JUMP = &HEA
Public Const VK_OEM_PA1 = &HEB
Public Const VK_OEM_PA2 = &HEC
Public Const VK_OEM_PA3 = &HED
Public Const VK_OEM_WSCTRL = &HEE
Public Const VK_OEM_CUSEL = &HEF
Public Const VK_OEM_ATTN = &HF0
Public Const VK_OEM_FINNISH = &HF1
Public Const VK_OEM_COPY = &HF2
Public Const VK_OEM_AUTO = &HF3
Public Const VK_OEM_ENLW = &HF4
Public Const VK_OEM_BACKTAB = &HF5
Public Const VK_ATTN = &HF6
Public Const VK_CRSEL = &HF7
Public Const VK_EXSEL = &HF8
Public Const VK_EREOF = &HF9
Public Const VK_PLAY = &HFA
Public Const VK_ZOOM = &HFB
Public Const VK_NONAME = &HFC
Public Const VK_PA1 = &HFD
Public Const VK_OEM_CLEAR = &HFE

' Constantes para el mouse
'--------------------------------------------------------------
Public Const WM_MOUSEWHEEL = &H20A
'Private Const MK_CONTROL = &H8
Public Const MK_LBUTTON = &H201
'Private Const MK_MBUTTON = &H10
Public Const MK_RBUTTON = &H204
'Private Const MK_SHIFT = &H4
'Private Const MK_XBUTTON1 = &H20
'Private Const MK_XBUTTON2 = &H40
Public Const MK_MOVIMIENTO = &H200

'
Private Const LLKHF_ALTDOWN As Long = &H20&
'
' Códigos para los hooks (la acción a tomar en el hook del teclado)
Private Const HC_ACTION As Long = 0&


'--------------------------------------------------------------
' Funciones de la API de Windows
'--------------------------------------------------------------

' Para asignar un hook
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, _
                                                                                  ByVal dwThreadId As Long) As Long
' Para quitar el hok creado con SetWindowsHookEx
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
    
' Para llamar al siguiente hook
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
' Para saber si se ha pulsado en una tecla
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

' Teclas especiales    
Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

' Para copiar la estructura en un long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)



' La función a usar para el hook de teclado
'--------------------------------------------------------------
' Cuando el hook está activo, esta función se ejecutará al 
' detectarse una pulsación de teclado.
Public Function LLKeyBoardProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim pkbhs As tagKBDLLHOOKSTRUCT

    ' Copiar el parámetro en la estructura
    CopyMemory pkbhs, ByVal lParam, Len(pkbhs)
    
    If nCode = HC_ACTION Then
        ' Capturamos tecla en evento key_down
        If wParam = WM_KEYDOWN Then
            ' Llamamos a la función EventoTecla del formulario para pasarle la tecla pulsada
            frmMain.EventoTecla ObtenerTecla(pkbhs.vkCode), pkbhs.vkCode
        End If
        ' Capturamos tecla en evento key_up
        If wParam = WM_KEYUP Then
            ' Llamamos a la función EventoTecla del formulario para pasarle la tecla pulsada
            Select Case pkbhs.vkCode
            Case 164 'Alt
                frmMain.EventoTecla ObtenerTecla(164), 164
            Case 165  'AltGr
                frmMain.EventoTecla ObtenerTecla(165), 165
            'Case Else
            '    frmMain.EventoTecla pkbhs.vkCode, 0
            End Select
        End If
    End If
    ' Devolvemos el control continuando con la ejecución justo después de la interrupción
    LLKeyBoardProc = CallNextHookEx(mHook, nCode, wParam, lParam)
End Function

' La función a usar para el hoock del mouse
'--------------------------------------------------------------
Public Function LLMouseProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim pkbhs As MSLLHOOKSTRUCT

    ' copiar el parámetro en la estructura
    CopyMemory pkbhs, ByVal lParam, Len(pkbhs)
    '
    If nCode = HC_ACTION Then
        ' Capturamos evento
            frmMain.EventoMouse ObtenerTecla(wParam), pkbhs.pt.X, pkbhs.pt.Y
    End If
    ' Devolvemos el control continuando con la ejecución justo después de la interrupción
    LLMouseProc = CallNextHookEx(mmHook, nCode, wParam, ByVal lParam)
End Function

' Función para instalar los hooks
'--------------------------------------------------------------
Public Sub InstalarHooks(ByVal hMod As Long)
    ' instalar el gancho para el teclado y ratón
    ' hMod será el valor de App.hInstance de la aplicación
    mHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LLKeyBoardProc, hMod, 0&)
    mmHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf LLMouseProc, hMod, 0&)
End Sub

' Función para desinstalar los hooks
'--------------------------------------------------------------
Public Sub DesinstalarHooks()
    ' Es importante hacer esto antes de finalizar la aplicación,
    ' normalmente en el evento Unload o QueryUnload
    If mHook <> 0 Then
        UnhookWindowsHookEx mHook
    End If
    If mmHook <> 0 Then
        UnhookWindowsHookEx mmHook
    End If
End Sub

Function ObtenerTecla(ByVal X As Integer) As String
    Dim Tecla As String
    Select Case X
        Case MK_MOVIMIENTO                  'movimiento del ratón
            Tecla = "[MOUSE_MOVE]"
        Case WM_MOUSEWHEEL
            Tecla = "[MOUSE_WHEEL]"
        Case 513                            'botón izquierdo del ratón
            Tecla = "[MOUSE_LEFT_BTN_DOWN]"
        Case 514                            'botón izquierdo del ratón levantado
            Tecla = "[MOUSE_LEFT_BTN_UP]"
        Case 516                            'botón derecho del ratón
            Tecla = "[MOUSE_RIGHT_BTN_DOWN]"
        Case 517                            'botón derecho del ratón levantado
            Tecla = "[MOUSE_RIGHT_BTN_UP]"
        Case 519                            'botón medio del ratón
            Tecla = "[MOUSE_MID_BTN_DOWN]"
        Case 520                            'botón medio del ratón
            Tecla = "[MOUSE_MID_BTN_UP]"
        Case 3                              'VK_CANCEL 'break interrumpir
            Tecla = "INTRRUMPIR"
        Case 8                              'VK_BACK
            Tecla = "RETROCESO"
        Case 9                              'VK_TAB
            Tecla = "TAB"
        Case 13                             'VK_RETURN
            Tecla = "INTRO"
        Case 92                             'VK_CLEAR '5 en keypad sin numlook
        Case 19                             'VK_PAUSE 'Pausa
            Tecla = "PAUSA"
        Case 20
            Tecla = "BLOQ.MAY"
        Case 32                             'VK_SPACE
            Tecla = "ESPACIO"
        Case 27                             'VK_ESC 'escape
            Tecla = "ESC"
        Case 33                             'VK_PRIOR
            Tecla = "RE.PAG"
        Case 34                             'VK_NEXT
            Tecla = "AV.PAG"
        Case 35                             'VK_END
            Tecla = "FIN"
        Case 36                             'VK_HOME
            Tecla = "INICIO"
        Case 37                             'VK_LEFT
            Tecla = "IZQ"
        Case 38                             'VK_UP
            Tecla = "ARR"
        Case 39                             'VK_RIGHT
            Tecla = "DER"
        Case 40                             'VK_DOWN
            Tecla = "ABA"
        Case 44                             'Imprimir Pantalla
            Tecla = "IMP.PANT"
        Case 45, VK_Insert
            Tecla = "INS"
        Case 46, VK_Delete
            Tecla = "SUPR"
        Case 48 To 57                       'VK_0 - VK_9
            If Not ShiftPulsado Then        'si no se ha cambiado tecla de shift
                Tecla = Str$(X - 48)        'poner en tecla el nº correspondiente
            Else
                Tecla = Mid$("!""""·$%&/()=", X - 47, 1) 'extraer el caracter correspondiente
            End If
            If AltGr Then
                If X = 49 Then              'alt gr + 1
                    Tecla = "|"
                ElseIf X = 50 Then          'alt gr + 2
                    Tecla = "@"
                ElseIf X = 51 Then          'alt gr + 3
                    Tecla = "#"
                ElseIf X = 54 Then          'alt gr +6
                    Tecla = "¬"
                End If
            End If
        Case 65 To 90                       'letras VK_A - VK_Z
            If BloqMayus Then
                Tecla = IIf(ShiftPulsado, LCase$(Chr(X)), UCase$(Chr(X)))
            Else
                Tecla = IIf(ShiftPulsado, UCase$(Chr(X)), LCase$(Chr(X)))
            End If
        Case 91
            Tecla = "WINDOWS"
        Case 93
            Tecla = "MENU"
        Case 96 To 105                      'numpad VK_NUMPAD0 - VK_NUMPAD9'
            If Not NumLock Then
                Tecla = "NUM." & LTrim$(Str$(X - 96)) 'obtener número correspondiente a teclado numpad
            Else
                Tecla = "NUM." & ObtenerTecla(X - 48) 'obtener valor correspondiente a numpad sin numlock
            End If
        Case 106                            'VK_MULTIPLY
            Tecla = "*"
        Case 107                            'VK_NUMPADADD
            Tecla = "+"
        Case 110                            'VK_NUMPADDECIMAL
            Tecla = "."
        Case 111                            'VK_NUMPADDIVIDE
            Tecla = "/"
        Case 109                            'VK_SUBSTRACKT
            Tecla = "-"
        Case 112 To 123                     'VK_F1 - VK_F12
            Tecla = "F" & X - 111
        Case 144
            Tecla = "BLOQ.NUM"
        Case 145                            'VK_SCROLL 'Bloq Despl
            Tecla = "BLOQ.DESP"
        Case 160
            Tecla = "SHIFT.IZQ"
        Case 161
            Tecla = "SHIFT.DER"
        Case 162
            Tecla = "CTRL.IZQ"
        Case 163
            Tecla = "CTRL.DER"
        Case 164
            Tecla = "ALT"
        Case 165
            Tecla = "ALT.GR"
        Case 172
            Tecla = "NAVEGADOR"
        Case 173
            Tecla = "SILENCIO.ON/OFF"
        Case 180
            Tecla = "EMAIL"
        Case 182
            Tecla = "EXPLORADOR"
        Case 186 '^`
            Tecla = IIf(ShiftPulsado, "^", "`")
            Tecla = IIf(AltGr, "[", Tecla)
        Case 187 '+ *
            Tecla = IIf(ShiftPulsado, "*", "+")
            Tecla = IIf(AltGr, "]", Tecla)
        Case 188 '; ,
            Tecla = IIf(ShiftPulsado, ";", ",")
        Case 189 '- _ )
            Tecla = IIf(ShiftPulsado, "_", "-")
        Case 190 ': .
            Tecla = IIf(ShiftPulsado, ":", ".")
        Case 191 'ç Ç
            Tecla = IIf(ShiftPulsado, "Ç", "ç")
            Tecla = IIf(AltGr, "}", Tecla)
        Case 192 '~ '
            Tecla = IIf(ShiftPulsado, "~", "'")
        Case 219 '? '
            Tecla = IIf(ShiftPulsado, "?", "'")
        Case 220 '| \
            Tecla = IIf(ShiftPulsado, "ª", "º")
            Tecla = IIf(AltGr, "\", Tecla)
        Case 221 '¿ ¡
            Tecla = IIf(ShiftPulsado, "¿", "¡")
        Case 222 ' ¨ ´
            Tecla = IIf(ShiftPulsado, "¨", "´")
            Tecla = IIf(AltGr, "{", Tecla)
        Case 226 ' < >
            Tecla = IIf(ShiftPulsado, ">", "<")
        Case Else
            'Tecla = Trim$(Str$(X))
    End Select
    ' Verifica el tamaño y elimina espacios inecesarios
    If Len(Tecla) > 1 Then
        Tecla = StrFilter(Tecla, " ")
    End If
    ObtenerTecla = Tecla
End Function

' Devuelve True si está pulsado Shift
Private Function ShiftPulsado() As Boolean
    ShiftPulsado = IIf(GetKeyState(16) < 0, True, False) 'VK_SHIFT
End Function

' Devuelve True si está pulsado BloqMayus
Private Function BloqMayus() As Boolean
    BloqMayus = IIf(GetKeyState(20) < 0, True, False)   'VK_CAPSLOCK
End Function

' Devuelve True si está pulsado BloqNum
Private Function NumLock() As Boolean
    NumLock = IIf(GetKeyState(144) < 0, True, False)    'VK_NUMLOCK
End Function

' Devuelve True si está pulsado Alt Gr
Private Function AltGr() As Boolean
    AltGr = IIf(GetKeyState(165) < 0, True, False) 'VK_RMENU
End Function

' Elimina los caracteres indicados de una cadena de texto
Public Function StrFilter(ByVal Text As String, ByVal Chars As String) As String
    ' Eliminamos los caracteres en Chars de la cadena Text
    Dim p As Integer
    For p = 1 To Len(Chars)
        Text = Replace(Text, Mid$(Chars, p, 1), "")
    Next
    StrFilter = Text
End Function
