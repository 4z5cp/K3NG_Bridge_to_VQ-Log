; ============================================================================
; K3NG Bridge for VQ-Log
; Constants.pbi - Константы и перечисления
; ============================================================================

; === DDE Constants ===
; Определяем только те константы, которых нет в PureBasic
CompilerIf Not Defined(APPCLASS_STANDARD, #PB_Constant)
  #APPCLASS_STANDARD = 0
CompilerEndIf

CompilerIf Not Defined(CP_WINANSI, #PB_Constant)
  #CP_WINANSI = 1004
CompilerEndIf

CompilerIf Not Defined(DNS_REGISTER, #PB_Constant)
  #DNS_REGISTER = 1
CompilerEndIf

CompilerIf Not Defined(DNS_UNREGISTER, #PB_Constant)
  #DNS_UNREGISTER = 2
CompilerEndIf

CompilerIf Not Defined(DMLERR_NO_ERROR, #PB_Constant)
  #DMLERR_NO_ERROR = 0
CompilerEndIf

; DDE Transaction Types
CompilerIf Not Defined(XTYP_CONNECT, #PB_Constant)
  #XTYP_CONNECT = $1062
CompilerEndIf

CompilerIf Not Defined(XTYP_DISCONNECT, #PB_Constant)
  #XTYP_DISCONNECT = $10C2
CompilerEndIf

CompilerIf Not Defined(XTYP_REQUEST, #PB_Constant)
  #XTYP_REQUEST = $20B0
CompilerEndIf

CompilerIf Not Defined(XTYP_POKE, #PB_Constant)
  #XTYP_POKE = $4090
CompilerEndIf

CompilerIf Not Defined(XTYP_ADVSTART, #PB_Constant)
  #XTYP_ADVSTART = $1030
CompilerEndIf

CompilerIf Not Defined(XTYP_ADVSTOP, #PB_Constant)
  #XTYP_ADVSTOP = $8040
CompilerEndIf

CompilerIf Not Defined(XTYP_ADVREQ, #PB_Constant)
  #XTYP_ADVREQ = $2022
CompilerEndIf

CompilerIf Not Defined(XTYP_WILDCONNECT, #PB_Constant)
  #XTYP_WILDCONNECT = $2062
CompilerEndIf

CompilerIf Not Defined(XTYP_ADVDATA, #PB_Constant)
  #XTYP_ADVDATA = $4010
CompilerEndIf

; DDE Clipboard Formats
CompilerIf Not Defined(CF_TEXT, #PB_Constant)
  #CF_TEXT = 1
CompilerEndIf

; DDE Acknowledgement
CompilerIf Not Defined(DDE_FACK, #PB_Constant)
  #DDE_FACK = $8000
CompilerEndIf

; DDE Callback Filter Flags
CompilerIf Not Defined(APPCMD_CLIENTONLY, #PB_Constant)
  #APPCMD_CLIENTONLY = $00000010
CompilerEndIf

CompilerIf Not Defined(APPCLASS_STANDARD, #PB_Constant)
  #APPCLASS_STANDARD = $00000000
CompilerEndIf

; === Windows Constants ===
#ERROR_ALREADY_EXISTS = 183

; === Application Constants ===
#APP_NAME = "K3NG Bridge for VQ-Log"
#APP_VERSION = "1.0"
#CONFIG_FILE = "k3ng_bridge.ini"

; === Default Values ===
#DEFAULT_IP = "192.168.1.100"
#DEFAULT_PORT = 23
#DEFAULT_MODE = 2
#DEFAULT_POLL_INTERVAL = 1000
#MIN_POLL_INTERVAL = 1000     ; Минимум 1 секунда
#MAX_POLL_INTERVAL = 10000    ; Максимум 10 секунд
#CONNECT_TIMEOUT = 2000        ; Таймаут подключения 2 секунды (для быстрого закрытия)
#RECONNECT_DELAY = 3000        ; Задержка между попытками подключения (3 секунды)

; === Mode Constants ===
#MODE_CONTROLLER_TO_LOG = 0
#MODE_LOG_TO_CONTROLLER = 1
#MODE_BIDIRECTIONAL = 2

; === Window Enumeration ===
Enumeration Windows
  #MainWindow
EndEnumeration

; === Gadget Enumeration ===
Enumeration Gadgets
  ; Connection frame
  #FrameConnection
  #LabelIP
  #StringIP
  #LabelPort
  #StringPort
  #LabelMode
  #ComboMode

  ; Status frame
  #FrameStatus
  #LabelAzText
  #LabelAzValue
  #LabelElText
  #LabelElValue
  #LabelTCPStatus
  #LabelDDEStatus
  
  ; Manual control frame
  #FrameManual
  #LabelManualAz
  #StringManualAz
  #LabelManualAzDeg
  #LabelManualEl
  #StringManualEl
  #LabelManualElDeg
  #ButtonGo
  #ButtonStop

  ; Settings frame
  #FrameSettings
  #LabelPollInterval
  #StringPollInterval
  #LabelPollMs
  #ButtonApplyInterval
  #CheckStartMinimized

  ; Log frame
  #FrameLog
  #ListLog
EndEnumeration

; === Timer Enumeration ===
Enumeration Timers
  #TimerPoll
EndEnumeration