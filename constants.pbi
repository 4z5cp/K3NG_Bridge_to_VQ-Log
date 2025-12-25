; ============================================================================
; K3NG Bridge for VQ-Log
; Constants.pbi - Константы и перечисления
; ============================================================================

; === DDE Constants ===
#APPCLASS_STANDARD = 0
#CP_WINANSI = 1004
#DNS_REGISTER = 1
#DNS_UNREGISTER = 2
#DMLERR_NO_ERROR = 0

; === Application Constants ===
#APP_NAME = "K3NG Bridge for VQ-Log"
#APP_VERSION = "1.0"
#CONFIG_FILE = "k3ng_bridge.ini"

; === Default Values ===
#DEFAULT_IP = "192.168.1.100"
#DEFAULT_PORT = 23
#DEFAULT_MODE = 2
#DEFAULT_POLL_INTERVAL = 1000
#CONNECT_TIMEOUT = 5000

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
  #ButtonConnect
  
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
  
  ; Log frame
  #FrameLog
  #ListLog
EndEnumeration

; === Timer Enumeration ===
Enumeration Timers
  #TimerPoll
EndEnumeration
