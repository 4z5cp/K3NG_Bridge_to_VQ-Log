; ============================================================================
; K3NG Bridge for VQ-Log
; Procedures.pbi - Процедуры и функции
; ============================================================================

; === Structures ===
Structure AppConfig
  K3ngIP.s
  K3ngPort.i
  Mode.i
  PollInterval.i
  Connected.i
  WinX.i
  WinY.i
  StartMinimized.i
EndStructure

; === Global Variables ===
Global Config.AppConfig
Global DDEInst.l = 0
Global hszService.l = 0
Global hszTopic.l = 0
Global hszItemAz.l = 0
Global hszItemEl.l = 0
Global TCPConnection.i = 0
Global CurrentAzimuth.i = 0
Global CurrentElevation.i = 0
Global TargetAzimuth.i = -1
Global Mutex.i = 0

; === Procedure Declarations ===
Declare LogMsg(msg.s)
Declare ShowError(title.s, message.s)
Declare LoadConfig()
Declare SaveConfig()
Declare.i ConnectToK3NG()
Declare DisconnectK3NG()
Declare.s SendK3NGCommand(cmd.s)
Declare PollK3NGPosition()
Declare RotateToAzimuth(az.i)
Declare RotateToElevation(el.i)
Declare RotateToPosition(az.i, el.i)
Declare StopRotation()
Declare.i InitDDEServer()
Declare CleanupDDEServer()
Declare UpdateStatus()
Declare HandleConnectButton()
Declare HandleModeChange()
Declare HandleGoButton()
Declare HandleStopButton()
Declare HandleApplyInterval()
Declare HandleStartMinimizedToggle()
Declare HandleTimer()
Declare.i CheckSingleInstance()
Declare ReleaseSingleInstance()

; === Windows API Import ===
Import "user32.lib"
  DdeInitializeW(pidInst, pfnCallback, afCmd, ulRes)
  DdeUninitialize(idInst)
  DdeCreateStringHandleW(idInst, psz.p-unicode, iCodePage)
  DdeFreeStringHandle(idInst, hsz)
  DdeNameService(idInst, hsz1, hsz2, afCmd)
  DdeCreateDataHandle(idInst, pSrc, cb, cbOff, hszItem, wFmt, afCmd)
  DdeGetData(hData, pDst, cbMax, cbOff)
  DdeFreeDataHandle(hData)
  DdeAccessData(hData, pcbDataSize)
  DdeUnaccessData(hData)
EndImport

Import "kernel32.lib"
  CreateMutex_(lpMutexAttributes, bInitialOwner, lpName.p-unicode)
  CloseHandle_(hObject)
  GetLastError_()
EndImport

; ============================================================================
; LOGGING
; ============================================================================
Procedure LogMsg(msg.s)
  Protected timestamp.s = FormatDate("%hh:%ii:%ss", Date())
  If IsGadget(#ListLog)
    AddGadgetItem(#ListLog, -1, timestamp + " " + msg)
    SetGadgetState(#ListLog, CountGadgetItems(#ListLog) - 1)
    SendMessage_(GadgetID(#ListLog), #LB_SETTOPINDEX, CountGadgetItems(#ListLog) - 1, 0)
  EndIf
EndProcedure

; ============================================================================
; HELPER FUNCTIONS
; ============================================================================
Procedure ShowError(title.s, message.s)
  Protected msgW.i = 280, msgH.i = 130
  Protected winX.i, winY.i, winW.i, winH.i
  Protected msgX.i, msgY.i
  Protected msgWindow.i, btnOK.i, txtMsg.i
  Protected event.i
  
  winX = WindowX(#MainWindow)
  winY = WindowY(#MainWindow)
  winW = WindowWidth(#MainWindow)
  winH = WindowHeight(#MainWindow)
  
  msgX = winX + (winW - msgW) / 2
  msgY = winY + (winH - msgH) / 2
  
  msgWindow = OpenWindow(#PB_Any, msgX, msgY, msgW, msgH, title, #PB_Window_SystemMenu, WindowID(#MainWindow))
  If msgWindow
    txtMsg = TextGadget(#PB_Any, 10, 15, msgW - 20, 60, message)
    btnOK = ButtonGadget(#PB_Any, (msgW - 80) / 2, msgH - 45, 80, 30, "OK")
    
    SetActiveWindow(msgWindow)
    
    Repeat
      event = WaitWindowEvent()
      If event = #PB_Event_Gadget And EventGadget() = btnOK
        Break
      EndIf
      If event = #PB_Event_CloseWindow And EventWindow() = msgWindow
        Break
      EndIf
    ForEver
    
    CloseWindow(msgWindow)
  EndIf
EndProcedure

; ============================================================================
; CONFIGURATION
; ============================================================================
Procedure LoadConfig()
  Protected file.s = GetCurrentDirectory() + #CONFIG_FILE

  If OpenPreferences(file)
    Config\K3ngIP = ReadPreferenceString("K3ngIP", #DEFAULT_IP)
    Config\K3ngPort = ReadPreferenceLong("K3ngPort", #DEFAULT_PORT)
    Config\Mode = ReadPreferenceLong("Mode", #DEFAULT_MODE)
    Config\PollInterval = ReadPreferenceLong("PollInterval", #DEFAULT_POLL_INTERVAL)
    Config\WinX = ReadPreferenceLong("WinX", -1)
    Config\WinY = ReadPreferenceLong("WinY", -1)
    Config\StartMinimized = ReadPreferenceLong("StartMinimized", 0)
    ClosePreferences()
  Else
    Config\K3ngIP = #DEFAULT_IP
    Config\K3ngPort = #DEFAULT_PORT
    Config\Mode = #DEFAULT_MODE
    Config\PollInterval = #DEFAULT_POLL_INTERVAL
    Config\WinX = -1
    Config\WinY = -1
    Config\StartMinimized = 0
  EndIf
EndProcedure

Procedure SaveConfig()
  Protected file.s = GetCurrentDirectory() + #CONFIG_FILE

  If IsWindow(#MainWindow)
    Config\WinX = WindowX(#MainWindow)
    Config\WinY = WindowY(#MainWindow)
  EndIf

  If CreatePreferences(file)
    WritePreferenceString("K3ngIP", Config\K3ngIP)
    WritePreferenceLong("K3ngPort", Config\K3ngPort)
    WritePreferenceLong("Mode", Config\Mode)
    WritePreferenceLong("PollInterval", Config\PollInterval)
    WritePreferenceLong("WinX", Config\WinX)
    WritePreferenceLong("WinY", Config\WinY)
    WritePreferenceLong("StartMinimized", Config\StartMinimized)
    ClosePreferences()
  EndIf
EndProcedure

; ============================================================================
; TCP FUNCTIONS
; ============================================================================
Procedure ConnectThread(*dummy)
  Protected conn.i
  
  conn = OpenNetworkConnection(Config\K3ngIP, Config\K3ngPort, #PB_Network_TCP | #PB_Network_IPv4)
  
  If conn
    TCPConnection = conn
    Config\Connected = #True
  Else
    Config\Connected = #False
  EndIf
EndProcedure

Procedure.i ConnectToK3NG()
  Protected startTime.i
  Protected threadID.i

  If TCPConnection
    CloseNetworkConnection(TCPConnection)
    TCPConnection = 0
  EndIf

  Config\Connected = #False
  LogMsg("TCP: Connecting to " + Config\K3ngIP + ":" + Str(Config\K3ngPort) + "...")

  threadID = CreateThread(@ConnectThread(), 0)

  If Not threadID
    LogMsg("TCP: Thread creation error")
    ProcedureReturn #False
  EndIf

  startTime = ElapsedMilliseconds()
  While IsThread(threadID) And (ElapsedMilliseconds() - startTime) < #CONNECT_TIMEOUT
    Delay(50)
    While WindowEvent() : Wend
  Wend

  If IsThread(threadID)
    KillThread(threadID)
    Config\Connected = #False
    LogMsg("TCP: Connection timeout")
    ShowError("Error", "Connection timeout to " + Config\K3ngIP + #CRLF$ + "Host not responding.")
    ProcedureReturn #False
  ElseIf Config\Connected
    LogMsg("TCP: Connected")
    ProcedureReturn #True
  Else
    LogMsg("TCP: Connection failed")
    ShowError("Error", "Failed to connect to " + Config\K3ngIP + ":" + Str(Config\K3ngPort))
    ProcedureReturn #False
  EndIf
EndProcedure

Procedure DisconnectK3NG()
  If TCPConnection
    CloseNetworkConnection(TCPConnection)
    TCPConnection = 0
    Config\Connected = #False
    LogMsg("TCP: Disconnected")
  EndIf
EndProcedure

Procedure.s SendK3NGCommand(cmd.s)
  Protected *buffer, received.s = "", len.i
  Protected timeout.i = 1000, startTime.i
  Protected tempStr.s

  If Not TCPConnection
    ProcedureReturn ""
  EndIf

  ; Выделяем буфер для приёма данных
  *buffer = AllocateMemory(256)
  If Not *buffer
    LogMsg("TCP TX Error: Memory allocation failed")
    ProcedureReturn ""
  EndIf

  tempStr = cmd + #CR$
  If SendNetworkString(TCPConnection, tempStr, #PB_Ascii) = 0
    LogMsg("TCP TX Error: " + cmd)
    FreeMemory(*buffer)
    ProcedureReturn ""
  EndIf

  LogMsg("TCP TX: " + cmd)

  startTime = ElapsedMilliseconds()
  Repeat
    len = ReceiveNetworkData(TCPConnection, *buffer, 255)
    If len > 0
      tempStr = PeekS(*buffer, len, #PB_Ascii)
      received + tempStr
      If FindString(received, #CR$) Or FindString(received, #LF$)
        Break
      EndIf
    EndIf
    Delay(10)
  Until ElapsedMilliseconds() - startTime > timeout

  FreeMemory(*buffer)

  received = Trim(ReplaceString(ReplaceString(received, #CR$, ""), #LF$, ""))
  If received <> ""
    LogMsg("TCP RX: " + received)
  EndIf

  ProcedureReturn received
EndProcedure

Procedure PollK3NGPosition()
  Protected response.s, azPos.i, elPos.i
  
  If Not TCPConnection Or Config\Mode = #MODE_LOG_TO_CONTROLLER
    ProcedureReturn
  EndIf
  
  response = SendK3NGCommand("C2")
  
  If FindString(response, "AZ=")
    azPos = Val(Mid(response, FindString(response, "AZ=") + 3, 3))
    If azPos >= 0 And azPos <= 360
      CurrentAzimuth = azPos
    EndIf
    
    If FindString(response, "EL=")
      elPos = Val(Mid(response, FindString(response, "EL=") + 3, 3))
      If elPos >= 0 And elPos <= 180
        CurrentElevation = elPos
      EndIf
    EndIf
  EndIf
EndProcedure

; ============================================================================
; ROTATOR CONTROL
; ============================================================================
Procedure RotateToAzimuth(az.i)
  Protected cmd.s

  If Not TCPConnection
    ProcedureReturn
  EndIf

  If az >= 0 And az <= 360
    cmd = "M" + RSet(Str(az), 3, "0")
    SendK3NGCommand(cmd)
    TargetAzimuth = az
    LogMsg("Rotate Az: " + Str(az) + "°")
  EndIf
EndProcedure

Procedure RotateToElevation(el.i)
  Protected cmd.s

  If Not TCPConnection
    ProcedureReturn
  EndIf

  If el >= 0 And el <= 180
    cmd = "W" + RSet(Str(CurrentAzimuth), 3, "0") + " " + RSet(Str(el), 3, "0")
    SendK3NGCommand(cmd)
    LogMsg("Rotate El: " + Str(el) + "°")
  EndIf
EndProcedure

Procedure RotateToPosition(az.i, el.i)
  Protected cmd.s

  If Not TCPConnection
    ProcedureReturn
  EndIf

  If az >= 0 And az <= 360 And el >= 0 And el <= 180
    cmd = "W" + RSet(Str(az), 3, "0") + " " + RSet(Str(el), 3, "0")
    SendK3NGCommand(cmd)
    TargetAzimuth = az
    LogMsg("Rotate: Az=" + Str(az) + "° El=" + Str(el) + "°")
  EndIf
EndProcedure

Procedure StopRotation()
  If TCPConnection
    SendK3NGCommand("S")
    LogMsg("STOP command sent")
  EndIf
EndProcedure

; ============================================================================
; DDE SERVER
; ============================================================================
ProcedureDLL.l DDECallback(uType.l, uFmt.l, hconv.l, hsz1.l, hsz2.l, hdata.l, dwData1.l, dwData2.l)
  Protected result.l = 0
  Protected *data, dataSize.l, dataStr.s
  Protected cmdValue.i

  Select uType
    Case #XTYP_CONNECT
      LogMsg("DDE: Client connected")
      result = #True

    Case #XTYP_REQUEST
      If uFmt = #CF_TEXT
        dataStr = Str(CurrentAzimuth) + Chr(0)
        result = DdeCreateDataHandle(DDEInst, @dataStr, Len(dataStr) + 1, 0, hsz2, #CF_TEXT, 0)
        LogMsg("DDE: Request AZ=" + Str(CurrentAzimuth))
      EndIf

    Case #XTYP_POKE
      If hdata And (Config\Mode = #MODE_LOG_TO_CONTROLLER Or Config\Mode = #MODE_BIDIRECTIONAL)
        *data = DdeAccessData(hdata, @dataSize)
        If *data
          dataStr = PeekS(*data, dataSize, #PB_Ascii)
          DdeUnaccessData(hdata)

          LogMsg("DDE: Received: " + dataStr)

          If Left(UCase(dataStr), 3) = "GA:"
            cmdValue = Val(Mid(dataStr, 4))
            If cmdValue >= 0 And cmdValue <= 360
              RotateToAzimuth(cmdValue)
            EndIf
          ElseIf Left(UCase(dataStr), 3) = "GE:"
            cmdValue = Val(Mid(dataStr, 4))
            If cmdValue >= 0 And cmdValue <= 180
              RotateToElevation(cmdValue)
            EndIf
          EndIf
        EndIf
        result = #DDE_FACK
      EndIf

    Case #XTYP_ADVSTART
      LogMsg("DDE: Advise loop started")
      result = #True

    Case #XTYP_ADVSTOP
      LogMsg("DDE: Advise loop stopped")
      result = #True

    Case #XTYP_DISCONNECT
      LogMsg("DDE: Client disconnected")
      result = 0

  EndSelect

  ProcedureReturn result
EndProcedure

Procedure.i InitDDEServer()
  Protected result.l

  result = DdeInitializeW(@DDEInst, @DDECallback(), #APPCLASS_STANDARD, 0)
  If result <> #DMLERR_NO_ERROR
    LogMsg("DDE: Initialization error, code " + Str(result))
    ProcedureReturn #False
  EndIf

  hszService = DdeCreateStringHandleW(DDEInst, "ARSWIN", #CP_WINANSI)
  hszTopic = DdeCreateStringHandleW(DDEInst, "RCI", #CP_WINANSI)
  hszItemAz = DdeCreateStringHandleW(DDEInst, "AZIMUTH", #CP_WINANSI)
  hszItemEl = DdeCreateStringHandleW(DDEInst, "ELEVATION", #CP_WINANSI)

  If DdeNameService(DDEInst, hszService, 0, #DNS_REGISTER)
    LogMsg("DDE: Server started (ARSWIN|RCI)")
    ProcedureReturn #True
  Else
    LogMsg("DDE: Service registration error")
    ProcedureReturn #False
  EndIf
EndProcedure

Procedure CleanupDDEServer()
  If DDEInst
    DdeNameService(DDEInst, hszService, 0, #DNS_UNREGISTER)
    DdeFreeStringHandle(DDEInst, hszService)
    DdeFreeStringHandle(DDEInst, hszTopic)
    DdeFreeStringHandle(DDEInst, hszItemAz)
    DdeFreeStringHandle(DDEInst, hszItemEl)
    DdeUninitialize(DDEInst)
    DDEInst = 0
    LogMsg("DDE: Server stopped")
  EndIf
EndProcedure

; ============================================================================
; EVENT HANDLERS
; ============================================================================
Procedure HandleConnectButton()
  Config\K3ngIP = GetGadgetText(#StringIP)
  Config\K3ngPort = Val(GetGadgetText(#StringPort))
  Config\Mode = GetGadgetState(#ComboMode)
  SaveConfig()
  
  If Config\Connected
    DisconnectK3NG()
    SetGadgetText(#ButtonConnect, "Connect")
  Else
    If ConnectToK3NG()
      SetGadgetText(#ButtonConnect, "Disconnect")
    EndIf
  EndIf
  UpdateStatus()
EndProcedure

Procedure HandleModeChange()
  Config\Mode = GetGadgetState(#ComboMode)
  SaveConfig()
  LogMsg("Mode: " + GetGadgetItemText(#ComboMode, Config\Mode))
EndProcedure

Procedure HandleGoButton()
  Protected az.i = Val(GetGadgetText(#StringManualAz))
  Protected el.i = Val(GetGadgetText(#StringManualEl))
  RotateToPosition(az, el)
EndProcedure

Procedure HandleStopButton()
  StopRotation()
EndProcedure

Procedure HandleApplyInterval()
  Protected newInterval.i

  newInterval = Val(GetGadgetText(#StringPollInterval))

  ; Check range: minimum 100 ms, maximum 10000 ms (10 sec)
  If newInterval < 100
    newInterval = 100
    SetGadgetText(#StringPollInterval, "100")
  ElseIf newInterval > 10000
    newInterval = 10000
    SetGadgetText(#StringPollInterval, "10000")
  EndIf

  ; Apply new interval
  If newInterval <> Config\PollInterval
    Config\PollInterval = newInterval
    SaveConfig()

    ; Restart timer with new interval
    RemoveWindowTimer(#MainWindow, #TimerPoll)
    AddWindowTimer(#MainWindow, #TimerPoll, Config\PollInterval)

    LogMsg("Poll interval set: " + Str(Config\PollInterval) + " ms")
  EndIf
EndProcedure

Procedure HandleStartMinimizedToggle()
  Config\StartMinimized = GetGadgetState(#CheckStartMinimized)
  SaveConfig()
  If Config\StartMinimized
    LogMsg("Start minimized: enabled")
  Else
    LogMsg("Start minimized: disabled")
  EndIf
EndProcedure

Procedure HandleTimer()
  PollK3NGPosition()
  UpdateStatus()
EndProcedure

; ============================================================================
; SINGLE INSTANCE
; ============================================================================
Procedure.i CheckSingleInstance()
  Mutex = CreateMutex(#Null, #True, "Global\K3NG_Bridge_Mutex")

  If Mutex = 0
    ProcedureReturn #False
  EndIf

  If GetLastError_() = #ERROR_ALREADY_EXISTS
    ; Another instance is already running
    CloseHandle_(Mutex)
    Mutex = 0
    ProcedureReturn #False
  EndIf

  ProcedureReturn #True
EndProcedure

Procedure ReleaseSingleInstance()
  If Mutex
    CloseHandle_(Mutex)
    Mutex = 0
  EndIf
EndProcedure

; ============================================================================
; GUI UPDATE
; ============================================================================
Procedure UpdateStatus()
  SetGadgetText(#LabelAzValue, Str(CurrentAzimuth) + "°")
  SetGadgetText(#LabelElValue, Str(CurrentElevation) + "°")
  
  If Config\Connected
    SetGadgetText(#LabelTCPStatus, "TCP: Подключено")
    SetGadgetColor(#LabelTCPStatus, #PB_Gadget_FrontColor, RGB(0, 128, 0))
  Else
    SetGadgetText(#LabelTCPStatus, "TCP: Отключено")
    SetGadgetColor(#LabelTCPStatus, #PB_Gadget_FrontColor, RGB(192, 0, 0))
  EndIf
  
  If DDEInst
    SetGadgetText(#LabelDDEStatus, "DDE: ARSWIN|RCI")
    SetGadgetColor(#LabelDDEStatus, #PB_Gadget_FrontColor, RGB(0, 128, 0))
  Else
    SetGadgetText(#LabelDDEStatus, "DDE: Не запущен")
    SetGadgetColor(#LabelDDEStatus, #PB_Gadget_FrontColor, RGB(192, 0, 0))
  EndIf
EndProcedure