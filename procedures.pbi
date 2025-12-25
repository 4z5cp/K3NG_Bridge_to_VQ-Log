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

; === DDE API Import ===
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

; ============================================================================
; LOGGING
; ============================================================================
Procedure LogMsg(msg.s)
  Protected timestamp.s = FormatDate("%hh:%ii:%ss", Date())
  AddGadgetItem(#ListLog, -1, timestamp + " " + msg)
  SetGadgetState(#ListLog, CountGadgetItems(#ListLog) - 1)
  SendMessage_(GadgetID(#ListLog), #LB_SETTOPINDEX, CountGadgetItems(#ListLog) - 1, 0)
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
    ClosePreferences()
  Else
    Config\K3ngIP = #DEFAULT_IP
    Config\K3ngPort = #DEFAULT_PORT
    Config\Mode = #DEFAULT_MODE
    Config\PollInterval = #DEFAULT_POLL_INTERVAL
    Config\WinX = -1
    Config\WinY = -1
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
  LogMsg("TCP: Подключение к " + Config\K3ngIP + ":" + Str(Config\K3ngPort) + "...")
  
  threadID = CreateThread(@ConnectThread(), 0)
  
  If Not threadID
    LogMsg("TCP: Ошибка создания потока")
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
    LogMsg("TCP: Таймаут подключения")
    ShowError("Ошибка", "Таймаут подключения к " + Config\K3ngIP + #CRLF$ + "Хост не отвечает.")
    ProcedureReturn #False
  ElseIf Config\Connected
    LogMsg("TCP: Подключено")
    ProcedureReturn #True
  Else
    LogMsg("TCP: Ошибка подключения")
    ShowError("Ошибка", "Не удалось подключиться к " + Config\K3ngIP + ":" + Str(Config\K3ngPort))
    ProcedureReturn #False
  EndIf
EndProcedure

Procedure DisconnectK3NG()
  If TCPConnection
    CloseNetworkConnection(TCPConnection)
    TCPConnection = 0
    Config\Connected = #False
    LogMsg("TCP: Отключено")
  EndIf
EndProcedure

Procedure.s SendK3NGCommand(cmd.s)
  Protected buffer.s, received.s = "", len.i
  Protected timeout.i = 1000, startTime.i
  
  If Not TCPConnection
    ProcedureReturn ""
  EndIf
  
  buffer = cmd + #CR$
  If SendNetworkString(TCPConnection, buffer, #PB_Ascii) = 0
    LogMsg("TCP TX Error: " + cmd)
    ProcedureReturn ""
  EndIf
  
  LogMsg("TCP TX: " + cmd)
  
  startTime = ElapsedMilliseconds()
  Repeat
    len = ReceiveNetworkData(TCPConnection, @buffer, 255)
    If len > 0
      buffer = PeekS(@buffer, len, #PB_Ascii)
      received + buffer
      If FindString(received, #CR$) Or FindString(received, #LF$)
        Break
      EndIf
    EndIf
    Delay(10)
  Until ElapsedMilliseconds() - startTime > timeout
  
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
    LogMsg("Rotate Az to: " + Str(az) + "°")
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
    LogMsg("Rotate El to: " + Str(el) + "°")
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
    LogMsg("Rotate to: Az=" + Str(az) + "° El=" + Str(el) + "°")
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
          
          LogMsg("DDE: Poke received: " + dataStr)
          
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
    LogMsg("DDE: Init failed, error " + Str(result))
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
    LogMsg("DDE: Failed to register service")
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
