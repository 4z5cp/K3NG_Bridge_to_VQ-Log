; K3NG Bridge for VQ-Log
; DDE Server (ARSWIN emulation) + TCP Client for K3NG Rotator Controller
; PureBasic 6.x
;
; ВАЖНО: Компилировать с включённой опцией "Create threadsafe executable"
; (Compiler -> Compiler Options -> Create threadsafe executable)

EnableExplicit

; === DDE Constants (only those not built into PureBasic) ===
#APPCLASS_STANDARD = 0
#CP_WINANSI = 1004
#DNS_REGISTER = 1
#DNS_UNREGISTER = 2
#DMLERR_NO_ERROR = 0

; === GUI Enumerations ===
Enumeration Windows
  #MainWindow
EndEnumeration

Enumeration Gadgets
  #IPLabel
  #IPInput
  #PortLabel
  #PortInput
  #ModeLabel
  #ModeCombo
  #ConnectBtn
  #DisconnectBtn
  #StatusFrame
  #AzLabel
  #AzValue
  #ElLabel
  #ElValue
  #TCPStatus
  #DDEStatus
  #LogFrame
  #LogList
  #ManualFrame
  #ManualAzLabel
  #ManualAzInput
  #ManualElLabel
  #ManualElInput
  #ManualGoBtn
  #ManualStopBtn
  #PollTimer
EndEnumeration

; === Structures ===
Structure AppConfig
  K3ngIP.s
  K3ngPort.i
  Mode.i          ; 0=Controller→Log, 1=Log→Controller, 2=Bidirectional
  PollInterval.i  ; ms
  Connected.i
  WinX.i          ; Позиция окна X
  WinY.i          ; Позиция окна Y
EndStructure

; === Globals ===
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
Global gMainWindow.i
Global gLogList.i

; === DDE API ===
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

; === Logging ===
Procedure LogMsg(msg.s)
  Protected timestamp.s = FormatDate("%hh:%ii:%ss", Date())
  AddGadgetItem(gLogList, -1, timestamp + " " + msg)
  SetGadgetState(gLogList, CountGadgetItems(gLogList) - 1)
  SendMessage_(GadgetID(gLogList), #LB_SETTOPINDEX, CountGadgetItems(gLogList) - 1, 0)
EndProcedure

; === Helper Functions ===
Procedure ShowError(title.s, message.s)
  Protected msgW.i = 250, msgH.i = 120
  Protected winX.i, winY.i, winW.i, winH.i
  Protected msgX.i, msgY.i
  Protected msgWindow.i, btnOK.i, txtMsg.i
  Protected event.i
  
  ; Получаем позицию и размер основного окна
  winX = WindowX(#MainWindow)
  winY = WindowY(#MainWindow)
  winW = WindowWidth(#MainWindow)
  winH = WindowHeight(#MainWindow)
  
  ; Вычисляем центр
  msgX = winX + (winW - msgW) / 2
  msgY = winY + (winH - msgH) / 2
  
  ; Создаём окно сообщения
  msgWindow = OpenWindow(#PB_Any, msgX, msgY, msgW, msgH, title, #PB_Window_SystemMenu | #PB_Window_WindowCentered, WindowID(#MainWindow))
  If msgWindow
    txtMsg = TextGadget(#PB_Any, 10, 15, msgW - 20, 50, message)
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

; === TCP Functions ===
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
  Protected timeout.i = 5000  ; 5 секунд максимум
  Protected startTime.i
  Protected threadID.i
  
  If TCPConnection
    CloseNetworkConnection(TCPConnection)
    TCPConnection = 0
  EndIf
  
  Config\Connected = #False
  LogMsg("TCP: Подключение к " + Config\K3ngIP + ":" + Str(Config\K3ngPort) + "...")
  
  ; Запускаем подключение в отдельном потоке
  threadID = CreateThread(@ConnectThread(), 0)
  
  If Not threadID
    LogMsg("TCP: Ошибка создания потока")
    ProcedureReturn #False
  EndIf
  
  ; Ждём результат с таймаутом
  startTime = ElapsedMilliseconds()
  While IsThread(threadID) And (ElapsedMilliseconds() - startTime) < timeout
    Delay(50)
    ; Обрабатываем события окна чтобы не виснуть
    While WindowEvent() : Wend
  Wend
  
  ; Проверяем результат
  If IsThread(threadID)
    ; Таймаут - поток ещё работает
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
  
  ; Send command with CR
  buffer = cmd + #CR$
  If SendNetworkString(TCPConnection, buffer, #PB_Ascii) = 0
    LogMsg("TCP TX Error: " + cmd)
    ProcedureReturn ""
  EndIf
  
  LogMsg("TCP TX: " + cmd)
  
  ; Wait for response
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
  
  If Not TCPConnection Or Config\Mode = 1  ; Skip if Log→Controller only mode
    ProcedureReturn
  EndIf
  
  response = SendK3NGCommand("C2")
  
  ; Parse response: AZ=xxxEL=yyy or AZ=xxx
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
    ; Команда для элевации в GS232B
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
    ; Команда W для азимута и элевации одновременно
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

; === DDE Callback ===
ProcedureDLL.l DDECallback(uType.l, uFmt.l, hconv.l, hsz1.l, hsz2.l, hdata.l, dwData1.l, dwData2.l)
  Protected result.l = 0
  Protected *data, dataSize.l, dataStr.s
  Protected cmdValue.i
  
  Select uType
    Case #XTYP_CONNECT
      ; VQ-Log connecting - accept all
      LogMsg("DDE: Client connected")
      result = #True
      
    Case #XTYP_REQUEST
      ; VQ-Log requesting current azimuth
      If uFmt = #CF_TEXT
        dataStr = Str(CurrentAzimuth) + Chr(0)
        result = DdeCreateDataHandle(DDEInst, @dataStr, Len(dataStr) + 1, 0, hsz2, #CF_TEXT, 0)
        LogMsg("DDE: Request AZ=" + Str(CurrentAzimuth))
      EndIf
      
    Case #XTYP_POKE
      ; VQ-Log sending command (GA:xxx or GE:xxx)
      If hdata And (Config\Mode = 1 Or Config\Mode = 2)  ; Log→Controller or Bidirectional
        *data = DdeAccessData(hdata, @dataSize)
        If *data
          dataStr = PeekS(*data, dataSize, #PB_Ascii)
          DdeUnaccessData(hdata)
          
          LogMsg("DDE: Poke received: " + dataStr)
          
          ; Parse GA:xxx command (Go Azimuth)
          If Left(UCase(dataStr), 3) = "GA:"
            cmdValue = Val(Mid(dataStr, 4))
            If cmdValue >= 0 And cmdValue <= 360
              RotateToAzimuth(cmdValue)
            EndIf
          ; Parse GE:xxx command (Go Elevation)
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

; === DDE Server Init/Cleanup ===
Procedure.i InitDDEServer()
  Protected result.l
  
  result = DdeInitializeW(@DDEInst, @DDECallback(), #APPCLASS_STANDARD, 0)
  If result <> #DMLERR_NO_ERROR
    LogMsg("DDE: Init failed, error " + Str(result))
    ProcedureReturn #False
  EndIf
  
  ; Create string handles
  hszService = DdeCreateStringHandleW(DDEInst, "ARSWIN", #CP_WINANSI)
  hszTopic = DdeCreateStringHandleW(DDEInst, "RCI", #CP_WINANSI)
  hszItemAz = DdeCreateStringHandleW(DDEInst, "AZIMUTH", #CP_WINANSI)
  hszItemEl = DdeCreateStringHandleW(DDEInst, "ELEVATION", #CP_WINANSI)
  
  ; Register service
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

; === Configuration ===
Procedure LoadConfig()
  Protected file.s = GetCurrentDirectory() + "k3ng_bridge.ini"
  
  If OpenPreferences(file)
    Config\K3ngIP = ReadPreferenceString("K3ngIP", "192.168.1.100")
    Config\K3ngPort = ReadPreferenceLong("K3ngPort", 23)
    Config\Mode = ReadPreferenceLong("Mode", 2)
    Config\PollInterval = ReadPreferenceLong("PollInterval", 1000)
    Config\WinX = ReadPreferenceLong("WinX", -1)
    Config\WinY = ReadPreferenceLong("WinY", -1)
    ClosePreferences()
  Else
    Config\K3ngIP = "192.168.1.100"
    Config\K3ngPort = 23
    Config\Mode = 2
    Config\PollInterval = 1000
    Config\WinX = -1
    Config\WinY = -1
  EndIf
EndProcedure

Procedure SaveConfig()
  Protected file.s = GetCurrentDirectory() + "k3ng_bridge.ini"
  
  ; Сохраняем текущую позицию окна
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

; === GUI ===
Procedure CreateMainWindow()
  Protected flags.i = #PB_Window_SystemMenu | #PB_Window_MinimizeGadget
  Protected x.i, y.i
  
  ; Используем сохранённые координаты или центр экрана
  If Config\WinX >= 0 And Config\WinY >= 0
    x = Config\WinX
    y = Config\WinY
  Else
    x = (GetSystemMetrics_(#SM_CXSCREEN) - 420) / 2
    y = (GetSystemMetrics_(#SM_CYSCREEN) - 530) / 2
  EndIf
  
  gMainWindow = OpenWindow(#MainWindow, x, y, 420, 530, "K3NG Bridge for VQ-Log", flags)
  
  ; Connection settings
  FrameGadget(#PB_Any, 10, 10, 400, 100, "Подключение к K3NG")
  TextGadget(#IPLabel, 20, 35, 80, 20, "IP адрес:")
  StringGadget(#IPInput, 100, 32, 120, 24, Config\K3ngIP)
  TextGadget(#PortLabel, 230, 35, 40, 20, "Порт:")
  StringGadget(#PortInput, 275, 32, 60, 24, Str(Config\K3ngPort), #PB_String_Numeric)
  
  TextGadget(#ModeLabel, 20, 65, 80, 20, "Режим:")
  ComboBoxGadget(#ModeCombo, 100, 62, 180, 24)
  AddGadgetItem(#ModeCombo, 0, "Controller → Log")
  AddGadgetItem(#ModeCombo, 1, "Log → Controller")
  AddGadgetItem(#ModeCombo, 2, "Bidirectional")
  SetGadgetState(#ModeCombo, Config\Mode)
  
  ButtonGadget(#ConnectBtn, 300, 55, 100, 30, "Connect")
  
  ; Status
  FrameGadget(#StatusFrame, 10, 115, 400, 80, "Статус")
  TextGadget(#AzLabel, 20, 140, 60, 20, "Azimuth:")
  TextGadget(#AzValue, 85, 140, 50, 20, "---°", #PB_Text_Right)
  TextGadget(#ElLabel, 150, 140, 60, 20, "Elevation:")
  TextGadget(#ElValue, 215, 140, 50, 20, "---°", #PB_Text_Right)
  
  TextGadget(#TCPStatus, 20, 165, 180, 20, "TCP: Отключено")
  TextGadget(#DDEStatus, 210, 165, 180, 20, "DDE: Не запущен")
  
  ; Manual control
  FrameGadget(#ManualFrame, 10, 200, 400, 55, "Ручное управление")
  TextGadget(#ManualAzLabel, 20, 225, 25, 20, "Az:")
  StringGadget(#ManualAzInput, 45, 222, 50, 24, "0", #PB_String_Numeric)
  TextGadget(#PB_Any, 98, 225, 15, 20, "°")
  TextGadget(#ManualElLabel, 120, 225, 20, 20, "El:")
  StringGadget(#ManualElInput, 142, 222, 50, 24, "0", #PB_String_Numeric)
  TextGadget(#PB_Any, 195, 225, 15, 20, "°")
  ButtonGadget(#ManualGoBtn, 220, 220, 60, 28, "GO")
  ButtonGadget(#ManualStopBtn, 285, 220, 60, 28, "STOP")
  
  ; Log
  FrameGadget(#LogFrame, 10, 260, 400, 250, "Лог")
  ListViewGadget(#LogList, 20, 280, 380, 220)
  gLogList = #LogList
  
  ; Timer for polling
  AddWindowTimer(#MainWindow, #PollTimer, Config\PollInterval)
EndProcedure

Procedure UpdateStatus()
  SetGadgetText(#AzValue, Str(CurrentAzimuth) + "°")
  SetGadgetText(#ElValue, Str(CurrentElevation) + "°")
  
  If Config\Connected
    SetGadgetText(#TCPStatus, "TCP: Подключено")
    SetGadgetColor(#TCPStatus, #PB_Gadget_FrontColor, RGB(0, 128, 0))
  Else
    SetGadgetText(#TCPStatus, "TCP: Отключено")
    SetGadgetColor(#TCPStatus, #PB_Gadget_FrontColor, RGB(192, 0, 0))
  EndIf
  
  If DDEInst
    SetGadgetText(#DDEStatus, "DDE: ARSWIN|RCI")
    SetGadgetColor(#DDEStatus, #PB_Gadget_FrontColor, RGB(0, 128, 0))
  Else
    SetGadgetText(#DDEStatus, "DDE: Не запущен")
    SetGadgetColor(#DDEStatus, #PB_Gadget_FrontColor, RGB(192, 0, 0))
  EndIf
EndProcedure

; === Main ===
Procedure Main()
  Protected event.i, gadget.i, quit.i = #False
  
  LoadConfig()
  CreateMainWindow()
  
  ; Start DDE Server
  If InitDDEServer()
    LogMsg("Приложение запущено")
  Else
    LogMsg("Ошибка запуска DDE сервера")
  EndIf
  
  UpdateStatus()
  
  ; Main loop
  Repeat
    event = WaitWindowEvent()
    
    Select event
      Case #PB_Event_CloseWindow
        quit = #True
        
      Case #PB_Event_Timer
        If EventTimer() = #PollTimer
          PollK3NGPosition()
          UpdateStatus()
        EndIf
        
      Case #PB_Event_Gadget
        gadget = EventGadget()
        
        Select gadget
          Case #ConnectBtn
            Config\K3ngIP = GetGadgetText(#IPInput)
            Config\K3ngPort = Val(GetGadgetText(#PortInput))
            Config\Mode = GetGadgetState(#ModeCombo)
            SaveConfig()
            
            If Config\Connected
              DisconnectK3NG()
            Else
              ConnectToK3NG()
            EndIf
            UpdateStatus()
            
          Case #ModeCombo
            Config\Mode = GetGadgetState(#ModeCombo)
            SaveConfig()
            LogMsg("Режим: " + GetGadgetItemText(#ModeCombo, Config\Mode))
            
          Case #ManualGoBtn
            Protected az.i = Val(GetGadgetText(#ManualAzInput))
            Protected el.i = Val(GetGadgetText(#ManualElInput))
            RotateToPosition(az, el)
            
          Case #ManualStopBtn
            StopRotation()
            
        EndSelect
    EndSelect
    
  Until quit
  
  ; Cleanup
  DisconnectK3NG()
  CleanupDDEServer()
  SaveConfig()
EndProcedure

Main()
