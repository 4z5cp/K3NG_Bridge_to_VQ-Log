; ============================================================================
; K3NG Bridge for VQ-Log
; Procedures.pbi - Процедуры и функции
; ============================================================================
;
; ИСТОРИЯ РАЗРАБОТКИ И ВАЖНЫЕ МОМЕНТЫ:
;
; 1. DDE ПРОТОКОЛ:
;    - VQ-Log использует DDE для связи с программами управления ротаторами
;    - Регистрируются два сервиса: ARSWIN (старый) и ARSVCOM (современный)
;    - Топик: RCI (Rotator Control Interface)
;    - Элементы: AZIMUTH и ELEVATION
;
; 2. ПРОБЛЕМА С ЭЛЕВАЦИЕЙ:
;    - АЗИМУТ работает через стандартный ADVISE loop (VQ-Log подписывается)
;    - ЭЛЕВАЦИЯ НЕ РАБОТАЕТ: VQ-Log НЕ устанавливает ADVISE для ELEVATION
;    - VQ-Log отправляет POKE "RE:" для запроса элевации, но не принимает
;      данные обратно ни одним из стандартных методов DDE
;    - Попытки решения (все неудачные):
;      * Возврат data handle из XTYP_POKE - игнорируется
;      * DdePostAdvise для ELEVATION - ошибка DMLERR_NOTPROCESSED (нет ADVISE loop)
;      * DdeClientTransaction с XTYP_ADVDATA - ошибка 16390 (DMLERR_NOTPROCESSED)
;      * Отправка обоих значений "RA:355 RE:05" в AZIMUTH - не парсится
;    - ВЫВОД: VQ-Log использует нестандартный протокол для элевации,
;             нужна документация от разработчика
;
; 3. ФОРМАТ ДАННЫХ:
;    - Азимут: "RA:355" (всегда 3 цифры с ведущими нулями)
;    - Элевация: "RE:05" (всегда 2 цифры с ведущими нулями)
;    - Команды управления: "GA:180" (азимут), "GE:45" (элевация)
;
; 4. TCP/IP К K3NG КОНТРОЛЛЕРУ:
;    - Используется протокол GS232B
;    - Команды: C2 (запрос позиции), M<azimuth> (поворот азимута),
;               W<elevation> (поворот элевации), S (стоп)
;
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
Global hszService.l = 0          ; DDE string handle для ARSVCOM
Global hszService2.l = 0         ; DDE string handle для ARSWIN
Global hszTopic.l = 0            ; DDE string handle для RCI
Global hszItemAz.l = 0           ; DDE string handle для AZIMUTH
Global hszItemEl.l = 0           ; DDE string handle для ELEVATION
Global DDEConversation.l = 0     ; Хэндл текущего DDE подключения
Global LastRequestType.s = ""   ; Последний тип запроса: "RA" или "RE"
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
  DdePostAdvise(idInst, hszTopic, hszItem)
  DdeQueryStringW(idInst, hsz, psz, cchMax, iCodePage)
  DdeConnect(idInst, hszService, hszTopic, pCC)
  DdeDisconnect(hConv)
  DdeClientTransaction(pData, cbData, hConv, hszItem, wFmt, wType, dwTimeout, pdwResult)
  DdeGetLastError(idInst)
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

  ; LogMsg("TCP TX: " + cmd)  ; Отключено для уменьшения шума в логах

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
  ; If received <> ""
  ;   LogMsg("TCP RX: " + received)  ; Отключено для уменьшения шума в логах
  ; EndIf

  ProcedureReturn received
EndProcedure

Procedure PollK3NGPosition()
  Protected response.s, azPos.i, elPos.i
  Protected dataReceived.i = #False

  If Not TCPConnection Or Config\Mode = #MODE_LOG_TO_CONTROLLER
    ; Контроллер не подключен - сбрасываем значения в 0
    If CurrentAzimuth <> 0 Or CurrentElevation <> 0
      CurrentAzimuth = 0
      CurrentElevation = 0
      If DDEInst And hszTopic
        If hszItemAz
          DdePostAdvise(DDEInst, hszTopic, hszItemAz)
          LogMsg("DDE: Posted AZ update = 0 (controller disconnected)")
        EndIf
        If hszItemEl
          DdePostAdvise(DDEInst, hszTopic, hszItemEl)
          LogMsg("DDE: Posted EL update = 0 (controller disconnected)")
        EndIf
      EndIf
    EndIf
    ProcedureReturn
  EndIf

  response = SendK3NGCommand("C2")

  ; Проверяем получили ли валидный ответ от контроллера
  If response <> "" And FindString(response, "AZ=")
    dataReceived = #True

    azPos = Val(Mid(response, FindString(response, "AZ=") + 3, 3))
    If azPos >= 0 And azPos <= 360
      If CurrentAzimuth <> azPos
        CurrentAzimuth = azPos
        ; Уведомляем VQ-Log об изменении азимута через ADVISE
        If DDEInst And hszTopic And hszItemAz
          DdePostAdvise(DDEInst, hszTopic, hszItemAz)
          LogMsg("DDE: Posted AZ update = " + Str(CurrentAzimuth))
        EndIf
      EndIf
    EndIf

    If FindString(response, "EL=")
      elPos = Val(Mid(response, FindString(response, "EL=") + 3, 3))
      If elPos >= 0 And elPos <= 180
        If CurrentElevation <> elPos
          CurrentElevation = elPos
          ; Уведомляем VQ-Log об изменении элевации через ADVISE
          If DDEInst And hszTopic And hszItemEl
            DdePostAdvise(DDEInst, hszTopic, hszItemEl)
            LogMsg("DDE: Posted EL update = " + Str(CurrentElevation))
          EndIf
        EndIf
      EndIf
    EndIf
  Else
    ; Контроллер не ответил - сбрасываем в 0
    If CurrentAzimuth <> 0 Or CurrentElevation <> 0
      CurrentAzimuth = 0
      CurrentElevation = 0
      If DDEInst And hszTopic
        If hszItemAz
          DdePostAdvise(DDEInst, hszTopic, hszItemAz)
          LogMsg("DDE: Posted AZ update = 0 (no response)")
        EndIf
        If hszItemEl
          DdePostAdvise(DDEInst, hszTopic, hszItemEl)
          LogMsg("DDE: Posted EL update = 0 (no response)")
        EndIf
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
  Protected debugMsg.s

  ; Детальное логирование всех DDE транзакций для отладки
  debugMsg = "DDE Callback: type=$" + RSet(Hex(uType), 4, "0") + " fmt=$" + Hex(uFmt) + " hconv=$" + Hex(hconv)
  LogMsg(debugMsg)

  Select uType
    Case #XTYP_CONNECT
      ; hsz1 = Topic, hsz2 = Service
      ; Проверяем, что клиент подключается к ARSVCOM|RCI или ARSWIN|RCI
      If hsz1 = hszTopic And (hsz2 = hszService Or hsz2 = hszService2)
        If hsz2 = hszService
          LogMsg("DDE: Подключено к ARSVCOM")
        Else
          LogMsg("DDE: Подключено к ARSWIN")
        EndIf
        result = #True
      Else
        LogMsg("DDE: Соединение отклонено")
        result = #False
      EndIf

    Case #XTYP_REQUEST, #XTYP_ADVREQ
      ; Декодируем item для логирования
      Protected itemStr.s{128}
      DdeQueryStringW(DDEInst, hsz2, @itemStr, 128, #CP_WINUNICODE)

      Protected transType.s
      If uType = #XTYP_REQUEST
        transType = "REQUEST"
      Else
        transType = "ADVREQ"
      EndIf

      If uFmt = #CF_TEXT
        Protected *buffer, bufLen.i
        If hsz2 = hszItemAz
          ; AZIMUTH элемент используется для передачи и азимута и элевации
          ; Определяем что отправить по последнему POKE запросу (RA или RE)
          If LastRequestType = "RE"
            ; Был запрос элевации через RE: POKE
            dataStr = "RE:" + RSet(Str(CurrentElevation), 2, "0")
          Else
            ; Был запрос азимута через RA: POKE или первый запрос
            dataStr = "RA:" + RSet(Str(CurrentAzimuth), 3, "0")
          EndIf
          bufLen = Len(dataStr) + 1
          *buffer = AllocateMemory(bufLen)
          If *buffer
            PokeS(*buffer, dataStr, -1, #PB_Ascii)
            result = DdeCreateDataHandle(DDEInst, *buffer, bufLen, 0, hsz2, #CF_TEXT, 0)
            FreeMemory(*buffer)
            LogMsg("DDE: " + transType + " for AZIMUTH -> " + dataStr)
          EndIf
        ElseIf hsz2 = hszItemEl
          ; ЭЛЕВАЦИЯ: VQ-Log никогда не запрашивает этот элемент через REQUEST/ADVREQ
          ; Элемент зарегистрирован, но фактически не используется
          ; VQ-Log использует POKE "RE:" вместо REQUEST (см. ниже)
          dataStr = "RE:" + RSet(Str(CurrentElevation), 2, "0")
          bufLen = Len(dataStr) + 1
          *buffer = AllocateMemory(bufLen)
          If *buffer
            PokeS(*buffer, dataStr, -1, #PB_Ascii)
            result = DdeCreateDataHandle(DDEInst, *buffer, bufLen, 0, hsz2, #CF_TEXT, 0)
            FreeMemory(*buffer)
            LogMsg("DDE: " + transType + " for ELEVATION -> " + dataStr)
          EndIf
        Else
          LogMsg("DDE: " + transType + " for UNKNOWN item: " + itemStr)
        EndIf
      EndIf

    Case #XTYP_POKE
      ; Сохраняем conversation handle если еще не сохранен
      If DDEConversation = 0
        DDEConversation = hconv
      EndIf

      *data = DdeAccessData(hdata, @dataSize)
      If *data
        dataStr = PeekS(*data, dataSize, #PB_Ascii)
        DdeUnaccessData(hdata)

        LogMsg("DDE: Received: " + dataStr + " (hconv=" + Str(hconv) + ")")

        If Left(UCase(dataStr), 3) = "GA:"
          ; Команда управления - поворот азимута
          If Config\Mode = #MODE_LOG_TO_CONTROLLER Or Config\Mode = #MODE_BIDIRECTIONAL
            cmdValue = Val(Mid(dataStr, 4))
            If cmdValue >= 0 And cmdValue <= 360
              RotateToAzimuth(cmdValue)
            EndIf
          EndIf
          result = #DDE_FACK
        ElseIf Left(UCase(dataStr), 3) = "GE:"
          ; Команда управления - поворот элевации
          If Config\Mode = #MODE_LOG_TO_CONTROLLER Or Config\Mode = #MODE_BIDIRECTIONAL
            cmdValue = Val(Mid(dataStr, 4))
            If cmdValue >= 0 And cmdValue <= 180
              RotateToElevation(cmdValue)
            EndIf
          EndIf
          result = #DDE_FACK
        ElseIf Left(UCase(dataStr), 3) = "RA:"
          ; VQ-Log запрашивает азимут - сохраняем тип и отправляем обновление
          LogMsg("DDE: RA POKE request - current AZ=" + Str(CurrentAzimuth))
          LastRequestType = "RA"
          If hszItemAz
            DdePostAdvise(DDEInst, hszTopic, hszItemAz)
          EndIf
          result = #DDE_FACK
        ElseIf Left(UCase(dataStr), 3) = "RE:"
          ; VQ-Log запрашивает элевацию - сохраняем тип и отправляем обновление
          ; Ответ будет отправлен через элемент AZIMUTH в следующем ADVREQ
          LogMsg("DDE: RE POKE request - current EL=" + Str(CurrentElevation))
          LastRequestType = "RE"
          If hszItemAz
            DdePostAdvise(DDEInst, hszTopic, hszItemAz)
          EndIf
          result = #DDE_FACK
        Else
          result = #DDE_FACK
        EndIf
      Else
        result = #DDE_FACK
      EndIf

    Case #XTYP_ADVSTART, $80A2
      ; Декодируем topic и item
      Protected topicName.s{128}, itemName.s{128}
      DdeQueryStringW(DDEInst, hsz1, @topicName, 128, #CP_WINUNICODE)
      DdeQueryStringW(DDEInst, hsz2, @itemName, 128, #CP_WINUNICODE)
      ; Сохраняем conversation handle для отправки POKE
      DDEConversation = hconv
      LogMsg("DDE: Advise " + topicName + "|" + itemName + " (hconv=" + Str(hconv) + ")")
      result = #True

    Case #XTYP_ADVSTOP, $80D2
      LogMsg("DDE: Advise loop stopped")
      result = #True

    Case #XTYP_DISCONNECT
      LogMsg("DDE: Client disconnected (hconv=" + Str(hconv) + ")")
      If DDEConversation = hconv
        DDEConversation = 0
      EndIf
      result = 0

    Case #XTYP_WILDCONNECT
      ; Клиент делает wildcard connect - вернуть список доступных topic
      LogMsg("DDE: Wildcard connect request")
      ; Создаем список пар Service|Topic
      ; Формат: массив из HSZPAIR структур, завершающийся NULL парой
      Protected *hszPair, pairSize.i = 8 ; 2 DWORD по 4 байта
      *hszPair = AllocateMemory(pairSize * 2) ; одна пара + null пара
      If *hszPair
        PokeL(*hszPair, hszService)
        PokeL(*hszPair + 4, hszTopic)
        PokeL(*hszPair + 8, 0)
        PokeL(*hszPair + 12, 0)
        result = DdeCreateDataHandle(DDEInst, *hszPair, pairSize * 2, 0, 0, #CF_TEXT, 0)
        FreeMemory(*hszPair)
      EndIf

    Default
      ; Игнорируем неизвестные типы сообщений
      result = 0

  EndSelect

  ProcedureReturn result
EndProcedure

Procedure.i InitDDEServer()
  Protected result.l
  Protected ddeFlags.l

  ; Используем флаги для сервера: разрешаем все транзакции
  ddeFlags = 0  ; Без фильтрации callback

  result = DdeInitializeW(@DDEInst, @DDECallback(), ddeFlags, 0)
  If result <> #DMLERR_NO_ERROR
    LogMsg("DDE: Initialization error, code " + Str(result))
    ProcedureReturn #False
  EndIf

  hszService = DdeCreateStringHandleW(DDEInst, "ARSVCOM", #CP_WINUNICODE)
  hszService2 = DdeCreateStringHandleW(DDEInst, "ARSWIN", #CP_WINUNICODE)
  hszTopic = DdeCreateStringHandleW(DDEInst, "RCI", #CP_WINUNICODE)
  hszItemAz = DdeCreateStringHandleW(DDEInst, "AZIMUTH", #CP_WINUNICODE)
  hszItemEl = DdeCreateStringHandleW(DDEInst, "ELEVATION", #CP_WINUNICODE)

  ; Регистрируем оба сервиса для совместимости
  If DdeNameService(DDEInst, hszService, 0, #DNS_REGISTER)
    LogMsg("DDE: Сервер ARSVCOM|RCI зарегистрирован")
  Else
    LogMsg("DDE: Ошибка регистрации сервиса ARSVCOM")
    ProcedureReturn #False
  EndIf

  If DdeNameService(DDEInst, hszService2, 0, #DNS_REGISTER)
    LogMsg("DDE: Сервер ARSWIN|RCI зарегистрирован")
  Else
    LogMsg("DDE: Ошибка регистрации сервиса ARSWIN")
    ProcedureReturn #False
  EndIf

  ProcedureReturn #True
EndProcedure

; Функции для подключения к VQ-Log как клиент удалены - не используются
; Мы работаем только как DDE сервер

Procedure CleanupDDEServer()
  If DDEInst
    DdeNameService(DDEInst, hszService, 0, #DNS_UNREGISTER)
    DdeNameService(DDEInst, hszService2, 0, #DNS_UNREGISTER)
    DdeFreeStringHandle(DDEInst, hszService)
    DdeFreeStringHandle(DDEInst, hszService2)
    DdeFreeStringHandle(DDEInst, hszTopic)
    DdeFreeStringHandle(DDEInst, hszItemAz)
    DdeFreeStringHandle(DDEInst, hszItemEl)
    DdeUninitialize(DDEInst)
    DDEInst = 0
    LogMsg("DDE: Сервер остановлен")
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
  Protected lockFile.s = GetTemporaryDirectory() + "K3NG_Bridge.lock"

  ; Check if lock file exists
  If FileSize(lockFile) >= 0
    ; Lock file exists - another instance might be running
    ; Try to open it exclusively to verify
    Mutex = OpenFile(#PB_Any, lockFile, #PB_File_SharedRead)
    If Mutex = 0
      ; Can't open - another instance is running
      ProcedureReturn #False
    Else
      ; Could open - previous instance didn't clean up, overwrite
      CloseFile(Mutex)
      DeleteFile(lockFile)
    EndIf
  EndIf

  ; Create lock file
  Mutex = CreateFile(#PB_Any, lockFile)
  If Mutex = 0
    ProcedureReturn #False
  EndIf

  WriteStringN(Mutex, Str(GetCurrentProcessId_()))
  FlushFileBuffers(Mutex)

  ProcedureReturn #True
EndProcedure

Procedure ReleaseSingleInstance()
  Protected lockFile.s = GetTemporaryDirectory() + "K3NG_Bridge.lock"

  If Mutex
    CloseFile(Mutex)
    DeleteFile(lockFile)
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
    SetGadgetText(#LabelDDEStatus, "DDE: ARSWIN/ARSVCOM|RCI")
    SetGadgetColor(#LabelDDEStatus, #PB_Gadget_FrontColor, RGB(0, 128, 0))
  Else
    SetGadgetText(#LabelDDEStatus, "DDE: Не запущен")
    SetGadgetColor(#LabelDDEStatus, #PB_Gadget_FrontColor, RGB(192, 0, 0))
  EndIf
EndProcedure