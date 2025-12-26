; ============================================================================
; K3NG Bridge for VQ-Log
; Main.pb - Основной файл программы
; ============================================================================
;
; ОПИСАНИЕ:
;   Программа-мост между K3NG контроллером ротатора (TCP/IP, протокол GS232B)
;   и программой VQ-Log (DDE протокол ARSVCOM|RCI)
;
; ВАЖНО: Компилировать с включённой опцией "Create threadsafe executable"
; (Compiler -> Compiler Options -> Create threadsafe executable)
;
; СТРУКТУРА ПРОЕКТА:
;   Main.pb        - Основной файл (этот файл)
;   Constants.pbi  - Константы и перечисления
;   Procedures.pbi - Процедуры и функции (см. комментарии там для деталей)
;   Windows.pbi    - Окна и гаджеты
;
; ТЕКУЩЕЕ СОСТОЯНИЕ:
;   ✅ Азимут работает корректно (передача данных в VQ-Log)
;   ✅ Элевация работает корректно (передача данных в VQ-Log)
;   ✅ Команды управления GA:/GE: работают
;   ✅ TCP/IP связь с K3NG контроллером работает
;   ✅ DDE протокол ARSVCOM|RCI полностью функционален
;
; ============================================================================

EnableExplicit

; === Include Files ===
XIncludeFile "Constants.pbi"
XIncludeFile "Procedures.pbi"
XIncludeFile "Windows.pbi"

; ============================================================================
; MAIN
; ============================================================================
Procedure Main()
  Protected event.i, gadget.i, quit.i = #False
  Protected i.i, cmdLine.s

  ; Логируем параметры командной строки
  cmdLine = "Command line params: "
  For i = 0 To CountProgramParameters() - 1
    cmdLine + "[" + Str(i) + "]=" + ProgramParameter(i) + " "
  Next

  ; Проверка единственного экземпляра - ОТКЛЮЧЕНА для работы с VQ-Log
  ; VQ-Log может запускать приложение несколько раз для разных DDE сервисов
  ;If Not CheckSingleInstance()
  ;  ; Приложение уже запущено - тихо выходим
  ;  ProcedureReturn
  ;EndIf

  ; Загрузка конфигурации
  LoadConfig()

  ; ВАЖНО: Создаем окно первым делом для message loop
  CreateMainWindow()

  ; Выводим параметры командной строки в лог
  LogMsg(cmdLine)

  ; Запуск DDE сервера СРАЗУ после создания окна
  ; Это критично - VQ-Log ждет что DDE сервер будет доступен немедленно
  If InitDDEServer()
    LogMsg("Application started")
  Else
    LogMsg("DDE server start error")
  EndIf

  ; Обрабатываем сообщения Windows чтобы завершить регистрацию DDE
  ; Увеличиваем количество итераций и задержку для надежной инициализации
  For i = 1 To 50
    WindowEvent()
    Delay(20)
  Next

  LogMsg("DDE: Инициализация завершена, сервер готов к подключениям")

  ; Запускаем таймер опроса ПОСЛЕ инициализации DDE сервера
  AddWindowTimer(#MainWindow, #TimerPoll, Config\PollInterval)

  ; Минимизация окна при запуске, если опция включена
  If Config\StartMinimized
    SetWindowState(#MainWindow, #PB_Window_Minimize)
  EndIf

  ; Автоматическое подключение к контроллеру при запуске
  If ConnectToK3NG()
    ; Делаем первый опрос контроллера сразу после подключения
    ; чтобы LastAzimuth и LastElevation были инициализированы
    PollK3NGPosition()
  EndIf

  UpdateStatus()
  
  ; === Главный цикл событий ===
  Repeat
    event = WaitWindowEvent()
    
    Select event
      Case #PB_Event_CloseWindow
        AppShuttingDown = #True  ; Устанавливаем флаг для быстрого выхода
        RemoveWindowTimer(#MainWindow, #TimerPoll)  ; Останавливаем таймер немедленно
        quit = #True

      Case #PB_Event_Timer
        If EventTimer() = #TimerPoll
          HandleTimer()
        EndIf
        
      Case #PB_Event_Gadget
        gadget = EventGadget()

        Select gadget
          Case #ComboMode
            HandleModeChange()

          Case #ButtonGo
            HandleGoButton()

          Case #ButtonStop
            HandleStopButton()

          Case #ButtonApplyInterval
            HandleApplyInterval()

          Case #CheckStartMinimized
            HandleStartMinimizedToggle()

        EndSelect
    EndSelect
    
  Until quit
  
  ; === Очистка ===
  DisconnectK3NG()
  CleanupDDEServer()
  SaveConfig()
  ;ReleaseSingleInstance()  ; Отключено вместе с single instance check

EndProcedure

; === Точка входа ===
Main()