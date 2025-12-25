; ============================================================================
; K3NG Bridge for VQ-Log
; Main.pb - Основной файл программы
; ============================================================================
;
; ВАЖНО: Компилировать с включённой опцией "Create threadsafe executable"
; (Compiler -> Compiler Options -> Create threadsafe executable)
;
; Структура проекта:
;   Main.pb        - Основной файл (этот файл)
;   Constants.pbi  - Константы и перечисления
;   Procedures.pbi - Процедуры и функции
;   Windows.pbi    - Окна и гаджеты
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

  ; Проверка единственного экземпляра
  If Not CheckSingleInstance()
    ; Приложение уже запущено - тихо выходим
    ProcedureReturn
  EndIf

  ; Загрузка конфигурации
  LoadConfig()

  ; Создание главного окна
  CreateMainWindow()

  ; Минимизация окна при запуске, если опция включена
  If Config\StartMinimized
    SetWindowState(#MainWindow, #PB_Window_Minimize)
  EndIf

  ; Запуск DDE сервера
  If InitDDEServer()
    LogMsg("Application started")
  Else
    LogMsg("DDE server start error")
  EndIf

  ; Автоматическое подключение к контроллеру при запуске
  If ConnectToK3NG()
    SetGadgetText(#ButtonConnect, "Disconnect")
  EndIf

  UpdateStatus()
  
  ; === Главный цикл событий ===
  Repeat
    event = WaitWindowEvent()
    
    Select event
      Case #PB_Event_CloseWindow
        quit = #True
        
      Case #PB_Event_Timer
        If EventTimer() = #TimerPoll
          HandleTimer()
        EndIf
        
      Case #PB_Event_Gadget
        gadget = EventGadget()
        
        Select gadget
          Case #ButtonConnect
            HandleConnectButton()

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
  ReleaseSingleInstance()

EndProcedure

; === Точка входа ===
Main()