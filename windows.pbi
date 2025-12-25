; ============================================================================
; K3NG Bridge for VQ-Log
; Windows.pbi - Окна и гаджеты
; ============================================================================

#MAIN_WINDOW_WIDTH = 420
#MAIN_WINDOW_HEIGHT = 530

Procedure CreateMainWindow()
  Protected flags.i = #PB_Window_SystemMenu | #PB_Window_MinimizeGadget
  Protected x.i, y.i
  
  ; Используем сохранённые координаты или центр экрана
  If Config\WinX >= 0 And Config\WinY >= 0
    x = Config\WinX
    y = Config\WinY
  Else
    x = (GetSystemMetrics_(#SM_CXSCREEN) - #MAIN_WINDOW_WIDTH) / 2
    y = (GetSystemMetrics_(#SM_CYSCREEN) - #MAIN_WINDOW_HEIGHT) / 2
  EndIf
  
  ; === Main Window ===
  OpenWindow(#MainWindow, x, y, #MAIN_WINDOW_WIDTH, #MAIN_WINDOW_HEIGHT, #APP_NAME, flags)
  
  ; === Connection Frame ===
  FrameGadget(#FrameConnection, 10, 10, 400, 100, "Подключение к K3NG")
  
  TextGadget(#LabelIP, 20, 35, 80, 20, "IP адрес:")
  StringGadget(#StringIP, 100, 32, 120, 24, Config\K3ngIP)
  
  TextGadget(#LabelPort, 230, 35, 40, 20, "Порт:")
  StringGadget(#StringPort, 275, 32, 60, 24, Str(Config\K3ngPort), #PB_String_Numeric)
  
  TextGadget(#LabelMode, 20, 68, 80, 20, "Режим:")
  ComboBoxGadget(#ComboMode, 100, 65, 180, 24)
  AddGadgetItem(#ComboMode, #MODE_CONTROLLER_TO_LOG, "Controller → Log")
  AddGadgetItem(#ComboMode, #MODE_LOG_TO_CONTROLLER, "Log → Controller")
  AddGadgetItem(#ComboMode, #MODE_BIDIRECTIONAL, "Bidirectional")
  SetGadgetState(#ComboMode, Config\Mode)
  
  ButtonGadget(#ButtonConnect, 300, 62, 100, 30, "Connect")
  
  ; === Status Frame ===
  FrameGadget(#FrameStatus, 10, 115, 400, 80, "Статус")
  
  TextGadget(#LabelAzText, 20, 140, 60, 20, "Azimuth:")
  TextGadget(#LabelAzValue, 85, 140, 50, 20, "---°", #PB_Text_Right)
  
  TextGadget(#LabelElText, 150, 140, 60, 20, "Elevation:")
  TextGadget(#LabelElValue, 215, 140, 50, 20, "---°", #PB_Text_Right)
  
  TextGadget(#LabelTCPStatus, 20, 165, 180, 20, "TCP: Отключено")
  SetGadgetColor(#LabelTCPStatus, #PB_Gadget_FrontColor, RGB(192, 0, 0))
  
  TextGadget(#LabelDDEStatus, 210, 165, 180, 20, "DDE: Не запущен")
  SetGadgetColor(#LabelDDEStatus, #PB_Gadget_FrontColor, RGB(192, 0, 0))
  
  ; === Manual Control Frame ===
  FrameGadget(#FrameManual, 10, 200, 400, 55, "Ручное управление")
  
  TextGadget(#LabelManualAz, 20, 225, 25, 20, "Az:")
  StringGadget(#StringManualAz, 45, 222, 50, 24, "0", #PB_String_Numeric)
  TextGadget(#LabelManualAzDeg, 98, 225, 15, 20, "°")
  
  TextGadget(#LabelManualEl, 120, 225, 20, 20, "El:")
  StringGadget(#StringManualEl, 142, 222, 50, 24, "0", #PB_String_Numeric)
  TextGadget(#LabelManualElDeg, 195, 225, 15, 20, "°")
  
  ButtonGadget(#ButtonGo, 220, 220, 60, 28, "GO")
  ButtonGadget(#ButtonStop, 285, 220, 60, 28, "STOP")
  
  ; === Log Frame ===
  FrameGadget(#FrameLog, 10, 260, 400, 260, "Лог")
  ListViewGadget(#ListLog, 20, 280, 380, 230)
  
  ; === Timer ===
  AddWindowTimer(#MainWindow, #TimerPoll, Config\PollInterval)
  
EndProcedure
