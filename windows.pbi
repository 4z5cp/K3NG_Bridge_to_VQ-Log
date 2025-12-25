; ============================================================================
; K3NG Bridge for VQ-Log
; Windows.pbi - Windows and gadgets
; ============================================================================

#MAIN_WINDOW_WIDTH = 420
#MAIN_WINDOW_HEIGHT = 590

Procedure CreateMainWindow()
  Protected flags.i = #PB_Window_SystemMenu | #PB_Window_MinimizeGadget
  Protected x.i, y.i
  
  ; Use saved coordinates or center of screen
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
  FrameGadget(#FrameConnection, 10, 10, 400, 100, "K3NG Connection")
  
  TextGadget(#LabelIP, 20, 35, 80, 20, "IP Address:")
  StringGadget(#StringIP, 100, 32, 120, 24, Config\K3ngIP)
  
  TextGadget(#LabelPort, 230, 35, 40, 20, "Port:")
  StringGadget(#StringPort, 275, 32, 60, 24, Str(Config\K3ngPort), #PB_String_Numeric)
  
  TextGadget(#LabelMode, 20, 68, 80, 20, "Mode:")
  ComboBoxGadget(#ComboMode, 100, 65, 180, 24)
  AddGadgetItem(#ComboMode, #MODE_CONTROLLER_TO_LOG, "Controller -> Log")
  AddGadgetItem(#ComboMode, #MODE_LOG_TO_CONTROLLER, "Log -> Controller")
  AddGadgetItem(#ComboMode, #MODE_BIDIRECTIONAL, "Bidirectional")
  SetGadgetState(#ComboMode, Config\Mode)
  
  ButtonGadget(#ButtonConnect, 300, 62, 100, 30, "Connect")
  
  ; === Status Frame ===
  FrameGadget(#FrameStatus, 10, 115, 400, 80, "Status")
  
  TextGadget(#LabelAzText, 20, 140, 60, 20, "Azimuth:")
  TextGadget(#LabelAzValue, 85, 140, 50, 20, "---", #PB_Text_Right)
  
  TextGadget(#LabelElText, 150, 140, 60, 20, "Elevation:")
  TextGadget(#LabelElValue, 215, 140, 50, 20, "---", #PB_Text_Right)
  
  TextGadget(#LabelTCPStatus, 20, 165, 180, 20, "TCP: Disconnected")
  SetGadgetColor(#LabelTCPStatus, #PB_Gadget_FrontColor, RGB(192, 0, 0))
  
  TextGadget(#LabelDDEStatus, 210, 165, 180, 20, "DDE: Not running")
  SetGadgetColor(#LabelDDEStatus, #PB_Gadget_FrontColor, RGB(192, 0, 0))
  
  ; === Manual Control Frame ===
  FrameGadget(#FrameManual, 10, 200, 400, 55, "Manual Control")

  TextGadget(#LabelManualAz, 20, 225, 25, 20, "Az:")
  StringGadget(#StringManualAz, 45, 222, 50, 24, "0", #PB_String_Numeric)
  TextGadget(#LabelManualAzDeg, 98, 225, 15, 20, "deg")

  TextGadget(#LabelManualEl, 120, 225, 20, 20, "El:")
  StringGadget(#StringManualEl, 142, 222, 50, 24, "0", #PB_String_Numeric)
  TextGadget(#LabelManualElDeg, 195, 225, 15, 20, "deg")

  ButtonGadget(#ButtonGo, 220, 220, 60, 28, "GO")
  ButtonGadget(#ButtonStop, 285, 220, 60, 28, "STOP")

  ; === Settings Frame ===
  FrameGadget(#FrameSettings, 10, 260, 400, 55, "Settings")

  TextGadget(#LabelPollInterval, 20, 285, 90, 20, "Poll Interval:")
  StringGadget(#StringPollInterval, 110, 282, 70, 24, Str(Config\PollInterval), #PB_String_Numeric)
  TextGadget(#LabelPollMs, 185, 285, 20, 20, "ms")
  ButtonGadget(#ButtonApplyInterval, 220, 280, 80, 28, "Apply")
  CheckBoxGadget(#CheckStartMinimized, 310, 283, 90, 20, "Start Minimized")
  SetGadgetState(#CheckStartMinimized, Config\StartMinimized)

  ; === Log Frame ===
  FrameGadget(#FrameLog, 10, 320, 400, 260, "Log")
  ListViewGadget(#ListLog, 20, 340, 380, 230)
  
  ; === Timer ===
  AddWindowTimer(#MainWindow, #TimerPoll, Config\PollInterval)
  
EndProcedure