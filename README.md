# K3NG Bridge to VQ-Log

Bridge application between K3NG Rotator Controller (via TCP/IP) and VQ-Log logging software (via DDE).

This application emulates ARSWIN DDE server, allowing VQ-Log to control K3NG rotator controller over network without additional software.

## Features

- **DDE Server** — Emulates ARSWIN (Service: `ARSWIN`, Topic: `RCI`)
- **TCP Client** — Connects to K3NG controller via network (Telnet/TCP)
- **GS-232 Protocol** — Full support for Yaesu GS-232A/B commands
- **Bidirectional Control** — Three operating modes:
  - Controller → Log (read position from controller)
  - Log → Controller (send commands from VQ-Log)
  - Bidirectional (both directions)
- **Manual Control** — Azimuth and Elevation control from GUI
- **Configuration** — Settings saved to INI file

## Requirements

- Windows 7/10/11
- PureBasic 6.x (for compilation)
- K3NG Rotator Controller with network interface
- VQ-Log logging software

## Project Structure

```
K3NG_Bridge_to_VQ-Log/
├── Main.pb           # Main program file
├── Constants.pbi     # Constants and enumerations
├── Procedures.pbi    # Functions and procedures
├── Windows.pbi       # GUI windows and gadgets
└── README.md         # This file
```

## Compilation

1. Open `Main.pb` in PureBasic IDE
2. Go to **Compiler → Compiler Options**
3. Enable **"Create threadsafe executable"**
4. Compile (F5 or Ctrl+F5)

## Configuration

### VQ-Log Setup

1. Open VQ-Log
2. Go to **Configuration → Rotor Control → Sets**
3. Set path to compiled `K3NG_Bridge.exe`
4. Enable rotor control

### K3NG Bridge Setup

1. Enter K3NG controller IP address
2. Enter TCP port (default: 23 for Telnet)
3. Select operating mode
4. Click **Connect**

## GS-232 Commands Used

| Command | Description |
|---------|-------------|
| `C2` | Request current position (AZ=xxx EL=yyy) |
| `Mxxx` | Rotate to azimuth xxx |
| `Wxxx yyy` | Rotate to azimuth xxx and elevation yyy |
| `S` | Stop rotation |

## DDE Interface

| Parameter | Value |
|-----------|-------|
| Service | `ARSWIN` |
| Topic | `RCI` |
| Item | `AZIMUTH`, `ELEVATION` |

### DDE Commands (from VQ-Log)

- `GA:xxx` — Go to Azimuth xxx degrees
- `GE:xxx` — Go to Elevation xxx degrees

## Screenshots

*Coming soon*

## License

MIT License

## Author

4Z5CP

## Links

- [K3NG Rotator Controller](https://github.com/k3ng/k3ng_rotator_controller)
- [VQ-Log](https://www.dxmaps.com/vqlog.html)
- [ARSWIN / EA4TX](https://ea4tx.com/)
