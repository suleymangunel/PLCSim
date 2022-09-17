Attribute VB_Name = "Module1"
Option Explicit
Declare Function MapPhysToLin Lib "WinIo.dll" (ByVal PhysAddr As Long, ByVal PhysSize As Long, ByRef PhysMemHandle) As Long
Declare Function UnmapPhysicalMemory Lib "WinIo.dll" (ByVal PhysMemHandle, ByVal LinAddr) As Boolean
Declare Function GetPhysLong Lib "WinIo.dll" (ByVal PhysAddr As Long, ByRef PhysVal As Long) As Boolean
Declare Function SetPhysLong Lib "WinIo.dll" (ByVal PhysAddr As Long, ByVal PhysVal As Long) As Boolean
Declare Function GetPortVal Lib "WinIo.dll" (ByVal PortAddr As Integer, ByRef PortVal As Long, ByVal bSize As Byte) As Boolean
Declare Function SetPortVal Lib "WinIo.dll" (ByVal PortAddr As Integer, ByVal PortVal As Long, ByVal bSize As Byte) As Boolean
Declare Function InitializeWinIo Lib "WinIo.dll" () As Boolean
Declare Function ShutdownWinIo Lib "WinIo.dll" () As Boolean

Type N
 Active As Boolean
 Inputs(3) As Byte
 Outputs4In(3) As Byte
 OutActive(3) As Boolean
 OutOnDelay(3) As Integer
 OutOffDelay(3) As Integer
 OutRet(3) As Boolean
 OutNC(3) As Boolean
End Type

Type N2
 StartCount As Boolean
 OutControl(3) As Boolean
 OutOnDelay(3) As Integer
 OutOnDelayOk(3) As Boolean
 OutOffDelay(3) As Integer
 OutOffDelayOk(3) As Boolean
 OutChanged_1(3) As Boolean
 OutChanged_2(3) As Boolean
 StatusOkChanged_1(3) As Boolean
 StatusOkChanged_2(3) As Boolean
End Type

Type FN
 Active As Boolean
 Inputs(3) As Byte
 Outputs4In(3) As Byte
 OutActive(3) As Boolean
 OutOnDelay(3) As Integer
 OutOffDelay(3) As Integer
 OutRet(3) As Boolean
 OutNC(3) As Boolean
End Type

Global Network(255) As N
Global Netwrk2(255) As N2
Global FileNet As FN
Global PortNo, PortNo0, PortNo1, PortNoUser As Integer
Global BlkNo As Byte
Global pIn(3) As Byte
Global pOut(3) As Byte
Global PushButton As Integer

Sub SabitleriYukle()
 PushButton = 2
End Sub
