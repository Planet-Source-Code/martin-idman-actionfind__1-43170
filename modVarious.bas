Attribute VB_Name = "modVarious"
' Force variable declaration
Option Explicit

' Api declare
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

' Sub to sleep x seconds
Public Sub Sleep(lngSleep As Long)
   Dim lngSleepEnd As Long
   lngSleepEnd = GetTickCount + lngSleep * 1000
   While GetTickCount <= lngSleepEnd
      DoEvents
   Wend
End Sub

' Sub to freeze x seconds
Public Sub Freeze(lngFreeze As Long)
   Dim lngFreezeEnd As Long
   lngFreezeEnd = GetTickCount + lngFreeze * 1000
   While GetTickCount <= lngFreezeEnd
   Wend
End Sub

