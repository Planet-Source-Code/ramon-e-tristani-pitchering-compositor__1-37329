Attribute VB_Name = "modAppDeclarations"
Option Explicit


'*Module API Calls
Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration _
As Long) As Long

