Attribute VB_Name = "modRefineBorder"
'*******************************************************************
'*In order to make use of this module you must usr the hWnd
'*property of the control you wish to alter. Of course... the
'*control should have this property available first. I think the only
'*other control that doesn't have it is the line control. Anyway, to
'*use this use the following syntax:
'*
'*                            RefineBorder(controlname.hWnd)
'*
'*The best place to make this happen is in either a Sub Main(),
'*Form Load(), etc. Have fun!
'*******************************************************************


Option Explicit

'*API Calls
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
(ByVal hWnd As Long, ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
(ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
ByVal cy As Long, ByVal wFlags As Long) As Long

'*Module API constants
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000

Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4

Public Function RefineBorder(ByVal hWnd As Long)
'*Alters the border appereance of contained controls
   Dim l_lRetrieveVal As Long
   
   '*Retrieve the current border style
   l_lRetrieveVal = GetWindowLong(hWnd, GWL_EXSTYLE)

   '*Calculate border style to use
   l_lRetrieveVal = l_lRetrieveVal Or WS_EX_STATICEDGE And Not WS_EX_CLIENTEDGE

   '*Apply the changes
   SetWindowLong hWnd, GWL_EXSTYLE, l_lRetrieveVal
   SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or _
   SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
   
End Function



