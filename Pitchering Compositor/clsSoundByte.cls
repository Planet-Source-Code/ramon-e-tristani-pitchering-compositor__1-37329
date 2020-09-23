VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSoundByte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*Member variable declaration
Private m_lngPitch As Long, m_lngLength As Long

Public Property Get SoundPitch() As Long
'*Defines the pitch property

   SoundPitch = m_lngPitch

End Property

Public Property Let SoundPitch(ByVal lngDat As Long)
'*Receives the pitch value

   m_lngPitch = lngDat

End Property

Public Property Get SoundLength() As Long
'*Defines the pitch property

   SoundLength = m_lngLength

End Property

Public Property Let SoundLength(ByVal lngDat As Long)
'*Receives the length value

   m_lngLength = lngDat
   
End Property
