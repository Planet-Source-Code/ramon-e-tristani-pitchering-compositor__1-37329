VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCollectSounds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_colSounds As Collection
Private c_ClassDat As clsSoundByte

Public Sub SoundInclude(ByRef classData As clsSoundByte)
'*Includes the soundbit to the internal class collection

   m_colSounds.Add classData
   
End Sub

Private Sub Class_Initialize()
'*Creates a new instance of the collection
   
   Set m_colSounds = New Collection

End Sub

Public Sub PlayBack()
'*Plays back stored sound data
   
   '*Variable declarations
   Dim intIndex As Integer
   
   '*Cycles through the collection playing back stored sound data
   For intIndex = 1 To m_colSounds.Count
      With m_colSounds
         DoEvents
         Beep .Item(intIndex).SoundPitch, .Item(intIndex).SoundLength
      End With
   Next intIndex
   
End Sub

Public Sub RemoveSound(ByRef datIndex As Integer)
'*Removes sound data from the collection

   On Error GoTo ErrorTrap

   '*Removes specific item from collection
   m_colSounds.Remove (datIndex)
   
Exit Sub
ErrorTrap:
'*Alerts the user that an item must be selected

   MsgBox "Please select an item to remove.", vbOKOnly + vbExclamation, "Select"
   
End Sub

Public Sub ClearBuffer()
'*Clears the collection contents

   '*Variable declaration
   Dim intIndex As Integer
   
   '*Variable definitions
   intIndex = m_colSounds.Count
   
   '*Removes each item form the collection buffer
   Do While intIndex <> 0
      m_colSounds.Remove (intIndex)
      intIndex = intIndex - 1
   Loop
   
End Sub

Public Sub TestSound(ByRef classData As clsSoundByte)
'*Tests the sound byte
   
   '*Processes the sound data through the API call
   Beep classData.SoundPitch, classData.SoundLength
   
End Sub

Public Sub EditSound(ByVal datIndex As Integer, ByVal pitchDat As Long, ByVal lengthDat _
As Long)
'*Edits the sound stored in the collection buffer

   On Error GoTo ErrorTrap

   '*Edits specific item from collection
   With m_colSounds
      .Item(datIndex).SoundPitch = pitchDat
      .Item(datIndex).SoundLength = lengthDat
   End With
      
   
Exit Sub
ErrorTrap:
'*Alerts the user that an item must be selected

   MsgBox "Please select an item to edit.", vbOKOnly + vbExclamation, "Select"
   
End Sub

Public Sub ExportToFile(ByRef strData As String)
'*Exports collection data to file
   
   '*Variable declarations
   Dim intFIleNum As Integer, intIndex As Integer
   
   '*Variable definitions
   intFIleNum = FreeFile

   Open strData For Output As #intFIleNum
   
   '*Writes collection data to file
   For intIndex = 1 To m_colSounds.Count
      Write #intFIleNum, m_colSounds.Item(intIndex).SoundPitch, _
      m_colSounds.Item(intIndex).SoundLength
   Next intIndex
   
   '*Closes the file
   Close #intFIleNum
   
End Sub

Public Sub LoadFile(ByRef strData As String)
'*Loads sound data from file

   '*Variable declarations
   Dim intFIleNum As Integer, intIndex As Integer, lngPitch As Long, lngLength As Long
   
   '*Variable definitions
   intFIleNum = FreeFile
   
   '*Empties the collection is items are stored within it
   If m_colSounds.Count > 0 Then
      Call ClearBuffer
   End If
   
   '*Opens the data file
   Open strData For Input As #intFIleNum
   
   '*Loads the data from the file into the collection
   Do Until EOF(intFIleNum)
      '*Creates new instance of clsSoundByte
      Set c_ClassDat = New clsSoundByte
      Input #intFIleNum, lngPitch, lngLength
      c_ClassDat.SoundPitch = lngPitch
      c_ClassDat.SoundLength = lngLength
      m_colSounds.Add c_ClassDat
   Loop
    
   '*Closes the file
   Close #intFIleNum
   
End Sub

