VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pitchering Compositor "
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3615
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlFileOps 
      Left            =   240
      Top             =   4380
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraContainer 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4815
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton cmdLoad 
         Caption         =   "&Load SoundByte From Sequence File"
         Height          =   315
         Left            =   120
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Export sequence to a Visual Basic BAS module"
         Top             =   3900
         Width           =   3075
      End
      Begin VB.CheckBox chkUnlockPitch 
         Height          =   270
         Left            =   1380
         TabIndex        =   19
         ToolTipText     =   "Lock bars"
         Top             =   420
         Width           =   195
      End
      Begin VB.CheckBox chkUnlockLength 
         Height          =   270
         Left            =   2460
         TabIndex        =   18
         ToolTipText     =   "Lock bars"
         Top             =   420
         Width           =   195
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "&Quit "
         Height          =   315
         Left            =   2040
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Remove item from the list"
         Top             =   4380
         Width           =   1155
      End
      Begin VB.CheckBox chkLockBars 
         Height          =   270
         Left            =   2280
         TabIndex        =   16
         ToolTipText     =   "Lock bars"
         Top             =   1260
         Width           =   195
      End
      Begin VB.CommandButton cmdTest 
         BackColor       =   &H8000000C&
         Caption         =   "&Test Sound"
         Height          =   315
         Left            =   1380
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Test current sound bit"
         Top             =   2100
         Width           =   1815
      End
      Begin VB.CommandButton cmdModuleExport 
         Caption         =   "&Export Soundbyte to Song Module"
         Height          =   315
         Left            =   120
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Export sequence to a Visual Basic BAS module"
         Top             =   3540
         Width           =   3075
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit Sound"
         Height          =   315
         Left            =   1380
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Change sound bit settings"
         Top             =   2820
         Width           =   1815
      End
      Begin VB.CommandButton cmdAddSound 
         Caption         =   "&Add Sound"
         Height          =   315
         Left            =   1380
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Insert sound bit into sequence"
         Top             =   2460
         Width           =   1815
      End
      Begin VB.CommandButton cmdClearCollection 
         Caption         =   "&Clear Buffer"
         Height          =   315
         Left            =   1380
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Erase sound sequence"
         Top             =   3180
         Width           =   1815
      End
      Begin VB.CommandButton cmdRemoveItem 
         Caption         =   "&Remove"
         Height          =   315
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Remove item from the list"
         Top             =   3180
         Width           =   1155
      End
      Begin VB.CommandButton cmdPlaySequence 
         Appearance      =   0  'Flat
         Caption         =   "&Play Sequence"
         Height          =   675
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Play entire sound sequence"
         Top             =   2460
         Width           =   1155
      End
      Begin VB.ListBox lstSoundBits 
         Height          =   1950
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Sound bit"
         Top             =   420
         Width           =   1155
      End
      Begin VB.TextBox txtPitch 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   420
         Width           =   495
      End
      Begin VB.TextBox txtLength 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   420
         Width           =   495
      End
      Begin MSComctlLib.Slider sdrPitch 
         Height          =   1275
         Left            =   1500
         TabIndex        =   4
         ToolTipText     =   "Edit sound pitch"
         Top             =   720
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   2249
         _Version        =   393216
         Orientation     =   1
         Max             =   2000
         TickStyle       =   2
         TickFrequency   =   200
      End
      Begin MSComctlLib.Slider sdrLength 
         Height          =   1275
         Left            =   2640
         TabIndex        =   5
         ToolTipText     =   "Edit Sound duration"
         Top             =   720
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   2249
         _Version        =   393216
         Orientation     =   1
         Max             =   2000
         TickStyle       =   2
         TickFrequency   =   200
      End
      Begin VB.Label lblNoteSequence 
         AutoSize        =   -1  'True
         Caption         =   "Sequence :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   915
      End
      Begin VB.Label lblPitch 
         AutoSize        =   -1  'True
         Caption         =   "Pitch :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1380
         TabIndex        =   7
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblLength 
         AutoSize        =   -1  'True
         Caption         =   "Length :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2460
         TabIndex        =   6
         Top             =   120
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************
'*This code is the property of Ramon E. Tristani and PCS WebLabs.
'*As such, its use is restricted to non-comercial use by others. Sell
'*or sell or implement this implementation of this code without prior
'*authorization from the author is lawfully prohibited under the law
'*protecting intellectual property.
'*Copyright PCS WebLabs - Ramon E. Tristani
'********************************************************************************
'*Project Description:
'*
'*This application makes use of the Beep API function to create music
'*scores by changing sound pitch and length. It allows data to be saved
'*to disk, and thus loaded from a file. It supports editing of selected
'*sounds and play back.
'********************************************************************************


Option Explicit

'*Module constant declaration
Private Const MAX_VAL As Integer = 2000
Private Const NO_VAL As Integer = 0

'*Module variables and objects
Private mc_SoundCollect As clsCollectSounds
Private intListMemberIndex As Integer

Private Sub chkLockBars_Click()
'*Locks the slider bars together

   '*Locks the sliders according to the highest value
   If sdrPitch.Value > sdrLength.Value Then
      sdrLength.Value = sdrPitch.Value
      txtLength.Text = sdrLength.Value
   ElseIf sdrPitch.Value < sdrLength.Value Then
      sdrPitch.Value = sdrLength.Value
      txtPitch.Text = sdrLength.Value
   End If
   
End Sub

Private Sub chkUnlockLength_Click()
'*Unlocks the Length text box to allow for user input

   '*Locks and unlocks the text box depending on the check
   '*box value
   If chkUnlockLength.Value = vbChecked Then
      txtLength.Locked = False
      txtLength.BackColor = vbWhite
   Else
      txtLength.Locked = True
      txtLength.BackColor = vbButtonFace
   End If

End Sub

Private Sub chkUnlockPitch_Click()
'*Unlocks the Pitch text box to allow for user input

   '*Locks and unlocks the text box depending on the check
   '*box value
   If chkUnlockPitch.Value = vbChecked Then
      txtPitch.Locked = False
      txtPitch.BackColor = vbWhite
   Else
      txtPitch.Locked = True
      txtPitch.BackColor = vbButtonFace
   End If

End Sub

Private Sub cmdAddSound_Click()
'*Stores the SoundByte class in the collection and displays
'*the collection contents in the list box

   '*Variable and object declarations
   Dim c_SoundBit As clsSoundByte
   
   '*Create new instace of the class
   Set c_SoundBit = New clsSoundByte
   
   '*Class definition
   c_SoundBit.SoundPitch = CLng(Val(txtPitch.Text))
   c_SoundBit.SoundLength = CLng(Val(txtLength.Text))
   
   '*Adds the class to the collection
   mc_SoundCollect.SoundInclude c_SoundBit
   
   '*Outputs the class data to the list box
   lstSoundBits.AddItem c_SoundBit.SoundPitch & "," & Space(1) & c_SoundBit.SoundLength
   
   '*Enables the following controls if the buffer is populated
   If lstSoundBits.ListCount > 0 Then
      cmdEdit.Enabled = True
      cmdClearCollection.Enabled = True
      cmdPlaySequence.Enabled = True
      cmdRemoveItem.Enabled = True
      cmdModuleExport.Enabled = True
   End If
   
End Sub

Private Sub cmdClearCollection_Click()
'*Clears all items in the list box and collection buffer
   
   '*Clears the list box
   lstSoundBits.Clear
   
   '*Clears the collection buffer
   mc_SoundCollect.ClearBuffer
   
   '*Disables the controls if the buffer is empty
   Call NullControls
   
End Sub

Private Sub cmdEdit_Click()
'*Allows editing of specific sounds

   '*Variable declarations
   Dim lngPitch As Long, lngLength As Long, intIndexDat As Long
   
   '*Variable definition
   lngPitch = CLng(Val(txtPitch.Text))
   lngLength = CLng(Val(txtLength.Text))
   intIndexDat = intListMemberIndex + 1
   
  '*Error handler
  On Error GoTo ErrorTrap
  
   '*Edits the collection member
   mc_SoundCollect.EditSound intIndexDat, lngPitch, lngLength
   
   '*Updates the values listed in the list box
   lstSoundBits.RemoveItem (intListMemberIndex)
   lstSoundBits.AddItem lngPitch & "," & Space(1) & lngLength, intListMemberIndex
   
Exit Sub
ErrorTrap:
'*No event

End Sub

Private Sub cmdLoad_Click()
'*Loads the buffer and list box with file data
   
   '*Variable declarations
   Dim strFileName As String, intIndex As Integer, intFIleNum As Integer
   Dim lngPitchDat As Long, lngLengthDat As Long

   '*Control properties
   With cdlFileOps
      .InitDir = App.Path
      .Filter = "SoundByte Files (*.sbt)|*.sbt"
      .DialogTitle = "Save SoundByte File As..."
      .ShowOpen
   End With
   
   '*Variable definitions
   strFileName = cdlFileOps.FileName
   intFIleNum = FreeFile
   
   '*Opens file data
   mc_SoundCollect.LoadFile (strFileName)
   
   '*Re-opens the file for listing purposes
   Open strFileName For Input As #intFIleNum
   
   '*Populates the list box
   Do Until EOF(intFIleNum)
      Input #intFIleNum, lngPitchDat, lngLengthDat
      lstSoundBits.AddItem lngPitchDat & "," & Space(1) & lngLengthDat
   Loop
   
   '*Closes the file
   Close #intFIleNum
   
   '*Enables controls
   cmdEdit.Enabled = True
   cmdClearCollection.Enabled = True
   cmdPlaySequence.Enabled = True
   cmdRemoveItem.Enabled = True
   cmdModuleExport.Enabled = True
   
End Sub

Private Sub cmdModuleExport_Click()
'*Exports the sound data to a BAS module

   Dim strFileName As String
    
   '*Control properties
   With cdlFileOps
      .InitDir = App.Path
      .Filter = "SoundByte Files (*.sbt)|*.sbt"
      .DialogTitle = "Save SoundByte File As..."
      .ShowSave
   End With
   
   '*Variable definitions
   strFileName = cdlFileOps.FileName
   
   '*Writes soun data to file
   mc_SoundCollect.ExportToFile (strFileName)
   
End Sub

Private Sub cmdPlaySequence_Click()
'*Plays back the soundbyte sequence stored in the collection

   '*Plays back the sound collection
   mc_SoundCollect.PlayBack
   
End Sub

Private Sub cmdQuit_Click()
'*Exits the application

   '*Variable declarations
   Dim intResponse As Integer
   
   '*Variable definition
   intResponse = MsgBox("Are you sure you wish to exit Pitchering Compositor?", vbYesNo _
   + vbQuestion, "Exit?")
   
   If intResponse = vbYes Then Unload Me
   
End Sub

Private Sub cmdRemoveItem_Click()
'*Removes items from list box collection and buffer collection
   
   '*Variable declarations
   Dim intIndexDat As Integer
   
   '*Variable definitions
   intIndexDat = intListMemberIndex + 1
   
   '*Removes the item from the collection
   mc_SoundCollect.RemoveSound (intIndexDat)
   
   On Error GoTo ErrorTrap
   
   '*Removes the item from the list box
   lstSoundBits.RemoveItem (intListMemberIndex)
   
   '*Disables the controls if the buffer is empty
      Call NullControls
   
Exit Sub
ErrorTrap:
'*No event
   
End Sub

Private Sub NullControls()
'*Disables command buttons when the buffer is empty

   '*Disables the following controls if the buffer is empty
   If lstSoundBits.ListCount < 1 Then
      cmdEdit.Enabled = False
      cmdClearCollection.Enabled = False
      cmdPlaySequence.Enabled = False
      cmdRemoveItem.Enabled = False
      cmdModuleExport.Enabled = False
   End If
   
End Sub
Private Sub cmdTest_Click()
'*Assigns values to global UDT and calls the API wrapper to create
'*the beep with passed values
   
   '*Variable declaration
   Dim lngPitch As Long, lngLength As Long, lngSoundBit As Long
   
   '*Variable definition
   lngPitch = CLng(Val(txtPitch.Text))
   lngLength = CLng(Val(txtLength.Text))
   
   '*API wrapper call
   lngSoundBit = CreateSoundOutput(lngPitch, lngLength)

End Sub

Private Function CreateSoundOutput(ByVal lngDat1 As Long, ByVal lngDat2 As Long) As Long
'*Creates an API wrapper function that deals with the BEEP function

   '*Defines the function with the API call and passed parameters
   CreateSoundOutput = Beep(lngDat1, lngDat2)

End Function

Private Sub Form_Click()
PrintForm
End Sub

Private Sub Form_Load()
'*Sets initial application properties
   
   '*Variable declarations
   Dim objControl As Control
   
   '*Creates new instance of the collection object
   Set mc_SoundCollect = New clsCollectSounds
   
   '*Error trap incase controls don't support the hWnd
   '*property
   On Error Resume Next
   
   '*Cycles through each contained control on the form
   '*and applies the RefineBorder visual effect to all
   '*controls supporting the hWnd property.
   For Each objControl In Me.Controls
      If TypeOf objControl Is TextBox Or TypeOf objControl Is Frame _
      Or TypeOf objControl Is CommandButton Then
         RefineBorder (objControl.hWnd)
      End If
   Next objControl
   
   '*Disables controls
   cmdTest.Enabled = False
   cmdAddSound.Enabled = False
   cmdEdit.Enabled = False
   cmdClearCollection.Enabled = False
   cmdPlaySequence.Enabled = False
   cmdRemoveItem.Enabled = False
   cmdModuleExport.Enabled = False
   
End Sub

Private Sub lstSoundBits_Click()
'*Returns the value of the list box item clicked

   intListMemberIndex = lstSoundBits.ListIndex
   
End Sub

Private Sub sdrLength_Scroll()
'*Ensures the text box value remains updated
'*accurately throughout movement of the slider

   '*Displays the slider value in the text box
   txtLength.Text = sdrLength.Value
   
   '*Equalizes the values of the sliders when the
   '*sliders are locked together
   If chkLockBars.Value = vbChecked Then
      sdrPitch.Value = sdrLength.Value
      txtPitch.Text = sdrPitch.Value
   End If
   
End Sub

Private Sub sdrPitch_Scroll()
'*Ensures the text box value remains updated
'*accurately throughout movement of the slider

   '*Displays the slider value in the text box
   txtPitch.Text = sdrPitch.Value
   
   '*Equalizes the values of the sliders when the
   '*sliders are locked together
   If chkLockBars.Value = vbChecked Then
      sdrLength.Value = sdrPitch.Value
      txtLength.Text = sdrLength.Value
   End If
   
End Sub

Private Sub txtLength_Change()

   '*Sets the position of the slider to the value of the text box if the edit lock is disabled
   If chkUnlockLength.Value = vbChecked Then sdrLength.Value = Val(txtLength.Text)
   
   '*Prevents a value higher than "2000" in the text box
   If Val(txtLength.Text) > MAX_VAL Then txtLength.Text = MAX_VAL
   
   '*Resets the textbbox back to zero
   If txtLength.Text = "" Then txtLength.Text = NO_VAL
   
   '*Prevents from non-numeric characters in the text box
   If Not IsNumeric(txtLength.Text) Then txtLength.Text = ""
   
End Sub

Private Sub txtPitch_Change()

   '*Sets the position of the slider to the value of the text box if the edit lock is disabled
   If chkUnlockPitch.Value = vbChecked Then sdrPitch.Value = Val(txtPitch.Text)
   
   '*Prevents a value higher than "2000" in the text box
   If Val(txtPitch.Text) > MAX_VAL Then txtPitch.Text = MAX_VAL
   
   '*Resets the textbbox back to zero
   If txtPitch.Text = "" Then txtPitch.Text = NO_VAL
   
   '*Prevents from non-numeric characters in the text box
   If Not IsNumeric(txtPitch.Text) Then txtPitch.Text = ""
   
   '*Enables both the sound test and sound add
   If Val(txtPitch.Text) > 0 Then
      cmdTest.Enabled = True
      cmdAddSound.Enabled = True
   Else
      cmdTest.Enabled = False
      cmdAddSound.Enabled = False
   End If
   
End Sub


