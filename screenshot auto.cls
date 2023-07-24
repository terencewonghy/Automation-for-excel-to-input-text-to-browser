VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "¤u§@ªí1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
#If VBA7 Then
    Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As LongPtr)
#Else
    Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
#End If

Private Const VK_SNAPSHOT As Byte = &H2C
Private Const KEYEVENTF_KEYUP As Long = &H2

Sub BindKey()
    Application.OnKey "^q", "CaptureScreen"
End Sub

Sub CaptureScreen()
    Dim wdApp As Word.Application
    Dim wdDoc As Word.Document
    Dim wdRange As Word.Range
    
    On Error GoTo ErrorHandler
    
    'Create a new Word document
    Set wdApp = New Word.Application
    Set wdDoc = wdApp.Documents.Add
    Set wdRange = wdDoc.Range(Start:=0, End:=0)
    
    'Cap Screen to clipboard
    keybd_event VK_SNAPSHOT, 1, 0, 0
    keybd_event VK_SNAPSHOT, 0, KEYEVENTF_KEYUP, 0
    Application.Wait Now + TimeValue("0:00:01")
    
    'Paste screenshot from clipboard into Word document
    wdApp.Selection.Paste
    
    'Save and close the word doc
    wdDoc.SaveAs "C:\Users\Wong Ho Yin\Desktop\Screenshot.docx"
    wdDoc.Close
    
    'Clean up
    Set wdRange = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing
    
    MsgBox "Screenshot added to Screenshot.docx"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub



