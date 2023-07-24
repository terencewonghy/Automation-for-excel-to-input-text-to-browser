VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "¤u§@ªí1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    If ActiveSheet.Name <> "Sheet1" Then
        MsgBox "Please select the Sheet1 worksheet and try again."
        Exit Sub
    End If
    
    If Range("A1").Value = "" Then 'Cell boxes should be adjusted according to different cases, this is just a demo
        MsgBox "Please enter a valid ID in cell A1 and try again."
        Exit Sub
    End If
    
     If Range("A2").Value = "" Then 'Cell boxes should be adjusted according to different cases, this is just a demo
        MsgBox "Please enter a valid ID in cell A1 and try again."
        Exit Sub
    End If
    
    OpenChrome
    Application.Wait (Now + TimeValue("0:00:10")) 'Browser responsive time might vary according to different running computers, adjust according to needs
    Range("A1").Select
    SendKeys ActiveCell.Value, True
    
    Range("A2").Select
    SendKeys "{TAB}", True
    SendKeys ActiveCell.Value, True
    SendKeys "{Enter}"
End Sub

Sub OpenChrome()
    Dim chromePath As String
    chromePath = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" 'location of desired browser, Edge not yet tested
    Shell (chromePath & " -url https://www.google.com") 'users should change the url according to needs
End Sub

'This Macro is currently under developement and optimization, further improvements and changes may be made! (Ver 1.0)
