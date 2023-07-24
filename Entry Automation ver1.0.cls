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
    OpenChrome
    Application.Wait (Now + TimeValue("0:00:03"))
    Range("A1").Select
    SendKeys ActiveCell.Value, True
    
    Range("A2").Select
    SendKeys "{TAB}", True
    SendKeys ActiveCell.Value, True
    SendKeys "{Enter}"
End Sub

Sub OpenChrome()
    Dim chromePath As String
    chromePath = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    Shell (chromePath & " -url https://www.google.com")
End Sub

