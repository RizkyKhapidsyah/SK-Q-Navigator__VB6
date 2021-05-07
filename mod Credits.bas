Attribute VB_Name = "modCredits"
Public Const EM_GETLINECOUNT = &HBA
Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                            (ByVal hwnd As Long, _
                            ByVal wMsg As Long, _
                            ByVal wParam As Long, _
                            lParam As Any) As Long   ' <---
