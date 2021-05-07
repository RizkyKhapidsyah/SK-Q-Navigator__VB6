Attribute VB_Name = "Module3"
Global lLeft, lTop, lWidth, lHeight As Long
Public Function sudpre()
If Form1.TabStrip1.SelectedItem.index = 1 Then
Form1.TabStrip1.Tabs(Form1.TabStrip1.Tabs.Count).Selected = True
Else
Form1.TabStrip1.Tabs(Form1.TabStrip1.SelectedItem.index - 1).Selected = True
    Form1.TabStrip1.SetFocus
    Form1.TabStrip1.Refresh
End If
End Function
Public Function sudnex()
If Form1.TabStrip1.SelectedItem.index = Form1.TabStrip1.Tabs.Count Then
Form1.TabStrip1.Tabs(1).Selected = True
Else
Form1.TabStrip1.Tabs(Form1.TabStrip1.SelectedItem.index + 1).Selected = True
    Form1.TabStrip1.SetFocus
    Form1.TabStrip1.Refresh
End If
End Function


Sub RepositionProgressBar()
        'RESIZE and POSITION THE PROGRESS BAR
        lLeft = Form1.StatusBar.Panels(1).Left + 10
        lTop = Form1.StatusBar.Top + 80
        lWidth = Form1.StatusBar.Panels(1).Width - 20
        lHeight = Form1.StatusBar.Height - 70
        Form1.ProgressBar1.Move lLeft, lTop, lWidth, lHeight
Form1.ProgressBar2.Move lLeft, lTop, lWidth, lHeight
Form1.ProgressBar3.Move lLeft, lTop, lWidth, lHeight
Form1.ProgressBar4.Move lLeft, lTop, lWidth, lHeight
Form1.ProgressBar5.Move lLeft, lTop, lWidth, lHeight
Form1.ProgressBar6.Move lLeft, lTop, lWidth, lHeight
Form1.ProgressBar7.Move lLeft, lTop, lWidth, lHeight
Form1.ProgressBar8.Move lLeft, lTop, lWidth, lHeight
Form1.ProgressBar9.Move lLeft, lTop, lWidth, lHeight
Form1.ProgressBar10.Move lLeft, lTop, lWidth, lHeight

End Sub

