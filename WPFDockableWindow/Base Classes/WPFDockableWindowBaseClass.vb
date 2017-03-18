Imports Inventor

Public Class WPFDockableWindowBaseClass

    Public Sub New(InvApp As Inventor.Application, InternalName As String, WindowTitle As String, WPFControl As UIElement, Optional Height As Integer = 200.0, Optional ShowTitle As Boolean = True)
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        'Windows Form Settings
        Visible = True
        FormBorderStyle = Forms.FormBorderStyle.None
        'Add the WPF UIElemet to the ElementHost
        ElementHost1.Child = WPFControl
        'Create the Inventor DockableWindow
        Dim DocWin = InvApp.UserInterfaceManager.DockableWindows.Add(CLID, InternalName, WindowTitle)
        'Add the Windows Form to the Inventor DockableWindow
        DocWin.AddChild(Me.Handle)
        DocWin.ShowTitleBar = ShowTitle
        DocWin.Height = Height
        DocWin.Visible = True
    End Sub


    Private Function UnWrapWindow(WindowToUnwrap As Window) As UIElement
        'Need to investigate on how to "UnWrap" the contents of a WPF Window.

        'Best practice to avoid needing this method would be to create a WPF User control 
        'that could be use in a Window or a Dockable Window. But I will investigate this for 
        'the sake of using already existing Windows and ease of use going forward.

        Return Nothing
    End Function


End Class