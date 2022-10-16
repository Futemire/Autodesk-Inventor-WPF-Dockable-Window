Imports System.Text
Imports Inventor

Public Class WPFDockableWindowBaseClass

    Public Sub New(InvApp As Inventor.Application, InternalName As String, WindowTitle As String, WPFControl As UIElement, Optional Height As Integer = 200.0, Optional ShowTitle As Boolean = True)
        ' This call is required by the designer..
        BindingErrorTraceListener.SetTrace()
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

Public Class BindingErrorTraceListener
    Inherits DefaultTraceListener

    Private Shared _Listener As BindingErrorTraceListener

    Public Shared Sub SetTrace()
        SetTrace(SourceLevels.[Error], TraceOptions.None)
    End Sub

    Public Shared Sub SetTrace(ByVal level As SourceLevels, ByVal options As TraceOptions)
        If _Listener Is Nothing Then
            _Listener = New BindingErrorTraceListener()
            PresentationTraceSources.DataBindingSource.Listeners.Add(_Listener)
        End If

        _Listener.TraceOutputOptions = options
        PresentationTraceSources.DataBindingSource.Switch.Level = level
    End Sub

    Public Shared Sub CloseTrace()
        If _Listener Is Nothing Then
            Return
        End If

        _Listener.Flush()
        _Listener.Close()
        PresentationTraceSources.DataBindingSource.Listeners.Remove(_Listener)
        _Listener = Nothing
    End Sub

    Private _Message As StringBuilder = New StringBuilder()

    Private Sub New()
    End Sub

    Public Overrides Sub Write(ByVal message As String)
        _Message.Append(message)
    End Sub

    Public Overrides Sub WriteLine(ByVal message As String)
        _Message.Append(message)
        Dim final = _Message.ToString()
        _Message.Length = 0
        MessageBox.Show(final, "Binding Error", MessageBoxButton.OK, MessageBoxImage.[Error])
    End Sub
End Class