Class Button1Window

    ''' <summary>
    ''' AddinBaseClass object to be used referenced by this window.
    ''' </summary>
    Private _Addin As AddinBaseClass

    ''' <summary>
    ''' Just a sample event to show how to hook back into the Button1ControlDef. <para/>
    ''' This can be modified to fit your needs, or deleted.
    ''' </summary>
    Public Event MyEvent(message As String)

    ''' <summary>
    ''' Creates a new instance of the window and moves it to the last saved location.
    ''' </summary>
    ''' <param name="Addin">[AddinBaseClass] AddinBaseClass object to be used referenced by this window.</param>
    Sub New(ByVal Addin As AddinBaseClass)

        ' This call is required by the designer.
        InitializeComponent()

        'Link the Add-in Object
        _Addin = Addin

        'Set Inventor as the parent of this window.
        'This will allow your window to be minimized with Inventor.
        Call _Addin.SetInventorAsOwnerWindow(Me)

        'Set the saved window location.
        Me.Top = My.Settings.Button1WindowLocation.X
        Me.Left = My.Settings.Button1WindowLocation.Y

    End Sub

    Private Sub button_Click(sender As Object, e As RoutedEventArgs) Handles button.Click
        RaiseEvent MyEvent("And I think you like turtles too!")
    End Sub

    Private Sub button_Copy_Click(sender As Object, e As RoutedEventArgs) Handles button_Copy.Click
        'Display the balloon tip.
        _Addin.InvApp.UserInterfaceManager.BalloonTips("Fütemire_WPFDockableWindow_BalloonTip").Interval = 3000
        _Addin.InvApp.UserInterfaceManager.BalloonTips("Fütemire_WPFDockableWindow_BalloonTip").Display("I think he likes turtles....")

    End Sub

    ''' <summary>
    ''' Closes the window and saves its last location on screen to the app settings.
    ''' </summary>
    Private Sub Button1Window_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        My.Settings.Button1WindowLocation = New System.Drawing.Point(Me.Top, Me.Left)
        My.Settings.Save()
    End Sub
End Class