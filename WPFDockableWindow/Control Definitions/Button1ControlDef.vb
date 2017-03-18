Imports Inventor

''' <summary>
''' This is a sample class that demonstrates creating a Control Definition by inheriting from the ControlDefBaseClass. <para/>
''' The ControlDefBaseClass handles the repetitive work of creating/accessing Tabs and Panels as well as adding Control Definitions to specified Panels. 
''' You can modify this to fit your needs, or delete it.
''' </summary>
Public Class Button1CDef : Inherits ControlDefBaseClass

    'Inventor's button definition for this ControlDef.
    Public WithEvents ButtonDef As ButtonDefinition

    'The window that this control definition will display/interact with.
    Public WithEvents Button1Win As Button1Window

    'Use Inventor's Transaction manager to set undo checkpoints.
    Public TransMan As TransactionManager
    Public AddinTransAction As Transaction

    'Events to be used in the current ControlDef.
    Public WithEvents AppEvents As ApplicationEvents

    ''' <summary>
    ''' [Private] Active Inventor Document reference object.
    ''' </summary>
    Private _ActDoc As Document

    ''' <summary>
    ''' [Private] Document sub type such as "SolidPart" or "Sheetmetal".
    ''' </summary>
    Private _ComponentType As InventorDocumentSubType

    ''' <summary>
    ''' [Private] AddinBaseClass reference object.
    ''' </summary>
    Private _Addin As AddinBaseClass


    ''' <summary>
    ''' Creates the ComponentDefinition.
    ''' </summary>
    ''' <param name="Addin"></param>
    Public Sub New(ByVal Addin As AddinBaseClass)

        'Initialize the base class.
        'This MUST be called before anything else.
        MyBase.New(Addin)

        'Populate to the private object.
        _Addin = Addin

        'Create the balloon tips that will be displayed to the user.
        'This is less intrusive than constantly displaying message boxes.
        _Addin.InvApp.UserInterfaceManager.BalloonTips.Add("Fütemire_WPFDockableWindow_BalloonTip", '<---- Company Name, Addin Name, Balloon Tip Name, No Spaces
                                                           "Hey Now", '<---- Balloon Tip Title, Spaces OK
                                                           "" '<---- I leave this blank because I specify the message when I display the Balloon tip item.
                                                           )


        'Hookup the event objects needed for this control.
        AppEvents = _Addin.InvApp.ApplicationEvents

        'Define the button name & internal name. The Button Name can be anything but the Internal Name must be unique and cannot contain spaces.
        'Because there are nearly 2000 button definition already in Inventor I typically create Internal Names by using this format...
        'MyCompanyName_MyAddinSafeName_MyButtonName

        ButtonName = "WPFDockableWindow Button 1"
        LargeIcon = My.Resources.I32
        SmallIcon = My.Resources.I16

        ButtonInternalName = "Fütemire_WPFDockableWindow_Button1Window"
        ShowButtonText = True


        'General hover text to be displayed when user hovers over the button.
        'Note that this text will ignored if you decide to use the progressive tool tip.
        ToolTip = "Launch ""I Like Turtles."""

        'This feature is optional. If set to True then your button will display a more elaborate tool tip with a
        'title, space for a longer button description, and the option to add an image or video string.
        Progressive = False
        ProgressiveTitle = Nothing
        ProgressiveDesc = Nothing
        ProgressiveImage = Nothing
        ProgressiveVideo = Nothing

        'Create a new list of the ribbons you want to add the button to.
        IncludeInRibbons = New List(Of RibbonType)({RibbonType.Part, RibbonType.Assembly})

        'Specify the Internal Name of the Ribbon Tab that the button should be included in.
        'If Left blank then a control definition will be created but NOT added to UI.
        'If a Ribbon Tab with the specified InternalName is found then the existing object is used.
        'If a Ribbon Tab with the specified InternalName is not fount then it is created.
        'Use the same nomenclature as explained for the ButtonInternalName above.
        '<Required if adding button to UI.>
        TabInternalName = "Fütemire_WPFDockableWindow_Tab"

        'Friendly Name of Tab to create if InteranlName is not found. This name can contain spaces.
        TabName = "Fütemire"

        'Specify the Internal Name of the Ribbon Panel that the button should be included in.
        'If Left blank then a control definition will be created but NOT added to UI.
        'If a Ribbon Panel with the specified InternalName is found then the existing object is used.
        'If a Ribbon Panel with the specified InternalName is not fount then it is created.
        'Use the same nomenclature as explained for the ButtonInternalName above.
        '<Required if adding button to UI.>
        PanelInternalName = "Fütemire_WPFDockableWindow_Panel1"

        'Friendly Name of Tab to create if InteranlName is not found. This name can contain spaces.
        PanelName = "Some Panel Eh?"

        'Create the control button definition and link it to your Global Private WithEvents button def above.
        ButtonDef = CreateControlDefButton()

        'Make the call to add the button to the User Interface.
        Call AddButtonToUI()

    End Sub

    ''' <summary>
    ''' Launch the main window for the WPFDockableWindow add-in.
    ''' </summary>
    Private Sub ButtonDefinition_OnExecute(Context As NameValueMap) Handles ButtonDef.OnExecute
        Try

            'Don't allow multiple instances of the window.
            If Not Button1Win Is Nothing Then Exit Sub

            'Set the Active Document & Active RangeBox
            _ActDoc = _Addin.InvApp.ActiveDocument

            'Get the active document type.
            _ComponentType = _Addin.InvApp.ActiveDocument.GetDocumentSubType

            'Get the Active Document.
            Select Case True
                Case _ComponentType = InventorDocumentSubType.Assembly
                    _ActDoc = DirectCast(_ActDoc, AssemblyDocument)
                Case _ComponentType = InventorDocumentSubType.SolidPart
                    _ActDoc = DirectCast(_ActDoc, PartDocument)
                Case _ComponentType = InventorDocumentSubType.SheetMetal
                    _ActDoc = DirectCast(_ActDoc, PartDocument)
                Case Else
                    'We don't have a valid document type.
                    Exit Sub
            End Select


            'Transaction set undo checkpoints that can be called later if need be.
            'Transactions combine many changes into one undo, gives it a name and allows the user 
            'to click the undo button to undo everything in one click. In this addin example we
            'haven't changed anything, but I wanted to share how to implement Transactions.

            'Create the transaction manager and start a transaction.
            TransMan = _Addin.InvApp.TransactionManager
            AddinTransAction = TransMan.StartTransaction(_ActDoc, "WPFDockableWindow")

            'Create the window object.
            Button1Win = New Button1Window(_Addin)

            'You could create events in your MainWindow and add handlers here to handle user interactions.
            'It would look something like this...
            AddHandler Button1Win.MyEvent, AddressOf EventFromButton1Window
            AddHandler Button1Win.Closed, AddressOf OnButton1WinClose

            'Now show the window.
            Button1Win.Show()

        Catch ex As Exception
            'On error we need to abort the transaction if it is uncommitted.
            If Not AddinTransAction Is Nothing Then
                AddinTransAction.Abort()
                AddinTransAction = Nothing
            End If

            'Log the error.
            _Addin.Logging.WriteToErrorLog("There was an error opening the form!", _Addin, ex)
        End Try
    End Sub


#Region "Inventor Events"
    'An example of where you can handle events that this control def has hooked into.
    Private Sub AppEvents_OnActivateDocument(DocumentObject As _Document, BeforeOrAfter As EventTimingEnum,
                                             Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum) Handles AppEvents.OnActivateDocument
        Select Case BeforeOrAfter
            Case EventTimingEnum.kBefore

            Case EventTimingEnum.kAfter

            Case EventTimingEnum.kAbort
        End Select

    End Sub

    Private Sub AppEvents_OnSaveDocument(DocumentObject As _Document, BeforeOrAfter As EventTimingEnum,
                                         Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum) Handles AppEvents.OnSaveDocument
        Select Case BeforeOrAfter
            Case EventTimingEnum.kBefore

            Case EventTimingEnum.kAfter

            Case EventTimingEnum.kAbort

        End Select
    End Sub
#End Region

#Region "Sample Event Handler"
    ''' <summary>
    ''' Just a sample event handler to show how to handle events from your Button1Window. <para/>
    ''' This event is fired FROM the Button1Window when the MyEvent event is raised.
    ''' </summary>
    Public Sub EventFromButton1Window(message As String)
        MsgBox(message)
    End Sub

    ''' <summary>
    ''' This event is attached in this class and is called when the Button1Window closes.
    ''' </summary>
    Private Sub OnButton1WinClose()

        'Set the Button1Win to Nothing so we can open it again the next time the button is clicked.
        Button1Win = Nothing

        'The window closed so lets stop logging undo checkpoints. 
        'End the transaction so it is committed.
        If Not AddinTransAction Is Nothing Then
            AddinTransAction.End()
            AddinTransAction = Nothing
        End If

    End Sub


#End Region
End Class