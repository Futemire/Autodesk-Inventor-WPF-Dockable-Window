Imports System.ComponentModel
Imports System.Runtime.CompilerServices
Imports System.Windows.Threading
Imports System.Collections.ObjectModel
Imports Inventor
Public Class QuickProps : Implements INotifyPropertyChanged

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Private Sub NotifyPropertyChanged(<CallerMemberName> Optional ByVal propertyName As String = Nothing)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub

    Public Property iProps As ObservableCollection(Of KeyValuePair(Of String, String))
        Get
            Return _iProps
        End Get
        Set(value As ObservableCollection(Of KeyValuePair(Of String, String)))
            _iProps = value
            NotifyPropertyChanged()
            Refresh(Me)
        End Set
    End Property
    Private _iProps As ObservableCollection(Of KeyValuePair(Of String, String))

    Private _InvApp As Inventor.Application
    Public Sub New(InvApp As Inventor.Application)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        _InvApp = InvApp
        iProps = New ObservableCollection(Of KeyValuePair(Of String, String))

        AddHandler InvApp.ApplicationEvents.OnActivateDocument, AddressOf LoadProps



    End Sub

    Private Sub LoadProps()
        Dim TempList As New ObservableCollection(Of KeyValuePair(Of String, String))
        iProps.Clear()

        For Each Prop As [Property] In _InvApp.ActiveDocument.PropertySets("Inventor User Defined Properties")
            TempList.Add(New KeyValuePair(Of String, String)(Prop.DisplayName, Prop.Expression))
        Next

        iProps = TempList


        For Each ite As KeyValuePair(Of String, String) In iProps
            Debug.WriteLine("ITem Key: " & ite.Key & "   Item Value: " & ite.Value)
        Next


    End Sub

    Public EmptyDelegate As Action = Sub()
                                     End Sub

    Public Sub Refresh(UI_Element As UIElement)
        UI_Element.Dispatcher.Invoke(DispatcherPriority.Render, EmptyDelegate)
    End Sub

    Private Sub BTN_AddIpropertyBlock_Click(sender As Object, e As RoutedEventArgs) Handles BTN_AddIpropertyBlock.Click
        MsgBox("Imma be add'n some properties once I add the code to do so!   Yep...")
    End Sub
End Class
