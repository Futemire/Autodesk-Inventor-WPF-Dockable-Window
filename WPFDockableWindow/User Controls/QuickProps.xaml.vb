Imports System
Imports System.Collections.Generic
Imports System.Windows
Imports System.Windows.Data
Imports System.ComponentModel
Imports System.Runtime.CompilerServices
Imports System.Windows.Threading
Imports System.Collections.ObjectModel
Imports Inventor
Imports System.Runtime.InteropServices
Imports Microsoft.VisualBasic.ApplicationServices

Public Class QuickProps : Implements INotifyPropertyChanged

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private listViewSortCol As GridViewColumnHeader = Nothing
    Private listViewSortAdorner As SortAdorner = Nothing

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
        'add the DataContext here..?
        'DataContext = iProps

        AddHandler InvApp.ApplicationEvents.OnActivateDocument, AddressOf LoadProps

        'Dim view As CollectionView = CType(CollectionViewSource.GetDefaultView(LV_iPropertyBlocks.ItemsSource), CollectionView)
        'view.Filter = AddressOf PropertyFilter

    End Sub

    'Private Function PropertyFilter(obj As Object) As Boolean
    '    If String.IsNullOrEmpty(txtFilter.Text) Then
    '        Return True
    '    Else
    '        Return TryCast(SelectedPropertyValue, User).Name.IndexOf(txtFilter.Text, StringComparison.OrdinalIgnoreCase) >= 0
    '    End If
    'End Function

    Private Sub LoadProps(ByVal DocumentObject As _Document, ByVal BeforeOrAfter As EventTimingEnum, ByVal Context As NameValueMap, <Out> ByRef HandlingCode As HandlingCodeEnum)
        If BeforeOrAfter = EventTimingEnum.kAfter Then
            Dim TempList As New ObservableCollection(Of KeyValuePair(Of String, String))
            iProps.Clear()
            For Each Prop As [Property] In _InvApp.ActiveDocument.PropertySets("Inventor User Defined Properties")
                TempList.Add(New KeyValuePair(Of String, String)(Prop.DisplayName, Prop.Expression))
            Next

            iProps = TempList
            'add the DataContext here instead..?
            LV_iPropertyBlocks.DataContext = iProps
            'there's no way this works, right..?
            LV_iPropertyBlocks.ItemsSource = iProps

            For Each ite As KeyValuePair(Of String, String) In iProps
                Debug.WriteLine("ITem Key: " & ite.Key & "   Item Value: " & ite.Value)
            Next
            HandlingCode = HandlingCodeEnum.kEventHandled
        Else
            HandlingCode = HandlingCodeEnum.kEventNotHandled
        End If
    End Sub

    Public EmptyDelegate As Action = Sub()
                                     End Sub

    Public Sub Refresh(UI_Element As UIElement)
        UI_Element.Dispatcher.Invoke(DispatcherPriority.Render, EmptyDelegate)
    End Sub

    Private Sub BTN_AddIpropertyBlock_Click(sender As Object, e As RoutedEventArgs) Handles BTN_AddIpropertyBlock.Click
        MsgBox("Imma be add'n some properties once I add the code to do so!   Yep...")
    End Sub

    Protected Function SetProperty(Of T)(ByRef field As T, newValue As T, <CallerMemberName> Optional propertyName As String = Nothing) As Boolean
        If Not Equals(field, newValue) Then
            field = newValue
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
            Return True
        End If

        Return False
    End Function

    Private _selectedPropertyValue As Object

    Public Property SelectedPropertyValue As Object
        Get
            Return _selectedPropertyValue
        End Get
        Set(value As Object)
            SetProperty(_selectedPropertyValue, value)
            NotifyPropertyChanged()
        End Set
    End Property

    Private Sub txtFilter_TextChanged(sender As Object, e As TextChangedEventArgs)
        CollectionViewSource.GetDefaultView(LV_iPropertyBlocks.ItemsSource).Refresh()
    End Sub

    Private Sub lvUsersColumnHeader_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim column As GridViewColumnHeader = (TryCast(sender, GridViewColumnHeader))
        Dim sortBy As String = column.Tag.ToString()

        If listViewSortCol IsNot Nothing Then
            AdornerLayer.GetAdornerLayer(listViewSortCol).Remove(listViewSortAdorner)
            LV_iPropertyBlocks.Items.SortDescriptions.Clear()
        End If

        Dim newDir As ListSortDirection = ListSortDirection.Ascending
        If listViewSortCol Is column AndAlso listViewSortAdorner.Direction = newDir Then newDir = ListSortDirection.Descending
        listViewSortCol = column
        listViewSortAdorner = New SortAdorner(listViewSortCol, newDir)
        AdornerLayer.GetAdornerLayer(listViewSortCol).Add(listViewSortAdorner)
        LV_iPropertyBlocks.Items.SortDescriptions.Add(New SortDescription(sortBy, newDir))
    End Sub
End Class

Public Class SortAdorner
    Inherits Adorner

    Private Shared ascGeometry As Geometry = Geometry.Parse("M 0 4 L 3.5 0 L 7 4 Z")
    Private Shared descGeometry As Geometry = Geometry.Parse("M 0 0 L 3.5 4 L 7 0 Z")
    Public Property Direction As ListSortDirection

    Public Sub New(ByVal element As UIElement, ByVal dir As ListSortDirection)
        MyBase.New(element)
        Me.Direction = dir
    End Sub

    Protected Overrides Sub OnRender(ByVal drawingContext As DrawingContext)
        MyBase.OnRender(drawingContext)
        If AdornedElement.RenderSize.Width < 20 Then Return
        Dim transform As TranslateTransform = New TranslateTransform(AdornedElement.RenderSize.Width - 15, (AdornedElement.RenderSize.Height - 5) / 2)
        drawingContext.PushTransform(transform)
        Dim geometry As Geometry = ascGeometry
        If Me.Direction = ListSortDirection.Descending Then geometry = descGeometry
        drawingContext.DrawGeometry(Brushes.Black, Nothing, geometry)
        drawingContext.Pop()
    End Sub
End Class
