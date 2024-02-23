Imports System.Drawing
Imports System.Windows.Forms
Imports Office = Microsoft.Office.Core
Imports Visio = Microsoft.Office.Interop.Visio
Partial Public Class ThisAddIn
    Public Property AddinUI As AddinUI = new AddinUI()

    Protected Overrides Function CreateRibbonExtensibilityObject() As Office.IRibbonExtensibility
        Return _addinUI
    End Function

    ''' 
    ''' A simple command
    ''' 
    Public Sub Command1()
        MessageBox.Show(
            "Hello from command 1!",
            "visioTables")
    End Sub

    ''' 
    ''' A command to demonstrate conditionally enabling/disabling.
    ''' The command gets enabled only when a shape is selected
    ''' 
    Public Sub Command2()
        If Application Is Nothing OrElse Application.ActiveWindow Is Nothing OrElse Application.ActiveWindow.Selection Is Nothing Then Exit Sub

        MessageBox.Show(
            String.Format("Hello from (conditional) command 2! You have {0} shapes selected.", Application.ActiveWindow.Selection.Count),
            "visioTables")
    End Sub

    ''' 
    ''' Callback called by the UI manager when user clicks a button
    ''' Should do something meaninful wehn corresponding action is called.
    ''' 
    Public Sub OnCommand(commandId As String)
        Select Case commandId
            Case "Command1"
                Command1()
                Return

            Case "Command2"
                Command2()
                Return

        End Select
    End Sub

    ''' 
    ''' Callback called by UI manager.
    ''' Should return if corresponding command shoudl be enabled in the user interface.
    ''' By default, all commands are enabled.
    ''' 
    Public Function IsCommandEnabled(commandId As String) As Boolean
        Select Case commandId
            Case "Command1"
                ' make command1 always enabled
                Return True

            Case "Command2"
                ' make command2 enabled only if a window is opened
                Return Application IsNot Nothing AndAlso Application.ActiveWindow IsNot Nothing AndAlso Application.ActiveWindow.Selection.Count > 0
            Case Else
                Return True
        End Select
    End Function

    ''' 
    ''' Callback called by UI manager.
    ''' Should return if corresponding command (button) is pressed or not (makes sense for toggle buttons)
    ''' 
    Public Function IsCommandChecked(command As String) As Boolean
        Return False
    End Function

    ''' 
    ''' Callback called by UI manager.
    ''' Returns a label associated with given command.
    ''' We assume for simplicity taht command labels are named simply named as [commandId]_Label (see resources)
    ''' 
    Public Function GetCommandLabel(command As String) As String
        Return My.Resources.ResourceManager.GetString(command & "_Label")
    End Function

    ''' 
    ''' Returns a icon associated with given command.
    ''' We assume for simplicity that icon ids are named after command commandId.
    ''' 
    Public Function GetCommandBitmap(command As String) As Bitmap
        Return DirectCast(My.Resources.ResourceManager.GetObject(command), Bitmap)
    End Function


    Sub UpdateUI()
        AddinUI.UpdateRibbon()

    End Sub

    Public Sub Application_SelectionChanged(window As Visio.Window)
        UpdateUI()
    End Sub

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        AddHandler Application.SelectionChanged, AddressOf Application_SelectionChanged

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        RemoveHandler Application.SelectionChanged, AddressOf Application_SelectionChanged

    End Sub

End Class
