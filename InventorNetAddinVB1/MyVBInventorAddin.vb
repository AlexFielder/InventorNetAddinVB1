#Region "Namespaces"

Imports System.Text
Imports System.Linq
Imports System.Xml
Imports System.Reflection
Imports System.ComponentModel
Imports System.Collections
Imports System.Collections.Generic
Imports System.Windows
Imports System.Windows.Media.Imaging
Imports System.Windows.Forms
Imports System.Drawing
Imports System.IO
Imports System.Runtime.InteropServices
Imports Microsoft.Win32

Imports Inventor

#End Region

'Namespace InventorNetAddinVB1
<GuidAttribute("c9424b68-1aa4-49fb-a4f2-2fa6c267ce35")>
Public Class MyVBInventorAddin
    Implements Inventor.ApplicationAddInServer
    Public Sub New()
    End Sub

#Region "Implementations of the Interface Members"

    ''' <summary>
    ''' Do initializations in it such as caching the application object, registering event handlers, and adding ribbon buttons.
    ''' </summary>
    ''' <param name="siteObj">The entry object for the addin.</param>
    ''' <param name="loaded1stTime">Indicating whether the addin is loaded for the 1st time.</param>
    Public Sub Activate(siteObj As Inventor.ApplicationAddInSite, loaded1stTime As Boolean) Implements ApplicationAddInServer.Activate
        AddinGlobal.InventorApp = siteObj.Application

        Try
            AddinGlobal.GetAddinClassId(Me.GetType())

            Dim icon1 As New Icon(Assembly.GetExecutingAssembly().GetManifestResourceStream("InventorNetAddinVB1.addin.ico"))
            'Change it if necessary but make sure it's embedded.
            Dim button1 As New InventorButton("Button 1", "MyVBInventorAddin.Button_" & Guid.NewGuid().ToString(), "Button 1 description", "Button 1 tooltip", icon1, icon1,
                CommandTypesEnum.kShapeEditCmdType, ButtonDisplayEnum.kDisplayTextInLearningMode)
            button1.SetBehavior(True, True, True)
            button1.Execute = AddressOf ButtonActions.Button1_Execute

            If loaded1stTime = True Then
                Dim uiMan As UserInterfaceManager = AddinGlobal.InventorApp.UserInterfaceManager
                If uiMan.InterfaceStyle = InterfaceStyleEnum.kRibbonInterface Then
                    'kClassicInterface support can be added if necessary.
                    Dim ribbon As Inventor.Ribbon = uiMan.Ribbons("Part")
                    Dim tab As RibbonTab
                    Try
                        tab = ribbon.RibbonTabs("id_TabSheetMetal") 'Change it if necessary.
                    Catch
                        tab = ribbon.RibbonTabs.Add("id_TabSheetMetal", "id_Tabid_TabSheetMetal", Guid.NewGuid().ToString())
                    End Try
                    AddinGlobal.RibbonPanelId = "{51f8ccf4-5fc6-4592-b68d-e19c993f5faa}"
                    AddinGlobal.RibbonPanel = tab.RibbonPanels.Add("InventorNetAddin", "MyVBInventorAddin.RibbonPanel_" & Guid.NewGuid().ToString(), AddinGlobal.RibbonPanelId, String.Empty, True)

                    Dim cmdCtrls As CommandControls = AddinGlobal.RibbonPanel.CommandControls
                    cmdCtrls.AddButton(button1.ButtonDef, button1.DisplayBigIcon, button1.DisplayText, "", button1.InsertBeforeTarget)
                End If
            End If
        Catch e As Exception
            MessageBox.Show(e.ToString())
        End Try

        ' TODO: Add more initialization code below.

    End Sub

    ''' <summary>
    ''' Do cleanups in it such as releasing COM objects or forcing the GC to Collect when necessary.
    ''' </summary>
    Public Sub Deactivate() Implements ApplicationAddInServer.Deactivate
        ' Add more cleanup work below if necessary, e.g. remove even handlers, flush and close log files, etc.


        ' Release COM objects
        '

        ' Add more FinalReleaseComObject() calls for COM objects you know, e.g.
        'if (comObj != null) Marshal.FinalReleaseComObject(comObj);

        ' Release the COM objects maintained by InventorNetAddinWizard.
        For Each button As InventorButton In AddinGlobal.ButtonList
            If button.ButtonDef IsNot Nothing Then
                Marshal.FinalReleaseComObject(button.ButtonDef)
            End If
        Next
        If AddinGlobal.RibbonPanel IsNot Nothing Then
            Marshal.FinalReleaseComObject(AddinGlobal.RibbonPanel)
        End If
        If AddinGlobal.InventorApp IsNot Nothing Then
            Marshal.FinalReleaseComObject(AddinGlobal.InventorApp)
        End If
    End Sub

    ''' <summary>
    ''' Deprecated. Use the ControlDefinition instead to execute commands.
    ''' </summary>
    ''' <param name="commandID"></param>
    Public Sub ExecuteCommand(commandID As Integer) Implements ApplicationAddInServer.ExecuteCommand
    End Sub

    ''' <summary>
    ''' Implement it if wanting to expose your own automation interface. 
    ''' </summary>
    Public ReadOnly Property Automation() As Object Implements ApplicationAddInServer.Automation
        Get
            Return Nothing
        End Get
    End Property

#End Region


End Class
'End Namespace
