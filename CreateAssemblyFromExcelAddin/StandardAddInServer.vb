Imports Inventor
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Microsoft.Win32
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Text.RegularExpressions
Imports System.Drawing
Imports Microsoft.Office.Interop
Imports VDF = Autodesk.DataManagement.Client.Framework
Imports Microsoft.Web.Services3
Imports Autodesk.Connectivity.WebServices
Imports ADSK = Autodesk.Connectivity.WebServices
Imports Autodesk.Connectivity.WebServicesTools
Imports Autodesk.Connectivity.Explorer.ExtensibilityTools

Namespace CreateAssemblyFromExcelAddin
    <ProgIdAttribute("CreateAssemblyFromExcelAddin.StandardAddInServer"), _
    GuidAttribute("38d3b483-2797-4307-92ae-8456410d6801")> _
    Public Class StandardAddInServer
        Implements Inventor.ApplicationAddInServer

        ' Inventor application object.
        Private m_inventorApplication As Inventor.Application
        Public StartFolder As String
        Public ProjectCode As String
#Region "ApplicationAddInServer Members"

        Public Sub Activate(ByVal addInSiteObject As Inventor.ApplicationAddInSite, ByVal firstTime As Boolean) Implements Inventor.ApplicationAddInServer.Activate

            ' This method is called by Inventor when it loads the AddIn.
            ' The AddInSiteObject provides access to the Inventor Application object.
            ' The FirstTime flag indicates if the AddIn is loaded for the first time.

            ' Initialize AddIn members.
            m_inventorApplication = addInSiteObject.Application

            ' TODO:  Add ApplicationAddInServer.Activate implementation.
            ' e.g. event initialization, command creation etc.
            Dim conDefs As Inventor.ControlDefinitions = _
                m_inventorApplication. _
                CommandManager.ControlDefinitions

            ' our custom command ID
            Dim idCommand1 As String = "ID_COMAND_2"

            Try
                ' try get the existing command definition
                btnDef = conDefs.Item(idCommand1)
            Catch ex As Exception
                ' or create it
                btnDef = conDefs.AddButtonDefinition( _
                    "Create Assembly From Excel", idCommand1, _
                    CommandTypesEnum.kEditMaskCmdType, _
                    Guid.NewGuid().ToString(), _
                    "Creates the Assembly structure using Excel for data input and Vaulted files where possible!", _
                    "Use this at the start of the project to build the parent assembly structure", _
                    GetICOResource("CreateAssemblyFromExcelAddin.Icons8-Windows-8-Food-Cafe.ico"), _
                    GetICOResource("CreateAssemblyFromExcelAddin.Icons8-Windows-8-Food-Cafe.ico"))
            End Try

            If (firstTime) Then
                If (m_inventorApplication.UserInterfaceManager.
                    InterfaceStyle =
                    InterfaceStyleEnum.kRibbonInterface) Then

                    '1. access the Assembly ribbon
                    Dim ribbonAssembly As Inventor.Ribbon = m_inventorApplication.UserInterfaceManager.Ribbons.Item("Assembly")
                    Dim tabAssemblyFeatureCount As Inventor.RibbonTab = ribbonAssembly.RibbonTabs.Add("CreateAssemblyFromExcel", "ASSY_TAB_CAFE", Guid.NewGuid().ToString())
                    Dim MyAssemblyCommandsPanel As Inventor.RibbonPanel = tabAssemblyFeatureCount.RibbonPanels.Add("Create Assembly From Excel", "ASSY_PNL_RUN_CAFE", Guid.NewGuid().ToString())
                    MyAssemblyCommandsPanel.CommandControls.AddButton(btnDef, True)

                End If
            End If

            ' register the method that will be executed
            AddHandler btnDef.OnExecute, AddressOf Command1Method
        End Sub

        Public Sub Deactivate() Implements Inventor.ApplicationAddInServer.Deactivate

            ' This method is called by Inventor when the AddIn is unloaded.
            ' The AddIn will be unloaded either manually by the user or
            ' when the Inventor session is terminated.

            ' TODO:  Add ApplicationAddInServer.Deactivate implementation

            ' Release objects.
            m_inventorApplication = Nothing

            System.GC.Collect()
            System.GC.WaitForPendingFinalizers()
        End Sub

        Public ReadOnly Property Automation() As Object Implements Inventor.ApplicationAddInServer.Automation

            ' This property is provided to allow the AddIn to expose an API 
            ' of its own to other programs. Typically, this  would be done by
            ' implementing the AddIn's API interface in a class and returning 
            ' that class object through this property.

            Get
                Return Nothing
            End Get

        End Property

        Public Sub ExecuteCommand(ByVal commandID As Integer) Implements Inventor.ApplicationAddInServer.ExecuteCommand

            ' Note:this method is now obsolete, you should use the 
            ' ControlDefinition functionality for implementing commands.

        End Sub

#End Region



        Private btnDef As Inventor.ButtonDefinition
        Private Sub Command1Method()
            If TypeOf m_inventorApplication.ActiveDocument Is PartDocument Then
                MessageBox.Show("Needs to be run from the highest-level Assembly file!")
            Else
                RunCAFE(m_inventorApplication.ActiveDocument)
            End If
        End Sub

        ''' <summary>
        ''' Returns the relevant resource from the compiled .dll file
        ''' Means you don't need to Copy local .ico (or any other resource files)!
        ''' </summary>
        ''' <param name="icoResourceName"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetICOResource( _
                  ByVal icoResourceName As String) As Object
            Dim assemblyNet As System.Reflection.Assembly = _
              System.Reflection.Assembly.GetExecutingAssembly()
            Dim stream As System.IO.Stream = _
              assemblyNet.GetManifestResourceStream(icoResourceName)
            Dim ico As System.Drawing.Icon = _
              New System.Drawing.Icon(stream)
            Return PictureDispConverter.ToIPictureDisp(ico)
        End Function

        Private Sub RunCAFE(Document As _Document)
            Dim FilesArray As New ArrayList
            Dim PartsList As List(Of SubObjectCls)
            PartsList = New List(Of SubObjectCls)
            'pass the local variables to our external .dll
            'XTVB.InventorApplication = ThisApplication
            Dim ProjectCode As String = InputBox("Which project?", "4 Letter Project Code", "CODE")
            Dim filetab As String = ProjectCode + "-MODELLING-BASELINE"
            Dim oXL As Excel.Application
            Dim oWB As Excel.Workbook
            Dim oSheet As Excel.Worksheet
            'Dim oRng As Excel.Range
            
            ' Start Excel and get Application object.
            oXL = CreateObject("Excel.Application")
            oXL.Visible = True

            oWB = oXL.Workbooks.Open("C:\LEGACY VAULT WORKING FOLDER\Designs\Project Tracker.xlsx")
            oSheet = oXL.Sheets(filetab)
            oSheet.Activate()

            'FilesArray = GoExcel.CellValues("C:\LEGACY VAULT WORKING FOLDER\Designs\Project Tracker.xlsx", filetab, "A3", "A4") ' sets excel to the correct sheet!
            For MyRow As Integer = 3 To 5000 ' max limit = 50 rows for debugging purposes
                Dim SO As SubObjectCls
                'not sure if we should change this to Column C as it contains the files we know about from the Vault
                'if we did we could then have it insert that file if we linked this routine to Vault...?
                If oSheet.Cells(MyRow, 2).Value = "" Then Exit For
                'If GoExcel.CellValue("B" & MyRow) = "" Then Exit For 'exits when the value is empty!
                '    Dim tmpstr As String = GoExcel.CellValue("I" & MyRow) 'parent row
                '    If Not tmpstr.StartsWith("AS-") Then
                '        Continue For
                '    End If
                'some error checking since we don't always have parent assembly information in Excel:
                Dim PartNo As String = oSheet.Cells(MyRow, 2).Value
                Dim Descr As String = oSheet.Cells(MyRow, 11).Value
                Dim RevNumber As String = oSheet.Cells(MyRow, 12).Value
                Dim LegacyDrawingNumber As String = oSheet.Cells(MyRow, 13).Value
                Dim ParentAssembly As String = oSheet.Cells(MyRow, 9).Value
                If ParentAssembly = "" Then
                    ParentAssembly = "NA"
                End If
                SO = New SubObjectCls(PartNo, Descr, RevNumber, LegacyDrawingNumber, ParentAssembly)
                PartsList.Add(SO)
            Next
            'MessageBox.Show(PartsList.Count)
            'Call XTVB.PopulatePartsList(PartsList)
            StartFolder = System.IO.Path.GetDirectoryName(System.IO.Path.GetDirectoryName(m_inventorApplication.ActiveDocument.FullDocumentName))
            ProjectCode = ProjectCode
            PartsList = PartsList
            oWB.Close()
            oSheet = Nothing
            oWB = Nothing
            oXL.Quit()
            oXL = Nothing
            'XTVB.FilesArray = FilesArray
            'XTVB.GoExcel = GoExcel.Application
            Dim tr As Transaction
            tr = m_inventorApplication.TransactionManager.StartTransaction(m_inventorApplication.ActiveDocument, "Create Standard Parts From Excel")
            BeginCreateAssemblyStructure()
            tr.End()
        End Sub

        Public PartsList As List(Of SubObjectCls)
        Public CompleteList As List(Of SubObjectCls) = New List(Of SubObjectCls)
        Public r As List(Of SubObjectCls)
        Public ParentList As List(Of String)
        Public parentAssemblyFilename As String
        Public highestlevel As Long = 0
        Public foundfile As FileInfo
        Private Property files As IEnumerable(Of FileInfo)

        'Public Sub PopulatePartsList(ByVal iLogicPartsList As List(Of SubObjectCls))
        '    PartsList = iLogicPartsList
        'End Sub

        ''' <summary>
        ''' Begins our Create Assembly subroutine
        ''' </summary>
        ''' <remarks>Uses the PartsList object passed to it from Inventor (which was ulitmately created from Excel data.</remarks>
        Public Sub BeginCreateAssemblyStructure()
            'define the parent assembly
            If StartFolder = String.Empty Then
                StartFolder = System.Environment.SpecialFolder.MyComputer
            End If
            Dim DirStruct As DirectoryInfo = New DirectoryInfo(StartFolder)
            Dim dlg1 = New FolderBrowserDialog
            dlg1.Description = "Select a folder to work with:"
            dlg1.ShowNewFolderButton = True
            dlg1.SelectedPath = StartFolder
            'dlg1.RootFolder = StartFolder
            Dim result As DialogResult = dlg1.ShowDialog()
            If result = DialogResult.OK Then
                DirStruct = New DirectoryInfo(dlg1.SelectedPath)
            Else
                MessageBox.Show("You need to pick a folder for this to work")
                Exit Sub
            End If
            Dim asmDoc As AssemblyDocument
            Try
                asmDoc = m_inventorApplication.ActiveDocument
            Catch ex As Exception
                MessageBox.Show("You need to have an assembly open, not a part!")
            End Try
            parentAssemblyFilename = System.IO.Path.GetFileNameWithoutExtension(m_inventorApplication.ActiveDocument.DisplayName)
            Dim level As Long = 0
            ParseFolders(DirStruct, level)
            Dim res = From a As SubObjectCls In CompleteList
                      Where (a.Level > 1)
                      Order By a.Level
                      Group By groupKey = level
                      Into groupName = Group

            Dim grouped = CompleteList.OrderBy(Function(x) x.Level).GroupBy(Function(x) x.Level)

            For Each group In grouped
                If Not group.Key = 1 Then 'skip the first level as it's our top level assembly!
                    'MessageBox.Show("Level = " & group.Key & " of " & grouped.Count)
                    'If group.Key <= 3 Then
                    For Each subobj As SubObjectCls In group
                        CreateAssemblyStructure(subobj, subobj.ParentAssembly)
                    Next
                    'End If
                End If
            Next
            MessageBox.Show(CompleteList.Count)
        End Sub

        ''' <summary>
        ''' Creates dummy files to pre-populate Vault with or creates files named "Replace With " to denote where files already exist.
        ''' </summary>
        ''' <param name="subObject">The Subobject we need to create/replace</param>
        ''' <returns>Returns the filename of the newly created dummy Part/Assemly</returns>
        ''' <remarks></remarks>
        Private Function CreateAssemblyComponents(subObject As SubObjectCls) As String
            Dim basepartname As String = String.Empty
            Dim newfilename As String = String.Empty
            Try
                If subObject.PartNo.StartsWith("AS-") Then
                    newfilename = System.IO.Path.GetDirectoryName(m_inventorApplication.ActiveDocument.FullDocumentName) & "\" & subObject.PartNo & ".iam"
                    basepartname = "C:\LEGACY VAULT WORKING FOLDER\Designs\DT-99999-000.iam"
                ElseIf subObject.PartNo.StartsWith("DT-") Then
                    If subObject.LegacyDescr.Contains("ASSEMBLY") Or subObject.LegacyDescr.Contains("ASSY") Then
                        newfilename = System.IO.Path.GetDirectoryName(m_inventorApplication.ActiveDocument.FullDocumentName) & "\" & subObject.PartNo & ".iam"
                        basepartname = "C:\LEGACY VAULT WORKING FOLDER\Designs\DT-99999-000.iam"
                    Else
                        newfilename = System.IO.Path.GetDirectoryName(m_inventorApplication.ActiveDocument.FullDocumentName) & "\" & subObject.PartNo & ".ipt"
                        basepartname = "C:\LEGACY VAULT WORKING FOLDER\Designs\DT-99999-000.ipt"
                    End If
                ElseIf subObject.PartNo.StartsWith("DL-") Then
                    'technically this assembly is missing from the structure!
                    newfilename = System.IO.Path.GetDirectoryName(m_inventorApplication.ActiveDocument.FullDocumentName) &
                        "\" & GetFriendlyDirName(subObject.PartNo) & ".iam"
                    basepartname = "C:\LEGACY VAULT WORKING FOLDER\Designs\DT-99999-000.iam"
                End If
                'check if the file exists locally and copy a template to create it if not.
                If Not System.IO.File.Exists(newfilename) Then 'we need to create it - but we also might need to search the local working folder for it too...?
                    'MessageBox.Show("Looking for: " + newfilename, "Finding Files!")
                    Dim expectedfileextension As String = System.IO.Path.GetExtension(newfilename)
                    Dim tmpstr As String = FindFileInVWF(newfilename)
                    If tmpstr = String.Empty Then
                        'it doesn't exist anywhere else in the Local Vault Working Folder
                        System.IO.File.Copy(basepartname, newfilename)
                    Else
                        'it does exist and we (Currently) are creating a placeholder file to replace later, although this creates its own issues!
                        newfilename = foundfile.FullName
                        'newfilename = System.IO.Path.GetDirectoryName(m_inventorApplication.ActiveDocument.FullDocumentName) & "\Replace with " & foundfile.Name
                        If Not System.IO.Path.GetExtension(newfilename) = expectedfileextension Then
                            'correct the extension or Inventor will shit the bed
                            newfilename = Left(newfilename, newfilename.Length - 4) & expectedfileextension
                        End If
                        If Not System.IO.File.Exists(newfilename) Then
                            System.IO.File.Copy(basepartname, newfilename)
                        End If
                        'need to empty the foundfile object so it can be reused/found on the next file.
                        foundfile = Nothing
                    End If
                End If
            Catch ex As Exception
                MessageBox.Show("Exception was: " + ex.Message + vbCrLf + ex.StackTrace)
            End Try
            Return newfilename
        End Function

        ''' <summary>
        ''' Creates a new occurrence of our subObject if the "subObject.name:1" doesn't already exist
        ''' </summary>
        ''' <param name="subObject">The SubObject to Create (it it doesn't already exist!)</param>
        ''' <param name="parentName">The parent of subObject</param>
        ''' <remarks></remarks>
        Private Sub CreateAssemblyStructure(subObject As SubObjectCls, parentName As String)
            Dim asmDoc As AssemblyDocument = Nothing
            Dim realOcc As ComponentOccurrence = Nothing
            Dim realOccStr As String = String.Empty
            Dim PosnMatrix As Matrix
            Try
                Dim newfilename As String = CreateAssemblyComponents(subObject)
                If Not System.IO.Path.GetDirectoryName(newfilename).Contains(ProjectCode) Then
                    'file is outside the working folder for this assembly

                Else
                    'file is new.
                    Dim i As Integer = PartsList.FindIndex(Function(str As SubObjectCls) str.PartNo = subObject.PartNo)
                    If Not i = -1 Then
                        AligniPropertyValues(subObject, PartsList(i))
                    End If
                End If

                PosnMatrix = m_inventorApplication.TransientGeometry.CreateMatrix
                If parentName = System.IO.Path.GetFileNameWithoutExtension(m_inventorApplication.ActiveDocument.DisplayName) Then
                    'immediate descendants of the parent assembly
                    asmDoc = m_inventorApplication.ActiveDocument
                    Try
                        realOcc = asmDoc.ComponentDefinition.Occurrences.Add(newfilename, PosnMatrix)
                        realOccStr = realOcc.Name
                        If Not realOccStr.StartsWith("Replace With", StringComparison.OrdinalIgnoreCase) And _
                            System.IO.Path.GetDirectoryName(newfilename).Contains(ProjectCode) Then 'assign iproperties to new parts
                            AssignIProperties(realOcc, subObject)
                        End If
                    Catch ex As Exception
                        MessageBox.Show("Exception was: " + ex.Message + vbCrLf + ex.StackTrace)
                    End Try
                Else
                    'one of its grandchildren
                    Dim tmpdoc As Inventor.Document = Nothing
                    For Each doc As Inventor.Document In m_inventorApplication.ActiveDocument.AllReferencedDocuments
                        'at this point we should stop the insert of occurrences if the assembly is called "Replace With XXX" as it's doing work that's unecessary!
                        If doc.DisplayName = parentName & ".iam" Or doc.DisplayName.StartsWith("Replace With " & parentName, StringComparison.OrdinalIgnoreCase) Then
                            tmpdoc = doc
                            Exit For
                        End If
                    Next

                    Try
                        asmDoc = tmpdoc
                        For Each a As ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences
                            If a.Name = System.IO.Path.GetFileNameWithoutExtension(newfilename) & ":1" Then
                                realOcc = a
                            End If
                        Next

                        If realOcc Is Nothing Then 'only insert the occurrence once or we end up with a huge assembly containing multiple occurrences...
                            realOcc = asmDoc.ComponentDefinition.Occurrences.Add(newfilename, PosnMatrix)
                            realOccStr = realOcc.Name
                            If Not realOccStr.StartsWith("Replace With", StringComparison.OrdinalIgnoreCase) And _
                                System.IO.Path.GetDirectoryName(newfilename).Contains(ProjectCode) Then 'assign iproperties to new parts
                                AssignIProperties(realOcc, subObject)
                            End If
                        End If
                    Catch ex As Exception
                        MessageBox.Show("Exception was: " + ex.Message + vbCrLf + ex.StackTrace)
                    End Try
                End If
            Catch ex As Exception
                MessageBox.Show("Exception was: " + ex.Message + vbCrLf + ex.StackTrace)
            End Try
        End Sub

        ''' <summary>
        ''' Copies iProperty values to the subObjectCls for this file
        ''' </summary>
        ''' <param name="subObject">the subObjectCls to update</param>
        ''' <param name="PartsListObj">the subObjectCls containing the values from Excel</param>
        ''' <remarks></remarks>
        Private Sub AligniPropertyValues(ByRef subObject As SubObjectCls, ByVal PartsListObj As SubObjectCls)
            subObject.LegacyDescr = PartsListObj.LegacyDescr
            subObject.LegacyDrawingNo = PartsListObj.LegacyDrawingNo
            subObject.LegacyRev = PartsListObj.LegacyRev
        End Sub

        ''' <summary>
        ''' Assigns IProperties using the ComponentOccurrence and SubObject objects
        ''' </summary>
        ''' <param name="realocc">the occurrence we wish to edit</param>
        ''' <param name="subObject">the subobject containing our aligned iProperties (taken from Excel)</param>
        ''' <remarks></remarks>
        Private Sub AssignIProperties(ByVal realocc As ComponentOccurrence, ByVal subObject As SubObjectCls)
            Try
                Dim invProjProperties As PropertySet = realocc.Definition.Document.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}")
                Dim invSummaryiProperties As PropertySet = realocc.Definition.Document.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}")
                'Project iProperties
                invProjProperties.ItemByPropId(PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties).Value = subObject.PartNo 'part number
                invProjProperties.ItemByPropId(PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties).Value = subObject.LegacyDescr 'description
                invProjProperties.ItemByPropId(PropertiesForDesignTrackingPropertiesEnum.kProjectDesignTrackingProperties).Value = "A90.1" 'project
                'Summary iProperties
                invSummaryiProperties.ItemByPropId(PropertiesForSummaryInformationEnum.kRevisionSummaryInformation).Value = subObject.LegacyRev 'revision
                invSummaryiProperties.ItemByPropId(PropertiesForSummaryInformationEnum.kSubjectSummaryInformation).Value = subObject.LegacyDrawingNo 'subject
                invSummaryiProperties.ItemByPropId(PropertiesForSummaryInformationEnum.kTitleSummaryInformation).Value = subObject.LegacyDescr 'title
                invSummaryiProperties.ItemByPropId(PropertiesForSummaryInformationEnum.kCommentsSummaryInformation).Value = "MODELLED FROM DRAWINGS"
            Catch ex As Exception
                MessageBox.Show("Exception was: " + ex.Message + vbCrLf + ex.StackTrace)
            End Try
        End Sub

        ''' <summary>
        ''' Returns the name of a file found in the local Vault Working Folder
        ''' </summary>
        ''' <param name="newfilename">the @@-######-000 Filename to look for.</param>
        ''' <returns>Returns the filename we were searching for if it already exists</returns>
        ''' <remarks></remarks>
        Private Function FindFileInVWF(newfilename As String) As String
            Dim dir = New DirectoryInfo("C:\Legacy Vault Working Folder\Designs")
            Dim tmpstr = GetExistingFile(dir, System.IO.Path.GetFileNameWithoutExtension(newfilename))
            If tmpstr = "" And foundfile Is Nothing Then
                Return ""
            Else
                Return tmpstr
            End If
        End Function

        ''' <summary>
        ''' Searches for an existing file using the dir.EnumerateFiles() Method
        ''' </summary>
        ''' <param name="dir">The Top Level directory to search</param>
        ''' <param name="newfilename">the file to look for</param>
        ''' <returns>Returns foundfilename if it finds a match</returns>
        ''' <remarks>Also creates the files collection if it doesn't exist already</remarks>
        Private Function GetExistingFile(ByVal dir As DirectoryInfo, ByVal newfilename As String) As String
            Dim foundfilename As String = String.Empty
            Try
                'this way should create a large list of .ipt/iam files in one pass that we can keep for later reuse. It will likely be slow initially but faster later.
                If files Is Nothing Then
                    files = dir.EnumerateFiles("*.*", SearchOption.AllDirectories).Where(Function(s As FileInfo) _
                                                                                             s.Name.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) OrElse _
                                                                                             s.Name.EndsWith(".iam", StringComparison.OrdinalIgnoreCase))
                End If
                Dim file As FileInfo = (From f As FileInfo In files
                                       Where System.IO.Path.GetFileNameWithoutExtension(f.Name) = newfilename
                                       Select f).FirstOrDefault()
                If Not file Is Nothing Then
                    foundfile = file
                    foundfilename = file.Name
                End If

                'For Each file As FileInfo In files
                '    If System.IO.Path.GetFileNameWithoutExtension(file.Name) = newfilename Then
                '        foundfilename = file.Name
                '        foundfile = file 'set this in case we can't return foundfilename
                '        Return foundfilename
                '        Exit For
                '    End If
                'Next
            Catch ex As Exception
                MessageBox.Show("Exception was: " + ex.Message + vbCrLf + ex.StackTrace)
            End Try
            Return foundfilename
        End Function

#Region "Folder Structure Parser from console app"
        ''' <summary>
        ''' Parses a series of subfolders, ignoring directories called "Superseded", documents whose names contain "IL" or "DL"
        ''' </summary>
        ''' <param name="dir">The Directory to search</param>
        ''' <param name="level">the structure level we are currently at</param>
        ''' <remarks></remarks>
        Public Sub ParseFolders(ByVal dir As DirectoryInfo, ByVal level As Long)
            Try
                If Not dir.Name.Contains("Superseded") Then
                    Dim friendlydirname As String = GetFriendlyDirName(dir.Name)
                    Dim friendlyparentdirname As String = String.Empty
                    If Not dir.Parent Is Nothing Then
                        friendlyparentdirname = GetFriendlyDirName(dir.Parent.Name)
                    Else
                        friendlyparentdirname = parentAssemblyFilename
                    End If
                    'need to account for instances where the "AS-" file isn't the first we'll find
                    Dim thisAssy As FileInfo = (From a As FileInfo In dir.GetFiles()
                                               Where GetFriendlyName(a.Name) = friendlydirname And _
                                               Not a.Name.Contains("IL") And _
                                               Not a.Name.Contains("DL") And _
                                               Not a.Name.Contains("SP") And _
                                               (getsheetnum(a.Name) <= 1)
                                               Select a).FirstOrDefault
                    If Not thisAssy Is Nothing Then
                        CompleteList.Add(New SubObjectCls(m_partno:=GetFriendlyName(thisAssy.Name),
                                                m_legacydescr:="",
                                                m_legacyrev:="",
                                                m_legacydrawingno:="",
                                                m_parentassy:=friendlyparentdirname,
                                                m_level:=level + 1))
                    Else
                        'the assembly drawing probably exists somewhere else in the structure so for now we'll assume it's just "missing"
                        thisAssy = New FileInfo(dir.FullName & "\" & GetFriendlyDirName(dir.Name) & ".txt")
                        CompleteList.Add(New SubObjectCls(m_partno:=GetFriendlyName(thisAssy.Name),
                                                m_legacydescr:="",
                                                m_legacyrev:="",
                                                m_legacydrawingno:="",
                                                m_parentassy:=friendlyparentdirname,
                                                m_level:=level + 1))
                    End If
                    For Each file As FileInfo In dir.GetFiles()
                        If Not file.Name.Contains("IL") And Not file.Name.Contains("DL") And Not file.Name.Contains("SP") And Not file.Name = thisAssy.Name Then
                            'if the directory name is the same as the assembly name then the parentassembly is the folder above!
                            Dim friendlyfilename As String = GetFriendlyName(file.Name)
                            If friendlydirname = friendlyfilename Then 'parent assembly in this folder
                                If getsheetnum(file.Name) <= 1 Then
                                    CompleteList.Add(New SubObjectCls(m_partno:=friendlyfilename,
                                                              m_legacydescr:="",
                                                              m_legacyrev:="",
                                                              m_legacydrawingno:="",
                                                              m_parentassy:=friendlyparentdirname,
                                                              m_level:=level + 1))
                                End If
                            Else
                                If getsheetnum(file.Name) <= 1 Then
                                    CompleteList.Add(New SubObjectCls(m_partno:=friendlyfilename,
                                                              m_legacydescr:="",
                                                              m_legacyrev:="",
                                                              m_legacydrawingno:="",
                                                              m_parentassy:=friendlydirname,
                                                              m_level:=level + 2))
                                End If
                            End If
                        End If
                    Next
                    For Each subDir As DirectoryInfo In dir.GetDirectories()
                        If Not subDir.Name.Contains("Superseded") Then
                            ParseFolders(subDir, level + 1)
                        End If
                    Next
                End If
                highestlevel += 1
            Catch ex As Exception
                MessageBox.Show("Exception was: " + ex.Message + vbCrLf + ex.StackTrace)
            End Try
            'Return info
        End Sub

        ''' <summary>
        ''' Returns a "Friendly" filename for comparison-sake
        ''' </summary>
        ''' <param name="p">the String to match against</param>
        ''' <returns>Returns the matched String</returns>
        ''' <remarks></remarks>
        Public Function GetFriendlyName(p As String) As Object
            Dim f As String = String.Empty
            Dim r As New Regex("\w{2}-\d{5,}|\w{2}-\w\d{5,}")
            f = r.Match(p).Captures(0).ToString() + "-000"
            Console.WriteLine(f)
            Return f
        End Function

        ''' <summary>
        ''' Returns a "Friendly" directory name for comparison-sake and use as the "parentname"
        ''' </summary>
        ''' <param name="p1">the String to match against</param>
        ''' <returns>Returns the matched String</returns>
        ''' <remarks></remarks>
        Public Function GetFriendlyDirName(p1 As String) As Object
            If Not p1.Contains(":") Then
                Dim f As String = String.Empty
                Dim r As New Regex("\d{3,}|\w\d{3,}")
                f = "AS-" + r.Match(p1).Captures(0).ToString() + "-000"
                Return f
            Else
                'Return p1
                Return parentAssemblyFilename
            End If
        End Function

        ''' <summary>
        ''' Returns the Directory code without any spurious extras
        ''' </summary>
        ''' <param name="p1">the String to match against</param>
        ''' <returns>Returns the matched String</returns>
        ''' <remarks></remarks>
        Public Function GetDirName(p1 As String) As Object
            If Not p1.Contains(":") Then
                Dim f As String = String.Empty
                Dim r As New Regex("\d{3,}|\w\d{3,}") 'get just the @##### or ###### drawing number
                f = r.Match(p1).Captures(0).ToString()
                Return f
            Else
                Return p1
            End If
        End Function

        ''' <summary>
        ''' Returns the Sheetnum for each file we've found
        ''' </summary>
        ''' <param name="p1">the String to match against</param>
        ''' <returns>Returns the matched String</returns>
        ''' <remarks></remarks>
        Private Function getsheetnum(p1 As String) As Integer
            Dim f As String = String.Empty
            Dim pattern As String = "(.*)(sht-)(\d{3})(.*)"
            Dim matches As MatchCollection = Regex.Matches(p1, pattern)
            For Each m As Match In matches
                Dim g As System.Text.RegularExpressions.Group = m.Groups(3)
                f = CInt(g.Value)
            Next
            Return CInt(f)
        End Function
#End Region

    End Class
#Region "Sub Object Class"
    ''' <summary>
    ''' Our SubObject Class
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SubObjectCls
        Implements IComparable(Of SubObjectCls)
        Public PartNo As String
        Public LegacyDescr As String
        Public LegacyRev As String
        Public LegacyDrawingNo As String
        Public ParentAssembly As String
        Public HasChildren As Boolean
        Public Children As List(Of SubObjectCls)
        Public Level As Long

        ''' <summary>
        ''' Creates a new instance
        ''' </summary>
        ''' <param name="m_partno">Part Number (From Filename)</param>
        ''' <param name="m_legacydescr">Legacy Drawing Description</param>
        ''' <param name="m_legacyrev">Legacy Drawing Revision</param>
        ''' <param name="m_legacydrawingno">Legacy Drawing Number</param>
        ''' <param name="m_parentassy">Parent Assembly Name</param>
        ''' <param name="m_haschildren">Optional Boolean for whether this is a parent Assembly</param>
        ''' <param name="m_children">Optional Collection of Children used when HasChildren= True</param>
        ''' <param name="m_level">Part/Assembly level within the Top-Level Assembly Structure</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal m_partno As String,
                       ByVal m_legacydescr As String,
                       ByVal m_legacyrev As String,
                       ByVal m_legacydrawingno As String,
                       ByVal m_parentassy As String,
                             Optional ByVal m_haschildren As Boolean = False,
                             Optional ByVal m_children As List(Of SubObjectCls) = Nothing,
                       Optional ByVal m_level As Long = 0
                       )
            PartNo = m_partno
            LegacyDescr = m_legacydescr
            LegacyRev = m_legacyrev
            LegacyDrawingNo = m_legacydrawingno
            ParentAssembly = m_parentassy
            HasChildren = m_haschildren
            Children = m_children
            Level = m_level
        End Sub

        ''' <summary>
        ''' Allows us to use this Class with LINQ
        ''' </summary>
        ''' <param name="other">the other instance to compare to</param>
        ''' <returns>Returns the comparison requested</returns>
        ''' <remarks></remarks>
        Public Function CompareTo(other As SubObjectCls) As Integer Implements IComparable(Of SubObjectCls).CompareTo
            Return Me.CompareTo(other)
        End Function

        ''' <summary>
        ''' Split the Collection into a list of SubObjectCls
        ''' </summary>
        ''' <param name="source">The Source List to split</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Split(source As List(Of SubObjectCls)) As List(Of List(Of SubObjectCls))
            Return source.[Select](Function(x, i) New With { _
                Key .Index = i, _
                Key .Value = x _
            }).GroupBy(Function(x) x.Index).[Select](Function(x) x.[Select](Function(v) v.Value).ToList()).ToList()
        End Function

        ''' <summary>
        ''' Allows for grouping of the SubObjectCls by whichever variable we choose **Not Implemented**
        ''' </summary>
        ''' <param name="p1"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function GroupBy(p1 As Object) As Object
            Throw New NotImplementedException
        End Function

        ''' <summary>
        ''' Allows for sorting of the SubObjectCls class
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="KeySelector"></param>
        ''' <remarks></remarks>
        Public Sub Sort(Of T As IComparable)(KeySelector As Func(Of SubObjectCls, T))
            Dim lc As subObjectLevelComparer = New subObjectLevelComparer()
            Children.Sort(lc)
        End Sub
    End Class

    ''' <summary>
    ''' Allows us to compare any SubObjectCls class object to another
    ''' </summary>
    ''' <remarks></remarks>
    Public Class subObjectLevelComparer
        Implements IComparer(Of SubObjectCls)

        Public Function Compare(x As SubObjectCls, y As SubObjectCls) As Integer Implements IComparer(Of SubObjectCls).Compare
            Return x.Level.CompareTo(y.Level)
        End Function
    End Class
#End Region

#Region "PictureDispConverter"
    ''' <summary>
    ''' Class that converts our icons into something Inventor can use.
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class PictureDispConverter
        <DllImport("OleAut32.dll",
          EntryPoint:="OleCreatePictureIndirect",
          ExactSpelling:=True, PreserveSig:=False)> _
        Private Shared Function OleCreatePictureIndirect( _
    <MarshalAs(UnmanagedType.AsAny)> picdesc As Object, _
    ByRef iid As Guid, _
    <MarshalAs(UnmanagedType.Bool)> fOwn As Boolean _
    ) As IPictureDisp
        End Function

        Shared iPictureDispGuid As Guid = GetType( _
          IPictureDisp).GUID

        Private NotInheritable Class PICTDESC
            Private Sub New()
            End Sub
            'Picture Types
            Public Const PICTYPE_UNINITIALIZED As Short = -1
            Public Const PICTYPE_NONE As Short = 0
            Public Const PICTYPE_BITMAP As Short = 1
            Public Const PICTYPE_METAFILE As Short = 2
            Public Const PICTYPE_ICON As Short = 3
            Public Const PICTYPE_ENHMETAFILE As Short = 4

            <StructLayout(LayoutKind.Sequential)> _
            Public Class Icon
                Friend cbSizeOfStruct As Integer = Marshal.SizeOf( _
                  GetType(PICTDESC.Icon))
                Friend picType As Integer = PICTDESC.PICTYPE_ICON
                Friend hicon As IntPtr = IntPtr.Zero
                Friend unused1 As Integer
                Friend unused2 As Integer

                Friend Sub New(icon__1 As System.Drawing.Icon)
                    Me.hicon = icon__1.ToBitmap().GetHicon()
                End Sub
            End Class

            <StructLayout(LayoutKind.Sequential)> _
            Public Class Bitmap
                Friend cbSizeOfStruct As Integer = Marshal.SizeOf( _
                  GetType(PICTDESC.Bitmap))
                Friend picType As Integer = PICTDESC.PICTYPE_BITMAP
                Friend hbitmap As IntPtr = IntPtr.Zero
                Friend hpal As IntPtr = IntPtr.Zero
                Friend unused As Integer

                Friend Sub New(bitmap__1 As System.Drawing.Bitmap)
                    Me.hbitmap = bitmap__1.GetHbitmap()
                End Sub
            End Class
        End Class

        Public Shared Function ToIPictureDisp( _
                          icon As System.Drawing.Icon _
                          ) As IPictureDisp
            Dim pictIcon As New PICTDESC.Icon(icon)
            Return OleCreatePictureIndirect(pictIcon, _
                                            iPictureDispGuid, True)
        End Function

        Public Shared Function ToIPictureDisp( _
                          bmp As System.Drawing.Bitmap _
                          ) As IPictureDisp
            Dim pictBmp As New PICTDESC.Bitmap(bmp)
            Return OleCreatePictureIndirect(pictBmp, _
                                            iPictureDispGuid, True)
        End Function
    End Class
#End Region
#Region "Vault Command Class"
    ''' <summary>
    ''' The base class for a Vault Mirror command.
    ''' Each command is self contained.  In other words, nothing is cached between commands,
    ''' the client is re-logged in for each command.
    ''' </summary>
    Public MustInherit Class Command
        'private static long MAX_FILE_SIZE = 45 * 1024 * 1024; // 45 MB 

        Protected m_conn As VDF.Vault.Currency.Connections.Connection
        Protected m_vaultExplorer As IExplorerUtil
        Protected m_vault As String

        Public Display As IStatusDisplay

        Public Sub New(username As String, password As String, server As String, vault As String)
            m_vault = vault
            m_conn = VDF.Vault.Library.ConnectionManager.LogIn(server, vault, username, password, VDF.Vault.Currency.Connections.AuthenticationFlags.Standard, Nothing).Connection
            Display = Nothing
        End Sub

        Protected Sub DeleteFile(filePath As String)
            If Display IsNot Nothing Then
                Display.ChangeStatusMessage("Deleting " & filePath)
            End If

            System.IO.File.SetAttributes(filePath, FileAttributes.Normal)
            System.IO.File.Delete(filePath)

        End Sub

        Protected Sub DeleteFolder(dirPath As String)
            Dim subFolderPaths As String() = Directory.GetDirectories(dirPath)
            If subFolderPaths IsNot Nothing Then
                For Each subFolderPath As String In subFolderPaths
                    DeleteFolder(subFolderPath)
                Next
            End If

            Dim filePaths As String() = Directory.GetFiles(dirPath)
            If filePaths IsNot Nothing Then
                For Each filePath As String In filePaths
                    DeleteFile(filePath)
                Next
            End If

            Directory.Delete(dirPath)
        End Sub


        Protected Sub DownloadFile(file As ADSK.File, filePath As String)
            If Display IsNot Nothing Then
                Display.ChangeStatusMessage("Downloading " & Convert.ToString(file.Name))
            End If

            ' remove the read-only attribute
            If System.IO.File.Exists(filePath) Then
                System.IO.File.SetAttributes(filePath, FileAttributes.Normal)
            End If

            'm_vaultExplorer.DownloadFile(file, filePath);
            Dim settings As New VDF.Vault.Settings.AcquireFilesSettings(m_conn)
            settings.AddFileToAcquire(New VDF.Vault.Currency.Entities.FileIteration(m_conn, file), VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download, New VDF.Currency.FilePathAbsolute(filePath))
            m_conn.FileManager.AcquireFiles(settings)
        End Sub


#Region "Explicit Download"

        ' Downloading CAD files via the webservices is discouraged starting in Vault 2012.
        ' Using the ExtensibilityTools DLL is the preferred method since it will fix up any broken
        ' references.  The downside is that the Vault client must be installed.

        'protected void DownloadFile(ADSK.File file, string filePath)
        '{
        '    if (Display != null)
        '        Display.ChangeStatusMessage("Downloading " + file.Name);

        '    // remove the read-only attribute
        '    if (System.IO.File.Exists(filePath))
        '        System.IO.File.SetAttributes(filePath, FileAttributes.Normal);

        '    if (file.FileSize > MAX_FILE_SIZE)
        '        DownloadFileLarge(file, filePath);
        '    else
        '        DownloadFileStandard(file, filePath);

        '    // set the file to read-only
        '    System.IO.File.SetAttributes(filePath, FileAttributes.ReadOnly);
        '}

        'private void DownloadFileStandard(ADSK.File file, string filePath)
        '{
        '    using (FileStream stream = new FileStream(filePath, FileMode.Create, FileAccess.ReadWrite))
        '    {
        '        byte[] fileData;
        '        m_serviceManager.DocumentService.DownloadFile(file.Id, true, out fileData);

        '        stream.Write(fileData, 0, fileData.Length);
        '        stream.Close();
        '    }
        '}

        'private void DownloadFileLarge(ADSK.File file, string filePath)
        '{
        '    if (System.IO.File.Exists(filePath))
        '        System.IO.File.SetAttributes(filePath, FileAttributes.Normal);

        '    using (FileStream stream = new FileStream(filePath, FileMode.Create, FileAccess.ReadWrite))
        '    {
        '        long startByte = 0;
        '        long endByte = MAX_FILE_SIZE - 1;
        '        byte[] buffer;

        '        while (startByte < file.FileSize)
        '        {
        '            endByte = startByte + MAX_FILE_SIZE;
        '            if (endByte > file.FileSize)
        '                endByte = file.FileSize;

        '            buffer = m_serviceManager.DocumentService.DownloadFilePart(file.Id, startByte, endByte, true);
        '            stream.Write(buffer, 0, buffer.Length);
        '            startByte += buffer.Length;
        '        }
        '        stream.Close();
        '    }
        '}
#End Region

        Public Sub Execute()
            Execute_Impl()
            VDF.Vault.Library.ConnectionManager.LogOut(m_conn)
            m_conn = Nothing
            m_vaultExplorer = Nothing
        End Sub

        Public MustOverride Sub Execute_Impl()

    End Class

    Public Interface IStatusDisplay
        Sub ChangeStatusMessage(message As String)
    End Interface
#End Region
#Region "Vault Full Mirror Command Class"
    ''' <summary>
    ''' Summary description for FullMirrorCommand.
    ''' </summary>
    Public Class FullMirrorCommand
        Inherits Command
        Private m_outputFolder As String

        Public Sub New(username As String, password As String, server As String, vault As String, outputFolder As String)
            MyBase.New(username, password, server, vault)
            m_outputFolder = outputFolder
        End Sub

        Public Overrides Sub Execute_Impl()
            ' cycle through all of the files in the Vault and place them on disk if needed
            Dim root As Folder = m_conn.WebServiceManager.DocumentService.GetFolderRoot()

            ' add the vault name to the local path.  This prevents users from accedently
            ' wiping out their C drive or something like that.
            Dim localPath As String = System.IO.Path.Combine(m_outputFolder, m_vault)
            FullMirrorVaultFolder(root, localPath)

            ' cycle through all of the files on disk and make sure that they are in the Vault
            FullMirrorLocalFolder(root, localPath)

        End Sub

        Private Sub FullMirrorVaultFolder(folder As Folder, localFolder As String)
            If folder.Cloaked Then
                Return
            End If

            If Not Directory.Exists(localFolder) Then
                Directory.CreateDirectory(localFolder)
            End If

            Dim files As ADSK.File() = m_conn.WebServiceManager.DocumentService.GetLatestFilesByFolderId(folder.Id, True)
            If files IsNot Nothing Then
                For Each file As ADSK.File In files
                    If file.Cloaked Then
                        Continue For
                    End If

                    Dim filePath As String = System.IO.Path.Combine(localFolder, file.Name)
                    If System.IO.File.Exists(filePath) Then
                        If file.CreateDate <> System.IO.File.GetCreationTime(filePath) Then
                            DownloadFile(file, filePath)
                        End If
                    Else
                        DownloadFile(file, filePath)
                    End If
                Next
            End If

            Dim subFolders As Folder() = m_conn.WebServiceManager.DocumentService.GetFoldersByParentId(folder.Id, False)
            If subFolders IsNot Nothing Then
                For Each subFolder As Folder In subFolders
                    FullMirrorVaultFolder(subFolder, System.IO.Path.Combine(localFolder, subFolder.Name))
                Next
            End If
        End Sub

        Private Sub FullMirrorLocalFolder(folder As Folder, localFolder As String)
            If folder.Cloaked Then
                Return
            End If

            ' delete any files on disk that are not in the vault
            Dim localFiles As String() = Directory.GetFiles(localFolder)
            Dim vaultFiles As ADSK.File() = m_conn.WebServiceManager.DocumentService.GetLatestFilesByFolderId(folder.Id, True)

            If vaultFiles Is Nothing AndAlso localFiles IsNot Nothing Then
                For Each localFile As String In localFiles
                    DeleteFile(localFile)
                Next
            Else
                For Each localFile As String In localFiles
                    Dim fileFound As Boolean = False
                    Dim filename As String = System.IO.Path.GetFileName(localFile)
                    For Each vaultFile As ADSK.File In vaultFiles
                        If Not vaultFile.Cloaked AndAlso vaultFile.Name = filename Then
                            fileFound = True
                            Exit For
                        End If
                    Next

                    If Not fileFound Then
                        DeleteFile(localFile)
                    End If
                Next
            End If


            ' recurse the subdirectories and delete any folders not in the Vault
            Dim localFullPaths As String() = Directory.GetDirectories(localFolder)
            If localFullPaths IsNot Nothing Then
                For Each localFullPath As String In localFullPaths
                    Dim vaultPath As String = Convert.ToString(folder.FullName) & "/" & System.IO.Path.GetFileName(localFullPath)
                    Dim vaultSubFolder As Folder() = m_conn.WebServiceManager.DocumentService.FindFoldersByPaths(New String() {vaultPath})

                    If vaultSubFolder(0).Id < 0 Then
                        DeleteFolder(localFullPath)
                    Else
                        FullMirrorLocalFolder(vaultSubFolder(0), localFullPath)
                    End If
                Next
            End If
        End Sub

    End Class

#End Region
End Namespace

