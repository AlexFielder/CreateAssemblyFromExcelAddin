Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data
Imports System.IO
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Inventor
Imports Microsoft.Office.Interop

Imports ACW = Autodesk.Connectivity.WebServices
Imports ADSK = Autodesk.Connectivity.WebServices
Imports Autodesk.Connectivity.Explorer.ExtensibilityTools
Imports Autodesk.Connectivity.WebServices
Imports Autodesk.Connectivity.WebServicesTools
Imports Framework = Autodesk.DataManagement.Client.Framework
Imports Vault = Autodesk.DataManagement.Client.Framework.Vault
Imports VDF = Autodesk.DataManagement.Client.Framework

Namespace CreateAssemblyFromExcelAddin
    <ProgIdAttribute("CreateAssemblyFromExcelAddin.StandardAddInServer"), _
    GuidAttribute("38d3b483-2797-4307-92ae-8456410d6801")> _
    Public Class StandardAddInServer
        Implements Inventor.ApplicationAddInServer

        ' Inventor
        Private m_inventorApplication As Inventor.Application
        Private StartFolder As String
        Private ProjectCode As String
        Private PartsListFromExcel As List(Of SubObjectCls) = New List(Of SubObjectCls)
        Private CompleteListFromSystemDrive As List(Of SubObjectCls) = New List(Of SubObjectCls)
        Private EditedListFromSystemDrive As List(Of SubObjectCls) = New List(Of SubObjectCls)
        Public Shared NoMatch As Boolean = True
        Private parentAssemblyFilename As String
        Private highestlevel As Long = 0
        Private foundfile As FileInfo
        Private Property files As IEnumerable(Of FileInfo)
        'Vault
        Public m_conn As Framework.Vault.Currency.Connections.Connection = Nothing
        Private VaultedFileList As List(Of ACW.File) = New List(Of Autodesk.Connectivity.WebServices.File)
        Private FileIterations As List(Of Vault.Currency.Entities.FileIteration) = New List(Of Vault.Currency.Entities.FileIteration)
        Private FolderIdsToFolderEntities As IDictionary(Of Long, Vault.Currency.Entities.Folder)
        Private AssociationArrays As ACW.FileAssocArray() = Nothing
        Private FileAssocLiteArrays As ACW.FileAssocLite() = Nothing
        Private associationsByFile As New Dictionary(Of Long, List(Of Vault.Currency.Entities.FileIteration))()
        Public Shared selectedfile As ListBoxFileItem = Nothing
        Public Shared FoundList As List(Of ListBoxFileItem) = Nothing
#Region "ApplicationAddInServer Members"



        Public Sub Activate(ByVal AddInSiteObject As Inventor.ApplicationAddInSite, ByVal FirstTime As Boolean) Implements Inventor.ApplicationAddInServer.Activate

            ' This method is called by Inventor when it loads the AddIn.
            ' The AddInSiteObject provides access to the Inventor Application object.
            ' The FirstTime flag indicates if the AddIn is loaded for the first time.

            ' Initialize AddIn members.
            m_inventorApplication = AddInSiteObject.Application

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

            If (FirstTime) Then
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

        Public Sub ExecuteCommand(ByVal CommandID As Integer) Implements Inventor.ApplicationAddInServer.ExecuteCommand

            ' Note:this method is now obsolete, you should use the
            ' ControlDefinition functionality for implementing commands.

        End Sub

#End Region
#Region "CAFE"


        Private btnDef As Inventor.ButtonDefinition
        Private Sub Command1Method()
            If TypeOf m_inventorApplication.ActiveDocument Is PartDocument Then
                MessageBox.Show("Needs to be run from the highest-level Assembly file!")
            Else
                m_conn = Vault.Forms.Library.Login(Nothing)
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

        ''' <summary>
        ''' This is the main Method of our tool. It prompts the user to input a project code.
        ''' </summary>
        ''' <param name="Document">the parent assembly document we should be running from</param>
        ''' <remarks></remarks>
        Private Sub RunCAFE(Document As _Document)
            'Dim FilesArray As New ArrayList
            'pass the local variables to our external .dll
            'XTVB.InventorApplication = ThisApplication
            ProjectCode = InputBox("Which project?", "4 Letter Project Code", "CODE")
            If ProjectCode = "" Then
                MessageBox.Show("No project Code inserted, Exiting...")
                Exit Sub
            End If
            Dim filetab As String = ProjectCode + "-MODELLING-BASELINE"
            'start faster excel implementation (FEI)
            Dim XLFile As New FileToProcess()
            XLFile.FileName = "C:\LEGACY VAULT WORKING FOLDER\Designs\Project Tracker.xlsx"
            Dim percent As Double = Nothing

            Dim xlds As DataSet = XLFile.GetExcelData(filetab)
            If xlds Is Nothing Then
                MessageBox.Show("No matching tab could be found for the project code you used." & vbCrLf & "Suggest you try again!")
                Exit Sub
            End If
            Dim ExcelDataTable As DataTable = xlds.Tables(0)

            Dim rowNum As Integer = 3
            For Each row As DataRow In ExcelDataTable.Rows
                Dim SO As SubObjectCls = New SubObjectCls()
                For i = 0 To ExcelDataTable.Columns.Count - 1
                    Dim column As DataColumn = ExcelDataTable.Columns(i)
                    Select Case column.ColumnName
                        Case "DRAWING NUMBER"
                            SO.PartNo = row(column).ToString
                        Case "DRAWING TITLE"
                            If row(column).ToString = "N/A" Then
                                SO.LegacyDescr = "REFER TO PDF"
                            Else
                                SO.LegacyDescr = row(column).ToString
                            End If

                        Case "DRAWING REV"
                            If row(column).ToString = "N/A" Then
                                SO.LegacyRev = "REFER TO PDF"
                            Else
                                SO.LegacyRev = row(column).ToString
                            End If
                        Case "LEGACY DRAWING NUMBER"
                            If row(column).ToString = "N/A" Then
                                SO.LegacyDrawingNo = "REFER TO PDF"
                            Else
                                SO.LegacyDrawingNo = row(column).ToString
                            End If
                        Case "PARENT"
                            SO.ParentAssembly = row(column).ToString
                        Case "VAULTED NAME"
                            SO.FileName = row(column).ToString
                        Case Else
                    End Select
                Next
                PartsListFromExcel.Add(SO)
                rowNum += 1
            Next
            StartFolder = System.IO.Path.GetDirectoryName(System.IO.Path.GetDirectoryName(m_inventorApplication.ActiveDocument.FullDocumentName))

            Dim tr As Transaction
            tr = m_inventorApplication.TransactionManager.StartTransaction(m_inventorApplication.ActiveDocument, "Create Standard Parts From Excel")
            BeginCreateAssemblyStructure()
            tr.End()
        End Sub

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
            dlg1.Dispose()
            Dim asmDoc As AssemblyDocument
            Try
                asmDoc = m_inventorApplication.ActiveDocument
            Catch ex As Exception
                MessageBox.Show("You need to have an assembly open, not a part!")
            End Try
            parentAssemblyFilename = System.IO.Path.GetFileNameWithoutExtension(m_inventorApplication.ActiveDocument.DisplayName)
            Dim level As Long = 0
            ParseFolders(DirStruct, level)
            'here is where we need to query the vault/download files
            BeginRePopulateCompleteList()
            GetLatestFilesFromVault()
            EditedListFromSystemDrive = CompleteListFromSystemDrive
            Dim grouped = CompleteListFromSystemDrive.OrderBy(Function(x) x.Level).GroupBy(Function(x) x.Level)
            For Each Group In grouped
                If Group.Key > 1 Then
                    For Each item As SubObjectCls In Group
                        Dim File As ACW.File = (From f In VaultedFileList
                                               Where f.Name = item.FileName
                                               Select f).FirstOrDefault()
                        If Not File Is Nothing Then
                            Dim ids As Long() = FileAssocLiteArrays.Select(
                                Function(EachFileAssocLite) EachFileAssocLite).Where(
                                Function(FileAssocLiteToCheck) FileAssocLiteToCheck.ParFileId = File.Id).Select(
                                Function(ChildFileAssoc) ChildFileAssoc.CldFileId).ToArray()
                            If Not ids.Length = 0 Then
                                Dim Linkedfiles = VaultedFileList.Where(Function(x) ids.Contains(x.Id))
                                For Each RemoveableItem In Linkedfiles
                                    Dim FileToRemove = (From f In EditedListFromSystemDrive
                                                       Where f.FileName = RemoveableItem.Name
                                                       Select f).FirstOrDefault()
                                    EditedListFromSystemDrive.Remove(FileToRemove)
                                Next
                            End If
                        End If
                    Next
                End If
            Next

            Dim EditedGrouped = EditedListFromSystemDrive.OrderBy(Function(x) x.Level).GroupBy(Function(x) x.Level)
            Dim percent As Double = Nothing
            Dim i As Integer = 0
            For Each group In EditedGrouped
                If Not group.Key = 1 Then 'skip the first level as it's our top level assembly!
                    'MessageBox.Show("Level = " & group.Key & " of " & grouped.Count)
                    'If group.Key <= 3 Then
                    For Each subobj As SubObjectCls In group
                        percent = (CDbl(i) / EditedListFromSystemDrive.Count)
                        UpdateStatusBar(percent, "Creating Assembly Structure at Level {" & group.Key & "}... Please Wait")
                        CreateAssemblyStructure(subobj, subobj.ParentAssembly)
                        i += 1
                    Next
                    'End If
                End If
            Next
            MessageBox.Show(CompleteListFromSystemDrive.Count)
            PlayBackgroundSoundResource()
        End Sub

        Sub PlayBackgroundSoundResource()
            My.Computer.Audio.Play(My.Resources.fin, _
                AudioPlayMode.WaitToComplete)
        End Sub

        ''' <summary>
        ''' Copied from our QueryVault Excel addin.
        ''' Allows us to search for a bunch of filenames
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub BeginRePopulateCompleteList()
            DoInitialSearch()
            'Dim results As List(Of Autodesk.Connectivity.WebServices.File) = New List(Of ACW.File)()
            If VaultedFileList.Count > 0 Then
                VaultedFileList.RemoveAll(Function(x) x.Name.ToLower().StartsWith("replace with"))
                VaultedFileList.RemoveAll(Function(x) x.Name.ToLower().Contains("_"))
                VaultedFileList = VaultedFileList.FindAll(Function(x) x.Name.EndsWith(".ipt") Or x.Name.EndsWith(".iam"))
            End If
            FileAssocLiteArrays = m_conn.WebServiceManager.DocumentService.GetFileAssociationLitesByIds(
                VaultedFileList.Select(Function(x) x.Id).ToArray(),
                FileAssocAlg.Actual,
                FileAssociationTypeEnum.All,
                True,
                FileAssociationTypeEnum.All,
                True,
                False,
                False,
                False)
        End Sub

        ''' <summary>
        ''' Performs the vault equivalent of a GET on the vaulted files we found.
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub GetLatestFilesFromVault()
            Dim percent As Double = Nothing
            Dim i As Integer = 0
            For Each File As ADSK.File In VaultedFileList
                percent = (CDbl(i) / VaultedFileList.Count)
                UpdateStatusBar(percent, "Performing a GET from the Vault... Please Wait")
                If File.Cloaked Then Continue For
                Dim settings As Vault.Settings.AcquireFilesSettings = New Vault.Settings.AcquireFilesSettings(m_conn)
                settings.AddFileToAcquire(New VDF.Vault.Currency.Entities.FileIteration(m_conn, File), VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download)
                m_conn.FileManager.AcquireFiles(settings)
                i += 1
            Next
        End Sub

        ''' <summary>
        ''' Based on the implementation in the QueryVault Excel addin, this will attempt to take the CompleteList object, extract the "friendly" filenames
        ''' and search the vault for them.
        ''' Then we'll need to remove the files which are children of vaulted components from CompleteList
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub DoInitialSearch()
            UpdateStatusBar("Getting Initial File details (If available) From Vault... Please Wait")
            Dim NonVaultedFileNames As String() = CompleteListFromSystemDrive.Select(Function(x) x.PartNo).ToArray()
            NonVaultedFileNames = NonVaultedFileNames.Distinct.ToArray()
            'build the search based on our range of Distinct (potentially non-vaulted) filenames.
            Dim conditions As SrchCond() = New SrchCond(NonVaultedFileNames.Length - 1) {}
            Dim percent As Double = Nothing
            For i As Integer = 0 To NonVaultedFileNames.Length - 1
                'UpdateStatusBar(percent, "Building Initial Search Conditions... Please Wait")
                Dim searchCondition As New SrchCond()
                searchCondition.PropDefId = 9
                'filename
                searchCondition.PropTyp = PropertySearchType.SingleProperty
                searchCondition.SrchOper = 1
                searchCondition.SrchTxt = NonVaultedFileNames(i)
                searchCondition.SrchRule = SearchRuleType.May
                conditions(i) = searchCondition
            Next
            Dim bookmark As String = String.Empty
            Dim status As SrchStatus = Nothing
            VaultedFileList = New List(Of ACW.File)()

            While status Is Nothing OrElse VaultedFileList.Count < status.TotalHits
                Dim files As Autodesk.Connectivity.WebServices.File() = m_conn.WebServiceManager.DocumentService.FindFilesBySearchConditions(conditions, Nothing, Nothing, True, True, bookmark, _
                    status)
                If files IsNot Nothing Then
                    VaultedFileList.AddRange(files)
                    percent = (CDbl(VaultedFileList.Count) / status.TotalHits)
                    UpdateStatusBar(percent, "Searching the Vault... Please Wait")
                End If
            End While
            'get all FileIterations
            FileIterations = New List(Of Vault.Currency.Entities.FileIteration)(VaultedFileList.[Select](Function(result) New VDF.Vault.Currency.Entities.FileIteration(m_conn, result)))
            'FileIterations = VaultedFileList.Select(Function(result) New VDF.Vault.Currency.Entities.FileIteration(m_conn, result))
            'then get all the folders for these files
            FolderIdsToFolderEntities = m_conn.FolderManager.GetFoldersByIds(FileIterations.Select(Function(file) file.FolderId))
        End Sub



        ''' <summary>
        ''' Creates dummy files to pre-populate Vault with or creates files named "Replace With " to denote where files already exist.
        ''' </summary>
        ''' <param name="subObject">The Subobject we need to create/replace</param>
        ''' <returns>Returns the filename of the newly created dummy Part/Assemly</returns>
        ''' <remarks></remarks>
        Private Function CreateAssemblyComponents(SubObject As SubObjectCls) As String
            Dim basepartname As String = String.Empty
            Dim basecablepartname As String = String.Empty
            Dim newfilename As String = String.Empty
            Dim NewCablePartName As String = String.Empty
            Try
                Dim i As Integer = PartsListFromExcel.FindIndex(Function(str As SubObjectCls) str.PartNo = SubObject.PartNo)
                If Not i = -1 Then
                    AligniPropertyValues(SubObject, PartsListFromExcel(i))
                End If
                If SubObject.PartNo.StartsWith("AS-", StringComparison.Ordinal) Then
                    If Not SubObject.LegacyDescr Is Nothing Then
                        If SubObject.LegacyDescr.ToLower.Contains("cable") Then
                            newfilename = System.IO.Path.GetDirectoryName(m_inventorApplication.ActiveDocument.FullDocumentName) & "\" & SubObject.PartNo & ".iam"
                            basepartname = "C:\LEGACY VAULT WORKING FOLDER\Designs\DT-99999-000 CABLE.iam"
                            basecablepartname = "C:\LEGACY VAULT WORKING FOLDER\Designs\DT-99999-000 CABLE.ipt"
                            NewCablePartName = System.IO.Path.GetDirectoryName(m_inventorApplication.ActiveDocument.FullDocumentName) & "\" & SubObject.PartNo & ".ipt"
                        Else
                            newfilename = System.IO.Path.GetDirectoryName(m_inventorApplication.ActiveDocument.FullDocumentName) & "\" & SubObject.PartNo & ".iam"
                            basepartname = "C:\LEGACY VAULT WORKING FOLDER\Designs\DT-99999-000.iam"
                        End If
                    Else
                        newfilename = System.IO.Path.GetDirectoryName(m_inventorApplication.ActiveDocument.FullDocumentName) & "\" & SubObject.PartNo & ".iam"
                        basepartname = "C:\LEGACY VAULT WORKING FOLDER\Designs\DT-99999-000.iam"
                    End If
                ElseIf SubObject.PartNo.StartsWith("DT-", StringComparison.Ordinal) Then
                    If Not SubObject.LegacyDescr Is Nothing Then
                        If SubObject.LegacyDescr.Contains("ASSEMBLY") Or SubObject.LegacyDescr.Contains("ASSY") Then
                            newfilename = System.IO.Path.GetDirectoryName(m_inventorApplication.ActiveDocument.FullDocumentName) & "\" & SubObject.PartNo & ".iam"
                            basepartname = "C:\LEGACY VAULT WORKING FOLDER\Designs\DT-99999-000.iam"
                        Else
                            newfilename = System.IO.Path.GetDirectoryName(m_inventorApplication.ActiveDocument.FullDocumentName) & "\" & SubObject.PartNo & ".ipt"
                            basepartname = "C:\LEGACY VAULT WORKING FOLDER\Designs\DT-99999-000.ipt"
                        End If
                    Else
                        newfilename = System.IO.Path.GetDirectoryName(m_inventorApplication.ActiveDocument.FullDocumentName) & "\" & SubObject.PartNo & ".ipt"
                        basepartname = "C:\LEGACY VAULT WORKING FOLDER\Designs\DT-99999-000.ipt"
                    End If
                ElseIf SubObject.PartNo.StartsWith("DL-", StringComparison.Ordinal) Then
                    'technically this assembly is missing from the structure!
                    newfilename = System.IO.Path.GetDirectoryName(m_inventorApplication.ActiveDocument.FullDocumentName) &
                        "\" & GetFriendlyDirName(SubObject.PartNo) & ".iam"
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
                        If Not NewCablePartName = String.Empty Then
                            System.IO.File.Copy(basecablepartname, NewCablePartName)
                            NewCablePartName = Nothing
                        End If
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
        Private Sub CreateAssemblyStructure(SubObject As SubObjectCls, ParentName As String)
            Dim asmDoc As AssemblyDocument = Nothing
            Dim realOcc As ComponentOccurrence = Nothing
            Dim realOccStr As String = String.Empty
            Dim PosnMatrix As Matrix
            Dim newfilename As String = String.Empty
            Dim i As Integer = 0
            Try
                i = VaultedFileList.FindIndex(Function(x) x.Name = SubObject.FileName)
                If i = -1 Then
                    newfilename = CreateAssemblyComponents(SubObject)
                Else
                    Dim foldername As String = FolderIdsToFolderEntities.Select(Function(m) m).Where(Function(kvp) kvp.Key = VaultedFileList.Item(i).FolderId).Select(Function(k) k.Value).First().FullName
                    foldername = foldername.Replace("/", "\").Replace("$", "C:\Legacy Vault Working Folder\")
                    newfilename = foldername & "\" & VaultedFileList.Item(i).Name
                    'Dim FileId As Long = VaultedFileList.Item(i).Id
                    'VaultedFileList.RemoveAll(Function(x) )
                End If

                PosnMatrix = m_inventorApplication.TransientGeometry.CreateMatrix
                If ParentName = System.IO.Path.GetFileNameWithoutExtension(m_inventorApplication.ActiveDocument.DisplayName) Then
                    'immediate descendants of the parent assembly
                    asmDoc = m_inventorApplication.ActiveDocument
                    Try
                        If Not asmDoc.ComponentDefinition.Document.FullFileName.Tolower.Contains(ProjectCode.ToLower()) Then Exit Sub
                        realOcc = asmDoc.ComponentDefinition.Occurrences.Add(newfilename, PosnMatrix)
                        realOccStr = realOcc.Name
                        If Not realOccStr.StartsWith("Replace With", StringComparison.OrdinalIgnoreCase) And _
                            System.IO.Path.GetDirectoryName(newfilename).ToLower.Contains(ProjectCode.ToLower) Then 'assign iproperties to new parts
                            'cable assemblies in parent assembly
                            ReplaceCableParts(realOcc, SubObject)
                        End If
                    Catch ex As Exception
                        MessageBox.Show("Exception was: " + ex.Message + vbCrLf + ex.StackTrace)
                    End Try
                Else
                    'one of its grandchildren
                    Dim tmpdoc As Inventor.Document = Nothing
                    For Each doc As Inventor.Document In m_inventorApplication.ActiveDocument.AllReferencedDocuments
                        'at this point we should stop the insert of occurrences if the assembly is called "Replace With XXX" as it's doing work that's unecessary!
                        If doc.DisplayName = ParentName & ".iam" Or doc.DisplayName.StartsWith("Replace With " & ParentName, StringComparison.OrdinalIgnoreCase) Then
                            tmpdoc = doc
                            Exit For
                        End If
                    Next

                    Try
                        If Not tmpdoc Is Nothing Then
                            asmDoc = tmpdoc
                            For Each a As ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences
                                If a.Name = System.IO.Path.GetFileNameWithoutExtension(newfilename) & ":1" Then
                                    realOcc = a
                                End If
                            Next

                            If realOcc Is Nothing Then 'only insert the occurrence once or we end up with a huge assembly containing multiple occurrences...
                                'don't add anything to files outside of this project folder as we have to assume they are not checked out/are released.
                                If Not asmDoc.ComponentDefinition.Document.FullFileName.tolower.Contains(ProjectCode.ToLower()) Then Exit Sub
                                realOcc = asmDoc.ComponentDefinition.Occurrences.Add(newfilename, PosnMatrix)
                                realOccStr = realOcc.Name
                                If Not realOccStr.StartsWith("Replace With", StringComparison.OrdinalIgnoreCase) And _
                                    System.IO.Path.GetDirectoryName(newfilename).ToLower.Contains(ProjectCode.ToLower) Then 'assign iproperties to new parts
                                    ReplaceCableParts(realOcc, SubObject)
                                End If
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
        Private Sub AligniPropertyValues(ByRef SubObject As SubObjectCls, ByVal PartsListObj As SubObjectCls)
            SubObject.LegacyDescr = PartsListObj.LegacyDescr
            SubObject.LegacyDrawingNo = PartsListObj.LegacyDrawingNo
            SubObject.LegacyRev = PartsListObj.LegacyRev
        End Sub

        ''' <summary>
        ''' Assigns IProperties using the ComponentOccurrence and SubObject objects
        ''' </summary>
        ''' <param name="realocc">the occurrence we wish to edit</param>
        ''' <param name="subObject">the subobject containing our aligned iProperties (taken from Excel)</param>
        ''' <remarks></remarks>
        Private Sub AssignIProperties(ByVal RealOcc As ComponentOccurrence, ByVal SubObject As SubObjectCls)
            Try
                Dim invProjProperties As PropertySet = RealOcc.Definition.Document.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}")
                Dim invSummaryiProperties As PropertySet = RealOcc.Definition.Document.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}")
                invProjProperties.ItemByPropId(PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties).Value = SubObject.PartNo 'part number
                If SubObject.LegacyDescr Is Nothing Then
                    invProjProperties.ItemByPropId(PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties).Value = "POPULATED BY CAFE TOOL" 'description
                Else
                    invProjProperties.ItemByPropId(PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties).Value = SubObject.LegacyDescr 'description
                    invSummaryiProperties.ItemByPropId(PropertiesForSummaryInformationEnum.kTitleSummaryInformation).Value = SubObject.LegacyDescr 'title
                End If
                invProjProperties.ItemByPropId(PropertiesForDesignTrackingPropertiesEnum.kProjectDesignTrackingProperties).Value = "A90.1" 'project
                If SubObject.LegacyRev Is Nothing Then
                    invSummaryiProperties.ItemByPropId(PropertiesForSummaryInformationEnum.kRevisionSummaryInformation).Value = "A" 'revision
                Else
                    invSummaryiProperties.ItemByPropId(PropertiesForSummaryInformationEnum.kRevisionSummaryInformation).Value = SubObject.LegacyRev 'revision
                End If
                If SubObject.LegacyDrawingNo Is Nothing Then
                    invSummaryiProperties.ItemByPropId(PropertiesForSummaryInformationEnum.kSubjectSummaryInformation).Value = "HR/0/#####" 'subject
                Else
                    invSummaryiProperties.ItemByPropId(PropertiesForSummaryInformationEnum.kSubjectSummaryInformation).Value = SubObject.LegacyDrawingNo 'subject
                End If
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
        Private Function FindFileInVWF(NewFileName As String) As String
            Dim dir = New DirectoryInfo("C:\Legacy Vault Working Folder\Designs")
            Dim tmpstr = GetExistingFile(dir, System.IO.Path.GetFileNameWithoutExtension(NewFileName))
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
        Private Function GetExistingFile(ByVal Dir As DirectoryInfo, ByVal NewFileName As String) As String
            Dim foundfilename As String = String.Empty
            Try
                'this way should create a large list of .ipt/iam files in one pass that we can keep for later reuse. It will likely be slow initially but faster later.
                If files Is Nothing Then
                    files = Dir.EnumerateFiles("*.*", SearchOption.AllDirectories).Where(Function(s As FileInfo) _
                                                                                             s.Name.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) OrElse _
                                                                                             s.Name.EndsWith(".iam", StringComparison.OrdinalIgnoreCase))
                End If
                Dim file As FileInfo = (From f As FileInfo In files
                                       Where System.IO.Path.GetFileNameWithoutExtension(f.Name) = NewFileName
                                       Select f).FirstOrDefault()
                If Not file Is Nothing Then
                    foundfile = file
                    foundfilename = file.Name
                End If
            Catch ex As Exception
                MessageBox.Show("Exception was: " + ex.Message + vbCrLf + ex.StackTrace)
            End Try
            Return foundfilename
        End Function
        ''' <summary>
        ''' Replaces Cable parts with their required copy/renamed part files and adjusts the relevant parameters accordingly.
        ''' </summary>
        ''' <param name="realOcc">the occurrence we are editing</param>
        ''' <param name="SubObject">the subobject to edit</param>
        ''' <remarks></remarks>
        Private Sub ReplaceCableParts(ByVal realOcc As ComponentOccurrence, ByVal SubObject As SubObjectCls)
            If Not SubObject.LegacyDescr Is Nothing Then
                If SubObject.LegacyDescr.ToLower.Contains("cable") And SubObject.FileName.ToLower.EndsWith(".iam") Then
                    Dim occNum As Integer = 1
                    For Each SubcompOcc As ComponentOccurrence In realOcc.SubOccurrences
                        If SubcompOcc.Name.EndsWith(":" & occNum.ToString()) Then
                            Dim fileNameToSwapWith As String = System.IO.Path.GetDirectoryName(m_inventorApplication.ActiveDocument.FullFileName) & "\" & System.IO.Path.GetFileNameWithoutExtension(SubObject.FileName) & ".ipt"
                            If System.IO.File.Exists(fileNameToSwapWith) Then
                                SubcompOcc.Replace(fileNameToSwapWith, True)
                            Else
                                SubcompOcc.Definition.Document.saveas(fileNameToSwapWith, False)
                                SubcompOcc.Replace(fileNameToSwapWith, True)
                            End If
                            Dim CableCompDef As PartComponentDefinition = SubcompOcc.Definition
                            Dim CableIdParameter As Parameter = CableCompDef.Parameters.Item("CABLE_ID")
                            CableIdParameter.Value = SubObject.LegacyDrawingNo
                            AssignIProperties(SubcompOcc, SubObject)
                            occNum += 1
                        End If
                    Next
                Else
                    AssignIProperties(realOcc, SubObject)
                End If
            Else
                AssignIProperties(realOcc, SubObject)
            End If
        End Sub
        ''' <summary>
        ''' Updates the statusbar with a percentage value
        ''' </summary>
        ''' <param name="percent"></param>
        ''' <param name="Message"></param>
        ''' <remarks></remarks>
        Private Sub UpdateStatusBar(ByVal percent As Double, ByVal Message As String)
            m_inventorApplication.StatusBarText = Message + " (" + percent.ToString("P1") + ")"
        End Sub
        ''' <summary>
        ''' updates the statusbar with a string value.
        ''' </summary>
        ''' <param name="Message"></param>
        ''' <remarks></remarks>
        Private Sub UpdateStatusBar(ByVal Message As String)
            m_inventorApplication.StatusBarText = Message
        End Sub
#End Region
#Region "Folder Structure Parser from console app"
        ''' <summary>
        ''' Parses a series of subfolders, ignoring directories called "Superseded", documents whose names contain "IL" or "DL"
        ''' </summary>
        ''' <param name="dir">The Directory to search</param>
        ''' <param name="level">the structure level we are currently at</param>
        ''' <remarks></remarks>
        Public Sub ParseFolders(ByVal Dir As DirectoryInfo, ByVal Level As Long)
            'Try
            If Not Dir.Name.Contains("Superseded") Then
                Dim friendlydirname As String = GetFriendlyDirName(Dir.Name)
                Dim friendlyparentdirname As String = String.Empty
                If ContainsNumbers(Dir.Parent.Name) Then
                    friendlyparentdirname = GetFriendlyDirName(Dir.Parent.Name)
                Else
                    friendlyparentdirname = parentAssemblyFilename
                End If
                'need to account for instances where the "AS-" file isn't the first we'll find
                Dim thisAssy As FileInfo = (From a As FileInfo In Dir.GetFiles()
                                            Where GetFriendlyName(a.Name) = friendlydirname And
                                            Not a.Name.Contains("IL") And
                                            Not a.Name.Contains("DL") And
                                            Not a.Name.Contains("SP") And
                                            Not a.Name.ToLower.Contains("missing") And
                                            (getsheetnum(a.Name) <= 1)
                                            Select a).FirstOrDefault
                If Not thisAssy Is Nothing Then
                    CompleteListFromSystemDrive.Add(New SubObjectCls() With
                                                    {.PartNo = GetFriendlyName(thisAssy.Name),
                                                     .ParentAssembly = friendlyparentdirname,
                                                     .Level = Level + 1,
                                                     .FileName = IIf(GetFriendlyName(thisAssy.Name).ToLower.StartsWith("as-"),
                                                                     GetFriendlyName(thisAssy.Name) & ".iam",
                                                                     GetFriendlyName(thisAssy.Name) & ".ipt")
                                                    }
                                                )
                Else
                    'the assembly drawing probably exists somewhere else in the structure so for now we'll assume it's just "missing"
                    thisAssy = New FileInfo(Dir.FullName & "\" & GetFriendlyDirName(Dir.Name) & ".txt")
                    CompleteListFromSystemDrive.Add(New SubObjectCls() With
                                                    {.PartNo = GetFriendlyName(thisAssy.Name),
                                                     .ParentAssembly = friendlyparentdirname,
                                                     .Level = Level + 1,
                                                     .FileName = IIf(GetFriendlyName(thisAssy.Name).ToLower.StartsWith("as-"),
                                                                     GetFriendlyName(thisAssy.Name) & ".iam",
                                                                     GetFriendlyName(thisAssy.Name) & ".ipt")
                        }
                    )
                End If
                For Each file As FileInfo In Dir.GetFiles()
                    If Not file.Name.Contains("IL") And Not file.Name.Contains("DL") And Not file.Name.Contains("SP") And Not file.Name = thisAssy.Name And Not file.Name.ToLower.Contains("missing") And file.Name.EndsWith(".pdf") Then
                        'if the directory name is the same as the assembly name then the parentassembly is the folder above!
                        Dim friendlyfilename As String = GetFriendlyName(file.Name)
                        If friendlydirname = friendlyfilename Then 'parent assembly in this folder
                            If getsheetnum(file.Name) <= 1 Then
                                CompleteListFromSystemDrive.Add(New SubObjectCls() With
                                                                {.PartNo = friendlyfilename,
                                                                 .ParentAssembly = friendlyparentdirname,
                                                                 .Level = Level + 1,
                                                                 .FileName = IIf(GetFriendlyName(file.Name).ToLower.StartsWith("as-"),
                                                                                 GetFriendlyName(file.Name) & ".iam",
                                                                                 GetFriendlyName(file.Name) & ".ipt")
                                                                }
                                                            )
                            End If
                        Else
                            If getsheetnum(file.Name) <= 1 Then
                                CompleteListFromSystemDrive.Add(New SubObjectCls() With
                                                                {.PartNo = friendlyfilename,
                                                                 .ParentAssembly = friendlydirname,
                                                                 .Level = Level + 2,
                                                                 .FileName = IIf(file.Name.ToLower.StartsWith("AS-"),
                                                                                 GetFriendlyName(file.Name) & ".iam",
                                                                                 GetFriendlyName(file.Name) & ".ipt")
                                                                }
                                                            )
                            End If
                        End If
                    End If
                Next
                For Each subDir As DirectoryInfo In Dir.GetDirectories()
                    If Not subDir.Name.Contains("Superseded") Then
                        ParseFolders(subDir, Level + 1)
                    End If
                Next
            End If
            highestlevel += 1
            'Catch ex As Exception
            '    MessageBox.Show("Exception was: " + ex.Message + vbCrLf + ex.StackTrace)
            'End Try
            'Return info
        End Sub

        ''' <summary>
        ''' Checks whether our folder name contains 5 or more # characters
        ''' </summary>
        ''' <param name="Str"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function ContainsNumbers(Str As String) As Boolean
            Dim f As String = Nothing
            Dim r As New Regex("\d{5,}")
            If r.Match(Str).Captures.Count > 0 Then
                Return True
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' Returns a "Friendly" filename for comparison-sake
        ''' </summary>
        ''' <param name="Str">the String to match against</param>
        ''' <returns>Returns the matched String</returns>
        ''' <remarks></remarks>
        Public Function GetFriendlyName(Str As String) As String
            Dim f As String = String.Empty
            Dim r As New Regex("\w{2}-\d{5,}|\w{2}-\w\d{5,}")
            f = r.Match(Str).Captures(0).ToString() + "-000"
            Console.WriteLine(f)
            Return f
        End Function

        ''' <summary>
        ''' Returns a "Friendly" directory name for comparison-sake and use as the "parentname"
        ''' </summary>
        ''' <param name="Str">the String to match against</param>
        ''' <returns>Returns the matched String</returns>
        ''' <remarks></remarks>
        Public Function GetFriendlyDirName(Str As String) As String
            If Not Str.Contains(":") Then
                Dim f As String = String.Empty
                Dim r As New Regex("\d{3,}|\w\d{3,}")
                f = "AS-" + r.Match(Str).Captures(0).ToString() + "-000"
                Return f
            Else
                'Return p1
                Return parentAssemblyFilename
            End If
        End Function

        ''' <summary>
        ''' Returns the Directory code without any spurious extras
        ''' </summary>
        ''' <param name="Str">the String to match against</param>
        ''' <returns>Returns the matched String</returns>
        ''' <remarks></remarks>
        Public Function GetDirName(Str As String) As String
            If Not Str.Contains(":") Then
                Dim f As String = String.Empty
                Dim r As New Regex("\d{3,}|\w\d{3,}") 'get just the @##### or ###### drawing number
                f = r.Match(Str).Captures(0).ToString()
                Return f
            Else
                Return Str
            End If
        End Function

        ''' <summary>
        ''' Returns the Sheetnum for each file we've found
        ''' </summary>
        ''' <param name="Str">the String to match against</param>
        ''' <returns>Returns the matched String</returns>
        ''' <remarks></remarks>
        Private Function getsheetnum(Str As String) As Integer
            Dim f As String = String.Empty
            Dim pattern As String = "(.*)(sht-)(\d{3})(.*)"
            Dim matches As MatchCollection = Regex.Matches(Str, pattern)
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
        Private m_PartNo As String
        Public Property PartNo() As String
            Get
                Return m_PartNo
            End Get
            Set(ByVal value As String)
                m_PartNo = value
            End Set
        End Property
        Private m_LegacyDescr As String
        Public Property LegacyDescr() As String
            Get
                Return m_LegacyDescr
            End Get
            Set(ByVal value As String)
                m_LegacyDescr = value
            End Set
        End Property
        Private m_LegacyRev As String
        Public Property LegacyRev() As String
            Get
                Return m_LegacyRev
            End Get
            Set(ByVal value As String)
                m_LegacyRev = value
            End Set
        End Property
        Private m_LegacyDrawingNo As String
        Public Property LegacyDrawingNo() As String
            Get
                Return m_LegacyDrawingNo
            End Get
            Set(ByVal value As String)
                m_LegacyDrawingNo = value
            End Set
        End Property
        Private m_ParentAssembly As String
        Public Property ParentAssembly() As String
            Get
                Return m_ParentAssembly
            End Get
            Set(ByVal value As String)
                m_ParentAssembly = value
            End Set
        End Property
        Private m_FileName As String
        Public Property FileName() As String
            Get
                Return m_FileName
            End Get
            Set(ByVal value As String)
                m_FileName = value
            End Set
        End Property

        Private m_HasChildren As Boolean
        Public Property HasChildren() As Boolean
            Get
                Return m_HasChildren
            End Get
            Set(ByVal value As Boolean)
                m_HasChildren = value
            End Set
        End Property
        Private m_Children As List(Of SubObjectCls)
        Public Property Children() As List(Of SubObjectCls)
            Get
                Return m_Children
            End Get
            Set(ByVal value As List(Of SubObjectCls))
                m_Children = value
            End Set
        End Property
        Private m_Level As Long
        Public Property Level() As Long
            Get
                Return m_Level
            End Get
            Set(ByVal value As Long)
                m_Level = value
            End Set
        End Property


        ''' <summary>
        ''' Allows us to use this Class with LINQ
        ''' </summary>
        ''' <param name="other">the other instance to compare to</param>
        ''' <returns>Returns the comparison requested</returns>
        ''' <remarks></remarks>
        Public Function CompareTo(Other As SubObjectCls) As Integer Implements IComparable(Of SubObjectCls).CompareTo
            Return Me.CompareTo(Other)
        End Function

        ''' <summary>
        ''' Split the Collection into a list of SubObjectCls
        ''' </summary>
        ''' <param name="source">The Source List to split</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Split(Source As List(Of SubObjectCls)) As List(Of List(Of SubObjectCls))
            Return Source.[Select](Function(x, i) New With { _
                Key .Index = i, _
                Key .Value = x _
            }).GroupBy(Function(x) x.Index).[Select](Function(x) x.[Select](Function(v) v.Value).ToList()).ToList()
        End Function

        ''' <summary>
        ''' Allows for grouping of the SubObjectCls by whichever variable we choose **Not Implemented**
        ''' </summary>
        ''' <param name="P"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function GroupBy(P As Object) As Object
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
#Region "ListBoxFileItem"
    ''' <summary>
    ''' A list box item which contains a File object
    ''' </summary>
    Public Class ListBoxFileItem
        Private m_file As Vault.Currency.Entities.FileIteration
        Public ReadOnly Property File() As Vault.Currency.Entities.FileIteration
            Get
                Return m_file
            End Get
        End Property

        Public Sub New(f As Vault.Currency.Entities.FileIteration)
            m_file = f
        End Sub
        Public m_folder As VDF.Vault.Currency.Entities.Folder
        Public ReadOnly Property Folder() As VDF.Vault.Currency.Entities.Folder
            Get
                Return m_folder
            End Get
        End Property
        ''' <summary>
        ''' Determines the text displayed in the ListBox
        ''' </summary>
        Public Overrides Function ToString() As String
            Return Me.m_file.EntityName
        End Function

        Public Property FeatureCount() As Integer
            Get
                Return m_FeatureCount
            End Get
            Set(value As Integer)
                m_FeatureCount = Value
            End Set
        End Property
        Private m_FeatureCount As Integer

        Public Property OccurrenceCount() As Integer
            Get
                Return m_OccurrenceCount
            End Get
            Set(value As Integer)
                m_OccurrenceCount = Value
            End Set
        End Property
        Private m_OccurrenceCount As Integer

        Public Property ParameterCount() As Integer
            Get
                Return m_ParameterCount
            End Get
            Set(value As Integer)
                m_ParameterCount = Value
            End Set
        End Property
        Private m_ParameterCount As Integer

        Public Property ConstraintCount() As Integer
            Get
                Return m_ConstraintCount
            End Get
            Set(value As Integer)
                m_ConstraintCount = Value
            End Set
        End Property
        Private m_ConstraintCount As Integer

        Public Property Material() As String
            Get
                Return m_Material
            End Get
            Set(value As String)
                m_Material = Value
            End Set
        End Property
        Private m_Material As String

        Public Property Title() As String
            Get
                Return m_Title
            End Get
            Set(value As String)
                m_Title = Value
            End Set
        End Property
        Private m_Title As String

        Public Property RevNumber() As String
            Get
                Return m_RevNumber
            End Get
            Set(value As String)
                m_RevNumber = Value
            End Set
        End Property
        Private m_RevNumber As String

        Public Property LegacyDrawingNumber() As String
            Get
                Return m_LegacyDrawingNumber
            End Get
            Set(value As String)
                m_LegacyDrawingNumber = Value
            End Set
        End Property
        Private m_LegacyDrawingNumber As String

        'public VDF.Vault.Currency.Properties.PropertyValues propValues;
        'public VDF.Vault.Currency.Properties.PropertyValues PropValues
        '{
        '    get { return propValues; }
        '}

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
    End Class

#End Region
#Region "Excel helper Class"
    Friend Class ExcelWrapper
        Implements IDisposable
        Private Class Window
            <DllImport("user32.dll", SetLastError:=True)> _
            Private Shared Function FindWindow(lpClassName As String, lpWindowName As String) As IntPtr
            End Function

            <DllImport("user32.dll")> _
            Private Shared Function GetWindowThreadProcessId(hWnd As IntPtr, ByRef ProcessID As IntPtr) As IntPtr
            End Function

            Public Shared Function GetWindowThreadProcessId(hWnd As IntPtr) As IntPtr
                Dim processID As IntPtr
                Dim returnResult As IntPtr = GetWindowThreadProcessId(hWnd, processID)
                Return processID
            End Function

            Public Shared Function FindExcel(caption As String) As IntPtr
                Dim hWnd As IntPtr = FindWindow("XLMAIN", caption)
                Return hWnd
            End Function
        End Class

        Private m_excel As Excel.Application
        Private m_windowHandle As IntPtr
        Private m_processID As IntPtr
        Private Const ExcelWindowCaption As String = "Running From CAFE Tool O_o"

        Public Sub New()
            m_excel = CreateExcelApplication()
            m_windowHandle = Window.FindExcel(ExcelWindowCaption)
            m_processID = Window.GetWindowThreadProcessId(m_windowHandle)
        End Sub

        Private Function CreateExcelApplication() As Excel.Application
            Dim excel As New Excel.Application
            excel.Caption = ExcelWindowCaption
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AlertBeforeOverwriting = False
            excel.AskToUpdateLinks = False
            Return excel
        End Function

        Public ReadOnly Property Excel() As Excel.Application
            Get
                Return Me.m_excel
            End Get
        End Property

        Public ReadOnly Property ProcessID() As Integer
            Get
                Return Me.m_processID.ToInt32()
            End Get
        End Property

        Public ReadOnly Property WindowHandle() As Integer
            Get
                Return Me.m_windowHandle.ToInt32()
            End Get
        End Property

        Public Sub Dispose() Implements IDisposable.Dispose
            If m_excel IsNot Nothing Then
                m_excel.Workbooks.Close()
                m_excel.Quit()
                Marshal.ReleaseComObject(m_excel)
                m_excel = Nothing
                GC.Collect()
                GC.WaitForPendingFinalizers()
                GC.Collect()
                GC.WaitForPendingFinalizers()
                Try
                    Dim processToKill As Process = Process.GetProcessById(Me.ProcessID)

                    If processToKill IsNot Nothing Then
                        processToKill.Kill()
                    End If
                Catch
                    Throw
                End Try
            End If
        End Sub
    End Class
#End Region
#Region "Excel Helper Class implementation"
    Class FileToProcess
        Public Property FileName() As String
            Get
                Return m_FileName
            End Get
            Set(value As String)
                m_FileName = Value
            End Set
        End Property
        Private m_FileName As String

        Public Function GetExcelData() As DataSet
            Using wrapper = New ExcelWrapper()
                Dim app As Excel.Application = wrapper.Excel

                Dim workbook As Excel.Workbook = app.Workbooks.Open(Me.FileName, 0, True)

                Try
                    Dim excelData As New DataSet()

                    For Each worksheet As Excel.Worksheet In workbook.Worksheets
                        Dim columns As Integer = worksheet.UsedRange.Columns.Count
                        Dim rows As Integer = worksheet.UsedRange.Rows.Count

                        If columns > 0 Then
                            Dim sheetData As System.Data.DataTable = excelData.Tables.Add(worksheet.Name)

                            Dim RowHeaders As Excel.Range = DirectCast(worksheet.UsedRange.Rows(1), Excel.Range)

                            For j As Integer = 1 To columns
                                Dim columnName As String = RowHeaders.Columns(j).text.ToString()

                                If String.IsNullOrEmpty(columnName) Then
                                    Continue For
                                ElseIf sheetData.Columns.Contains(columnName) Then
                                    Dim i As Integer = 1
                                    Dim c As String

                                    Do
                                        c = columnName & i.ToString()
                                        i += 1
                                    Loop While sheetData.Columns.Contains(c)

                                    sheetData.Columns.Add(c, GetType(String))
                                Else
                                    sheetData.Columns.Add(columnName, GetType(String))
                                End If
                            Next

                            For i As Integer = 2 To rows
                                Dim sheetRow As DataRow = sheetData.NewRow()

                                For j As Integer = 1 To columns
                                    Dim columnName As String = RowHeaders.Columns(j).Text.ToString()

                                    Dim oRange As Excel.Range = DirectCast(worksheet.Cells(i, j), Excel.Range)

                                    If Not String.IsNullOrEmpty(columnName) Then
                                        sheetRow(columnName) = oRange.Text.ToString()
                                    End If
                                Next

                                sheetData.Rows.Add(sheetRow)
                            Next
                        End If
                    Next

                    Return excelData
                    'Catch generatedExceptionName As Exception
                    '    Throw
                Finally
                    workbook.Close()
                    workbook = Nothing
                    app.Quit()
                    app = Nothing
                    'wrapper.Dispose()
                End Try
            End Using
        End Function

        Function GetExcelData(filetab As String) As DataSet
            Using wrapper = New ExcelWrapper()
                Dim app As Excel.Application = wrapper.Excel

                Dim workbook As Excel.Workbook = app.Workbooks.Open(Me.FileName, 0, True)

                Try
                    Dim excelData As New DataSet()
                    Dim worksheet As Excel.Worksheet = app.Sheets(filetab)
                    Dim Columns As Integer = worksheet.UsedRange.Columns.Count
                    Dim Rows As Integer = worksheet.UsedRange.Rows.Count
                    Dim range As Excel.Range = Nothing
                    range = worksheet.UsedRange()
                    Dim values As Object(,) = range.Value2
                    Dim sheetData As System.Data.DataTable = excelData.Tables.Add(worksheet.Name)
                    'column names
                    For j As Integer = 1 To values.GetLength(1)
                        Console.Write(values(2, j))
                        Console.WriteLine()
                    Next

                    'Dim ColumnNames As String() =
                    For i = 1 To values.GetLength(1)
                        Dim ColumnName As String = values(2, i)
                        If String.IsNullOrEmpty(ColumnName) Then
                            Continue For
                        Else
                            sheetData.Columns.Add(ColumnName)
                        End If
                    Next
                    'row data
                    For i = 2 To values.GetLength(0)
                        Dim SheetRow As DataRow = sheetData.NewRow()
                        Dim ColumnNum As Integer = 0
                        For j = 1 To sheetData.Columns.Count
                            Dim column As DataColumn = sheetData.Columns(ColumnNum)
                            If column Is Nothing Then
                                Continue For
                            Else
                                If Not values(i, j) Is Nothing Then
                                    If Not String.IsNullOrEmpty(values(i, j).ToString()) Then
                                        SheetRow(column) = values(i, j).ToString()
                                        'Else
                                        '    Continue For
                                    End If
                                End If
                                ColumnNum += 1
                            End If
                        Next
                        sheetData.Rows.Add(SheetRow)
                    Next
                    Return excelData
                Finally
                    workbook.Close()
                    workbook = Nothing
                    app.Quit()
                    app = Nothing
                    wrapper.Dispose()
                End Try
            End Using
        End Function

    End Class
#End Region
End Namespace