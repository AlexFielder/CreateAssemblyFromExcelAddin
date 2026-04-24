Imports System.IO
Imports System.Reflection
Imports Microsoft.Extensions.Configuration
Imports Newtonsoft.Json.Linq
Imports VDF = Autodesk.DataManagement.Client.Framework

Namespace CreateAssemblyFromExcelAddin

    ''' <summary>
    ''' Controls how the assembly hierarchy is built.
    ''' </summary>
    Public Enum HierarchySourceMode
        ''' <summary>
        ''' Build the hierarchy directly from the Excel spreadsheet's PARENT/DRAWING NUMBER
        ''' columns. No folder scan is required — the spreadsheet is the single source of truth.
        ''' This is the recommended mode for new deployments.
        ''' </summary>
        Excel

        ''' <summary>
        ''' Legacy mode: scan a physical folder tree to infer the hierarchy using regex-based
        ''' filename/directory parsing (AS-/DT-/DL- prefixes, sheet numbers, etc.).
        ''' Use only when the legacy folder structure matches the original project conventions.
        ''' </summary>
        FolderScan
    End Enum

    ''' <summary>
    ''' Loads and persists add-in settings from an <c>appsettings.json</c> file located
    ''' alongside the compiled DLL. Patterned on the Metrol JobPublisher approach
    ''' (see <c>FilePathRuleFactory.CreateRulesFromFile</c>) which uses
    ''' <c>Microsoft.Extensions.Configuration</c> + <c>AddJsonFile</c>.
    ''' Writes use <c>Newtonsoft.Json.Linq</c> so existing unrelated keys are preserved.
    ''' </summary>
    ''' <remarks>
    ''' Settings schema:
    ''' <code>
    ''' {
    '''   "Vault":            { "WorkingFolder": "C:\\Legacy Vault Working Folder" },
    '''   "ProjectTracker":   { "Path": "..." },
    '''   "HierarchySource":  "Excel",
    '''   "Templates": {
    '''     "UseInventorDefaults": false,
    '''     "Assembly":       "DT-99999-000.iam",
    '''     "Part":           "DT-99999-000.ipt",
    '''     "CableAssembly":  "DT-99999-000 CABLE.iam",
    '''     "CablePart":      "DT-99999-000 CABLE.ipt"
    '''   }
    ''' }
    ''' </code>
    ''' </remarks>
    Public Class CafeSettings

        Public Const SettingsFileName As String = "appsettings.json"

        ' Defaults match the pre-existing hard-coded values in StandardAddInServer.vb
        ' so the add-in continues to behave the same if appsettings.json is missing.
        Public Const DefaultVaultWorkingFolder As String = "C:\Legacy Vault Working Folder"
        Public Const DefaultProjectTrackerPath As String = "C:\Legacy Vault Working Folder\Designs\Project Tracker.xlsx"

        ''' <summary>Root of the local Vault Working Folder (no trailing slash).</summary>
        Public Property VaultWorkingFolder As String = DefaultVaultWorkingFolder

        ''' <summary>Full path to the Project Tracker workbook read by RunCAFE.</summary>
        Public Property ProjectTrackerPath As String = DefaultProjectTrackerPath

        ''' <summary>
        ''' Determines whether the assembly hierarchy is derived from Excel data
        ''' or from a legacy folder scan. Defaults to <see cref="HierarchySourceMode.Excel"/>.
        ''' Set to <c>"FolderScan"</c> in <c>appsettings.json</c> to restore the original behaviour.
        ''' </summary>
        Public Property HierarchySource As HierarchySourceMode = HierarchySourceMode.Excel

        ' ── Template settings ────────────────────────────────────────────────

        ''' <summary>
        ''' When <see langword="True"/>, new parts/assemblies are created using Inventor's
        ''' default templates (<c>Documents.Add</c>) instead of copying custom template files.
        ''' Slower for bulk operations but removes the dependency on pre-staged template files.
        ''' </summary>
        Public Property UseInventorDefaults As Boolean = False

        ''' <summary>Template filename for new assemblies (relative to Designs folder).</summary>
        Public Property AssemblyTemplate As String = "DT-99999-000.iam"

        ''' <summary>Template filename for new parts (relative to Designs folder).</summary>
        Public Property PartTemplate As String = "DT-99999-000.ipt"

        ''' <summary>Template filename for cable assemblies (relative to Designs folder).</summary>
        Public Property CableAssemblyTemplate As String = "DT-99999-000 CABLE.iam"

        ''' <summary>Template filename for cable parts (relative to Designs folder).</summary>
        Public Property CablePartTemplate As String = "DT-99999-000 CABLE.ipt"

        ''' <summary>
        ''' Location of <c>appsettings.json</c>. Resolved next to the currently-loaded
        ''' assembly so it works when Inventor loads the add-in from
        ''' <c>C:\ProgramData\Autodesk\Inventor Addins\...</c>.
        ''' </summary>
        Public Shared ReadOnly Property SettingsPath As String
            Get
                Dim asmDir As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                Return Path.Combine(asmDir, SettingsFileName)
            End Get
        End Property

        ''' <summary>
        ''' Loads settings from <c>appsettings.json</c> beside the add-in DLL.
        ''' Returns a populated instance with sensible defaults for any missing keys.
        ''' Never throws — on any failure the defaults are returned.
        ''' </summary>
        Public Shared Function Load() As CafeSettings
            Dim result As New CafeSettings()
            Try
                If Not File.Exists(SettingsPath) Then
                    Return result
                End If

                Dim config As IConfigurationRoot = New ConfigurationBuilder() _
                    .AddJsonFile(SettingsPath, optional:=True, reloadOnChange:=False) _
                    .Build()

                Dim workingFolder As String = config("Vault:WorkingFolder")
                If Not String.IsNullOrWhiteSpace(workingFolder) Then
                    result.VaultWorkingFolder = workingFolder.TrimEnd("\"c, "/"c)
                End If

                Dim hierarchySource As String = config("HierarchySource")
                If Not String.IsNullOrWhiteSpace(hierarchySource) Then
                    Dim parsed As HierarchySourceMode
                    If [Enum].TryParse(Of HierarchySourceMode)(hierarchySource, True, parsed) Then
                        result.HierarchySource = parsed
                    End If
                End If

                Dim trackerPath As String = config("ProjectTracker:Path")
                If Not String.IsNullOrWhiteSpace(trackerPath) Then
                    result.ProjectTrackerPath = trackerPath
                Else
                    ' If only WorkingFolder was set, derive the tracker path from it.
                    result.ProjectTrackerPath = Path.Combine(result.VaultWorkingFolder, "Designs\Project Tracker.xlsx")
                End If

                ' ── Templates ────────────────────────────────────────────────
                Dim useDefaults As String = config("Templates:UseInventorDefaults")
                If Not String.IsNullOrWhiteSpace(useDefaults) Then
                    Dim b As Boolean
                    If Boolean.TryParse(useDefaults, b) Then result.UseInventorDefaults = b
                End If

                Dim asmTpl As String = config("Templates:Assembly")
                If Not String.IsNullOrWhiteSpace(asmTpl) Then result.AssemblyTemplate = asmTpl

                Dim prtTpl As String = config("Templates:Part")
                If Not String.IsNullOrWhiteSpace(prtTpl) Then result.PartTemplate = prtTpl

                Dim cableAsmTpl As String = config("Templates:CableAssembly")
                If Not String.IsNullOrWhiteSpace(cableAsmTpl) Then result.CableAssemblyTemplate = cableAsmTpl

                Dim cablePrtTpl As String = config("Templates:CablePart")
                If Not String.IsNullOrWhiteSpace(cablePrtTpl) Then result.CablePartTemplate = cablePrtTpl

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("CafeSettings.Load failed: " & ex.Message)
            End Try
            Return result
        End Function

        ''' <summary>
        ''' Writes <see cref="VaultWorkingFolder"/> (and, if set, <see cref="ProjectTrackerPath"/>)
        ''' back to <c>appsettings.json</c>. Any keys not managed by this class are preserved.
        ''' </summary>
        Public Sub Save()
            Dim root As JObject
            If File.Exists(SettingsPath) Then
                Try
                    root = JObject.Parse(File.ReadAllText(SettingsPath))
                Catch
                    root = New JObject()
                End Try
            Else
                root = New JObject()
            End If

            Dim vaultSection As JObject = TryCast(root("Vault"), JObject)
            If vaultSection Is Nothing Then
                vaultSection = New JObject()
                root("Vault") = vaultSection
            End If
            vaultSection("WorkingFolder") = VaultWorkingFolder

            If Not String.IsNullOrWhiteSpace(ProjectTrackerPath) Then
                Dim trackerSection As JObject = TryCast(root("ProjectTracker"), JObject)
                If trackerSection Is Nothing Then
                    trackerSection = New JObject()
                    root("ProjectTracker") = trackerSection
                End If
                trackerSection("Path") = ProjectTrackerPath
            End If

            File.WriteAllText(SettingsPath, root.ToString(Newtonsoft.Json.Formatting.Indented))
        End Sub

        ''' <summary>
        ''' Queries the logged-in Vault connection for the root working folder
        ''' (<c>$</c>) and persists it to <c>appsettings.json</c> so that subsequent
        ''' operations — and the references stored inside newly-created assemblies —
        ''' use the path the user is actually mapped to.
        ''' </summary>
        ''' <param name="connection">An active Vault connection (the result of a successful login).</param>
        ''' <returns>The resolved working folder, or <see langword="Nothing"/> on failure.</returns>
        Public Shared Function CaptureFromVault(connection As VDF.Vault.Currency.Connections.Connection) As String
            If connection Is Nothing Then Return Nothing
            Try
                Dim wfPath As String = connection.WorkingFoldersManager.GetWorkingFolder("$", True).FullPath
                If String.IsNullOrWhiteSpace(wfPath) Then Return Nothing

                Dim settings As CafeSettings = Load()
                settings.VaultWorkingFolder = wfPath.TrimEnd("\"c, "/"c)
                settings.Save()
                Return settings.VaultWorkingFolder
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("CafeSettings.CaptureFromVault failed: " & ex.Message)
                Return Nothing
            End Try
        End Function

    End Class

End Namespace