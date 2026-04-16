Imports System.IO
Imports System.Reflection
Imports Microsoft.Extensions.Configuration
Imports Newtonsoft.Json.Linq
Imports VDF = Autodesk.DataManagement.Client.Framework

Namespace CreateAssemblyFromExcelAddin

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
    '''   "Vault":          { "WorkingFolder": "C:\\Legacy Vault Working Folder" },
    '''   "ProjectTracker": { "Path":          "C:\\Legacy Vault Working Folder\\Designs\\Project Tracker.xlsx" }
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

                Dim trackerPath As String = config("ProjectTracker:Path")
                If Not String.IsNullOrWhiteSpace(trackerPath) Then
                    result.ProjectTrackerPath = trackerPath
                Else
                    ' If only WorkingFolder was set, derive the tracker path from it.
                    result.ProjectTrackerPath = Path.Combine(result.VaultWorkingFolder, "Designs\Project Tracker.xlsx")
                End If
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
        ''' <remarks>
        ''' Uses the API pattern shown in <c>Metrol-Mechanical-Engineering/PDFExport</c>
        ''' (<c>Connection.WorkingFoldersManager.GetWorkingFolder("$").ToString()</c>).
        ''' </remarks>
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
