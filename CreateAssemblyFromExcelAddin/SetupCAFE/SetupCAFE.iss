; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define MyAppName "Create Assembly From Excel"
#define MyAppVersion "1.0"
#define MyAppPublisher "Matrix Precision Engineering Ltd."
#define MyAppURL "http://www.matrixprecision.com/"

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{38885C3D-3F74-4788-80AD-D022318173E6}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName=C:\Program Files\Autodesk\Inventor 2014\Bin\CreateAssemblyFromExcelAddin
DisableDirPage=yes
DefaultGroupName=Create Application From Excel Addin
DisableProgramGroupPage=yes
OutputDir=C:\Users\alex\Documents\Visual Studio 2012\Projects\CreateAssemblyFromExcelAddin\CreateAssemblyFromExcelAddin\SetupCAFE
OutputBaseFilename=setupCAFE
SetupIconFile=C:\Users\alex\Documents\Visual Studio 2012\Projects\CreateAssemblyFromExcelAddin\CreateAssemblyFromExcelAddin\Icons8-Windows-8-Food-Cafe.ico
Compression=lzma
SolidCompression=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "C:\Program Files\Autodesk\Inventor 2014\Bin\CreateAssemblyFromExcelAddin\CreateAssemblyFromExcelAddin.pdb"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Program Files\Autodesk\Inventor 2014\Bin\CreateAssemblyFromExcelAddin\CreateAssemblyFromExcelAddin.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\alex\Documents\Visual Studio 2012\Projects\CreateAssemblyFromExcelAddin\CreateAssemblyFromExcelAddin\Autodesk.CreateAssemblyFromExcelAddin.Inventor.addin"; DestDir: "C:\ProgramData\Autodesk\Inventor 2014\Addins\"; Flags: ignoreversion
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Icons]
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"

