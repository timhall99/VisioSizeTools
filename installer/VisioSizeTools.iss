; Inno Setup script for the Visio Size Tools COM add-in.
; CI compiles it with:  ISCC /DAppVersion=x.y.z /DSourceDir=<build output> /DOutputDir=<dir> installer\VisioSizeTools.iss
; The defaults below apply when compiling by hand after:  dotnet build VisioSizeTools.sln -c Release
;
; Installs per user (no admin rights) and registers the add-in under HKCU, exactly like scripts\Register.ps1.
; HKCU\Software\Classes is shared by 32- and 64-bit Office, so one set of keys works for either bitness.

#ifndef SourceDir
  #define SourceDir "..\src\VisioSizeTools\bin\Release\net48"
#endif
#ifndef AppVersion
  #define AppVersion GetVersionNumbersString(SourceDir + "\VisioSizeTools.dll")
#endif
#ifndef OutputDir
  #define OutputDir "..\artifacts"
#endif

#define AppName      "Visio Size Tools"
; "{{" escapes the brace inside Inno constants-aware sections
#define ClsId        "{{4A11EACB-6B73-4915-AD54-D92ECF41F340}"
#define ProgId       "VisioSizeTools.Connect"
#define ClassName    "VisioSizeTools.Connect"
; Must match the DLL's AssemblyName.FullName (AssemblyVersion is the 4-part file version from the SDK build).
#define AsmVersion   GetVersionNumbersString(SourceDir + "\VisioSizeTools.dll")
#define AsmFullName  "VisioSizeTools, Version=" + AsmVersion + ", Culture=neutral, PublicKeyToken=null"
#define ClassesKey   "Software\Classes"

[Setup]
AppId={{9C1E2B7D-4F3A-4E8B-9A61-2D5C7F0B8E14}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher=macmedia.tv
AppCopyright=Copyright © macmedia.tv
VersionInfoVersion={#AsmVersion}
PrivilegesRequired=lowest
DefaultDirName={localappdata}\Programs\VisioSizeTools
DisableDirPage=yes
DisableProgramGroupPage=yes
OutputDir={#OutputDir}
OutputBaseFilename=VisioSizeTools_Setup_{#AppVersion}
UninstallDisplayName={#AppName}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
; Visio locks the DLL while it is running; let Restart Manager offer to close it.
CloseApplications=yes
RestartApplications=no

[Files]
Source: "{#SourceDir}\VisioSizeTools.dll"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
; --- COM class registration (what regasm /codebase would write, but per user) ---
Root: HKCU; Subkey: "{#ClassesKey}\CLSID\{#ClsId}"; ValueType: string; ValueData: "{#ClassName}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "{#ClassesKey}\CLSID\{#ClsId}\ProgId"; ValueType: string; ValueData: "{#ProgId}"
Root: HKCU; Subkey: "{#ClassesKey}\CLSID\{#ClsId}\Implemented Categories\{{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"
Root: HKCU; Subkey: "{#ClassesKey}\CLSID\{#ClsId}\InprocServer32"; ValueType: string; ValueData: "mscoree.dll"
Root: HKCU; Subkey: "{#ClassesKey}\CLSID\{#ClsId}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"
Root: HKCU; Subkey: "{#ClassesKey}\CLSID\{#ClsId}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "{#ClassName}"
Root: HKCU; Subkey: "{#ClassesKey}\CLSID\{#ClsId}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AsmFullName}"
Root: HKCU; Subkey: "{#ClassesKey}\CLSID\{#ClsId}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU; Subkey: "{#ClassesKey}\CLSID\{#ClsId}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetCodeBase}"
Root: HKCU; Subkey: "{#ClassesKey}\{#ProgId}"; ValueType: string; ValueData: "{#ClassName}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "{#ClassesKey}\{#ProgId}\CLSID"; ValueType: string; ValueData: "{#ClsId}"

; --- Tell Visio to load the add-in at startup ---
Root: HKCU; Subkey: "Software\Microsoft\Visio\Addins\{#ProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AppName}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Visio\Addins\{#ProgId}"; ValueType: string; ValueName: "Description"; ValueData: "Size to Tallest / Shortest / Both Larger / Both Smaller on Home > Arrange."
Root: HKCU; Subkey: "Software\Microsoft\Visio\Addins\{#ProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"

[Messages]
FinishedLabel=Setup has installed [name] on your computer. Restart Visio to load the add-in.

[Code]
{ Unescaped file URI, as regasm writes it - "%20" for spaces breaks the CLR loader. }
function GetCodeBase(Param: string): string;
begin
  Result := ExpandConstant('{app}\VisioSizeTools.dll');
  StringChangeEx(Result, '\', '/', True);
  Result := 'file:///' + Result;
end;
