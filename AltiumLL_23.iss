[Setup]
AppName=Altium Library Loader
AppVersion=2.3.0
AppPublisher=SamacSys / unofficial patch by @lethargicpillow
AppPublisherURL=https://supplyframe.com/samacsys/
AppSupportURL=https://supplyframe.com/samacsys/
AppUpdatesURL=https://supplyframe.com/samacsys/
DefaultDirName={userdocs}\AltiumLL
DisableProgramGroupPage=yes
OutputDir=.
OutputBaseFilename=AltiumLL_2.3_Setup
Compression=lzma
SolidCompression=yes
PrivilegesRequired=lowest
SetupIconFile=userdocs\AltiumLL\SamacSys.ico
UninstallDisplayIcon={userdocs}\AltiumLL\SamacSys.ico
UninstallDisplayName=Altium Library Loader
VersionInfoVersion=2.3.0
VersionInfoCompany=SamacSys / unofficial patch by @lethargicpillow
VersionInfoDescription=Altium Library Loader
VersionInfoProductName=Altium Library Loader
WizardSmallImageFile=tmp\WizModernSmallImage.bmp
WizardImageFile=tmp\AltiumLL.bmp
DisableFinishedPage=yes

[Files]
Source: "userdocs\AltiumLL\AltiumLL.dfm"; DestDir: "{userdocs}\AltiumLL"; Flags: ignoreversion
Source: "userdocs\AltiumLL\AltiumLL.PrjScr"; DestDir: "{userdocs}\AltiumLL"; Flags: ignoreversion
Source: "userdocs\AltiumLL\AltiumLL.vbs"; DestDir: "{userdocs}\AltiumLL"; Flags: ignoreversion
Source: "userdocs\AltiumLL\SamacSys.ico"; DestDir: "{userdocs}\AltiumLL"; Flags: ignoreversion
Source: "userdocs\AltiumLL\SamacSys.PcbLib"; DestDir: "{userdocs}\AltiumLL"; Flags: ignoreversion
Source: "userdocs\AltiumLL\SamacSys.png"; DestDir: "{userdocs}\AltiumLL"; Flags: ignoreversion
Source: "userdocs\AltiumLL\SamacSys.SchLib"; DestDir: "{userdocs}\AltiumLL"; Flags: ignoreversion
Source: "userdocs\AltiumLL\LOAD_3D.png"; DestDir: "{userdocs}\AltiumLL"; Flags: ignoreversion
Source: "userdocs\AltiumLL\NO_3D.png"; DestDir: "{userdocs}\AltiumLL"; Flags: ignoreversion
Source: "userdocs\AltiumLL\SamacSys_Images\*"; DestDir: "{userdocs}\AltiumLL\SamacSys_Images"; Flags: ignoreversion recursesubdirs skipifsourcedoesntexist
Source: "userdocs\AltiumLL\Temp\*"; DestDir: "{userdocs}\AltiumLL\Temp"; Flags: ignoreversion recursesubdirs skipifsourcedoesntexist
Source: "tmp\AltiumLL.bmp"; Flags: dontcopy

[Dirs]
Name: "{userdocs}\AltiumLL\Temp"
Name: "{userdocs}\AltiumLL\SamacSys_Images"

[Code]

function FindAltiumProfile(): String;
var
  AltiumDir: String;
  FindRec: TFindRec;
begin
  Result := '';
  AltiumDir := ExpandConstant('{userappdata}\Altium');

  if FindFirst(AltiumDir + '\Altium Designer {*} ', FindRec) then begin
    Result := AltiumDir + '\' + FindRec.Name;
    FindClose(FindRec);
  end;

  if Result = '' then begin
    if FindFirst(AltiumDir + '\Altium Designer {*}', FindRec) then begin
      Result := AltiumDir + '\' + FindRec.Name;
      FindClose(FindRec);
    end;
  end;
end;

procedure PatchDXPRCS();
var
  ProfileDir: String;
  RcsFile: String;
  Content: String;
  Block: String;
  ProjectPath: String;
  ImagePath: String;
  Lines: TStringList;
begin
  ProfileDir := FindAltiumProfile();

  if ProfileDir = '' then begin
    MsgBox('Altium profile was not found in AppData. Run Altium once, then run this installer again.', mbInformation, MB_OK);
    exit;
  end;

  RcsFile := ProfileDir + '\DXP.RCS';
  ProjectPath := ExpandConstant('{userdocs}\AltiumLL\AltiumLL.PrjScr');
  ImagePath := ExpandConstant('{userdocs}\AltiumLL\SamacSys.png');

  if FileExists(RcsFile) then begin
    Lines := TStringList.Create;
    try
      Lines.LoadFromFile(RcsFile);
      Content := Lines.Text;
    finally
      Lines.Free;
    end;
  end else begin
    Content := '';
  end;

  if Pos('PL AltiumLL Command=', Content) > 0 then
    exit;

  Block :=
    'Tree EditScriptVBSCustom Caption=''[Custom]'' TopLevel=''True''  End'#13#10 +
    'PL AltiumLL Command=''ScriptingSystem:RunScript'' Params=''ProjectName=' + ProjectPath + '|ProcName=AltiumLL.vbs>Prechecks'' Caption=''Symbols | Footprints | 3D Models'' Image=''' + ImagePath + ''' DefaultChecked=0  End'#13#10 +
    'Insertion _User1 TargetID=''MNNoDocument_File'' RefID0=''NoDoc'''#13#10 +
    '    Link _User2 PLID=''AltiumLL''      End'#13#10 +
    'End'#13#10 +
    'Insertion _User3 TargetID=''MNSchematic_File'' RefID0=''SchDoc_File'''#13#10 +
    '    Link _User4 PLID=''AltiumLL''      End'#13#10 +
    'End'#13#10 +
    'Insertion _User5 TargetID=''MNSchematic_Tools10'' RefID0=''SchDoc_Tools'''#13#10 +
    '    Link _User6 PLID=''AltiumLL''      End'#13#10 +
    'End'#13#10 +
    'Insertion _User7 TargetID=''MNSchematic_SchLibMenu10'' RefID0=''SchLib_File'''#13#10 +
    '    Link _User8 PLID=''AltiumLL''      End'#13#10 +
    'End'#13#10 +
    'Insertion _User9 TargetID=''MNSchematic_SchLibMenuTools10'' RefID0=''SchLib_Tools'''#13#10 +
    '    Link _User10 PLID=''AltiumLL''      End'#13#10 +
    'End'#13#10 +
    'Insertion _User11 TargetID=''MNPCBLib_10'' RefID0=''PcbLib_File'''#13#10 +
    '    Link _User12 PLID=''AltiumLL''      End'#13#10 +
    'End'#13#10 +
    'Insertion _User13 TargetID=''MNPCBLib_Tools10'' RefID0=''PcbLib_Tools'''#13#10 +
    '    Link _User14 PLID=''AltiumLL''      End'#13#10 +
    'End'#13#10 +
    'Insertion _User15 TargetID=''MNPCB_File'' RefID0=''PcbDoc_File'''#13#10 +
    '    Link _User16 PLID=''AltiumLL''      End'#13#10 +
    'End'#13#10 +
    'Insertion _User17 TargetID=''MNPCB_Tools10'' RefID0=''PcbDoc_Tools'''#13#10 +
    '    Link _User18 PLID=''AltiumLL''      End'#13#10 +
    'End'#13#10 +
    'Insertion _User19 TargetID=''MNSchematicMenu'' InsertType=''After'' RefID0=''MNSchematic_Help10'''#13#10 +
    '    Link _User20 PLID=''AltiumLL''      End'#13#10 +
    'End'#13#10 +
    'Insertion _User21 TargetID=''MNPCBMenu'' InsertType=''After'' RefID0=''MNPCB_Help10'''#13#10 +
    '    Link _User22 PLID=''AltiumLL''      End'#13#10 +
    'End'#13#10#13#10;

  SaveStringToFile(RcsFile + '.bak', Content, False);
  SaveStringToFile(RcsFile, Block + Content, False);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    PatchDXPRCS();
end;

var
  FinalPage: TWizardPage;
  FinalImage: TBitmapImage;
  FinalText: TNewStaticText;

procedure InitializeWizard();
begin
  FinalPage := CreateCustomPage(
    wpInstalling,
    'Installation Complete',
    'Altium Library Loader is ready to use'
  );

  ExtractTemporaryFile('AltiumLL.bmp');

  FinalImage := TBitmapImage.Create(FinalPage);
  FinalImage.Parent := FinalPage.Surface;
  FinalImage.Bitmap.LoadFromFile(ExpandConstant('{tmp}\AltiumLL.bmp'));
  FinalImage.AutoSize := True;

  FinalImage.Left := (FinalPage.SurfaceWidth - FinalImage.Width) div 2;
  FinalImage.Top := 8;

  FinalText := TNewStaticText.Create(FinalPage);
  FinalText.Parent := FinalPage.Surface;
  FinalText.Caption :=
    'The Library Loader button has been added to the Altium Designer toolbar.'#13#10#13#10 +
    'Open a schematic document (.SchDoc) or PCB document (.PcbDoc).'#13#10 +
    'The button will appear in the top-right toolbar when a supported document is active.';
  FinalText.Left := 16;
  FinalText.Top := FinalImage.Top + FinalImage.Height + 16;
  FinalText.Width := FinalPage.SurfaceWidth - 32;
  FinalText.WordWrap := True;
end;

procedure CurPageChanged(CurPageID: Integer);
begin
  if CurPageID = FinalPage.ID then
    WizardForm.NextButton.Caption := SetupMessage(msgButtonFinish);
end;