// ============================================================================
// MIL-Altium Team Scripts
// ============================================================================
// This file contains multiple script procedures that can be run independently.
// Each procedure appears as a separate script in DXP -> Run Script menu.
// ============================================================================

// ----------------------------------------------------------------------------
// CONFIGURATION CONSTANTS
// ----------------------------------------------------------------------------

const
  CONFIG_FILE = 'MIL-Altium-Config.ini';

// Get the path to the config file (stored in C:\Temp)
function GetConfigFilePath(): String;
begin
  // Store config in MIL-Altium repo
  Result := 'C:\MIL-Altium\Configuration\' + CONFIG_FILE;

  // Ensure the configuration directory exists
  if not DirectoryExists('C:\MIL-Altium\Configuration\') then
    CreateDir('C:\MIL-Altium\Configuration\');
end;

function LoadRepoPath(): String;
var
  ConfigFile : TextFile;
  FilePath : String;
  Line : String;
begin
  Result := '';
  FilePath := GetConfigFilePath();

  if not FileExists(FilePath) then
    Exit;

  try
    AssignFile(ConfigFile, FilePath);
    Reset(ConfigFile);

    while not EOF(ConfigFile) do
    begin
      ReadLn(ConfigFile, Line);

      if Pos('RepoPath=', Line) = 1 then
      begin
        Result := Copy(Line, 10, Length(Line) - 9);
        Break;
      end;
    end;

    CloseFile(ConfigFile);
  except
    Result := '';
  end;
end;

procedure SaveRepoPath(RepoPath: String);
var
  ConfigFile : TextFile;
  FilePath : String;
begin
  FilePath := GetConfigFilePath();

  try
    AssignFile(ConfigFile, FilePath);
    Rewrite(ConfigFile);
    WriteLn(ConfigFile, '[MIL-Altium]');
    WriteLn(ConfigFile, 'RepoPath=' + RepoPath);
    CloseFile(ConfigFile);
  except
    ShowMessage('Error: Could not save configuration file.');
  end;
end;

function IsRepoPathValid(): Boolean;
var
  RepoPath : String;
  TempPath : String;
  LastChar : String;
begin
  Result := False;
  RepoPath := LoadRepoPath();

  if RepoPath = '' then
    Exit;

  LastChar := Copy(RepoPath, Length(RepoPath), 1);
  if LastChar <> '\' then
    RepoPath := RepoPath + '\';

  TempPath := RepoPath;
  if not DirectoryExists(TempPath) then
    Exit;

  TempPath := RepoPath + 'Part Libraries\';
  if not DirectoryExists(TempPath) then
    Exit;

  Result := True;
end;

// ----------------------------------------------------------------------------
// SCRIPT 1: Initialize MIL-Altium Configuration
// Run this FIRST to set up the repository path
// ----------------------------------------------------------------------------

procedure InitializeMILAltiumConfig();
var
  RepoPath : String;
  OpenDialog : TOpenDialog;
  TempPath : String;
  LastChar : String;
  Response : Integer;
begin
  // Check if we already have a valid path
  if IsRepoPathValid() then
  begin
    RepoPath := LoadRepoPath();

    Response := MessageDlg('MIL-Altium repository path is already configured:' + Chr(13) + Chr(10) + Chr(13) + Chr(10) +
                           RepoPath + Chr(13) + Chr(10) + Chr(13) + Chr(10) +
                           'Do you want to change it?',
                           mtConfirmation, mbYesNo, 0);

    if Response = mrNo then
    begin
      ShowMessage('Configuration unchanged.');
      Exit;
    end;
  end;

  // Prompt user to select the repo folder
  OpenDialog := TOpenDialog.Create(nil);

  try
    OpenDialog.Title := 'Browse to MIL-Altium Repository and select Crowned-Schwartz.png';
    OpenDialog.Filter := 'All Files (*.*)|*.*';
    OpenDialog.InitialDir := 'C:\';

    ShowMessage('REPOSITORY CONFIGURATION:' + Chr(13) + Chr(10) + Chr(13) + Chr(10) +
                '1. Navigate to your MIL-Altium repository ROOT folder' + Chr(13) + Chr(10) +
                '2. Select Crowned-Schwartz.png' + Chr(13) + Chr(10) +
                '3. Click Open' + Chr(13) + Chr(10) + Chr(13) + Chr(10) +
                'This only needs to be done once!');

    if OpenDialog.Execute then
    begin
      TempPath := OpenDialog.FileName;
      RepoPath := ExtractFilePath(TempPath);
    end
    else
    begin
      ShowMessage('Configuration cancelled.');
      Exit;
    end;

  finally
    OpenDialog.Free;
  end;

  // Ensure path ends with backslash
  LastChar := Copy(RepoPath, Length(RepoPath), 1);
  if LastChar <> '\' then
    RepoPath := RepoPath + '\';

  // Validate the path
  TempPath := RepoPath + 'Part Libraries\';
  if not DirectoryExists(TempPath) then
  begin
    ShowMessage('Error: Part Libraries folder not found in:' + Chr(13) + Chr(10) +
                RepoPath + Chr(13) + Chr(10) + Chr(13) + Chr(10) +
                'Please ensure you selected a file in the MIL-Altium repository root.');
    Exit;
  end;

  // Save the path
  SaveRepoPath(RepoPath);

  ShowMessage('Configuration saved successfully!' + Chr(13) + Chr(10) + Chr(13) + Chr(10) +
              'Repository path: ' + RepoPath + Chr(13) + Chr(10) + Chr(13) + Chr(10) +
              'Config file location: ' + GetConfigFilePath() + Chr(13) + Chr(10) + Chr(13) + Chr(10) +
              'All MIL-Altium scripts will now use this location.');
end;

// ----------------------------------------------------------------------------
// SCRIPT 2: Import Multiple Libraries
// Imports all 19 part libraries from the configured repository
// ----------------------------------------------------------------------------

procedure ImportMultipleLibraries();
var
  PartLibrariesPath : String;
  LibraryPath : String;
  InstalledCount : Integer;
  FailedCount : Integer;
  LibraryPaths : Array[0..18] of String;
  i : Integer;
  LastChar : String;

begin
  LibraryPaths[0]  := 'Capacitors\Project Outputs for Capacitors\Capacitors.IntLib';
  LibraryPaths[1]  := 'Connectors\Project Outputs for Connectors\Connectors.IntLib';
  LibraryPaths[2]  := 'Crystals and Oscillators\Project Outputs for Crystals and Oscillators\Crystals and Oscillators.IntLib';
  LibraryPaths[3]  := 'Devices\Project Outputs for Devices\Devices.IntLib';
  LibraryPaths[4]  := 'Diodes\Project Outputs for Diodes\Diodes.IntLib';
  LibraryPaths[5]  := 'Ferrite Beads\Project Outputs for Ferrite Beads\Ferrite Beads.IntLib';
  LibraryPaths[6]  := 'Fuses\Project Outputs for Fuses\Fuses.IntLib';
  LibraryPaths[7]  := 'ICs\Project Outputs for ICs\ICs.IntLib';
  LibraryPaths[8]  := 'Inductors\Project Outputs for Inductors\Inductors.IntLib';
  LibraryPaths[9]  := 'LEDs\Project Outputs for LEDs\LEDs.IntLib';
  LibraryPaths[10] := 'MCUs\Project Outputs for MCUs\MCUs.IntLib';
  LibraryPaths[11] := 'Memory\Project Outputs for Memory\Memory.IntLib';
  LibraryPaths[12] := 'Misc Components\Project Outputs for Misc Components\Misc Components.IntLib';
  LibraryPaths[13] := 'Mounts, Pads, Vias\Project Outputs for Mounts, Pads, Vias\Mounts, Pads, Vias.IntLib';
  LibraryPaths[14] := 'Resistors\Project Outputs for Resistors\Resistors.IntLib';
  LibraryPaths[15] := 'Switches\Project Outputs for Switches\Switches.IntLib';
  LibraryPaths[16] := 'Test Points\Project Outputs for Test Points\Test Points.IntLib';
  LibraryPaths[17] := 'Transistors\Project Outputs for Transistors\Transistors.IntLib';
  LibraryPaths[18] := 'Voltage Regulators\Project Outputs for Voltage Regulators\Voltage Regulators.IntLib';

  // Check if repo path is configured
  if not IsRepoPathValid() then
  begin
    ShowMessage('MIL-Altium repository path is not configured.' + Chr(13) + Chr(10) + Chr(13) + Chr(10) +
                'Please run the "InitializeMILAltiumConfig" script first.');
    Exit;
  end;

  // Load the repo path from config
  PartLibrariesPath := LoadRepoPath();

  // Ensure path ends with backslash
  LastChar := Copy(PartLibrariesPath, Length(PartLibrariesPath), 1);
  if LastChar <> '\' then
    PartLibrariesPath := PartLibrariesPath + '\';

  // Append Part Libraries folder
  PartLibrariesPath := PartLibrariesPath + 'Part Libraries\';

  InstalledCount := 0;
  FailedCount := 0;

  ShowMessage('Starting library import...' + Chr(13) + Chr(10) +
              'Part Libraries folder: ' + PartLibrariesPath + Chr(13) + Chr(10) +
              'Installing 19 libraries...');

  for i := 0 to 18 do
  begin
    LibraryPath := PartLibrariesPath + LibraryPaths[i];

    if FileExists(LibraryPath) then
    begin
      try
        IntegratedLibraryManager.InstallLibrary(LibraryPath);
        Inc(InstalledCount);
      except
        Inc(FailedCount);
      end;
    end
    else
    begin
      Inc(FailedCount);
    end;
  end;

  RunProcess('IntegratedLibrary:RefreshInstalledLibraries');

  if InstalledCount > 0 then
    ShowMessage('Library import complete!' + Chr(13) + Chr(10) +
                'Successfully installed: ' + IntToStr(InstalledCount) + ' libraries' + Chr(13) + Chr(10) +
                'Failed: ' + IntToStr(FailedCount) + ' libraries')
  else
    ShowMessage('Library import failed!' + Chr(13) + Chr(10) +
                'No libraries were installed.' + Chr(13) + Chr(10) +
                'Failed: ' + IntToStr(FailedCount) + ' libraries');
end;


