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
  // Store config in C:\Temp (should exist on all Windows systems)
  Result := 'C:\MIL-Altium\' + CONFIG_FILE;

  // Ensure the Temp directory exists
  if not DirectoryExists('C:\MIL-Altium') then
    CreateDir('C:\MIL-Altium');
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
    OpenDialog.Title := 'Browse to MIL-Altium Repository and Select ANY File';
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
// Imports all 15 part libraries from the configured repository
// ----------------------------------------------------------------------------

procedure ImportMultipleLibraries();
var
  PartLibrariesPath : String;
  LibraryPath : String;
  InstalledCount : Integer;
  FailedCount : Integer;
  LibraryPaths : Array[0..14] of String;
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
  LibraryPaths[11] := 'Misc Components\Project Outputs for Misc Components\Misc Components.IntLib';
  LibraryPaths[12] := 'Resistors\Project Outputs for Resistors\Resistors.IntLib';
  LibraryPaths[13] := 'Switches\Project Outputs for Switches\Switches.IntLib';
  LibraryPaths[14] := 'Voltage Regulators\Project Outputs for Voltage Regulators\Voltage Regulators.IntLib';

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
              'Installing 15 libraries...');

  for i := 0 to 14 do
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

// ----------------------------------------------------------------------------
// SCRIPT 3: Import JLC Rules
// Imports PCB design rules from the configured repository
// ----------------------------------------------------------------------------

procedure ImportJLCRules();
var
  RepoPath : String;
  RulesFilePath : String;
  Board : IPCB_Board;
  LastChar : String;
  Response : Integer;

begin
  Board := PCBServer.GetCurrentPCBBoard;

  if Board = nil then
  begin
    ShowMessage('Error: No PCB document is currently open.' + Chr(13) + Chr(10) +
                'Please open a PCB file first.');
    Exit;
  end;

  // Check if repo path is configured
  if not IsRepoPathValid() then
  begin
    ShowMessage('MIL-Altium repository path is not configured.' + Chr(13) + Chr(10) + Chr(13) + Chr(10) +
                'Please run the "InitializeMILAltiumConfig" script first.');
    Exit;
  end;

  // Load the repo path
  RepoPath := LoadRepoPath();

  // Ensure path ends with backslash
  LastChar := Copy(RepoPath, Length(RepoPath), 1);
  if LastChar <> '\' then
    RepoPath := RepoPath + '\';

  // Construct path to rules file
  RulesFilePath := RepoPath + 'Rules and Stackups\JLC Rules.RUL';

  if not FileExists(RulesFilePath) then
  begin
    ShowMessage('Error: JLC Rules.RUL file not found at:' + Chr(13) + Chr(10) +
                RulesFilePath);
    Exit;
  end;

  Response := MessageDlg('This will OVERWRITE all existing design rules in the current PCB.' + Chr(13) + Chr(10) + Chr(13) + Chr(10) +
                         'Rules file: JLC Rules.RUL' + Chr(13) + Chr(10) + Chr(13) + Chr(10) +
                         'Do you want to continue?',
                         mtWarning, mbYesNo, 0);

  if Response = mrNo then
  begin
    ShowMessage('Rules import cancelled.');
    Exit;
  end;

  try
    // Use the RunProcess command to import rules
    // This is the same as manually going to Design -> Rules -> Import Rules
    RunProcess('PCB:ImportRules|FileName=' + RulesFilePath);

    // Refresh the PCB display
    Board.GraphicallyInvalidate;

    ShowMessage('Design rules imported successfully!' + Chr(13) + Chr(10) + Chr(13) + Chr(10) +
                'JLC Rules have been applied to the current PCB.');
  except
    ShowMessage('Error: Failed to import design rules.');
  end;
end;

// ----------------------------------------------------------------------------
// SCRIPT 4: Import Board Stackup
// Imports a 2-layer or 4-layer stackup from the configured repository
// ----------------------------------------------------------------------------

procedure ImportBoardStackup();
var
  RepoPath : String;
  StackupFilePath : String;
  Board : IPCB_Board;
  LastChar : String;
  Response : Integer;
  StackupFolder : String;

begin
  Board := PCBServer.GetCurrentPCBBoard;

  // Check if a PCB is open
  if Board = nil then
  begin
    ShowMessage('Error: No PCB document is currently open.' + Chr(13) + Chr(10) +
                'Please open a PCB file first.');
    Exit;
  end;

  // Check if repo path is configured
  if not IsRepoPathValid() then
  begin
    ShowMessage('MIL-Altium repository path is not configured.' + Chr(13) + Chr(10) + Chr(13) + Chr(10) +
                'Please run the "InitializeMILAltiumConfig" script first.');
    Exit;
  end;

  // Load the repo path
  RepoPath := LoadRepoPath();

  // Ensure path ends with backslash
  LastChar := Copy(RepoPath, Length(RepoPath), 1);
  if LastChar <> '\' then
    RepoPath := RepoPath + '\';

  // Construct path to the stackup folder
  StackupFolder := RepoPath + 'Rules and Stackups\';

  // Verify that the Rules and Stackups folder exists
  if not DirectoryExists(StackupFolder) then
  begin
    ShowMessage('Error: Rules and Stackups folder not found at:' + Chr(13) + Chr(10) +
                StackupFolder + Chr(13) + Chr(10) + Chr(13) + Chr(10) +
                'Please ensure the MIL-Altium repository is configured correctly.');
    Exit;
  end;

  // Ask the user which stackup they want to import
  Response := MessageDlg('Which board stackup would you like to import?' + Chr(13) + Chr(10) + Chr(13) + Chr(10) +
                         'Yes = 2-Layer Stackup' + Chr(13) + Chr(10) +
                         'No  = 4-Layer Stackup',
                         mtConfirmation, mbYesNo, 0);

  // Set the stackup file path based on the user's choice
  if Response = mrYes then
  begin
    // User chose 2-layer stackup
    StackupFilePath := StackupFolder + '2-Layer.stackup';
  end
  else
  begin
    // User chose 4-layer stackup
    StackupFilePath := StackupFolder + '4-Layer.stackup';
  end;

  // Verify that the selected stackup file exists
  if not FileExists(StackupFilePath) then
  begin
    ShowMessage('Error: Stackup file not found at:' + Chr(13) + Chr(10) +
                StackupFilePath + Chr(13) + Chr(10) + Chr(13) + Chr(10) +
                'Please ensure the stackup file exists in the Rules and Stackups folder.' + Chr(13) + Chr(10) + Chr(13) + Chr(10) +
                'Expected filename: ' + ExtractFileName(StackupFilePath));
    Exit;
  end;

  // Confirm before overwriting existing stackup
  Response := MessageDlg('This will OVERWRITE the existing board stackup in the current PCB.' + Chr(13) + Chr(10) + Chr(13) + Chr(10) +
                         'Stackup file: ' + ExtractFileName(StackupFilePath) + Chr(13) + Chr(10) + Chr(13) + Chr(10) +
                         'Do you want to continue?',
                         mtWarning, mbYesNo, 0);

  if Response = mrNo then
  begin
    ShowMessage('Stackup import cancelled.');
    Exit;
  end;

  // Import the stackup file
  try
    // Open the Layer Stack Manager
    RunProcess('LayerStackManager:OpenPcbLayerStack');

    // Tell the user exactly where to find the file
    ShowMessage('The Layer Stack Manager is now open.' + Chr(13) + Chr(10) + Chr(13) + Chr(10) +
                'When the file dialog opens, navigate to:' + Chr(13) + Chr(10) + Chr(13) + Chr(10) +
                StackupFilePath + Chr(13) + Chr(10) + Chr(13) + Chr(10) +
                'Click OK to open the file dialog now.');

    // Open the Load Template file dialog
    RunProcess('LayerStackManager:LoadTemplateFromFile');

    // Refresh the PCB display to show the new stackup
    Board.GraphicallyInvalidate;

    ShowMessage('Stackup import complete!');
  except
    ShowMessage('Error: Failed to import board stackup.');
  end;
end;
