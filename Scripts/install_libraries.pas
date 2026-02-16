procedure ImportMultipleLibraries();
var
  // Path to the MIL-Altium repository root (user will provide this)
  MILAltiumRepoRoot : String;

  // Full path to each library file
  LibraryPath : String;

  // Counter for successfully installed libraries
  InstalledCount : Integer;

  // Counter for failed installations
  FailedCount : Integer;

  // Array of library paths relative to Part Libraries folder
  LibraryPaths : Array[0..14] of String;

  // Loop counter
  i : Integer;

  // Folder selection dialog
  OpenDialog : TOpenDialog;

  // Temporary string for path manipulation
  TempPath : String;
  LastChar : String;

begin
  // Define all library paths relative to the Part Libraries folder
  // Each path follows the pattern: FolderName/Project Outputs for FolderName/LibraryName.IntLib
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

  // Create a file open dialog for folder selection
  OpenDialog := TOpenDialog.Create(nil);

  try
    // Configure the dialog to select folders (not files)
    OpenDialog.Title := 'Select MIL-Altium Repository Folder';
    OpenDialog.Filter := 'All Files (*.*)|*.*';
    OpenDialog.InitialDir := 'C:\';

    // Show a message to guide the user
    ShowMessage('FOLDER SELECTION INSTRUCTIONS:' + Chr(13) + Chr(10) + Chr(13) + Chr(10) +
                '1. In the next dialog, navigate to your MIL-Altium repository' + Chr(13) + Chr(10) +
                '2. Select "Crowned-Schwartz.png"' + Chr(13) + Chr(10) +
                '3. Click Open' + Chr(13) + Chr(10) + Chr(13) + Chr(10) +
                'This script will import the MIL Altium libraries.');


    // Show the dialog and check if user clicked OK
    if OpenDialog.Execute then
    begin
      // Extract the directory path from the selected file
      TempPath := OpenDialog.FileName;
      MILAltiumRepoRoot := ExtractFilePath(TempPath);
    end
    else
    begin
      // User cancelled the dialog
      ShowMessage('Library import cancelled.');
      Exit;
    end;

  finally
    // Clean up the dialog object
    OpenDialog.Free;
  end;

  // Ensure the path ends with a backslash for consistency
  // Get the last character of the path
  LastChar := Copy(MILAltiumRepoRoot, Length(MILAltiumRepoRoot), 1);
  if LastChar <> '\' then
    MILAltiumRepoRoot := MILAltiumRepoRoot + '\';

  // Verify that the repository root exists
  TempPath := MILAltiumRepoRoot;
  if not DirectoryExists(TempPath) then
  begin
    ShowMessage('Error: Directory not found:' + Chr(13) + Chr(10) +
                MILAltiumRepoRoot + Chr(13) + Chr(10) + Chr(13) + Chr(10) +
                'Please check the path and try again.');
    Exit;
  end;

  // Verify that the Part Libraries folder exists
  TempPath := MILAltiumRepoRoot + 'Part Libraries\';
  if not DirectoryExists(TempPath) then
  begin
    ShowMessage('Error: Part Libraries folder not found in:' + Chr(13) + Chr(10) +
                MILAltiumRepoRoot + Chr(13) + Chr(10) + Chr(13) + Chr(10) +
                'Please ensure you selected a file in the MIL-Altium repository root folder.');
    Exit;
  end;

  // Initialize the counters
  InstalledCount := 0;
  FailedCount := 0;

  // Show a status message to the user
  ShowMessage('Starting library import...' + Chr(13) + Chr(10) +
              'MIL-Altium repository: ' + MILAltiumRepoRoot + Chr(13) + Chr(10) +
              'Installing 15 libraries...');

  // Loop through each library path in the array
  for i := 0 to 14 do
  begin
    // Construct the full absolute path to the library file
    // Format: MILAltiumRepoRoot + Part Libraries\ + RelativePath
    LibraryPath := MILAltiumRepoRoot + 'Part Libraries\' + LibraryPaths[i];

    // Check if the library file exists before attempting to install
    if FileExists(LibraryPath) then
    begin
      // Attempt to install the library
      try
        // Install the library into the workspace
        IntegratedLibraryManager.InstallLibrary(LibraryPath);

        // Increment the success counter
        Inc(InstalledCount);

      except
        // If installation fails, increment the failure counter
        // This allows the script to continue with remaining libraries
        Inc(FailedCount);
      end;
    end
    else
    begin
      // If the file doesn't exist, increment the failure counter
      Inc(FailedCount);
    end;
  end;

  // Refresh the installed libraries list in Altium
  // This updates the UI to show the newly installed libraries
  RunProcess('IntegratedLibrary:RefreshInstalledLibraries');

  // Show completion message with statistics
  if InstalledCount > 0 then
    ShowMessage('Library import complete!' + Chr(13) + Chr(10) +
                'Successfully installed: ' + IntToStr(InstalledCount) + ' libraries' + Chr(13) + Chr(10) +
                'Failed: ' + IntToStr(FailedCount) + ' libraries')
  else
    ShowMessage('Library import failed!' + Chr(13) + Chr(10) +
                'No libraries were installed.' + Chr(13) + Chr(10) +
                'Failed: ' + IntToStr(FailedCount) + ' libraries' + Chr(13) + Chr(10) +
                'Repository location: ' + MILAltiumRepoRoot);
end;
