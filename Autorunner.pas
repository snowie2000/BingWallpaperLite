unit Autorunner;

interface
// You must add System.Win.Registry to your 'uses' clause.

uses
  SysUtils, Registry, Dialogs, Windows;

function SetAutoRun(const AppName, AppPath: string; AllUsers: Boolean = False): Boolean;

function RemoveAutoRun(const AppName: string; AllUsers: Boolean = False): Boolean;

function IsAutoRunSet(const AppName: string; AllUsers: Boolean = False): Boolean;

implementation

const
  // The registry path for startup applications.
  RUN_KEY_PATH = 'Software\Microsoft\Windows\CurrentVersion\Run';

// -----------------------------------------------------------------------------
// SetAutoRun
//
// Adds (or updates) an entry in the Windows startup registry to automatically
// run an application when a user logs in.
//
// Parameters:
//   AppName: The name for the registry value (e.g., 'MyAwesomeApp').
//            This should be unique to your application.
//   AppPath: The full path to the executable to run. Use quotes to handle
//            paths with spaces.
//   AllUsers: If False (default), adds to startup for the CURRENT USER only.
//             If True, adds to startup for ALL USERS on the machine.
//             NOTE: Setting for AllUsers requires Administrator privileges.
//
// Returns:
//   True on success, False on failure.
// -----------------------------------------------------------------------------
function SetAutoRun(const AppName, AppPath: string; AllUsers: Boolean = False): Boolean;
var
  Registry: TRegistry;
begin
  Result := False;
  Registry := TRegistry.Create;
  try
    // Select the root key based on whether this is for all users or just the current one.
    if AllUsers then
      Registry.RootKey := HKEY_LOCAL_MACHINE
    else
      Registry.RootKey := HKEY_CURRENT_USER;

    // Open the 'Run' key. The 'True' parameter will create the key if it doesn't exist.
    // This will return True on success.
    if Registry.OpenKey(RUN_KEY_PATH, True) then
    begin
      // Write the application path as a string value.
      // E.g., Key: 'MyAwesomeApp', Value: '"C:\Program Files\MyApp\MyApp.exe"'
      Registry.WriteString(AppName, AppPath);
      Registry.CloseKey;
      Result := True;
    end;
  except
    on E: Exception do
    begin
      // An exception here is often due to insufficient privileges,
      // especially when trying to write to HKEY_LOCAL_MACHINE.
      // You can log the error or show a message.
      // For this example, we just ensure Result remains False.
      Result := False;
    end;
  end;
  Registry.Free;
end;

// -----------------------------------------------------------------------------
// RemoveAutoRun
//
// Removes an application's entry from the Windows startup registry.
//
// Parameters:
//   AppName:  The name of the registry value to remove (must match what was
//             used in SetAutoRun).
//   AllUsers: If False (default), removes from CURRENT USER's startup.
//             If True, removes from ALL USERS' startup (requires Admin privileges).
//
// Returns:
//   True on success or if the value didn't exist, False on failure.
// -----------------------------------------------------------------------------
function RemoveAutoRun(const AppName: string; AllUsers: Boolean = False): Boolean;
var
  Registry: TRegistry;
begin
  Result := False;
  Registry := TRegistry.Create;
  try
    if AllUsers then
      Registry.RootKey := HKEY_LOCAL_MACHINE
    else
      Registry.RootKey := HKEY_CURRENT_USER;

    // Open the key, but don't create it if it doesn't exist.
    if Registry.OpenKey(RUN_KEY_PATH, False) then
    begin
      // Only try to delete if the value actually exists.
      if Registry.ValueExists(AppName) then
      begin
        Registry.DeleteValue(AppName);
      end;
      Registry.CloseKey;
      Result := True; // Success, whether it was deleted or didn't exist.
    end
    else
    begin
      // If the key doesn't even exist, our work is done.
      Result := True;
    end;
  except
    on E: Exception do
    begin
      Result := False;
    end;
  end;
  Registry.Free;
end;

// -----------------------------------------------------------------------------
// IsAutoRunSet
//
// Checks if an application entry exists in the Windows startup registry.
//
// Parameters:
//   AppName:  The name of the registry value to check for.
//   AllUsers: If False (default), checks in the CURRENT USER's startup.
//             If True, checks in the ALL USERS' startup.
//
// Returns:
//   True if the entry exists, False otherwise.
// -----------------------------------------------------------------------------
function IsAutoRunSet(const AppName: string; AllUsers: Boolean = False): Boolean;
var
  Registry: TRegistry;
begin
  Result := False;
  Registry := TRegistry.Create;
  try
    if AllUsers then
      Registry.RootKey := HKEY_LOCAL_MACHINE
    else
      Registry.RootKey := HKEY_CURRENT_USER;

    // Open the key in read-only mode, as we are not writing anything.
    if Registry.OpenKeyReadOnly(RUN_KEY_PATH) then
    begin
      // ValueExists is the perfect method for this check.
      Result := Registry.ValueExists(AppName);
      Registry.CloseKey;
    end;
  except
    on E: Exception do
    begin
      // On any error (e.g., permissions), assume it's not set.
      Result := False;
    end;
  end;
  Registry.Free;
end;

end.

