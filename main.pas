unit main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms, Dialogs, CoolTrayIcon, Menus, ExtCtrls,
  IniFiles, uBWHandler, IdBaseComponent, IdAntiFreezeBase, IdAntiFreeze, ActiveX, ShlObj, ShObjIdl, ShellAPI, DateUtils;

type
  TfrmMain = class(TForm)
    trayIcon1: TCoolTrayIcon;
    pm1: TPopupMenu;
    Refresheveryday1: TMenuItem;
    StartwithWindows1: TMenuItem;
    N1: TMenuItem;
    Quit1: TMenuItem;
    tmr1: TTimer;
    idntfrz1: TIdAntiFreeze;
    UpdateNow1: TMenuItem;
    N2: TMenuItem;
    Openwallpaperfolder1: TMenuItem;
    Region1: TMenuItem;
    procedure Quit1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure tmr1Timer(Sender: TObject);
    procedure UpdateNow1Click(Sender: TObject);
    procedure Refresheveryday1Click(Sender: TObject);
    procedure StartwithWindows1Click(Sender: TObject);
    procedure Openwallpaperfolder1Click(Sender: TObject);
  private
    { Private declarations }
    FConfig: TIniFile;
    FNextUpdate: TDateTime;
    FBWHandler: TBingWallpaperHandler;
    FWallPaperFolder: string;
    FRegionMap: TStringList;
    procedure onRegionMenuClick(Sender: TObject);
    procedure regionMenuSetup();
    procedure uiSetup();
    procedure updateWallpaper(force: Boolean = false);
    procedure setWallpaper(filename: string);
    function is4kMonitor(): Boolean;
    procedure SetNextUpdate(const Value: TDateTime);
    property NextUpdate: TDateTime read FNextUpdate write SetNextUpdate;
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

uses
  Autorunner;

{$R *.dfm}

function NowUTC: TDateTime;
var
  system_datetime: TSystemTime;
begin
  GetSystemTime(system_datetime);
  Result := SystemTimeToDateTime(system_datetime);
end;

function OpenDirectoryInExplorer(const ADirectory: string): Boolean;
var
  ResultCode: Integer;
begin
  Result := False;
  if not DirectoryExists(ADirectory) then
  begin
    // The directory doesn't exist, so we can't open it.
    Exit;
  end;

  ResultCode := ShellExecute(0, 'open', PChar(ADirectory), nil, nil, SW_SHOWNORMAL);

  Result := (ResultCode > 32);
end;

procedure TfrmMain.FormCreate(Sender: TObject);
begin
  FRegionMap := TStringList.Create;
  FConfig := TIniFile.Create(ExtractFilePath(Application.ExeName) + 'config.ini');
  FBWHandler := TBingWallpaperHandler.Create;
  uiSetup();
end;

procedure TfrmMain.FormDestroy(Sender: TObject);
begin
  FConfig.Free;
  FRegionMap.Free;
end;

function TfrmMain.is4kMonitor: Boolean;
begin
  Result := (Monitor.Height > 1080) or (Monitor.Width > 1920);
end;

procedure TfrmMain.onRegionMenuClick(Sender: TObject);
var
  code: string;
  menu: TMenuItem;
begin
  if Sender is TMenuItem then
  begin
    menu := Sender as TMenuItem;
    if menu.Checked then
      Exit;
    code := FRegionMap[menu.Tag];
    with Region1.GetEnumerator do
    try
      while MoveNext do
        Current.Checked := False;   // uncheck all regions
    finally
      Free;
    end;
    menu.Checked := True;    // check me
    FBWHandler.Region := code;
    FConfig.WriteString('general', 'region', code);
    updateWallpaper(True);
  end;
end;

procedure TfrmMain.Openwallpaperfolder1Click(Sender: TObject);
begin
  OpenDirectoryInExplorer(FWallPaperFolder);
end;

procedure TfrmMain.Quit1Click(Sender: TObject);
begin
  Application.Terminate;
end;

procedure TfrmMain.Refresheveryday1Click(Sender: TObject);
begin
  Refresheveryday1.Checked := not Refresheveryday1.Checked;
  tmr1.Enabled := Refresheveryday1.Checked;
  FConfig.WriteBool('general', 'refresh', Refresheveryday1.Checked);
end;

procedure TfrmMain.regionMenuSetup;
var
  menu: TMenuItem;
  i: Integer;
const
  country: array[0..9] of string = ('United States', 'China', 'Japan', 'Germany', 'United Kingdom', 'France',
    'Australia', 'Canada', 'Brazil', 'India');
  countryCode: array[0..9] of string = ('en-US', 'zh-CN', 'ja-JP', 'de-DE', 'en-GB', 'fr-FR', 'en-AU', 'en-CA', 'pt-BR', 'en-IN');
begin
  for i := low(country) to high(country) do
  begin
    menu := TMenuItem.Create(self);
    menu.Caption := country[i];
    menu.Tag := i;
    menu.OnClick := onRegionMenuClick;
    FRegionMap.Add(countryCode[i]);
    Region1.Add(menu);
  end;
end;

procedure TfrmMain.SetNextUpdate(const Value: TDateTime);
begin
  FNextUpdate := Value;
  FConfig.WriteInteger('general', 'nextupdate', DateTimeToUnix(Value) div 10);
end;

procedure TfrmMain.setWallpaper(filename: string);
var
  desktopWallpaper: IDesktopWallpaper;
  ResultCode: HRESULT;
begin
  // In a real application, CoInitialize/CoUninitialize are typically called
  // once per thread or at application startup/shutdown.
  // For this self-contained function, we'll manage it here.
  ResultCode := CoInitialize(nil);
  if Failed(ResultCode) then
    raise Exception.Create('Failed to initialize COM library.');

  try
    // Create an instance of the DesktopWallpaper COM object.
    ResultCode := CoCreateInstance(CLSID_DesktopWallpaper, nil, CLSCTX_LOCAL_SERVER, IID_IDesktopWallpaper, desktopWallpaper);

    if Succeeded(ResultCode) and (desktopWallpaper <> nil) then
    begin
      // Set the wallpaper position to DWPOS_FILL.
      // DWPOS_FILL = 4, defined in Winapi.ShlObj. This is the equivalent of "Fill".
      // Other options include: DWPOS_FIT, DWPOS_STRETCH, DWPOS_CENTER, etc.
      desktopWallpaper.SetPosition(DWPOS_FILL);

        // Set the wallpaper image itself.
        // The first parameter (nil) means apply to all monitors.
        // To target a specific monitor, you would first get its ID with GetMonitorDevicePathAt.
      desktopWallpaper.SetWallpaper(nil, PWideChar(filename));

        // Check if the final operation was successful.
//        if Succeeded(ResultCode) then
//          Result := True;
    end
    else
    begin
      // If we failed to get the interface, it's likely because the
      // OS is older than Windows 8.
      raise Exception.CreateFmt('Failed to create IDesktopWallpaper interface. HRESULT: %x. This function requires Windows 8 or newer.',
        [ResultCode]);
    end;
  finally
    // Release the COM object and uninitialize COM.
    desktopWallpaper := nil;
    CoUninitialize;
  end;
end;

procedure TfrmMain.StartwithWindows1Click(Sender: TObject);
begin
  if StartwithWindows1.Checked then
    RemoveAutoRun('BingWallpaperLite')
  else
    SetAutoRun('BingWallpaperLite', Application.ExeName);
  StartwithWindows1.Checked := IsAutoRunSet('BingWallpaperLite');
end;

procedure TfrmMain.tmr1Timer(Sender: TObject);
begin
  if NowUTC() - FNextUpdate > 0 then
    updateWallpaper();
end;

procedure TfrmMain.uiSetup;
begin
  regionMenuSetup();
  with FConfig do
  begin
    tmr1.Enabled := ReadBool('general', 'refresh', True);
    FBWHandler.Region := ReadString('general', 'region', 'en-US');
    Refresheveryday1.Checked := tmr1.Enabled;
    StartwithWindows1.Checked := IsAutoRunSet('BingWallpaperLite');
    FNextUpdate := UnixToDateTime(Int64(ReadInteger('general', 'nextupdate', 0)) * 10);
    with Region1.GetEnumerator do
    try
      while MoveNext do
        Current.Checked := FRegionMap[Current.Tag] = FBWHandler.Region;   // check current region
    finally
      Free;
    end;
  end;
  FWallPaperFolder := IncludeTrailingPathDelimiter(ExtractFilePath(Application.ExeName) + 'wallpapers');
  ForceDirectories(FWallPaperFolder);

  if tmr1.Enabled then
    tmr1.OnTimer(nil);  // update wallpaper on launch
end;

procedure TfrmMain.UpdateNow1Click(Sender: TObject);
begin
  updateWallpaper(true);
end;

procedure TfrmMain.updateWallpaper(force: Boolean);
var
  urls: TStringList;
  s, filename: string;
  mem: TMemoryStream;
  next: TDateTime;
begin
  filename := FWallPaperFolder + FormatDateTime('YYYY-MM-DD', Now()) + '.jpg';
  if FileExists(filename) and not force then
    Exit; // do not update if the wallpaper is already uptodate

  urls := FBWHandler.GetWallpaperUrl(0, is4kMonitor(), next);
  mem := TMemoryStream.Create;
  try
    for s in urls do
    begin
      if FBWHandler.GetImage(s, mem) then
        Break;
    end;
    if mem.Size > 0 then
    begin
      // save wallpaper using today's date
      filename := FWallPaperFolder + FormatDateTime('YYYY-MM-DD', Now()) + '.jpg';
      mem.SaveToFile(filename);
      setWallpaper(filename);
      NextUpdate := next;
    end;
  finally
    urls.Free;
    mem.Free;
  end;
end;

end.

