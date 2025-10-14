unit main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms, Dialogs, CoolTrayIcon, Menus, ExtCtrls,
  IniFiles, uBWHandler, IdBaseComponent, IdAntiFreezeBase, IdAntiFreeze, ActiveX, ShlObj, ShObjIdl;

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
    procedure Quit1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure tmr1Timer(Sender: TObject);
    procedure UpdateNow1Click(Sender: TObject);
    procedure Refresheveryday1Click(Sender: TObject);
    procedure StartwithWindows1Click(Sender: TObject);
  private
    { Private declarations }
    FConfig: TIniFile;
    FLastUpdate: TDateTime;
    FBWHandler: TBingWallpaperHandler;
    FWallPaperFolder: string;
    procedure uiSetup();
    procedure updateWallpaper(force: Boolean = false);
    procedure setWallpaper(filename: string);
    function is4kMonitor(): Boolean;
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

uses
  Autorunner;

{$R *.dfm}

procedure TfrmMain.FormCreate(Sender: TObject);
begin
  FConfig := TIniFile.Create(ExtractFilePath(Application.ExeName) + 'config.ini');
  FBWHandler := TBingWallpaperHandler.Create;
  uiSetup();
end;

procedure TfrmMain.FormDestroy(Sender: TObject);
begin
  FConfig.Free;
end;

function TfrmMain.is4kMonitor: Boolean;
begin
  Result := (Monitor.Height > 1080) or (Monitor.Width > 1920);
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
  if Now() - FLastUpdate > 1 then
    updateWallpaper();
end;

procedure TfrmMain.uiSetup;
begin
  with FConfig do
  begin
    tmr1.Enabled := ReadBool('general', 'refresh', True);
    FBWHandler.Region := ReadString('general', 'region', 'en-US');
    Refresheveryday1.Checked := tmr1.Enabled;
    StartwithWindows1.Checked := IsAutoRunSet('BingWallpaperLite');
  end;
  FWallPaperFolder := IncludeTrailingPathDelimiter(ExtractFilePath(Application.ExeName) + 'wallpapers');
  ForceDirectories(FWallPaperFolder);

  if tmr1.Enabled then
    updateWallpaper();  // update wallpaper on launch
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
begin
  filename := FWallPaperFolder + FormatDateTime('YYYY-MM-DD', Now()) + '.jpg';
  if FileExists(filename) and not force then
    Exit; // do not update if the wallpaper is already uptodate

  urls := FBWHandler.GetWallpaperUrl(0, is4kMonitor());
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
      FLastUpdate := Now();
    end;
  finally
    urls.Free;
    mem.Free;
  end;
end;

end.

