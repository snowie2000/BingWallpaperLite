program BingWallpaperLite;

uses
  uDPIAware in 'uDPIAware.pas',
  Forms,
  main in 'main.pas' {frmMain},
  Execute.SChannel in 'Indy.SChannel\Execute.SChannel.pas',
  Execute.WinSSPI in 'Indy.SChannel\Execute.WinSSPI.pas',
  Execute.IdSSLSChannel in 'Indy.SChannel\Execute.IdSSLSChannel.pas',
  uBWHandler in 'uBWHandler.pas' {$R *.res},
  ShObjIdl in 'ShObjIdl.pas',
  Autorunner in 'Autorunner.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.ShowMainForm := False;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
