unit uDPIAware;

interface

uses
  Windows;

implementation

type
  DPI_AWARENESS_CONTEXT = type THandle;

const
  DPI_AWARENESS_CONTEXT_UNAWARE = DPI_AWARENESS_CONTEXT(-1);
  DPI_AWARENESS_CONTEXT_SYSTEM_AWARE = DPI_AWARENESS_CONTEXT(-2);
  DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE = DPI_AWARENESS_CONTEXT(-3);
  DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = DPI_AWARENESS_CONTEXT(-4);
  DPI_AWARENESS_CONTEXT_UNAWARE_GDISCALED = DPI_AWARENESS_CONTEXT(-5);
  PROCESS_DPI_UNAWARE = 0;
  PROCESS_SYSTEM_DPI_AWARE = 1;
  PROCESS_PER_MONITOR_DPI_AWARE = 2;

var
  pSetProcessDpiAwarenessContext: function(value: DPI_AWARENESS_CONTEXT): BOOL; stdcall;
  pSetProcessDpiAwareness: function(value: DWORD): HRESULT; stdcall;
  pSetProcessDPIAware: function(): BOOL; stdcall;

procedure InitDPIAware();
var
  hUser32: HMODULE;
begin
      // locate procs
  hUser32 := GetModuleHandle('user32.dll');
  pSetProcessDpiAwarenessContext := GetProcAddress(hUser32, 'SetProcessDpiAwarenessContext');
  pSetProcessDpiAwareness := GetProcAddress(hUser32, 'SetProcessDpiAwareness');
  pSetProcessDPIAware := GetProcAddress(hUser32, 'SetProcessDPIAware');

  if @pSetProcessDpiAwarenessContext <> nil then
  begin
    pSetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE);
    exit;
  end;
  if @pSetProcessDpiAwareness <> nil then
  begin
    pSetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE);
    exit;
  end;
  if @pSetProcessDPIAware <> nil then
  begin
    pSetProcessDPIAware();
    exit;
  end;
end;

initialization
  InitDPIAware();

end.

