unit uBWHandler;

interface

uses
  Windows, Sysutils, Classes, IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient, IdHTTP;

type
  TBingWallpaperHandler = class
  private
    FRequester: TIdHttp;
    FRegion: string;
  public
    function GetWallpaperUrl(idx: Integer; UHD: Boolean): TStringList;
    function GetImage(url: string; strm: TStream): Boolean;
    constructor Create;
    destructor Destroy; override;
    property Region: string read FRegion write FRegion;
  end;

implementation

uses
  Execute.IdSSLSChannel, superobject;

{ TBingWallpaperHandler }

constructor TBingWallpaperHandler.Create;
var
  SChannel: TIdSSLIOHandlerSocketSChannel;
begin
  FRequester := TIdHTTP.Create(nil);
  with FRequester do
  begin
    SChannel := TIdSSLIOHandlerSocketSChannel.Create;
    IOHandler := SChannel;
    Request.UserAgent :=
      'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/62.0.3202.94 Safari/537.36';
    HandleRedirects := True;
  end;

  FRegion := 'en-US';
end;

destructor TBingWallpaperHandler.Destroy;
begin
  FRequester.IOHandler.Free;
  FRequester.Free;
  inherited;
end;

function TBingWallpaperHandler.GetImage(url: string; strm: TStream): Boolean;
begin
  Result := False;
  try
    strm.Seek(0, soBeginning);
    strm.Size := 0;
    FRequester.Get(url, strm);
    Result := strm.Size > 0;
  except
  end;
end;

function TBingWallpaperHandler.GetWallpaperUrl(idx: Integer; UHD: Boolean): TStringList;
var
  url: string;
  resp: string;
  json: ISuperObject;
  images: TSuperArray;
begin
  Result := TStringList.Create;
  try
    url := format('https://www.bing.com/HPImageArchive.aspx?format=js&idx=%d&n=1&mkt=%s', [idx, Region]);
    resp := FRequester.Get(url);
    json := SO(resp);
    if json.IsType(stObject) then
    begin
      images := json.ForcePath('images', stArray).AsArray();
      if images.Length > 0 then
      begin
        if UHD then
          Result.Add(Format('https://www.bing.com%s_UHD.jpg', [images.O[0].S['urlbase']]));     //4k version
        Result.Add(Format('https://www.bing.com%s_1920x1080.jpg', [images.O[0].S['urlbase']]));   // 1080p version
      end;
    end;
  except
  end;
end;

end.

