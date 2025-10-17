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
    function GetWallpaperUrl(idx: Integer; UHD: Boolean; var nextUpdate: TDateTime): TStringList;
    function GetImage(url: string; strm: TStream): Boolean;
    constructor Create;
    destructor Destroy; override;
    property Region: string read FRegion write FRegion;
  end;

implementation

uses
  Execute.IdSSLSChannel, superobject;

function parseDatetime(const S: string): TDateTime;
var
  LYear, LMonth, LDay, LHour, LMinute, LSecond, LMilliSecond: Word;
begin
  // Ensure the string is the correct length before parsing
  if Length(S) <> 12 then
  begin
    Result := 0;
  end;

  try
    // Extract each part of the date and time
    LYear := StrToInt(Copy(S, 1, 4));
    LMonth := StrToInt(Copy(S, 5, 2));
    LDay := StrToInt(Copy(S, 7, 2));
    LHour := StrToInt(Copy(S, 9, 2));
    LMinute := StrToInt(Copy(S, 11, 2));
    LSecond := 0; // The string format has no seconds
    LMilliSecond := 0;

    // Use EncodeDate and EncodeTime to build the TDateTime value
    // This also validates that the numbers form a valid date (e.g., month is 1-12)
    Result := EncodeDate(LYear, LMonth, LDay) + EncodeTime(LHour, LMinute, LSecond, LMilliSecond);
  except
    // If StrToInt fails or EncodeDate/EncodeTime raise an error (e.g., for an invalid date like month 13)
    // an exception is raised to signal the conversion failed.
    on E: Exception do
    begin
      // Re-raise the exception with more context
      Result := 0;
    end;
  end;
end;


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

function TBingWallpaperHandler.GetWallpaperUrl(idx: Integer; UHD: Boolean; var nextUpdate: TDateTime): TStringList;
var
  url: string;
  resp: string;
  json: ISuperObject;
  images: TSuperArray;
  startdate, fullstartdate, enddate, fullenddate: string;
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
          // analyse the end date from fullstartdate
        startdate := images.O[0].S['startdate'];
        fullstartdate := images.O[0].S['fullstartdate'];
        enddate := images.O[0].S['enddate'];
        fullenddate := enddate + Copy(fullstartdate, length(startdate)+1, length(fullstartdate) - length(startdate));
        nextUpdate := parseDatetime(fullenddate);
        if nextUpdate = 0 then
          nextUpdate := Now() + 1;
      end;
    end;
  except
  end;
end;

end.

