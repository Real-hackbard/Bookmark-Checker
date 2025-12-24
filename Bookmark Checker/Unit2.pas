{$A+,B-,C+,D+,E-,F-,G+,H+,I+,J-,K-,L+,M-,N+,O+,P+,Q-,R-,S-,T-,U-,V+,W-,X+,Y+,Z1}
{$MINSTACKSIZE $00004000}
{$MAXSTACKSIZE $00100000}
{$IMAGEBASE $00400000}
{$APPTYPE GUI}

unit Unit2;

interface

uses
  Windows, ShellAPI, ShlObj, ActiveX, Classes, SysUtils, ComCtrls, IdHTTP;


function GetIEFavourites(const favpath: string): TStringList;
procedure FreePidl(pidl: PItemIDList);
procedure ExtractBookmarksFromFile(BookmarkFile: string; Bookmarks: TStrings);
procedure AddLinkToLV(LV: TListview; const Link: string; ResponseCode: Integer;
  const ResponseString: string);
procedure CheckBookmark(IdHTTP: TIdHTTP; const Link: string; var ResponseCode:
  Integer; var
  ResponseString: string);
function DeleteDeadFavLink(const favpath: string; DeadLink: string): Boolean;
function RewriteBookmarkFile(sl: TStrings; Filename: string): Cardinal;
function MakeLogHTML(Filename: String; Links: TStrings): Boolean;


implementation

function NextPos(SubStr: AnsiString; Str: AnsiString; LastPos: DWORD
  = 0): DWORD;
type
  StrRec = packed record
    allocSiz: Longint;
    refCnt: Longint;
    length: Longint;
  end;

const
  skew = sizeof(StrRec);

asm
  // Search-String passed?
  TEST    EAX,EAX
  JE      @@noWork

  // Sub-String passed?
  TEST    EDX,EDX
  JE      @@stringEmpty

  // Save registers affected
  PUSH    ECX
  PUSH    EBX
  PUSH    ESI
  PUSH    EDI

  // Load Sub-String pointer
  MOV     ESI,EAX
  // Load Search-String pointer
  MOV     EDI,EDX
  // Save Last Position in EBX
  MOV     EBX,ECX

  // Get Search-String Length
  MOV     ECX,[EDI-skew].StrRec.length
  // subtract Start Position
  SUB     ECX,EBX
  // Save Start Position of Search String to return
  PUSH    EDI
  // Adjust Start Position of Search String
  ADD     EDI,EBX

  // Get Sub-String Length
  MOV     EDX,[ESI-skew].StrRec.length
  // Adjust
  DEC     EDX
  // Failed if Sub-String Length was zero
  JS      @@fail
  // Pull first character of Sub-String for SCASB function
  MOV     AL,[ESI]
  // Point to second character for CMPSB function
  INC     ESI

  // Load character count to be scanned
  SUB     ECX,EDX
  // Failed if Sub-String was equal or longer than Search-String
  JLE     @@fail
@@loop:
  // Scan for first matching character
  REPNE   SCASB
  // Failed, if none are matching
  JNE     @@fail
  // Save counter
  MOV     EBX,ECX
  PUSH    ESI
  PUSH    EDI
  // load Sub-String length
  MOV     ECX,EDX
  // compare all bytes until one is not equal
  REPE    CMPSB
  // restore counter
  POP     EDI
  POP     ESI
  // all byte were equal, search is completed
  JE      @@found
  // restore counter
  MOV     ECX,EBX
  // continue search
  JMP     @@loop
@@fail:
  // saved pointer is not needed
  POP     EDX
  XOR     EAX,EAX
  JMP     @@exit
@@stringEmpty:
  // return zero - no match
  XOR     EAX,EAX
  JMP     @@noWork
@@found:
  // restore pointer to start position of Search-String
  POP     EDX
  // load position of match
  MOV     EAX,EDI
  // difference between position and start in memory is
  //   position of Sub
  SUB     EAX,EDX
@@exit:
  // restore registers
  POP     EDI
  POP     ESI
  POP     EBX
  POP     ECX
@@noWork:
end;

function GetIEFavourites(const favpath: string): TStringList;
var
  searchrec: TSearchRec;
  str: TStringList;
  path, dir, FileName: string;
  Buffer: array[0..2047] of Char;
  found: Integer;
begin
  str := TStringList.Create;
  // Get all file names in the favourites path
  path := FavPath + '\*.url';
  dir := ExtractFilepath(path);
  found := FindFirst(path, faAnyFile, searchrec);
  while found = 0 do
  begin
    // Get now URLs from files in variable files
    Setstring(FileName, Buffer, GetPrivateProfilestring('InternetShortcut',
      PChar('URL'), nil, Buffer, SizeOf(Buffer), PChar(dir + searchrec.Name)));
    str.Add(FileName);
    found := FindNext(searchrec);
  end;
  // find Subfolders
  found := FindFirst(dir + '\*.*', faAnyFile, searchrec);
  while found = 0 do
  begin
    if ((searchrec.Attr and faDirectory) > 0) and (searchrec.Name[1] <> '.')
      then
      str.Addstrings(GetIEFavourites(dir + '\' + searchrec.Name));
    found := FindNext(searchrec);
  end;
  FindClose(searchrec);
  Result := str;
end;

procedure FreePidl(pidl: PItemIDList);
var
  allocator: IMalloc;
begin
  if Succeeded(SHGetMalloc(allocator)) then
  begin
    allocator.Free(pidl);
{$IFDEF VER100}
    allocator.Release;
{$ENDIF}
  end;
end;

{------------------------------------------------------------------------------}
{  Extracts the links from the bookmarkfile                                    }
{------------------------------------------------------------------------------}

procedure ExtractBookmarksFromFile(BookmarkFile: string; Bookmarks: TStrings);
const
  BEGINLINK = 'HREF="';
  ENDLINK = '"';
var
  F: TextFile;
  s: string;
  StartPos, EndPos: Integer;
begin
  AssignFile(F, BookmarkFile);
{$I-}
  Reset(F);
{$I+}
  if IOResult = 0 then
  begin
    while not EOF(F) do
    begin
      readln(F, s);
      StartPos := AnsiPos(BEGINLINK, Uppercase (s));
      if StartPos > 0 then
      begin
        Delete (s, 1, StartPos + Length (BEGINLINK) -1);
        EndPos := AnsiPos(ENDLINK, s);
        if EndPos > 0 then
        begin
          Delete (S, EndPos, Length (S));
          If length(s) > 0 then
            Bookmarks.Add(s);
        end;
      end;
    end;
    CloseFile(F);
  end;
end;

{------------------------------------------------------------------------------}
{  Adds the link to a Listview                                                 }
{------------------------------------------------------------------------------}

procedure AddLinkToLV(LV: TListview; const Link: string; ResponseCode: Integer;
  const ResponseString: string);
var
  NewItem: TListItem;
begin
  NewItem := LV.Items.Add;
  NewItem.Caption := Link;
  NewItem.SubItems.Add(IntToStr(ResponseCode));
  NewItem.SubItems.Add(ResponseString);
  NewItem.MakeVisible(False);
end;

{------------------------------------------------------------------------------}
{  Checks the links via IdHTTP (out: ResponseCode & ResponseString of Server   }
{------------------------------------------------------------------------------}

procedure CheckBookmark(IdHTTP: TIdHTTP; const Link: string; var ResponseCode: 
  Integer; var 
  ResponseString: string); 
begin 
  try 
    IdHTTP.Head(Link); 
    ResponseCode := IdHTTP.ResponseCode; 
    ResponseString := IdHTTP.ResponseText; 
  except 
     // bei einer Exception könnte IdHTTP.ResponseText leer sein 
     // dann wird die Exception-Message als ResponseString zurückgeliefert 
     // (z.B. wenn eine Firewall den Port 80 geblockt hat) 
     on E:Exception do 
     begin 
       ResponseCode := IdHTTP.ResponseCode; 
       ResponseString := IdHTTP.ResponseText; 
       if ResponseString = '' then 
          ResponseString := E.Message; 
     end; 
  end; 
end;

{------------------------------------------------------------------------------}
{  Deletes a dead IE Link                                                      }
{------------------------------------------------------------------------------}

function DeleteDeadFavLink(const favpath: string; DeadLink: string): Boolean;
var
  searchrec: TSearchRec;
  Path, Dir, Filename: string;
  Buffer: array[0..1024] of Char;
  Found: Integer;
begin
  result := False;
  Path := favpath + '\*.url';
  Dir := ExtractFilePath(Path);
  Found := FindFirst(Path, faAnyFile, searchrec);
  while (Found = 0) and (result = False) do
  begin
    SetString(Filename, Buffer, GetPrivateProfileString('InternetShortcut',
      'URL', nil, Buffer, sizeof(Buffer), Pointer(dir + searchrec.Name)));
    if Filename = DeadLink then
    begin
      result := DeleteFile(dir + searchrec.Name);
      exit;
    end;
    found := FindNext(searchrec);
  end;
  found := FindFirst(dir + '\*.*', faAnyFile, searchrec);
  while (Found = 0) and (result = False) do
  begin
    if ((searchrec.Attr and faDirectory) > 0) and (searchrec.Name[1] <> '.')
      then
    begin
      DeleteDeadFavLink(dir + searchrec.Name, DeadLink);
    end;
    found := FindNext(searchrec);
  end;
  FindClose(searchrec);
  exit;
end;

{------------------------------------------------------------------------------}
{  Rewrites the bookmarkfile without the dead links                            }
{------------------------------------------------------------------------------}

function RewriteBookmarkFile(sl: TStrings; Filename: String): Cardinal;
var
  slBookmarkFile: TStringList;
  outerLoop, innerLoop: Cardinal;
  Count: Cardinal;
begin
  Count := 0;
  slBookmarkFile := TStringList.Create;
  try
    slBookmarkFile.LoadFromFile(Filename);
    for outerLoop := 0 to sl.Count-1 do
    begin
      for innerLoop := slBookmarkFile.Count-1 downto 0 do
      begin
        if pos(sl.Strings[outerLoop], slBookmarkFile.Strings[innerLoop]) > 0 then
        begin
          slBookmarkFile.Delete(innerLoop);
          Inc(Count);
        end;
      end;
    end;
    slBookmarkFile.SaveToFile(Filename);
  finally
    FreeAndNil(slBookmarkFile);
  end;
  result := Count;
end;

{------------------------------------------------------------------------------}
{  Log deleted links into an HTML file                                         }
{------------------------------------------------------------------------------}

const
  HTMLPAGE = '<html>'+#13#10+
    '<head>'+#13#10+
    '<title>BookmarkChecker-Logfile</title>'+#13#10+
    '<meta name="GENERATOR" content="BookmarkChecker">'+#13#10+
    '</head>'+#13#10+
    '<body>'+#13#10+
    '<h1>BookmarkChecker-Logfile</h1>'+#13#10;

resourcestring
  rsDeadLinks = 'gelöschte Links';

function MakeLogHTML(Filename: String; Links: TStrings): Boolean;
var
  sl: TStringList;
  Loop: Integer;
begin
  result := False;
  sl := TStringList.Create;
  try
    if FileExists(Filename) then
    begin
      sl.LoadFromFile(Filename);
      sl.Delete(sl.Count-1);
      sl.Delete(sl.Count-1);
    end
    else
      sl.Add(HTMLPAGE);
    sl.Add('<dl>');
    sl.Add('<dt><h2>'+DateTimeToStr(now)+'</h2>');
    sl.Add('<dl>');
    sl.Add('<dt><h3>'+rsDeadLinks+'</h3>');
    sl.Add('<dl>');
    for Loop := 0 to Links.Count-1 do
    begin
      sl.Add('<dt><a target="_blank" href="'+Links.Strings[Loop]+'">'+Links.Strings[Loop]+'</a></dt><br>');
    end;
    sl.Add('</dl>');
    sl.Add('</dt>');
    sl.Add('</dl>');
    sl.Add('</dt>');
    sl.Add('</dl>');
    sl.Add('</body>');
    sl.Add('</html>');
    sl.SaveToFile(Filename);
  finally
    FreeAndNil(sl);
  end;
end;

end.

