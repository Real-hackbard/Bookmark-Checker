{$A+,B-,C+,D+,E-,F-,G+,H+,I+,J-,K-,L+,M-,N+,O+,P+,Q-,R-,S-,T-,U-,V+,W-,X+,Y+,Z1}

{$MINSTACKSIZE $00004000}
{$MAXSTACKSIZE $00100000}
{$IMAGEBASE $00400000}
{$APPTYPE GUI}

unit Unit1;

{.$DEFINE DEBUG}

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, CheckLst, ComCtrls, IdBaseComponent, IdComponent,
  IdTCPConnection, IdTCPClient, IdHTTP, IdAntiFreezeBase, IdAntiFreeze,
  ExtCtrls, XPMan;

type
  TForm1 = class(TForm)
    btnBack: TButton;
    btnNext: TButton;
    IdHTTP1: TIdHTTP;
    IdAntiFreeze1: TIdAntiFreeze;
    btnCancel: TButton;
    Label1: TLabel;
    OpenDialog1: TOpenDialog;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    edtFilename: TEdit;
    btnOpen: TButton;
    rdbIEBookmarks: TRadioButton;
    rdbLinkFile: TRadioButton;
    TabSheet2: TTabSheet;
    lblCheckingLink: TLabel;
    lblLink: TLabel;
    lblProgressLabel: TLabel;
    ProgressBar1: TProgressBar;
    ListView1: TListView;
    TabSheet3: TTabSheet;
    lblDeadLinks: TLabel;
    ListView2: TListView;
    Bevel1: TBevel;
    Bevel2: TBevel;
    TabSheet4: TTabSheet;
    Image1: TImage;
    lblWellcome: TLabel;
    lblWizard: TLabel;
    lblHowToUse: TLabel;
    btnAbout: TButton;
    lblProgressDetail: TLabel;
    lblConnection: TLabel;
    cbConnection: TComboBox;
    lblTimeout: TLabel;
    edtTimeout: TEdit;
    chkLogFile: TCheckBox;
    SaveDialog1: TSaveDialog;
    stcFilename: TStaticText;
    chkCheckAll: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure btnNextClick(Sender: TObject);
    procedure btnBackClick(Sender: TObject);
    procedure PageControl1Change(Sender: TObject);
    procedure PageControl1Changing(Sender: TObject;
      var AllowChange: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnOpenClick(Sender: TObject);
    procedure rdbIEBookmarksClick(Sender: TObject);
    procedure rdbLinkFileClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure btnAboutClick(Sender: TObject);
    procedure cbConnectionChange(Sender: TObject);
    procedure edtTimeoutKeyPress(Sender: TObject; var Key: Char);
    procedure ListView2DblClick(Sender: TObject);
    procedure chkLogFileClick(Sender: TObject);
    procedure chkCheckAllClick(Sender: TObject);
  private
    { Private-Deklarationen }
    LastPage: Boolean;
    Cancel: Boolean;
    procedure DeadLinks(var sl: TStringList);
  public
    { Public-Deklarationen }
  end;

const
  APPNAME                = 'Bookmark Checker';
  INFO_TEXT              = APPNAME + ' %s' + #13#10 +
                           'Copyright © Your Name' + #13#10 +
                           'https://github.com/';

var
  Form1: TForm1;

implementation

{$R *.dfm}

uses
  Unit2, ShlObj, ActiveX, ShellAPI;

resourcestring
  rsWellcome =          'Welcome';
  rsWizard =            'The bookmark checker checks your bookmarks/favorites for ' +
                        'pages that no longer exist, and offers you the opportunity to access them. ' +
                        'to be removed automatically.';
  rsHowToUse =          'Click "Next" to continue the process. ' +
                        'or click "Cancel" to end the assistant.';
  rsBtnAbout =          'About';

  rsRdbIE =             '&Internet Explorer Favorites';
  rsRdbLinkFile =       '&Bookmark file in HTML format (Mozilla, Chrome, etc.)';
  rsBtnOpen =           'O&pen';
  rsLblConnection =     'Connection:';
  rscbConnectionModem = 'Modem / ISDN';
  rscbConnectionDSL =   'DSL or Faster';
  rsCbConnectionUser =  'Custom';
  rsLblTimeout =        'Timeout in milliseconds';

  rsTimeoutWrong =      'Specify a value for the timeout.' + #13#10 +
                        'or select an entry from the list.';

  rsLblCheckingLink =   'Check Link: ';
  rsLblProgress =       'Progress:';
  rsProgress =          '%d from %d Links checked';

  rsLblDeadLinks =      'Unreachable links: ';
  rsChkCheckAll =       '&Mark all';
  rsChkLogFile =        'Write to &log file';
  rsSaveAsDlgCaption =  'Save log file as...';

  rsBack =              '&Back';
  rsNext =              '&Next';
  rsFinish =            '&Finish';
  rsCancel =            '&Abort';
  rsExit =              '&Close';

  rsEndMessage =        'The selected links have been deleted.';
  rsNoAction =          'No links were deleted.';


function GetVersion: string;
var
  VerInfoSize            : DWORD;
  VerInfo                : Pointer;
  VerValueSize           : DWORD;
  VerValue               : PVSFixedFileInfo;
  Dummy                  : DWORD;
begin
  VerInfoSize := GetFileVersionInfoSize(PChar(ParamStr(0)), Dummy);
  GetMem(VerInfo, VerInfoSize);
  try
    GetFileVersionInfo(PChar(ParamStr(0)), 0, VerInfoSize, VerInfo);
    VerQueryValue(VerInfo, '\', Pointer(VerValue), VerValueSize);
    with VerValue^ do
    begin
      Result := IntToStr(dwFileVersionMS shr 16);
      Result := Result + '.' + IntToStr(dwFileVersionMS and $FFFF);
      Result := Result + '.' + IntToStr(dwFileVersionLS shr 16);
      Result := Result + '.' + IntToStr(dwFileVersionLS and $FFFF);
    end;
  finally
    FreeMem(VerInfo, VerInfoSize);
  end;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  Caption := APPNAME;
  Label1.Caption := APPNAME;

  lblWellcome.Caption := rsWellcome;
  lblWizard.Caption := rsWizard;
  lblHowToUse.Caption := rsHowToUse;
  btnBack.Caption := rsBack;
  btnNext.Caption := rsNext;
  btnCancel.Caption := rsCancel;
  btnAbout.Caption := rsBtnAbout;

  rdbIEBookmarks.Caption := rsRdbIE;
  rdbLinkFile.Caption := rsRdbLinkFile;
  btnOpen.Caption := rsBtnOpen;
  lblConnection.Caption := rsLblConnection;
  cbConnection.Items.Add(rscbConnectionModem);
  cbConnection.Items.Add(rscbConnectionDSL);
  cbConnection.Items.Add(rsCbConnectionUser);
  cbConnection.ItemIndex := 0;
  lblTimeout.Caption := rsLblTimeout;
  edtTimeout.Text := '5000';

  lblCheckingLink.Caption := rsLblCheckingLink;
  lblProgressLabel.Caption := rsLblProgress;

  lblDeadLinks.Caption := rsLblDeadLinks;
  chkCheckAll.Caption := rsChkCheckAll;
  chkLogFile.Caption := rsChkLogFile;
  SaveDialog1.Title := rsSaveAsDlgCaption;
  SaveDialog1.FileName := APPNAME+'.html';

  Application.Title := APPNAME;
  lblProgressDetail.Left := lblProgressLabel.Left + lblProgressLabel.Width + 2;
  PageControl1.ActivePageIndex := 0;
{$IFDEF DEBUG}
  rdbLinkFile.Checked := True;
  edtFilename.Text :=
    'E:\Delphi\Programme\VCL\BookmarkChecker1_0\Source\test.html';
{$ENDIF}
end;

procedure TForm1.DeadLinks(var sl: TStringList);
var
  Loop: Cardinal;
begin
  for Loop := 0 to Listview2.Items.Count - 1 do
  begin
    if Listview2.Items[Loop].Checked then
      sl.Add(Listview2.Items[Loop].Caption);
  end;
end;

procedure TForm1.btnNextClick(Sender: TObject);
var
  AllowChange: Boolean;
  sl: TStringList;
  Loop: Cardinal;
  pidl: PItemIDList;
  FavPath: array[0..MAX_PATH] of Char;
  Count: Cardinal;
  s: string;
begin
  if LastPage then
  begin
    sl := TStringList.Create;
    try
      Count := 0;
      DeadLinks(sl);
      // User selected links in the listview
      if sl.Count > 0 then
      begin
        if chkLogFile.Checked then
        begin
          MakeLogHTML(stcFilename.Caption, sl);
        end;
        // IE Bookmarks
        if rdbIEBookmarks.Checked then
        begin
          for Loop := 0 to sl.Count - 1 do
          begin
            if Succeeded(ShGetSpecialFolderLocation(0, CSIDL_FAVORITES, pidl))
              then
            begin
              if ShGetPathfromIDList(pidl, FavPath) then
              begin
                if DeleteDeadFavLink(FavPath, sl.Strings[Loop]) then
                  Inc(Count);
              end;
              FreePIDL(pidl);
            end;
          end;
        end
        // Bookmark-File
        else
        begin
          Count := RewriteBookmarkFile(sl, edtFilename.Text);
        end;
        s := Format(rsEndMessage, [Count]);
        Messagebox(Handle, pointer(s), APPNAME, MB_ICONINFORMATION);
      end
      else
      begin
        s := rsNoAction;
        Messagebox(Handle, pointer(s), APPNAME, MB_ICONINFORMATION);
      end;
    finally
      FreeAndNil(sl);
    end;
    Close;
  end;
  PageControl1.ActivePageIndex := PageControl1.ActivePageIndex + 1;
  PageControl1.OnChange(Sender);
  PageControl1.OnChanging(Sender, AllowChange);
end;

procedure TForm1.btnBackClick(Sender: TObject);
var
  AllowChange: Boolean;
begin
  Cancel := PageControl1.ActivePageIndex = 0;
  PageControl1.ActivePageIndex := PageControl1.ActivePageIndex - 1;
  PageControl1.OnChange(Sender);
  PageControl1.OnChanging(Sender, AllowChange);
end;

procedure TForm1.PageControl1Change(Sender: TObject);
begin
  LastPage := False;
  btnBack.Enabled := PageControl1.ActivePageIndex > 0;
  btnNext.Enabled := PageControl1.ActivePageIndex < PageControl1.PageCount;
  case PageControl1.ActivePageIndex of
    0: ;
    1:
      begin
        LastPage := False;
        Cancel := True;
        btnNext.Caption := rsNext;
      end;
    2:
      begin
        LastPage := False;
        Cancel := False;
        btnNext.Caption := rsNext;
      end;
    3:
      begin
        if Listview2.Items.Count = 0 then
        begin
          btnNext.Enabled := False;
          btnCancel.Caption := rsExit;
        end
        else
        begin
          LastPage := True;
          Cancel := False;
          btnNext.Caption := rsFinish;
        end;
      end;
  end;
end;

procedure TForm1.PageControl1Changing(Sender: TObject;
  var AllowChange: Boolean);
var
  Bookmarks: TStringList;
  Loop: Integer;
  pidl: PItemIDList;
  FavPath: array[0..MAX_PATH] of Char;
  ResponseCode: Integer;
  ResponseString: string;
begin
  Application.ProcessMessages;
  case PageControl1.ActivePageIndex of
    2:
      begin
        if edtTimeOut.Text = '' then
        begin
          PageControl1.ActivePageIndex := 1;
          Messagebox(Handle, pointer(rsTimeoutWrong), APPNAME, MB_ICONWARNING);
          exit;
        end;
        if (not FileExists(edtFilename.Text)) and (rdbLinkFile.Checked) then
        begin
          PageControl1.ActivePageIndex := 1;
          Messagebox(Handle, pointer(SysErrorMessage(GetLastError())), APPNAME, MB_ICONWARNING);
          exit;
        end;
        IdHTTP1.ReadTimeout := StrToInt(edtTimeout.Text);
        Bookmarks := TStringList.Create;
        try
          //  Bookmarkfile or IE Favorites
          if rdbLinkFile.Checked then
          begin
            // Extract bookmarks from the file
            if FileExists(edtFilename.Text) then
              ExtractBookmarksFromFile(edtFilename.Text, Bookmarks);
          end
          else
          begin
            // Get the IE Favorites
            if Succeeded(ShGetSpecialFolderLocation(Handle, CSIDL_FAVORITES,
              pidl)) then
            begin
              if ShGetPathfromIDList(pidl, FavPath) then
                Bookmarks := GetIEFavourites(StrPas(FavPath));
              FreePIDL(pidl);
            end;
          end;
          // only if there are bookmarks to check
          if Bookmarks.Count > 0 then
          begin
            btnNext.Enabled := False;
            Progressbar1.Max := BookMarks.Count;
            ProgressBar1.Position := 0;
            Listview1.Items.Clear;
            Listview2.Items.Clear;
            for Loop := 0 to Bookmarks.Count - 1 do
            begin
              lblLink.Caption := Bookmarks.Strings[Loop];
              if Cancel then
                exit;
              CheckBookmark(IdHTTP1, Bookmarks[Loop], ResponseCode,
                ResponseString);
              Progressbar1.StepIt;
              lblProgressDetail.Caption := Format(rsProgress, [Loop + 1,
                Bookmarks.Count]);
              AddLinkToLV(Listview1, Bookmarks.Strings[Loop], ResponseCode,
                ResponseString);
              if (ResponseCode = 403) or (ResponseCode = 404)
                or (ResponseCode = 410) or (ResponseCode = 500) then
              begin
                AddLinkToLV(Listview2, Bookmarks.Strings[Loop], ResponseCode,
                  ResponseString);
              end;
            end;
          end;
        finally
          FreeAndNil(Bookmarks);
        end;
        btnNext.Enabled := True;
      end;
  end;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Cancel := True;
end;

procedure TForm1.btnOpenClick(Sender: TObject);
begin
  if OpenDialog1.Execute then
    edtFilename.Text := OpenDialog1.Filename;
end;

procedure TForm1.rdbIEBookmarksClick(Sender: TObject);
begin
  edtFilename.Enabled := not rdbIEBookmarks.Checked;
  btnOpen.Enabled := not rdbIEBookmarks.Checked;
end;

procedure TForm1.rdbLinkFileClick(Sender: TObject);
begin
  edtFilename.Enabled := rdbLinkFile.Checked;
  btnOpen.Enabled := rdbLinkFile.Checked;
end;

procedure TForm1.btnCancelClick(Sender: TObject);
begin
  Close;
end;

procedure TForm1.btnAboutClick(Sender: TObject);
var
  s: String;
begin
  s := Format(INFO_TEXT, [GetVersion]);
  Messagebox(Handle, PChar(s), APPNAME, MB_ICONINFORMATION);
end;

procedure TForm1.cbConnectionChange(Sender: TObject);
begin
  case cbConnection.ItemIndex of
    0:
      begin
        edtTimeout.Text := '5000';
        edtTimeout.Enabled := False;
      end;
    1:
      begin
        edtTimeout.Text := '1000';
        edtTimeout.Enabled := False;
      end;
    2: edtTimeout.Enabled := True;
  end;
end;

procedure TForm1.edtTimeoutKeyPress(Sender: TObject; var Key: Char);
begin
  if not (Key in [#48..#57, #8]) then
    Key := #0;
end;

procedure TForm1.ListView2DblClick(Sender: TObject);
var
  URL: String;
begin
  Url := ListView2.Items[ListView2.Itemindex].Caption;
  ShellExecute(Handle, 'open', pointer(URL), nil, nil, SW_NORMAL);
end;

procedure TForm1.chkLogFileClick(Sender: TObject);
begin
  if chkLogFile.Checked then
  begin
    if SaveDialog1.Execute then
    begin
      stcFilename.Caption := SaveDialog1.FileName;
    end
    else
      chkLogFile.Checked := False;
  end;
end;

procedure TForm1.chkCheckAllClick(Sender: TObject);
var
  i: Integer;
begin
  if chkCheckAll.Checked then
  begin
    for i :=  0 to Listview2.Items.Count - 1 do
      ListView2.Items[i].Checked := True;
  end
  else
  begin
        for i :=  0 to Listview2.Items.Count - 1 do
      ListView2.Items[i].Checked := False;
  end;
end;

end.

