unit frmLogWindow;
// simple window that shows a ListBox field
// the ListBox can be accessed through methods
// logging is only enabled when the window is open (through the Enabled-Property

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Menus,ClipBrd;

type
  TLogWindow = class(TForm)
    ListBox: TListBox;
    PopupMenu1: TPopupMenu;
    Clear1: TMenuItem;
    toclipboard1: TMenuItem;
    procedure Clear1Click(Sender: TObject);
    procedure toclipboard1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
    LoggingEnabled : boolean;   //is logging enabled?   default is false

  public
    { Public declarations }
    /// <summary>Flag whether or not logging is enables</summary>
    /// <comments>logging is always enabled when log window is open</comments>
    property Enabled : boolean read LoggingEnabled write LoggingEnabled;  //property that reads or sets the internal variable whether or not logging is enabled
    /// <summary>adds an entry to the log window ListBox</summary>
    /// <comments>entries are only logged when the log window is open</comments>
    procedure addLogEntry(const logstring : string);
    /// <summary>clears the ListBox in the Log Window</summary>
    procedure Clear;
  end;

var
  LogWindow: TLogWindow;

implementation

procedure TLogWindow.Clear;
begin
  ListBox.Clear;
end;

procedure TLogWindow.addLogEntry(const logstring : string);
/// add line to the LitsBox
///only if logging is enabled (typically when WIndow is open)
begin
  If LoggingEnabled = true THEN   // only log when logging is enabled
  begin
    ListBox.Items.Add(logstring);
  end;
end;

{$R *.dfm}

procedure TLogWindow.Clear1Click(Sender: TObject);
begin
ListBox.Clear;
end;

procedure TLogWindow.FormClose(Sender: TObject; var Action: TCloseAction);
begin
 LogWindow.Enabled:=false; //enable logging when window is being displayed
end;

procedure TLogWindow.FormShow(Sender: TObject);
begin
 LogWindow.Enabled:=true; //enable logging when window is being displayed
end;

procedure TLogWindow.toclipboard1Click(Sender: TObject);
begin
// save selected line to clipboard
ClipBoard.AsText:=Listbox.Items[ListBox.Itemindex];
end;

end.
