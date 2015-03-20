unit _LogSQL;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, JvExExtCtrls, JvNetscapeSplitter, ComCtrls;

type
  TfrmLogSql = class(TForm)
    ListViewSQL: TListView;
    JvNetscapeSplitter1: TJvNetscapeSplitter;
    SynEdit1: TMemo;
    procedure ListViewSQLChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
  private
    { Private declarations }
  public
    procedure AddLogSQL(Command,CommandType,Status,CursorType,LockType:String;RecordsAffected:Integer);
  end;

var
  frmLogSql: TfrmLogSql;

implementation

{$R *.dfm}

{ TfrmLogSql }

procedure TfrmLogSql.AddLogSQL(Command, CommandType, Status, CursorType,
  LockType: String; RecordsAffected: Integer);
var
  item : TListItem;
begin
    ListViewSQL.Items.BeginUpdate;
  try
    item:=ListViewSQL.Items.Add;
    item.Caption:=FormatDateTime('DD/MM/YYYY HH:NN:SS.ZZZ',Now);
    item.SubItems.Add(CommandType);
    item.SubItems.Add(Command);
    item.SubItems.Add(Status);
    item.SubItems.Add(IntToStr(RecordsAffected));
    item.SubItems.Add(CursorType);
    item.SubItems.Add(LockType);
  finally
    ListViewSQL.Items.EndUpdate;
  end;
  ListViewSQL.Items.Item[ListViewSQL.Items.Count-1].MakeVisible(false); //Scroll to the last line
end;

procedure TfrmLogSql.ListViewSQLChange(Sender: TObject; Item: TListItem;
  Change: TItemChange);
begin
    if ListViewSQL.Selected<>nil then
    SynEdit1.Lines.Text:=ListViewSQL.Selected.SubItems[1];
end;

end.
