unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, XPMan, ToolWin, StdCtrls, CheckLst, ImgList, ExtCtrls,
  Grids, Menus, Clipbrd,OleServer, ExcelXP,Excel2000,ComObj;

type TMass=array of string; //Массив из строк (стлбцы в нашем сучае)

     TTable=record  //Тип - таблица
      Name:string;  //Название таблицы
      Index:integer; //ее порядковый номер
      Mass:TMass;    //Массив из столбцов, Length которого равен длине строки
     end;

     TTables=array of TTable; //массив из таблиц

type
  TForm1 = class(TForm)
    PageControl1: TPageControl;
    ToolBar1: TToolBar;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    ToolButton5: TToolButton;
    XPManifest1: TXPManifest;
    ImageList1: TImageList;
    ToolButton6: TToolButton;
    ToolButton7: TToolButton;
    Splitter1: TSplitter;
    StatusBar1: TStatusBar;
    ToolButton8: TToolButton;
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    ToolButton9: TToolButton;
    ToolButton10: TToolButton;
    ToolButton11: TToolButton;
    ListView1: TListView;
    OpenDialog1: TOpenDialog;
    SaveDialog1: TSaveDialog;
    N5: TMenuItem;
    N6: TMenuItem;
    N7: TMenuItem;
    N8: TMenuItem;
    N9: TMenuItem;
    N10: TMenuItem;
    N11: TMenuItem;
    XLApp: TExcelApplication;
    OpenDialog2: TOpenDialog;
    ToolButton13: TToolButton;
    ToolButton12: TToolButton;
    procedure NewDoc;   //Очистить всё
    procedure Filtration;  //Фильтрация
    function  inExcl(Str:string; var Ex:TMass):boolean; //сравнение комбинации с исключениями
    procedure ClearGrid(Grid:TStringGrid); //очистка табицы
    procedure SetEditText(Sender: TObject; ACol, ARow: Integer; const Value: String); //здесь делаем так, чтобы нельзя было ввести в ячейку более 1 символа
    procedure GenTables(vrnt,strk,stlb:integer);  //Генерируем случайные таблицы
    procedure FromExcel(FileName:string; var ToGrid:TStringGrid);  //Импорт таблиц из Excel
    procedure ExportToExcel(FileName:string);
    function  CreateTable(Cap:string; Chk:boolean):TTabSheet; //Создаем таблицу
    procedure DeleteTable; //Удаляем
    procedure CreateData;  //Все данные из таблиц переносим в массивы для последующей обработки
    procedure LoadTables(FileName:string);  //сохранение всех таблиц
    procedure SaveTables(FileName:string);  //их загрузка
    procedure ToolButton4Click(Sender: TObject);
    procedure ToolButton5Click(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure ToolButton9Click(Sender: TObject);
    procedure ToolButton11Click(Sender: TObject);
    procedure ToolButton1Click(Sender: TObject);
    procedure ToolButton3Click(Sender: TObject);
    procedure ToolButton2Click(Sender: TObject);
    procedure N7Click(Sender: TObject);
    procedure N6Click(Sender: TObject);
    procedure N8Click(Sender: TObject);
    procedure ToolButton8Click(Sender: TObject);
    procedure ToolButton13Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure N5Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  Tables:TTables; //Массив в котором будут все таблицы
  Exc:TMass;   //Массив исключений, берется из Form2.CheckListBox1
  CurrentFile:string; //текущий открытый файл
implementation

uses Unit2, Unit3;

{$R *.dfm}

{procedure Trial;
var
Buf:PChar;
Size:Cardinal;
Path,FileName:string;
count:integer;
begin
 GetMem(Buf, MAX_PATH);            //Выделяем память для Buf оъемом MAX_PATH (константа в модуле Windows)
 GetSystemDirectory(Buf,MAX_PATH); //эта функция вернет системную папку Windows
 Path:=Buf;
 FreeMem(Buf, MAX_PATH);
 FileName:=Path+'\mstrd.txt';
 if FileExists(FileName)=false then begin
  AssignFile(output,FileName);
  Rewrite(output);
  Writeln(1);
  CloseFile(output);
 end else begin
  AssignFile(input,FileName);
  Reset(input);
  Readln(count);
  CloseFile(input);
  if count>2 then Application.Terminate else begin
   AssignFile(output,FileName);
   Rewrite(output);
   Writeln(count+1);
   CloseFile(output);
  end;
 end;
end;}

procedure Recursia(x,i:integer; str:string; var Ex:TMass); //сюда передаем номер столбца и номер варианта в котором находится этот столбец, текущая комбинация, массив исклюений
var
k,h,len,p,u:integer;
com:string;
begin
str:=str+Tables[i].mass[x];
if i=High(Tables) then begin  //если достигли последнего варианта, то
 h:=Length(Tables[0].mass[0]);    //количество элементов
 len:=Length(str) div h;       //Длина элемента
//У нас здесь получается так:
//если например есть три таблицы, по 3 строки в каждом, то
//получается например str=111222121, при h=3, len=3
//теперь далее мы будем разбивать str
//в нашем примере это будет выглядеть так:
//com=121,потом будет com=122  и следующий com=121
 for k:=0 to h-1       do begin
  for p:=0 to len-1     do begin
   u:=h*p+k;
   com:=com+str[u+1];
  end;
  if Length(Ex)>0 then begin
   if Form1.InExcl(com,Ex)=false then writeln(com);
  end else writeln(com);
 Com:='';
 application.ProcessMessages;
 end;
 exit;
end;
//спускаемся на уровень ниже
 for k:=0 to High(Tables[i+1].Mass) do Recursia(k,i+1,str,Ex);
end;

function TForm1.inExcl(Str:string; var Ex:TMass):boolean;
var
i:integer;
begin
//проверка строки на исключение
 for i:=0 to High(Ex) do
  if Str=Ex[i] then begin
   Result:=true;
   exit;
  end;
Result:=false;
end;

function GetExclusions:TMass;
var
i:integer;
Res:TMass;
begin
//из Form2.CheckListBox'a1 массив исключений
for i:=0 to Form2.CheckListBox1.Count-1 do
 if Form2.CheckListBox1.Checked[i]=true then begin
  SetLength(Res,Length(Res)+1);
  Res[High(Res)]:=Form2.CheckListBox1.Items[i];
 end;
Result:=Res; 
end;

procedure TForm1.Filtration;
var
i:integer;
Exc:TMass;
Bar:TProgressbar;
begin
Bar:=TProgressbar.Create(Statusbar1); //динамически создаем Progressbar
Statusbar1.Panels[1].Text:='перебор вариантов...';
application.ProcessMessages;
try
 With Bar do begin
  width:=StatusBar1.Panels[0].Width;
  height:=Statusbar1.Height-4;
  top:=(Statusbar1.Height-height) div 2;
  left:=1;
  Max:=Length(Tables[0].Mass);
  Parent:=Statusbar1;
 end;
Exc:=GetExclusions;

//файл Report.txt для записи полученных данных
AssignFile(output,ExtractFilePath(ParamStr(0))+'Report.txt');
Rewrite(output);
for i:=0 to High(Tables[0].Mass) do begin
 Bar.Position:=Bar.Position+1;
 Recursia(i,0,'',Exc);
end;
finally
 closefile(output);
 Bar.Free;
 Statusbar1.Panels[1].Text:='';
end;
end;

procedure PasteToColumn(var Buf:string; x,count:integer; Var Sheet:variant);
var
IR2,IR1: ExcelRange;
begin
  Clipboard.AsText:=Buf;
  IDispatch(IR1):=Sheet.Cells.Item[1, x];
  IDispatch(IR2):=Sheet.Cells.Item[count, x];
  OLEVariant(Sheet.Range[IR1, IR2]).PasteSpecial;
  Clipboard.Clear;
  Buf:='';
end;

procedure TForm1.ExportToExcel(FileName:string);
var
XLApp,Sheet:Variant;
index,x:Integer;
str,Buf:string;
begin
try
Statusbar1.Panels[0].Text:='Экспорт данных в Excel...';
application.ProcessMessages;
Buf:='';
XLApp:= CreateOleObject('Excel.Application');  //Создаем Excel Application
XLApp.Visible:=true;                           //Показываем его
XLApp.Workbooks.Add(-4167);                    //добавляем книгу
XLApp.Workbooks[1].WorkSheets[1].Name:='Отчёт';
Sheet:=XLApp.Workbooks[1].WorkSheets['Отчёт'];

 AssignFile(input,filename);
 Reset(input);                //открываем файл на чтение
 index:=1;
 x:=1;
  while not eof(input) do begin
   Readln(str);
   Buf:=Buf+str+#9#10;
   if index mod 65536=0 then begin //если дошли до конца столбца (Excel), то копируем то что в буфере в этот столбец
    PasteToColumn(Buf,x,65536,Sheet);
    x:=x+1;  //далее будем добавлять в другой столбец
    index:=0;
    application.ProcessMessages;
   end;
   index:=index+1;
  end;

  PasteToColumn(Buf,x,index,Sheet); //если что то осталось, то вставляем оставшееся

  CloseFile(input);
  Statusbar1.Panels[0].Text:='';
except
 Statusbar1.Panels[0].Text:='';
 CloseFile(input);
 showmessage('Ошибка при экспорте данных');
end;
end;

procedure Sort;
var
i:integer;
Tab:TTable;
flag:boolean;
begin
//сортируем варинты по индексам в порядке возрастания
repeat
flag:=true;
 for i:=0 to High(Tables)-1 do
  if Tables[i].Index>Tables[i+1].Index then begin
   Tab:=Tables[i];
   Tables[i]:=Tables[i+1];
   Tables[i+1]:=Tab;
   flag:=false;
  end;
until flag=true;
end;

procedure TForm1.NewDoc;
var
i:integer;
begin
for i:=PageControl1.PageCount-1 downto 0 do PageControl1.Pages[i].Free;
 SetLength(Tables,0);
 Listview1.Clear;
 Statusbar1.Panels[0].Text:='';
end;

procedure TForm1.SaveTables(FileName:string);
var
i,j,k:integer;
begin
 AssignFile(output,FileName);
 Rewrite(output);
 Writeln(Listview1.items.count);
 SetLength(Tables,0);
 CreateData;
 for k:=0 to High(Tables) do begin
 Writeln(Tables[k].Name);
 Writeln(Tables[k].index);
 Writeln(integer(Listview1.Items[k].Checked));
 Writeln(High(Tables[k].Mass));
  for i:=0 to High(Tables[k].Mass) do begin
 Writeln(Length(Tables[k].Mass[i]));
   for j:=1 to Length(Tables[k].Mass[i]) do begin
    writeln(Tables[k].mass[i][j]);
   end;
  end;
 end;
 Writeln(Form2.CheckListBox1.Count);
 for i:=0 to Form2.CheckListBox1.Count-1 do begin
  writeln(Form2.CheckListBox1.items[i]);
  writeln(integer(Form2.CheckListBox1.Checked[i]));
 end;
 closefile(output);
end;

procedure TForm1.LoadTables(FileName:string);
var
i,j,k,x,y,count,h1,h2:integer;
n:string;
Page:TTabSheet;
Grid:TStringGrid;
begin
 AssignFile(input,filename);
 Reset(input);
 Readln(count);
 for k:=0 to Count-1 do begin
  Readln(n); //Name
  Readln(i); //Index
  Readln(j); //Checked
  Readln(h1);//High(Tables[k].Mass)
  Page:=CreateTable(n,Bool(j));
  Listview1.Items[Listview1.Items.Count-1].SubItems.text:=inttostr(i);
  Grid:=TStringGrid(PageControl1.Pages[k].FindComponent('Grid'));
  for x:=0 to h1 do begin
  Readln(h2);  //Length(Tables[k].Mass[i])
   for y:=1 to h2 do begin
    Readln(n);
    Grid.Cells[x,y-1]:=n;
   end;
  end;
 end;
 Readln(count);
 for i:=1 to count do begin
  Readln(n);
  Readln(j);
  Form2.CheckListBox1.Items.Add(n);
  Form2.CheckListBox1.Checked[Form2.CheckListBox1.Count-1]:=Bool(j);
 end;
CloseFile(input);
end;

function TForm1.CreateTable(Cap:string; Chk:boolean):TTabSheet;
var
Page:TTabSheet;
Listitem:TListitem;
begin
 Listitem:=Listview1.Items.Add;
 Listitem.Caption:=Cap;
 Listitem.SubItems.Add(inttostr(Listview1.Items.Count));
 Listitem.Checked:=Chk;
  Page:=TTabSheet.Create(PageControl1);
   with Page do begin  //создаем страницу
      PageControl := PageControl1;
      Caption := Cap;
   end;
  Result:=Page;
  with TStringGrid.create(Page) do begin  //создаем таблицу
   fixedcols:=0;
   fixedrows:=0;
   DefaultColWidth:=20;
   DefaultRowHeight:=20;
   RowCount:=1000;
   ColCount:=1000;
   Align:=alclient;
   Name:='Grid';
   OnSetEditText:=SetEditText;
   Options:=[goEditing,goHorzLine,goVertLine,goAlwaysShowEditor];
   Parent:=Page;
  end;
  Statusbar1.Panels[0].Text:='Вариантов: '+inttostr(Listview1.Items.Count);
end;

procedure TForm1.SetEditText(Sender: TObject; ACol, ARow: Integer; const Value: String);
begin
if Value<>'' then
(Sender as TStringGrid).Cells[Acol,Arow]:=Value[1];
end;

procedure TForm1.DeleteTable;
begin
 if Listview1.Selected<>nil then begin
  PageControl1.Pages[Listview1.Itemindex].Free;
  Listview1.DeleteSelected;
 end;
end;

function GridToTable(Grid:TStringGrid; Ind:integer):TTable;
var
i,x,y:integer;
str:string;
Tab:TTable;
begin
//копируем данные из таблицы в массив
Tab.Name:=Form1.Listview1.Items[Ind].Caption;
for x:=0 to Grid.ColCount-1 do begin
 SetLength(Tab.Mass,Length(Tab.Mass)+1);
 for y:=0 to Grid.RowCount-1 do begin
  if (y=0) and (Trim(Grid.Cells[x,y])='') then begin
   SetLength(Tab.Mass,Length(Tab.Mass)-1);
   Result:=Tab;
   exit;
  end else if Trim(Grid.Cells[x,y])='' then break;
   str:=str+Grid.cells[x,y][1];
 end;
 Tab.Mass[x]:=str;
 str:='';
end;
end;

function DataIsRight:boolean;
var
i,j,hi:integer;
begin
//проверка столбцов на одинаковый размер
 hi:=length(Tables[0].mass[0]);
 for i:=0 to High(Tables) do
  for j:=0 to High(Tables[i].mass) do
   if Length(Tables[i].Mass[j])<>hi then begin
    Result:=false;
    exit;
   end;
Result:=true;
end;

procedure TForm1.CreateData;
var
i,j:integer;
Grid:TStringGrid;
Page:TTabSheet;
ListItem:TListItem;
begin
SetLength(Tables,0);
//заносим все данные из таблиц в Tables
for i:=0 to PageControl1.PageCount-1 do //пробегаемся по всем страницам
 if Listview1.Items[i].Checked=true then begin //и если эту страницу нужно посчитать, то
 Grid:=TStringGrid(PageControl1.Pages[i].FindComponent('Grid')); //ищем на ней таблицу
 SetLength(Tables,Length(Tables)+1); //увеличеваем размер tables на 1, для новой таблицы
 Tables[High(Tables)]:=GridToTable(Grid,i); //копируем значения таблицы
 Tables[i].Index:=0; //пока равен нулю
 for j:=0 to Listview1.Items.Count-1 do //далее получаем номер таблицы из Listview'a
 if Listview1.Items[j].Caption=PageControl1.Pages[i].Caption then begin
  Tables[i].Index:=StrToInt(Trim(Listview1.Items[j].SubItems.Text));
  break;
 end;
end;
end;

procedure TForm1.ClearGrid(Grid:TStringGrid);
var
x,y:integer;
begin
  for x:=1 to Grid.ColCount do
    for y:=1 to Grid.RowCount do
    Grid.cells[x,y]:='';
end;

procedure TForm1.GenTables(vrnt,strk,stlb:integer);
var
Page:TTabSheet;
i,x,y:integer;
Grid:TStringGrid;
begin
Randomize;
 for i:=1 to vrnt do begin
  Page:=CreateTable(chr(i+64),true);
  Grid:=TStringGrid(PageControl1.Pages[i-1].FindComponent('Grid'));
  for y:=0 to stlb-1 do
   for x:=0 to strk-1 do
   Grid.Cells[x,y]:=inttostr(random(2)+1);
 end;
end;



procedure TForm1.FromExcel(FileName:string; var ToGrid:TStringGrid);
var
 WorkBk: _WorkBook;
 WorkSheet: _WorkSheet;
 K,R: integer;
 IIndex: OleVariant;
 NomFich: WideString;
begin
ClearGrid(ToGrid);
 NomFich:=FileName;
 IIndex:=1;
 XLApp.Connect;
 // Открываем файл Excel
 XLApp.WorkBooks.Open(NomFich,EmptyParam,EmptyParam,EmptyParam,EmptyParam,
       EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,
                                                EmptyParam,EmptyParam,0);
 WorkBk:=XLApp.WorkBooks.Item[IIndex];
 WorkSheet:=WorkBk.WorkSheets.Get_Item(1) as _WorkSheet;
   for K:=1 to ToGrid.ColCount do begin
    for R:=1 to ToGrid.RowCount do begin
     if (R=0) and (Trim(XLApp.Cells.Item[K,R])='') then begin
      XLApp.Quit;
      XLApp.Disconnect;
      exit;
     end else if Trim(XLApp.Cells.Item[K,R])='' then break;
     ToGrid.Cells[R-1,K-1]:=XLApp.Cells.Item[K,R];
    end;
   end;
 XLApp.Quit;
 XLApp.Disconnect;
end;

procedure TForm1.ToolButton4Click(Sender: TObject);
var
i:integer;
str:string;
begin
 if InputQuery('Создание таблицы-варианта','Введите название таблицы:',str) then begin
  for i:=0 to Listview1.Items.Count-1 do
   if Listview1.Items[i].Caption=str then begin
    showmessage('Недопустимое значение');
    exit;
   end;
CreateTable(Str,true);
str:=inttostr(Listview1.Items.Count);
  if InputQuery('Ввод номера','Введите порядковый номер варианта:',str) then
   if (TryStrToInt(str,i)=true)  then begin
    Listview1.Items[Listview1.Items.Count-1].SubItems.Text:=inttostr(i);
   end else showmessage('Недопустимое значение');
end;
end;

procedure TForm1.ToolButton5Click(Sender: TObject);
begin
DeleteTable;
end;

procedure TForm1.N3Click(Sender: TObject);
begin
Form1.Close;
end;

procedure TForm1.ToolButton9Click(Sender: TObject);
begin
Form2.show;
end;

procedure TForm1.ToolButton11Click(Sender: TObject);
begin
try
 CreateData;
 if DataIsRight=true then begin
  Sort;
  Filtration;
  if ToolButton12.Down=true then ExportToExcel(ExtractFilePath(ParamStr(0))+'Report.txt');
 end else showmessage('Длина столбцов должна быть одинакова');
except
 showmessage('произошла ошибка');
end;
Statusbar1.Panels[1].Text:='';
end;

procedure TForm1.ToolButton1Click(Sender: TObject);
begin
NewDoc;
end;

procedure TForm1.ToolButton3Click(Sender: TObject);
begin
if SaveDialog1.Execute then begin
 SaveTables(SaveDialog1.FileName+'.dat');
 CurrentFile:=SaveDialog1.FileName+'.dat';
end;
end;

procedure TForm1.ToolButton2Click(Sender: TObject);
begin
if OpenDialog1.Execute then begin
 NewDoc;
 try
 LoadTables(OpenDialog1.Filename);
 CurrentFile:=OpenDialog1.Filename;
 except
  showmessage('Ошибка при открытии файла');
 end;
end;
end;

procedure TForm1.N7Click(Sender: TObject);
begin
if CurrentFile<>'' then SaveTables(CurrentFile) else
 if SaveDialog1.Execute then begin
  SaveTables(SaveDialog1.FileName+'.dat');
  CurrentFile:=SaveDialog1.FileName+'.dat';
 end;
end;

procedure TForm1.N6Click(Sender: TObject);
begin
if OpenDialog1.Execute then begin
 NewDoc;
 try
 LoadTables(OpenDialog1.Filename);
 CurrentFile:=OpenDialog1.Filename;
 except
  showmessage('Ошибка при открытии файла');
 end;
end;
end;

procedure TForm1.N8Click(Sender: TObject);
begin
if SaveDialog1.Execute then begin
 SaveTables(SaveDialog1.FileName+'.dat');
 CurrentFile:=SaveDialog1.FileName+'.dat';
end;
end;

procedure TForm1.ToolButton8Click(Sender: TObject);
var
Grid:TStringGrid;
begin
if OpenDialog2.Execute then begin
 try
 Grid:=TStringGrid(PageControl1.ActivePage.FindComponent('Grid'));
 FromExcel(OpenDialog2.Filename,Grid);
 except
  showmessage('Ошибка при открытии файла');
 end;
end;
end;

procedure TForm1.ToolButton13Click(Sender: TObject);
begin
Form3.show;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
//Trial;
end;

procedure TForm1.N5Click(Sender: TObject);
begin
NewDoc;
end;

end.
