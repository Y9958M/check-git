unit uMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Data.Win.ADODB, Vcl.StdCtrls,
  Vcl.ComCtrls, Vcl.Menus, Vcl.Grids, Vcl.DBGrids, Vcl.ExtCtrls, Vcl.Buttons,
  ComObj,Clipbrd, frxClass, frxDBSet, frxExportText, frxExportPDF,
  frxExportXLS,frxCross, frxDesgn, frxADOComponents, frxPreview, cxGraphics,
  cxControls, cxLookAndFeels, cxLookAndFeelPainters, cxStyles, dxSkinsCore,
  dxSkinBlack, dxSkinBlue, dxSkinBlueprint, dxSkinCaramel, dxSkinCoffee,
  dxSkinDarkRoom, dxSkinDarkSide, dxSkinDevExpressDarkStyle,
  dxSkinDevExpressStyle, dxSkinFoggy, dxSkinGlassOceans, dxSkinHighContrast,
  dxSkiniMaginary, dxSkinLilian, dxSkinLiquidSky, dxSkinLondonLiquidSky,
  dxSkinMcSkin, dxSkinMoneyTwins, dxSkinOffice2007Black, dxSkinOffice2007Blue,
  dxSkinOffice2007Green, dxSkinOffice2007Pink, dxSkinOffice2007Silver,
  dxSkinOffice2010Black, dxSkinOffice2010Blue, dxSkinOffice2010Silver,
  dxSkinPumpkin, dxSkinSeven, dxSkinSevenClassic, dxSkinSharp, dxSkinSharpPlus,
  dxSkinSilver, dxSkinSpringTime, dxSkinStardust, dxSkinSummer2008,
  dxSkinTheAsphaltWorld, dxSkinsDefaultPainters, dxSkinValentine, dxSkinVS2010,
  dxSkinWhiteprint, dxSkinXmas2008Blue, dxSkinscxPCPainter, cxCustomData,
  cxFilter, cxData, cxDataStorage, cxEdit, cxDBData, cxGridCustomTableView,
  cxGridTableView, cxGridDBTableView, cxGridLevel, cxClasses, cxGridCustomView,
  cxGrid, frxBarcode, cxContainer, cxProgressBar, cxTextEdit, cxNavigator,
  Data.SqlExpr, dxDateRanges, Data.DbxSqlite, frxExportBaseDialog, DBAccess, Uni,
  MemDS, SQLiteUniProvider, UniProvider, MySQLUniProvider;

type
  TMain = class(TForm)
    pgc1: TPageControl;
    ts2: TTabSheet;
    ds1: TDataSource;
    con1: TADOConnection;
    qry1_bak: TADOQuery;
    strngrd1: TStringGrid;
    edt2: TEdit;
    btn2: TButton;
    btn3: TButton;
    pnl1: TPanel;
    btn1: TBitBtn;
    lbl1: TLabel;
    ts1: TTabSheet;
    lbl3: TLabel;
    btn6: TButton;
    dtp3: TDateTimePicker;
    dtp4: TDateTimePicker;
    chk2: TCheckBox;
    lbl4: TLabel;
    cbb3: TComboBox;
    frxrprt1: TfrxReport;
    frxdbdtst1: TfrxDBDataset;
    dlgSave1: TSaveDialog;
    frxlsxprt1: TfrxXLSExport;
    frxpdfxprt1: TfrxPDFExport;
    frxsmpltxtxprt1: TfrxSimpleTextExport;
    frxdsgnr1: TfrxDesigner;
    frxcrsbjct1: TfrxCrossObject;
    lbl7: TLabel;
    lbl8: TLabel;
    lbl9: TLabel;
    cbb1: TComboBox;
    cbb4: TComboBox;
    lbl5: TLabel;
    qry2: TADOQuery;
    btn8: TButton;
    edt1: TEdit;
    ds2: TDataSource;
    qry3: TADOQuery;
    ds3: TDataSource;
    qry4: TADOQuery;
    ds4: TDataSource;
    frxdbdtst2: TfrxDBDataset;
    frxdbdtst3: TfrxDBDataset;
    qry5: TADOQuery;
    ds5: TDataSource;
    frxdbdtst4: TfrxDBDataset;
    btn4: TButton;
    cxgrdbtblvwGrid1DBTableView1: TcxGridDBTableView;
    cxgrdlvlGrid1Level1: TcxGridLevel;
    cxgrd1: TcxGrid;
    frxrprt2: TfrxReport;
    frBarCode1: TfrxBarCodeObject;
    pm1: TPopupMenu;
    mniN1: TMenuItem;
    mniN2: TMenuItem;
    mniN3: TMenuItem;
    mniN4: TMenuItem;
    frxrprt3: TfrxReport;
    cGcGrid1DBTableView1p_No: TcxGridDBColumn;
    cGcGrid1DBTableView1p_C: TcxGridDBColumn;
    cGcGrid1DBTableView1barcode: TcxGridDBColumn;
    cGcGrid1DBTableView1proname: TcxGridDBColumn;
    cGcGrid1DBTableView1spec: TcxGridDBColumn;
    cGcGrid1DBTableView1normalprice: TcxGridDBColumn;
    cGcGrid1DBTableView1unit: TcxGridDBColumn;
    cGcGrid1DBTableView1area: TcxGridDBColumn;
    cGcGrid1DBTableView1proid: TcxGridDBColumn;
    cGcGrid1DBTableView1create_date: TcxGridDBColumn;
    lbl2: TLabel;
    qrySY_bak: TADOQuery;
    ts3: TTabSheet;
    btn5: TButton;
    mmo1: TMemo;
    cmd1: TADOCommand;
    tmr1: TTimer;
    qry0_bak: TADOQuery;
    conQ_bak: TADOConnection;
    mniN5: TMenuItem;
    qryUp: TADOQuery;
    cmdUP: TADOCommand;
    qryI: TADOQuery;
    xprgbarpb1: TcxProgressBar;
    cmd2: TADOCommand;
    Mmo2: TMemo;
    conQ: TUniConnection;
    qrySY: TUniQuery;
    MySQLUniProvider1: TMySQLUniProvider;
    SQLiteUniProvider1: TSQLiteUniProvider;
    sqliteconn1: TUniConnection;
    qry1: TUniQuery;
    qry0: TUniQuery;

    procedure FormCreate(Sender: TObject);

    procedure cbb3KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure btn1Click(Sender: TObject);
    procedure btn2Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure edt2KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure btn3Click(Sender: TObject);
    procedure cbb3DblClick(Sender: TObject);
    procedure dbgrd1TitleClick(Column: TColumn);
    procedure btn6Click(Sender: TObject);
    procedure chk2Click(Sender: TObject);
    procedure dbgrd2TitleClick(Column: TColumn);
    procedure dbgrd3TitleClick(Column: TColumn);
    procedure btn10Click(Sender: TObject);
    procedure edt1KeyPress(Sender: TObject; var Key: Char);
    procedure btn8Click(Sender: TObject);
    procedure btn9Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure cbb1DblClick(Sender: TObject);
    procedure cbb1KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure edt1KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure btn4Click(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure strngrd1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure cbb4KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure cbb4Exit(Sender: TObject);
    procedure cbb4KeyPress(Sender: TObject; var Key: Char);
    procedure edt2KeyPress(Sender: TObject; var Key: Char);
    procedure mniN3Click(Sender: TObject);
    procedure mniN1Click(Sender: TObject);
    procedure mniN4Click(Sender: TObject);
    procedure Downs(sender:Tobject);
    procedure con1BeforeConnect(Sender: TObject);
    procedure btn5Click(Sender: TObject);
    procedure LsTime;
    procedure tmr1Timer(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Main: TMain;
  j,tis1:Integer;
  File_Path:string;


implementation

{$R *.dfm}

uses  writedata, uhelp, UTTH;

procedure TMain.btn10Click(Sender: TObject);
begin
 //writedata.ExportToExcel(dbgrd1);
end;

procedure TMain.btn1Click(Sender: TObject);
var ii,jj:Integer;
begin
cbb1.Text:='';
for ii := 0 to 6 do
for jj := 1 to j do
strngrd1.Cells[ii,jj]:='';     //赋初值
j:=0;    //赋初值
//x:=0;   //赋初值
pnl1.Caption:='0';
strngrd1.RowCount:=1;
//x:=Length(arr);
//SetLength(arr,0);
//SetLength(arr,x);
cbb1.SetFocus;
end;

procedure TMain.btn2Click(Sender: TObject);
//var j:Integer;
begin
  if edt2.Text<>'' then
  begin
    qry1.Close;
    qry1.SQL.Clear;
    qry1.SQL.Text:='select barcode,proname,spec,normalprice,classid,proid,proflag,area,promtflag,measurename from product_barcode';
    qry1.SQL.Add('where barcode='''+edt2.Text+''' or proid= ''' +edt2.Text+'''');
    qry1.Open;
    if qry1.RecordCount >0 then
    begin
      if qry1.FieldByName('proflag').AsString='1' then
      begin
        case Application.MessageBox('联营商品，是否录入（ESC取消）','联营商品,友情提示：按ESC取消',MB_OKCANCEL or MB_DEFBUTTON2 or MB_ICONQUESTION) of
        ID_OK:
          begin
            j:=j+1;
            strngrd1.Cells[0,j]:=IntToStr(j);
            strngrd1.Cells[1,j]:=qry1.FieldByName('barcode').AsString;
            strngrd1.Cells[2,j]:=qry1.FieldByName('proname').AsString;
            strngrd1.Cells[3,j]:=qry1.FieldByName('spec').AsString;
            strngrd1.Cells[4,j]:=qry1.FieldByName('normalprice').AsString;
            //if qry1.FieldByName('promtflag').AsInteger<>0 then  strngrd1.Cells[5,j]:='促销';
            strngrd1.Cells[6,j]:=qry1.FieldByName('measurename').AsString;
            strngrd1.Cells[7,j]:=qry1.FieldByName('area').AsString;
            strngrd1.Cells[8,j]:=cbb4.Text;
            strngrd1.RowCount:=strngrd1.RowCount+1;
            edt2.Text:='';
            pnl1.Caption:=IntToStr(j);
            strngrd1.Row:=j;
          end;
        end;
      end else
      begin
        j:=j+1;
        strngrd1.Cells[0,j]:=IntToStr(j);
        strngrd1.Cells[1,j]:=qry1.FieldByName('barcode').AsString;
        strngrd1.Cells[2,j]:=qry1.FieldByName('proname').AsString;
        strngrd1.Cells[3,j]:=qry1.FieldByName('spec').AsString;
        strngrd1.Cells[4,j]:=qry1.FieldByName('normalprice').AsString;
        //if qry1.FieldByName('promtflag').AsInteger<>0 then  strngrd1.Cells[5,j]:='促销';
        strngrd1.Cells[6,j]:=qry1.FieldByName('measurename').AsString;
        strngrd1.Cells[7,j]:=qry1.FieldByName('area').AsString;
        strngrd1.Cells[8,j]:=cbb4.Text;
        strngrd1.RowCount:=strngrd1.RowCount+1;
        strngrd1.Row:=j;
        edt2.Text:='';
        pnl1.Caption:=IntToStr(j);
      end;
    end else
    begin
      edt2.Text:='';
      ShowMessage('无此商品条码！');
    end;
  end;
  if strngrd1.RowCount>1 then
  strngrd1.FixedRows:=1;
end;

procedure TMain.btn3Click(Sender: TObject);
var i:Integer;
begin
  if cbb1.Text='' then
  begin
    ShowMessage('请输入货架编号！');
  end else
  begin
    qry1.Close;
    qry1.SQL.Clear;
    qry1.SQL.Text:='select nid,P_No,p_C,proname,normalprice,spec,barcode,measurename,area,promtflag from v_cp where P_No='''+trim(cbb1.Text)+''' order by p_C';
    qry1.Open;
    if qry1.RecordCount>0 then
    begin
      //if Application.MessageBox('货架编号存在，是否调出？','调出提示',mb_OkCancel+mb_IconQuestion)=IDOk  then
      case Application.MessageBox('货架编号存在  是:覆盖，否:调出','请先调出，按否',MB_YESNOCANCEL or MB_DEFBUTTON3 or MB_ICONQUESTION) of
        ID_YES:
        begin
          if j<1 then
          begin
            ShowMessage('NO,你真厉害，这样也行！');
          end else
          begin
          qry1.Close;
          qry1.SQL.Clear;
          qry1.SQL.Text:='delete from checks where p_no='''+trim(cbb1.Text)+'''';
          qry1.ExecSQL;
            for I := 1 to j do
            begin
              qry1.Close;
              qry1.SQL.Text:='select nid,p_no,p_c,barcode,create_date from checks where 1=2';
              qry1.Open;
              qry1.Insert;
              qry1.FieldByName('nid').AsString:=strngrd1.Cells[0,i];
              qry1.FieldByName('P_No').AsString:=cbb1.Text;
              qry1.FieldByName('p_C').AsString:=strngrd1.Cells[8,i];
              qry1.FieldByName('barcode').AsString:=strngrd1.Cells[1,i];
              qry1.FieldByName('create_date').AsDateTime:=Now();
              qry1.Post;
            end;
          end;
        Main.btn1.Click;
        edt2.Text:='';
        cbb1.Text:='';
        cbb4.Text:='1';
        end;
        ID_NO:
        begin
          //strngrd1.Row:=strngrd1.RowCount-1;
          //strngrd1.RowCount:=strngrd1.RowCount+1;
          j:=qry1.RecordCount;
          for I := 1 to qry1.RecordCount do
            begin
              strngrd1.Cells[0,i]:=IntToStr(i);
              strngrd1.Cells[1,i]:=qry1.FieldByName('barcode').AsString;
              strngrd1.Cells[2,i]:=qry1.FieldByName('proname').AsString;
              strngrd1.Cells[3,i]:=qry1.FieldByName('spec').AsString;
              strngrd1.Cells[4,i]:=qry1.FieldByName('normalprice').AsString;
              //if qry1.FieldByName('promtflag').AsInteger<>0 then  strngrd1.Cells[5,j]:='促销';
              strngrd1.Cells[6,i]:=qry1.FieldByName('measurename').AsString;
              strngrd1.Cells[7,i]:=qry1.FieldByName('area').AsString;
              strngrd1.Cells[8,i]:=qry1.FieldByName('P_C').AsString;
              qry1.Next;
            end;
          //strngrd1.Row:=j-1;
          strngrd1.RowCount:=qry1.RecordCount+1;
          edt2.SetFocus;
        end;
        ID_CANCEL:;
      end;
      //ShowMessage('货架编号：'+edt2.Text+'存在。'+qry1.FieldByName('create_date').AsString );
      //edt1.Text:='';
    end
    else
    begin
      if j<1 then
      begin
        ShowMessage('亲，还没扫商品呢！');
      end else
      begin
        for I := 1 to j do
        begin
          qry1.Close;
          qry1.SQL.Text:='select nid,p_no,p_c,barcode,create_date from checks where 1=2';
          qry1.Open;
          qry1.Insert;
          qry1.FieldByName('nid').AsString:=strngrd1.Cells[0,i];
          qry1.FieldByName('P_No').AsString:=cbb1.Text;
          qry1.FieldByName('p_C').AsString:=strngrd1.Cells[8,i];
          qry1.FieldByName('barcode').AsString:=strngrd1.Cells[1,i];
          qry1.FieldByName('create_date').AsDateTime:=Now();
          qry1.Post;
        end;
      end;
      Main.btn1.Click;
      edt2.Text:='';
      cbb1.Text:='';
      cbb4.Text:='1';
    end;
  end;
end;


procedure TMain.btn4Click(Sender: TObject);
begin
  if cbb1.Text<>'' then
  begin
    qry1.Close;
    qry1.SQL.Clear;
    qry1.SQL.Text:='select P_No from checks where P_No='''+trim(cbb1.Text)+'''';
    qry1.Open;
    if qry1.RecordCount>0 then
    begin
      case Application.MessageBox(PChar('真的要删除此货架，里面还有 '+inttostr(qry1.RecordCount)+' 条商品信息？'),'警告',MB_ICONWARNING or MB_OKCANCEL or MB_DEFBUTTON2) of
      ID_OK:
        begin
          qry1.Close;
          qry1.SQL.Clear;
          qry1.SQL.Text:='delete from checks where p_no='''+trim(cbb1.Text)+'''';
          qry1.ExecSQL;
        end;
      end;
    end else
    begin
      ShowMessage('无此货架！');
    end;
  end else
  begin
    ShowMessage('货架编号为空！');
  end;
  edt2.SetFocus;
end;

procedure TMain.btn5Click(Sender: TObject);
begin
  try
    mmo1.Lines.Add('尝试连接 商品资料库  按网速情况 请等待一分钟左右````');
    conQ.Connected:=True;
    //contest.Connected:=True;
  except
    ShowMessage('请插上网线，确保能PING通服务器。无法解决请联系信息员');
    mmo1.Lines.Add('无法连接商品资料库！'+datetimetostr(now()));
    //Application.Terminate;
  end;
  if conQ.Connected=True then
  //if contest.Connected then
  begin
    mmo1.Lines.Add('成功 连接');
    UTTH.TTH.Create;
    //if writedata.DBTableExists('v_y_pro',conQ) then   //判断表是否存在
    //if writedata.DBTableExists('v_y_pro',contest) then UTTH.TTH.Create;  //判断表是否存在
    //begin
    //  UTTH.TTH.Create;  //判断表是否存在
    //end else
    //begin
    // ShowMessage('v_y_pro视图不存在');
    // Application.Terminate;
    //end;
  end;
end;

procedure TMain.Downs(sender:Tobject);//更新商品资料
var
i,i1,i2:Integer;
d0,d1,d2,sqltext,tmpsql:string;     //当天日期

begin
  d0:='第一次使用';
  d1:=datetostr(Now());
  d2:=DateTimeToStr(Now());
  sqltext:= 'select lastDate,isdo,ps from product_u ORDER BY ID DESC LIMIT 1;';
  tmpsql:='insert into product_u (lastDate,isdo,ps) values('''+d2+ ''',0,'''+d1+''')';
  i2:=0;
  xprgbarpb1.Properties.Min:=0;
  xprgbarpb1.Properties.Max:=0;
  xprgbarpb1.Position:=0;
  //pb1.Min:=0;
  //pb1.Max:=0;
  //pb1.Position:=0;
  //以上初始化

  //qry1.Close;
  //qry1.SQL.Clear;
  //qry1.SQL.Add('select max(ID) maxID from product_u; ');
  //qry1.Open;
  //if qry1.RecordCount = 1 then
  //begin
  //  maxId := qry1.FieldByName('maxID').AsInteger +1;
  //end;
  try
  qry1.Close;
  qry1.SQL.Clear;
  qry1.SQL.Text :=sqltext;
  qry1.Open;
  if qry1.RecordCount > 0 then
  begin
    d0:=qry1.FieldByName('lastDate').AsString;
    if qry1.FieldByName('ps').AsString = d1 then
    begin
       tmpsql:='update product_u set isdo =0 where ps = '''+d1+'''';
    end;

  end;
  mmo1.Lines.Add('上次更新时间：'+d0);

  qry1.Close;
  qry1.SQL.Clear;
  qry1.SQL.Add(tmpsql);
  //mmo1.Lines.Add(qry1.SQL.Text);
  qry1.ExecSQL;

  //以上插入更新记录

  qry0.DisableControls;
  mmo1.Lines.Add('    开始更新:'+ datetimetostr(now()) +'    正在初始化,删除全部商品数据，稍=');
  qry0.Close;
  qry0.SQL.Clear;
  qry0.SQL.Text:='delete from product_barcode;';
  qry0.ExecSQL;

  qrySY.Close;
  qrySY.SQL.Clear;
  qrySY.SQL.Text:='SELECT supid,divId,pty3Id,barcode,proname,spec,uName,avgprice,normalprice,proid,area,promtflag,proflag,statusId,udt FROM v_Y_pro ';
  qrySY.Open;

  i1:=qrySY.RecordCount-1;
  xprgbarpb1.Properties.Max:=i1;

  btn5.Caption:='更新商品数：'+inttostr(i1+1);
  btn5.Enabled:=False;

  qry0.Close;
  qry0.SQL.Clear;
  qry0.SQL.Text:='select supid,divid,pty3id,barcode,proname,spec,uname,avgprice,normalprice,proid,area,promtflag,proflag,statusid,udt from product_barcode where 1=2';
  qry0.Open;
  qry0.Append;
  for i := 0 to i1 do
  begin
    qry0.Insert;
    qry0.FieldByName('divId').AsString:=  qrysy.FieldByName('divId').AsString;
    qry0.FieldByName('pty3Id').AsString:=  qrysy.FieldByName('pty3Id').AsString;
    qry0.FieldByName('barcode').AsString:=  qrysy.FieldByName('barcode').AsString;
    qry0.FieldByName('proname').AsString:=  qrysy.FieldByName('proname').AsString;
    qry0.FieldByName('spec').AsString:=     qrysy.FieldByName('spec').AsString;
    qry0.FieldByName('uname').AsString:=     qrysy.FieldByName('uname').AsString;
    qry0.FieldByName('avgprice').AsString:= qrysy.FieldByName('avgprice').AsString;
    qry0.FieldByName('normalprice').AsString:=qrysy.FieldByName('normalprice').AsString;
    qry0.FieldByName('proid').AsString:=    qrysy.FieldByName('proid').AsString;
    qry0.FieldByName('area').AsString:=     qrysy.FieldByName('area').AsString;
    qry0.FieldByName('proflag').AsString:=  qrysy.FieldByName('proflag').AsString;
    qry0.FieldByName('promtflag').AsString:=qrysy.FieldByName('promtflag').AsString;
    qry0.FieldByName('statusId').AsString:=   qrysy.FieldByName('statusId').AsString;
    qry0.FieldByName('supid').AsString:=    qrysy.FieldByName('supid').AsString;
    qry0.FieldByName('udt').AsString:=      qrysy.FieldByName('udt').AsString;
    mmo1.Lines.Add('更新商品：' +qry0.FieldByName('proname').AsString);
    xprgbarpb1.Position:=i;
    i2:=i2+1;
    if i2=2 then
    begin
       //qry0.UpdateBatch();
       qry0.Post;
       i2:=0;
    end;
    if i=i1 then
    begin
      //cmd1.CommandText:='update product_u set isdo=1,lastdate='''+d2+''' where ps='''+d1+'''';
      //cmd1.Execute;
      qry1.Close;
      qry1.SQL.Clear;
      qry1.SQL.Text:= 'update product_u set isdo=1,lastdate='''+d2+''' where ps='''+d1+'''';
      qry1.ExecSQL;
      //qry1.SQL.Text:='select isdo,lastDate,ps from product_u where ps='''+d1+'''';
      //qry1.Open;
      //qry1.Edit;
      //qry1.FieldByName('isdo').AsInteger :=1;
      //qry1.FieldByName('lastdate').AsString := d2;
      //qry1.Post;
      btn5.Enabled:=True;
    end;
    qrySY.Next;
  end;
  //qry0.UpdateBatch();
  qry0.post;
  qry0.EnableControls;

  finally

  end;
  mmo1.Lines.Add('共'+inttostr(i1+1)+'条记录更新完成！');
  mmo1.Lines.Add('    结束时间：'+datetimetostr(now()));
  btn5.Caption:='完成更新';
  btn5.Enabled:=True;
  ts1.TabVisible:=true;
  ts2.TabVisible:=true;
  pgc1.ActivePageIndex:=0;
  edt2.SetFocus;

end;

procedure TMain.btn6Click(Sender: TObject);
var
w1:widestring;
begin
  qry2.Close;
  qry2.SQL.Clear;
  w1:='select nid,id,p_No,p_C,barcode,proname,spec,promtflag ,iif(spec=''散称'',normalprice/2,normalprice) as normalprice,proid,iif(spec=''散称'',''斤'',measurename) as measurename,area ,proflag,classid ,create_date  from V_CP where';
  qry2.SQL.Text:=w1;
  qry3.Close;
  qry3.SQL.Clear;
  qry3.SQL.Text:='select distinct p_no as p_no from v_cp where';

  if dtp3.Visible then
  begin
    case dtp3.DateTime>dtp4.DateTime of
      true:
      begin
        qry2.SQL.Add('create_date>=#'+datetostr(dtp4.datetime) +'# and create_date<= #'+datetostr(dtp3.datetime+1)+'#');
        qry3.SQL.Add('create_date>=#'+datetostr(dtp4.datetime) +'# and create_date<= #'+datetostr(dtp3.datetime+1)+'#');
      end;
      False:
      begin
        qry2.SQL.Add('create_date>=#'+datetostr(dtp3.datetime) +'# and create_date<= #'+datetostr(dtp4.datetime+1)+'#');
        qry3.SQL.Add('create_date>=#'+datetostr(dtp3.datetime) +'# and create_date<= #'+datetostr(dtp4.datetime+1)+'#');
      end;
    end;
  end
  else
  begin
    qry2.SQL.Add('create_date>=#'+datetostr(dtp4.datetime)+'# and create_date<#'+datetostr(dtp4.datetime+1)+'#');
    qry3.SQL.Add('create_date>=#'+datetostr(dtp4.datetime)+'# and create_date<#'+datetostr(dtp4.datetime+1)+'#');
  end;
  if cbb3.Text<>'' then
  begin
    qry2.SQL.Add('and p_no like '''+cbb3.Text+'%''');
    qry3.SQL.Add('and p_no like '''+cbb3.Text+'%''');
  end;
  if edt1.Text<>'' then
  begin
    qry2.SQL.Add('and barcode like ''%'+edt1.Text+'''');
    qry3.SQL.Add('and barcode like ''%'+edt1.Text+'''');
  end;
  qry2.SQL.Add('order by p_no,p_C,nid');
  //mmo2.Text:=qry2.SQL.Text;
  qry2.Open;
  qry3.SQL.Add('order by p_no');
  qry3.Open;
end;

procedure TMain.btn8Click(Sender: TObject);
begin
if j>0 then
begin
j:=j-1;
strngrd1.Cells[0,j+1]:='';
strngrd1.Cells[1,j+1]:='';
strngrd1.Cells[2,j+1]:='';
strngrd1.Cells[3,j+1]:='';
strngrd1.Cells[4,j+1]:='';
strngrd1.Cells[5,j+1]:='';
strngrd1.Cells[6,j+1]:='';
strngrd1.Cells[7,j+1]:='';
strngrd1.Cells[8,j+1]:='';
strngrd1.Cells[9,j+1]:='';
strngrd1.RowCount:=strngrd1.RowCount-1;
end;
edt2.SetFocus;
end;

procedure TMain.btn9Click(Sender: TObject);
begin
//frxrprt1.LoadFromFile('chceks.fp3');
frxrprt1.ShowReport;
end;

procedure TMain.cbb3DblClick(Sender: TObject);
begin
cbb1.Text:='';
end;

procedure TMain.cbb3KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
case Key of
VK_RETURN:btn6Click(nil);
VK_RIGHT:edt1.SetFocus;
VK_DOWN:;
end;
end;

procedure TMain.cbb4Exit(Sender: TObject);
begin
  if cbb4.Text='' then
  cbb4.Text:='1';
end;

procedure TMain.cbb4KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
case Key of
VK_RETURN:edt2.SetFocus;
end;
end;

procedure TMain.cbb4KeyPress(Sender: TObject; var Key: Char);
begin
if not (Key in['1'..'9',#8]) then
key:=#0;
end;

procedure TMain.cbb1DblClick(Sender: TObject);
begin
cbb1.Text:='';
end;

procedure TMain.cbb1KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
case Key of
VK_RETURN:edt2.SetFocus;
VK_UP:edt2.SetFocus;
VK_DOWN:btn3.SetFocus;
end;
end;

procedure TMain.chk2Click(Sender: TObject);
begin
case chk2.Checked of
True:dtp3.Visible:=True;
False:dtp3.Visible:=False;
end;
end;

procedure TMain.con1BeforeConnect(Sender: TObject);
var
constr:string;
begin

constr:='Provider=Microsoft.ACE.OLEDB.12.0;Mode=Share Deny None;Jet OLEDB:Engine Type=6;';
constr:=constr+'User ID=Admin;Jet OLEDB:Database Password="";';
constr:=constr+'Data Source='+file_path+'check.accdb;';
//constr:=constr+'';

con1.ConnectionString:=constr;
end;

procedure TMain.dbgrd1TitleClick(Column: TColumn);
begin
writedata.gdtitle(Column);
end;

procedure TMain.dbgrd2TitleClick(Column: TColumn);
begin
writedata.gdtitle(Column);
end;

procedure TMain.dbgrd3TitleClick(Column: TColumn);
begin
writedata.gdtitle(Column);
end;


procedure TMain.edt1KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
case Key of
VK_RETURN:btn6Click(nil);
VK_LEFT:cbb3.SetFocus;
VK_RIGHT:btn6.SetFocus;
end;
end;

procedure TMain.edt1KeyPress(Sender: TObject; var Key: Char);
begin
//if not (Key in['0'..'9',#8]) then
//key:=#0
end;

procedure TMain.edt2KeyPress(Sender: TObject; var Key: Char);
begin
//if not (Key in['0'..'9',#8]) then
//key:=#0
end;

procedure TMain.edt2KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
case Key of
VK_RETURN:btn2Click(nil);
VK_UP:
begin
if cbb4.ItemIndex>0 then
  cbb4.ItemIndex:=cbb4.ItemIndex-1;
end;
VK_DOWN:
begin
if cbb4.ItemIndex<10 then
  cbb4.ItemIndex:=cbb4.ItemIndex+1;
end;
VK_RIGHT:
begin
  btn8Click(nil);
end;
VK_LEFT:
begin
  cbb1.SetFocus;
end;
end;
end;

procedure TMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Application.Terminate;
end;

procedure TMain.FormCreate(Sender: TObject);
begin
File_Path:=ExtractFilePath(ParamStr(0));
//tis1:=0;

strngrd1.Cols[0].Add('序号');
strngrd1.ColWidths[0]:=40;
//strngrd1.Cells[0,0]:='序号';
strngrd1.Cols[1].Add('条码');
strngrd1.ColWidths[1]:=160;
strngrd1.Cols[2].Add('品名');
strngrd1.ColWidths[2]:=320;
strngrd1.Cols[3].Add('规格');
strngrd1.ColWidths[3]:=70;
strngrd1.Cols[4].Add('正常售价');
strngrd1.ColWidths[4]:=70;
strngrd1.Cols[5].Add(''); //天天苏杭修改--有促销
strngrd1.ColWidths[5]:=70;
strngrd1.Cols[6].Add('单位');
strngrd1.ColWidths[6]:=50;
strngrd1.Cols[7].Add('产地');
strngrd1.ColWidths[7]:=80;
strngrd1.Cols[8].Add('层数');
strngrd1.ColWidths[8]:=40;
//strngrd1.Cols[5].;
strngrd1.Cells[0,1]:='1';
cbb4.ItemIndex:=0;
dtp4.DateTime:=Now();
Main.Caption:= Main.Caption+'   版本号：Ver'+writedata.GetBuildInfo;

end;

procedure TMain.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
case Key of
VK_RETURN:edt2.SetFocus;
end;
end;

procedure TMain.FormShow(Sender: TObject);
var
d1:string;     //当天日期
begin
  d1:=datetostr(Now());
  dtp3.Date:=Now()-7;
  dtp4.Date:=Now();
tmr1.Enabled:=False;
pgc1.ActivePageIndex:=0;
edt2.SetFocus;

  qry1.Close;
  qry1.SQL.Clear;
  qry1.SQL.Text:='select lastdate,isdo,ps from product_u where ps='''+d1+'''';
  qry1.Open;
  if qry1.RecordCount=0 then          //无更新记录 ,插入
  begin
    cmd1.CommandText:='insert into product_U(isdo,ps) values(0,'''+d1+''')';
    cmd1.Execute;
  end

end;

procedure TMain.LsTime;
begin
  tis1:=tis1+1;
  //mmo1.Lines.Add('XX'+inttostr(tis1));
  if tis1=3 then
  begin
    btn5.Enabled:=False;
    tmr1.Enabled:=False;
    Self.btn5.Click;
    tis1:=0;
  end;
end;

procedure TMain.tmr1Timer(Sender: TObject);
begin
  LsTime;
  if tis1=0 then
  tmr1.Enabled:=False;
end;

procedure TMain.mniN1Click(Sender: TObject);
begin
frxrprt2.ShowReport;
end;

procedure TMain.mniN3Click(Sender: TObject);
begin
frxrprt1.ShowReport;
end;

procedure TMain.mniN4Click(Sender: TObject);
begin
frxrprt3.ShowReport;
end;

procedure TMain.strngrd1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
case Key of
VK_RETURN:edt2.SetFocus;
end;
end;

end.
