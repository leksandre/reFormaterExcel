program excel_report;

{$APPTYPE CONSOLE}

uses
  Windows,
  Messages,
  SysUtils,
  Variants,
  //Classes,
  //Graphics,
 // Controls,
  Forms,
  Dialogs,
  Math,
  //StdCtrls,
  //ExtCtrls,
  //ComCtrls,
  comObj,
  ExcelXP
  //,Excel2000,
  //Excel97
    ;
type
    mas=record
      param:Shortstring;
      value:String;
      end;
    OblList=record
       name:string;
       num:integer;
       Lx:integer;
       Ly:integer;
       end;

var patchS: TextFile;
patchSourse,PathFResults:WideString;
left,tegPar,TegParHRON,typle: Shortstring;
    masName,MasVal,sotrh:String;
    MasGen: array [1..1024] of mas;
   // MasInf:array[1..1024, 1..1024] of String;
    //Prom:array[1..1024]of String ;
w,h,lenMas,i,j,mi,lI,li2,y,Ybak,w1,h1,lenghSt,countList:integer;
REct:array [1..1024] of OblList;            excaptWas,sotrB:boolean ;
ArrayData,ArraydDat2,MasInf1,XLApp,Workbook,zona1,zonZ9:Variant;
workSheet,Range: OLEVariant;
   ASheet: ExcelWorksheet;
//////////////////////////////////////

function LinesCount(const Filename: string): Integer;  //единственная процедура которую скатал из нета, считает кол строк в файле
var
  HFile: THandle;
  FSize, WasRead, i: Cardinal;
  Buf: array[1..4096] of byte;
begin
  Result := 0;
  HFile := CreateFile(Pchar(FileName), GENERIC_READ, FILE_SHARE_READ, nil,
    OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0);
  if HFile <> INVALID_HANDLE_VALUE then
  begin
    FSize := GetFileSize(HFile, nil);
    if FSize > 0 then
    begin
      Inc(Result);
      ReadFile(HFile, Buf,4096 {32786}, WasRead, nil);
      repeat
        for i := WasRead downto 1 do
          if Buf[i] = 10 then
            Inc(Result);
        ReadFile(HFile, Buf, 4096, WasRead, nil);
      until WasRead = 0;
    end;
  end;
  CloseHandle(HFile); 
end;

procedure createLitDinM; // мас для значений под таблицей
begin     //   if li <= countList then begin
    zona1:=WorkBook.WorkSheets[li].Cells[w1+1,h1];    //проброс области в переменный масс
    zonZ9:=WorkBook.WorkSheets[li].Cells[w1+(w-w1),h1+lenMas-1];
    Range := WorkBook.WorkSheets[li].Range[zonA1, zonZ9];

    ArraydDat2 := VarArrayCreate([1, (w-w1)+1, 1, lenMas+1], varVariant);
    ArraydDat2:=Range.Value;     //    end;
end;

Function Ok(sotr:string;newp:char):boolean; // надо ли разбивать строку иф встречается запятая, (типа если запятая между "" тогда считаеться за один элемент)
begin
result:=true;
if newp=',' then
begin
if ((trim(sotr))[1]='"') then   
 begin
 if  ((trim(sotr))[Length(trim(sotr))]='"') then result:=false
                                            else result:=true;
 end


                           else   result:=false;
end           else   result:=true;
if (newp=',') and (sotrh=',')  then result:=sotrB;
sotrh:=newp; sotrB:=result;
end;

//////////////////////////////////////
function connect:boolean;//типа есть ли файлы вобще
var confFile:TextFile;
begin
Result:=False;
if not FileExists('XLReport.ini') then  begin  ShowMessage('нужен файл "XLReport.ini"'); exit; end
else
        begin
AssignFile(confFile,'XLReport.ini');
Reset(confFile) ;
ReadLn(confFile,left);   patchSourse:='________';      PathFResults:='________';
         while ('dataset' <>AnsiLowerCase(copy(trim(patchSourse),1,7))) and  not eof(confFile)  do
         ReadLn(confFile, patchSourse);
         while ('templat' <>AnsiLowerCase(copy(trim(PathFResults),1,7))) and  not eof(confFile)   do
         ReadLn(confFile, PathFResults);
CloseFile(confFile) ;
patchSourse:=copy(patchSourse,Pos('=',patchSourse)+1,length(patchSourse)-Pos('=',patchSourse)+1);
PathFResults:=copy(PathFResults,Pos('=',PathFResults)+1,length(PathFResults)-Pos('=',PathFResults)+1);
if not FileExists(patchSourse) then begin ShowMessage('нету файла источника "'+patchSourse+'"'); exit; end
else if not FileExists(PathFResults) then begin ShowMessage('нету файла шаблона "'+PathFResults+'"'); exit; end
else Result:=true;        end;
end;

procedure RUNd;  // поехали
var testS:textFile;  hd:HWND;
begin
CmdShow:=0;

//Application.FreeOnRelease;
//if CmdShow = SW_SHOWMINNOACTIVE then WindowState := wsMinimized;

//ShowWindow(hd, cmdShow);
//PostMessage(hd, WM_SYSCOMMAND, SC_MINIMIZE, 0);
 //  sleep(1000);

 lenghSt:=LinesCount(patchSourse);
AssignFile(patchS,patchSourse);
CoInitializeEx(Nil, 0);
try
    XLApp := GetActiveOleObject('Excel.Application');
  except
    try
      XLApp := CreateOleObject('Excel.Application');
    except
      ShowMessage('vs не мочь запустить Excel? Проверьте установлен ли он?');
      Exit;
   end;
   end;
     // ShellExecute(hInstance,'open',PANsiChar(PathFResults),NiL, NiL, SW_SHOWNORMAL);
  XLApp := GetActiveOleObject('Excel.Application');
  XLApp.Application.EnableEvents := false;

  try
  Workbook := XLApp.Workbooks.open(PathFResults);
  except
  MessageDlg('ошибка открытия документа, перезапустите',mtInformation,[mbOK,mbYes,mbHelp],0);
  Exit;
  end;  
  Reset(patchS);

  {countList:=WorkBook.Sheets.Count-1;
  if CountList=(-1) then begin showMessage('отсутствуют листы в документе'); exit; end;
  if CountList=(0) then countList:=1; }

  //ускорямся
PostMessage(Application.Handle, WM_SYSCOMMAND, SC_MINIMIZE, 0);
PostMessage((FindWindow(nil, pAnsiChar('Microsoft Execel - '+ExtractFileName(PathFResults)))), WM_SYSCOMMAND, SC_MINIMIZE, 0);
// запретить перерисовку экрана
XLapp.ScreenUpdating := False;
// отменить автоматическую калькуляцию формул
XLapp.Calculation := xlCalculationManual;
// отменить проверку автоматическую ошибок в ячейках (для XP и выше)
XLapp.ErrorCheckingOptions.BackgroundChecking := False;
XLapp.ErrorCheckingOptions.NumberAsText := False;
XLapp.ErrorCheckingOptions.InconsistentFormula := False;
//PostMessage(XLapp.Handle, WM_SYSCOMMAND, SC_MINIMIZE, 0);



end;

Procedure TyplZ; //  еденичные параметры или таблица значений
begin
left:='';
while (AnsiUpperCase(copy(trim(left),1,5))<>'RANGE') and  (not eof(patchS)) do
ReadLn(patchS,left);
if length(left)<=6 then typle:='only' else typle:=copy(left,7,length(left)-6);
end;

Procedure ParVal;//В смысле Parametr(Name), Value
Begin
      masname:=masname+'_____________________';    masval:= '';
        while  ('field' <> AnsiLowerCase(copy(trim(masName),1,5)))  and( masname[1] <> '[' ) and  (not eof(patchS)) do
        ReadLn(patchS, masname);
        if 'field' = AnsiLowerCase(copy(trim(masName),1,5)) then
        while  (masval = '') and  not eof(patchS)  do ReadLn(patchS, masval);
        if (masVal[1]<> '[') and  not eof(patchS) then

      masName:=copy(masName,Pos('=',masName)+1,length(masName)-Pos('=',masName)+1);
      masVal:=copy(masVal,Pos('=',masVal)+1,length(masVal)-Pos('=',masVal)+1);
end;

procedure workMasgen;   //работа с масом masGen хранит строку с патраметрами и 1ю строку значеий из текстовика
begin
      j:=1;   // зануляем массив (параметр значение)
      for i:=1 to lenMas do
      begin MasGen[i].param:=''; MasGen[i].value:='';end;

      //забиваем значения в единый массив
      for i:=1 to length(masName) do
      if  ok(MasGen[j].param,masName[i])
      then  MasGen[j].param:=MasGen[j].param + masName[i]
      else inc(j);

      lenMas:=j; j:=1;

      for i:=1 to length(MasVal) do
      if ok(MasGen[j].value,MasVal[i])
      then  MasGen[j].value:=MasGen[j].value + MasVal[i]
      else
      begin
       if MasGen[j].value[1]='"'
       then MasGen[j].value[1]:=' ';

       if MasGen[j].value[Length(MasGen[j].value)]='"'
       then MasGen[j].value[length(MasGen[j].value)]:=' ';

       inc(j);
      end;
      if MasGen[j].value[1]='"'
      then MasGen[j].value[1]:=' ';

      if MasGen[j].value[Length(MasGen[j].value)]='"'
      then MasGen[j].value[length(MasGen[j].value)]:=' ';

      if lenMas<>j then MessageDlg('Ошибка после тега '+tegPar+', количество полей после "FieldNames", и количество значений,'+#13#10+'в первой стрке, не совпадают, значения в резултатах могут не совпадать',mtInformation,[mbOK,mbYes],0);
end;

Procedure  MasCreatExcapt(li1:integer;li2:integer); // Заносим в док табличную часть
var R: OLEVariant;    hy:integer;
begin
    ////////////////////////////Форматирование таблицы
for i:=1 to lenMas do
    begin
      WorkBook.WorkSheets[li1].Cells[w1,h1+i-1].Copy(EmptyParam);
      zona1:=WorkBook.WorkSheets[li2].Cells[w1,h1+i-1];
      zonZ9:=WorkBook.WorkSheets[li2].Cells[w1+y-2,h1+i-1];
      WorkBook.WorkSheets[li2].Range[zonA1, zonZ9].PasteSpecial(xlPasteFormats,xlNone,False,False);
    end;

    zona1:=WorkBook.WorkSheets[li1].Cells[1,h1];          //копируем формат ячеек
    zonZ9:=WorkBook.WorkSheets[li1].Cells[w1,h1+lenMas-1];
 WorkBook.WorkSheets[li1].Range[zonA1, zonZ9].Copy(EmptyParam);

 zona1:=WorkBook.WorkSheets[li2].Cells[1,h1];        //всtавляем формат ячеек
    zonZ9:=WorkBook.WorkSheets[li2].Cells[w1,h1+lenMas-1];
 WorkBook.WorkSheets[li2].Range[zonA1, zonZ9].PasteSpecial(xlPasteALL,xlNone,False,False);


    zona1:=WorkBook.WorkSheets[li2].Cells[w1+y-1,h1];   //заполнение области под  заполненой таблицей
    zonZ9:=WorkBook.WorkSheets[li2].Cells[w1+y+(w-w1)-2,h1+lenMas-1];
    Range := WorkBook.WorkSheets[li2].Range[zonA1, zonZ9];
    Range.value := ArraydDat2;

    if eof(patchS) then begin hy:=y; y:=yBak; end;
    ////////////////////////////     //форматируем
    zona1:=WorkBook.WorkSheets[li1].Cells[w1+y-1,h1];          //копируем формат ячеек
    zonZ9:=WorkBook.WorkSheets[li1].Cells[w1+y+(w-w1)-3,h1+lenMas-1];
 WorkBook.WorkSheets[li1].Range[zonA1, zonZ9].Copy(EmptyParam);
    //////////////////////////////выше и ниже простой копи паст формата ячеек

  if eof(patchS) then begin y:=hy; end;

  zona1:=WorkBook.WorkSheets[li2].Cells[w1+y-1,h1];        //всtавляем формат ячеек
    zonZ9:=WorkBook.WorkSheets[li2].Cells[w1+y+(w-w1)-3,h1+lenMas-1];
 WorkBook.WorkSheets[li2].Range[zonA1, zonZ9].PasteSpecial(xlPasteFormats,xlNone,False,False);
    /////////////////////////////

 {  zona1:=WorkBook.WorkSheets[li1].Cells[w1+y-1,h1];   //заполнение области под  заполненой таблицей
    zonZ9:=WorkBook.WorkSheets[li1].Cells[w1+y+(w-w1)-2,h1+lenMas-1];
    Range := WorkBook.WorkSheets[li1].Range[zonA1, zonZ9];

     zona1:=WorkBook.WorkSheets[li2].Cells[w1+y-1,h1];   //заполнение области под  заполненой таблицей
    zonZ9:=WorkBook.WorkSheets[li2].Cells[w1+y+(w-w1)-2,h1+lenMas-1];
    R := WorkBook.WorkSheets[li2].Range[zonA1, zonZ9]; }

    ybak:=y;
    
end;






Procedure  MasCreat; // Заносим в док табличную часть
begin

   if not(ExcaptWas) then    begin   // XLApp.Workbooks.Add(xlWBATWorksheet, GetUserDefaultLCID);   //LCID
    zona1:=WorkBook.WorkSheets[li].Cells[w1,h1];
    zonZ9:=WorkBook.WorkSheets[li].Cells[w1+y,h1+lenMas-1];
    Range := WorkBook.WorkSheets[li].Range[zonA1, zonZ9];
    Range.Value := ArrayData;   end;
   //arrayData.free;
    ////////////////////////////
   // zona1:=WorkBook.WorkSheets[li].Cells[w1+1,h1];  //   копирование формата нижней строки в незаполненой таблице
   // zonZ9:=WorkBook.WorkSheets[li].Cells[w1+1,h1+lenMas-1];
   // WorkBook.WorkSheets[li].Range[zonA1, zonZ9].Copy(EmptyParam);
    ////////////////////////////
   // zona1:=WorkBook.WorkSheets[li].Cells[w1+y-1,h1];  //вставка этого формата в заполненую
   // zonZ9:=WorkBook.WorkSheets[li].Cells[w1+y-1,h1+lenMas-1];
   // {Range := }WorkBook.WorkSheets[li].Range[zonA1, zonZ9].PasteSpecial(xlPasteFormats,xlNone,False,False);
    ////////////////////////////
   // for i:=1 to lenMas do WorkBook.WorkSheets[li].Cells[w1+y-1,h1+i-1].formula := Prom[i];  // здесь пробивали строчку после значений таблицы, процедура была заменена на копирование всей активной области после строки с табл параметрами.
   ////////////////////////////
    zona1:=WorkBook.WorkSheets[li].Cells[w1+y-1,h1];   //заполнение области под  заполненой таблицей
    zonZ9:=WorkBook.WorkSheets[li].Cells[w1+y+(w-w1)-2,h1+lenMas-1];
    Range := WorkBook.WorkSheets[li].Range[zonA1, zonZ9];
    Range.value := ArraydDat2;
    ////////////////////////////     //форматируем
    zona1:=WorkBook.WorkSheets[li].Cells[w1+1,h1];          //копируем формат ячеек
    zonZ9:=WorkBook.WorkSheets[li].Cells[w1+(w-w1)-1,h1+lenMas-1];
    WorkBook.WorkSheets[li].Range[zonA1, zonZ9].Copy(EmptyParam);
    ////////////////////////////                             //выше и ниже простой копи паст формата ячеек
    zona1:=WorkBook.WorkSheets[li].Cells[w1+y-1,h1];        //всtавляем формат ячеек
    zonZ9:=WorkBook.WorkSheets[li].Cells[w1+y+(w-w1)-3,h1+lenMas-1];
    WorkBook.WorkSheets[li].Range[zonA1, zonZ9].PasteSpecial(xlPasteFormats,xlNone,False,False);
    /////////////////////////////



    ////////////////////////////Форматирование таблицы
for i:=1 to lenMas do
    begin
      WorkBook.WorkSheets[li].Cells[w1,h1+i-1].Copy(EmptyParam);
      zona1:=WorkBook.WorkSheets[li].Cells[w1,h1+i-1];
      zonZ9:=WorkBook.WorkSheets[li].Cells[w1+y-2,h1+i-1];
      WorkBook.WorkSheets[li].Range[zonA1, zonZ9].PasteSpecial(xlPasteFormats,xlNone,False,False);
    end;
end;

Procedure FindRectr(var Num:Integer; var x:Integer; var y:Integer);   //поиск активной области на листе
var R: variant; 
begin

 //if li <= countList then
    if (rect[Num].Lx=0) or (rect[Num].Ly=0)  then
    begin
     r:=Workbook.Worksheets[Num].Cells.SpecialCells(xlCellTypeLastCell,EmptyParam);
     x:=r.row;
     y:=r.column;
     rect[Num].Lx:=x;
     rect[Num].Ly:=y;
     rect[Num].num:=num;
     rect[Num].name:=Workbook.Worksheets[Num].Name;
    end else
    begin
     x:=rect[Num].Lx;
     y:=rect[Num].Ly;
    end;
end;

Procedure CellsProd ; //Работа с массом содержащим значения ячеек листа до обработки
Begin
if li <> li2 then
begin 
     // for i:=1 to w do for j:=1 to h  do MasInf1[i,j]:=''; // зануляем массив (Хранящий значения ячеек в активной области листа Excel)

            FindRectr(li,w,h);//поиск активной области
      // забивка значений ячеек в освобожденный массив
   MasInf1 := VarArrayCreate([1, w, 1, h], varVariant);
     { zona1:=WorkBook.WorkSheets[li].Cells[1,1];
    zonZ9:=WorkBook.WorkSheets[li].Cells[w,h];
    Range := WorkBook.WorkSheets[li].Range[zonA1, zonZ9];
   MasInf1:= Range.value ; }

  for i:=1 to w do for j:=1 to h  do MasInf1[i,j]:=Workbook.workSheets[li].Cells[i,j].Formula;
  li2:=li;

end;
End;

function seashLF:boolean; //поиск координат левого верхнего элемента таблицы
var et:string;   R: variant;//ExcelRange;
begin
et:=('='+AnsiLowerCase(copy(tegPar,2,length(tegPar)-2))+'_'+MasGen[1].param);
      w1:=0; h1:=0;
      for i:=1 to w do
      for j:=1 to h  do
      if (w1=0) and (h1=0)
      then if et = (AnsiLowerCase(copy(MasInf1[i,j],1,length(tegPar)))+(copy(MasInf1[i,j],length(tegPar)+1,length(MasInf1[i,j])-length(tegPar))))
      then begin w1:=i ; h1:=j;end;

if (w1=0) or (h1=0)
then result:=false else result:=true;
end;

Procedure Excapt;
var hron,NameListGen:String;
    ArraydDat3:Variant;
    CountList,ListGenN,KolList:integer;
begin
  CellsProd;   //Работа с массом содержащим значения ячеек листа до обработки
       if typle<>'only' then  if seashLF then  MasCreat;
    ListGenN:=li;   NameListGen:=Workbook.WorkSheets[li].name;
excaptWas:=true;
Workbook.sheets.Add;  Inc(listGenN);   kolList:=2;

ArraydDat3 := VarArrayCreate([1, MIN(lenghSt,65536), 1, lenMas], varVariant);

    j:=1;    y:=2;
    ReadLn(patchS, masval);
    if  (not eof(patchS))and ((Trim(masVal)[1]) <> '[') then
repeat
  begin
    for i:=1 to length(MasVal) do
      if Ok(Hron,MasVal[i]) and (I<=(length(MasVal)-1))
      then hron:=hron + MasVal[i]
      else
        if (y<=lenghSt) and (j<=lenMas)
        then
        begin

          if hron[1]='"'
          then  hron[1]:=' ';

          if hron[Length(hron)]='"'
          then hron[length(hron)]:=' ';

          ArraydDat3[y,j] := hron;
          inc(j);
          hron:='';
       end;
     inc(y); j:=1;  ReadLn(patchS, masval);
     if ((y+w)>=65536)
     then
     Begin
    zona1:=WorkBook.WorkSheets[li].Cells[w1,h1];
    zonZ9:=WorkBook.WorkSheets[li].Cells[w1+y,h1+lenMas-1];
    Range := WorkBook.WorkSheets[li].Range[zonA1, zonZ9];
    Range.Value := ArraydDat3;

 if ParamCount>0 then if ParamStr(1)='fm' then MasCreatExcapt(ListGenN,li);

   // WorkBook.WorkSheets[li].Move(EmptyParam,Workbook.WorkSheets[Workbook.WorkSheets.Count]);

     WorkBook.WorkSheets[li].Name:=NameListGen+' '+ IntToStr(KolList);
     Workbook.sheets.Add;  Inc(KolList);    Inc(listGenN);


     Y:=2; j:=1;
     end;
  end;
until (masVal[1] = '[') or (eof(patchS)) ;

if eof (patchS) then
Begin //if конец файла добиваем последнюю строку
 for i:=1 to length(MasVal) do  if Ok(Hron,MasVal[i]) and (I<=(length(MasVal)-1))
      then hron:=hron + MasVal[i] else  if (y<=lenghSt) and (j<=lenMas)
        then begin if hron[1]='"' then  hron[1]:=' '; if hron[Length(hron)]='"'
          then hron[length(hron)]:=' '; ArrayData[y,j] := hron;
          inc(j); hron:=''; end;  inc(y);
end;
 
    WorkBook.WorkSheets[li].Name:=NameListGen+' '+ IntToStr(KolList);
    zona1:=WorkBook.WorkSheets[li].Cells[w1,h1];
    zonZ9:=WorkBook.WorkSheets[li].Cells[w1+y,h1+lenMas-1];
    Range := WorkBook.WorkSheets[li].Range[zonA1, zonZ9];
    Range.Value := ArraydDat3;
    
    if ParamCount>0 then if ParamStr(1)='fm' then MasCreatExcapt(ListGenN,li);
  

//  if ParamCount>0 then if ParamStr(1)='fm' then MasCreatExcapt(ListGenN,li);

   //MasCreatExcapt(ListGenN,li);



end;
 //w-row(x); h-column(y);
procedure createDinMas;  //крутой мас для забивки    (два маса, второй для переноса данных)
var hron:String;
begin
    createLitDinM;
    ArrayData := VarArrayCreate([1,MIN(lenghSt,65536), 1, lenMas], varVariant);

    for i:=1 to lenMas do
      ArrayData[1,i] := MasGen[i].value;   //первая строка из маса хранященго проверяющего соответствие параметр - значение

    j:=1;    y:=2;
    ReadLn(patchS, masval);
    if  (not eof(patchS))and ((Trim(masVal)[1]) <> '[') then
repeat
  begin
    for i:=1 to length(MasVal) do
      if Ok(Hron,MasVal[i]) and (I<=(length(MasVal)-1))
      then hron:=hron + MasVal[i]
      else
        if (y<=lenghSt) and (j<=lenMas)
        then
        begin

          if hron[1]='"'
          then  hron[1]:=' ';

          if hron[Length(hron)]='"'
          then hron[length(hron)]:=' ';

          ArrayData[y,j] := hron;
          inc(j);
          hron:='';
       end;
     inc(y); j:=1;  ReadLn(patchS, masval);
  end;
until ((masVal[1] = '[') or (eof(patchS)) or ((y+w)>=65536));

if (eof (patchS))  and (not (excaptWas)) then
Begin //if конец файла добиваем последнюю строку
 for i:=1 to length(MasVal) do  if Ok(Hron,MasVal[i]) and (I<=(length(MasVal)-1))
      then hron:=hron + MasVal[i] else  if (y<=lenghSt) and (j<=lenMas)
        then begin if hron[1]='"' then  hron[1]:=' '; if hron[Length(hron)]='"'
          then hron[length(hron)]:=' '; ArrayData[y,j] := hron;
          inc(j); hron:=''; end;  inc(y);
end ;   //if не конец обрабатываем следующий тег

if not eof (patchS) then   TegParHRON:=(Trim(masVal));
if ((y+w)>=65536) then Excapt;  //if массив больше 65536 тога...


end;

Procedure OnlyParam; //забиваем  то что не в таблицах
begin
      for i:=1 to w do
        for j:=1 to h do
           for mi:=1 to lenMas do
        if length(MasInf1[i,j])>3 then
        if  '='+AnsiLowerCase(copy(tegPar,2,length(tegPar)-2))+'_'+MasGen[mi].param=AnsiLowerCase(copy(MasInf1[i,j],1,length(tegPar)))+(copy(MasInf1[i,j],length(tegPar)+1,length(MasInf1[i,j])-length(tegPar)))
        then
        XLApp.Sheets[li].Cells[i,j].Formula:=MasGen[mi].value;

end;


//формирует файл  с  инфой для проверки корректности работы
{Procedure past; var confFile2:TextFile;
begin
AssignFile(confFile2,'tuy.txt');
Rewrite(confFile2) ;
writeln(confFile2) ;
writeln(confFile2,patchSourse) ;
writeln(confFile2,PathFResults) ;
writeln(confFile2,ExtractFileName(patchSourse));
writeln(confFile2,ExtractFileName(PathFResults));

writeln(confFile2) ;
writeln(confFile2,inttostr(w)+' ') ;
write(confFile2,inttostr(h)+' ') ;
write(confFile2,inttostr(w1)+' ') ;
write(confFile2,inttostr(w1)+' ') ;
writeln(confFile2) ;
for mi:=1 to lenMas do  begin
writeln(confFile2,MasGen[mi].param) ; write(confFile2,MasGen[mi].Value) ;       end;

writeln(confFile2) ;writeln(confFile2) ;     i:=1;j:=1;mi:=1;
writeln(confFile2,AnsiLowerCase(copy(MasInf1[i,j],1,length(tegPar)))+(copy(MasInf1[i,j],length(tegPar)+1,length(MasInf1[i,j])-length(tegPar))));
writeln(confFile2,'='+AnsiLowerCase(copy(tegPar,2,length(tegPar)-2))+'_'+MasGen[mi].param);
writeln(confFile2,left);
writeln(confFile2,typle);
for i:=1 to w do  begin writeln(confFile2);
for j:=1 to h  do begin
write(confFile2,MasInf1[i,j]);
end; end;

Writeln (confFile2);
Writeln(confFile2,TegPar);
Writeln(confFile2,TegParHRON);
Writeln(confFile2,intTostr(lI));  Writeln(confFile2,intTostr(lI2));
CloseFile(confFile2) ;
end;     }
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////

begin
if connect then
begin
rund; li:=1;  li2:=0;  TegParHRON:='0'; excaptWas:=false;
While (not eof(patchS)) do
begin//открытие цикла while (not eof(patchS)) do
tegPar:='  ';
                  if TegParHRON<>'0'
                  then   tegPar:= TegParHRON
                  else
                  while (tegPar[1]<>'[') and (not eof(patchS))  do
                  ReadLn(patchS,tegPar);
TyplZ;
ParVal;

workmasgen; //работа с масом masGen хранит строку с патраметрами и 1ю строку значеий из текстовика
if (typle<>'only')  then  if seashLF then  createDinMas;
//searshF-ищет первый табличный параметр на листе и возвращает его в переменных w1 h1

      //  countList :=  WorkBook.Sheets.Count;
       // countList := WorkBook.WorkSheets.Count;
   //   for lI := 1 to countList do
    //  if li <= countList then
    if (not( excaptWas))  then
      begin

       CellsProd;   //Работа с массом содержащим значения ячеек листа до обработки
       if typle<>'only' then  if seashLF then  MasCreat; // Заносим в док табличную часть
       if typle='only' then  OnlyParam;  //забиваем  то что не в таблицах
      end;
    end;   // завершение while (not eof(patchS)) do
CloseFile(patchS);

XLApp.Application.EnableEvents := true;
XLapp.ScreenUpdating := True;
XLapp.Calculation := xlCalculationAutomatic;
XLapp.ErrorCheckingOptions.BackgroundChecking := False;
XLapp.ErrorCheckingOptions.NumberAsText := False;
XLapp.ErrorCheckingOptions.InconsistentFormula := False;
XLapp.Visible:= true;
PostMessage((FindWindow(nil, pAnsiChar('Microsoft Execel - '+ExtractFileName(PathFResults)))), WM_SYSCOMMAND, SC_MAXIMIZE, 0);

SetForegroundWindow(XLapp.Hwnd);

//past;
end;
end.

   



 { function _PasteSpecial(Paste: XlPasteType; Operation: XlPasteSpecialOperation;
                           SkipBlanks: OleVariant; Transpose: OleVariant; out RHS: OleVariant): HResult; stdcall; }


 //  for i:=1 to lenMas do    begin
      // Range := XLApp.Find(masGen[i].param,xlFormulas,xlNext);
              //  if not VarIsEmpty(Range) then

//WorkSheet.Range['B1', 'C10'].Interior.Color := RGB(223, 123, 123);
               //  Address := Range.Address;    Address:string;
        //   end;

//Sheet := Workbook.Worksheets.Item[1];

 // TableVals := VarArrayCreate([0, queSelectRecCount - 1,  0, queSelectFieldsCount - 1], varOleStr);

  //  XLApp.Range['E2', 'E2'].value formula:= 'Sum(a2:d2)';

//UsedRange := ISheet.UsedRange[xlLCID];
//Range := UsedRange.Find(What:='Text', LookIn := xlValues, SearchDirection := xlNext);

{  TExcelApplication;
 XLApp.Workbooks.Add(-4167);
 XLApp.Workbooks[1].WorkSheets[1].Name:='Отчёт';
 Colum:=XLApp.Workbooks[1].WorkSheets['Отчёт'].Columns;
 Colum.Columns[1].ColumnWidth:=12;
 Colum.Columns[2].ColumnWidth:=17; }
//SetLength  (m1,a,b);
//SetLength  (m2,b,a);
{read(f,m1[i]);
readln(f);
AssignFile(f2,'output.txt');
Rewrite(f2) ;
Write(f2,m1[i][j],m1[i][j]);
writeln(F2);
end;}

