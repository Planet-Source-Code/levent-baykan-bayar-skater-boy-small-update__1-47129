program pakextr;

{$APPTYPE CONSOLE}

uses
  SysUtils;
 var levent: file of byte;

    Magic:array [1..4] of byte;
    DirOffSet:longint;
    DirSize:longint;

    FileName:array[1..56] of char;//shortstring;
    offset:Longint;
    size:Longint;


    total : longint;

    WriteData: array[1..65536] of byte;
    toget:longint;
    writeh:longint;
   
    kt:smallint;
    tfile:text;
function RightStr(const AText: string; const ACount: Integer): string;
begin
  Result := Copy(AText, Length(AText) + 1 - ACount, ACount);
end;
procedure Extract(PakName:string;FileIndex:integer);
    var
    rfile,wfile:file of byte;

    fnq:string;
    k:integer;
    label
    jump;
    begin

    assignfile(rfile,pakname);
    reset(rfile);
    BlockRead(rfile,magic,4);
    BlockRead(rfile,DiroffSet,4);
    BlockRead(rfile,DirSize,4);
    seek(rfile,diroffset);
    total:= (dirsize div 64);

    for k:=1 to total do
    begin
           blockread(rfile,filename,56);
           blockread(rfile,offset,4);
           blockread(rfile,size,4);
    if k=FileIndex then goto jump;
    end;
    jump:
    toget:=size;
    seek(rfile,offset);

    fnq:=filename;

    //writeln( rightstr(trim(fnq),4));
    writeln(tfile,'file'+inttostr(fileindex)+rightstr(trim(fnq),4) +'  -->  '+ fnq);
    assignfile(wfile,'file'+inttostr(fileindex)+rightstr(trim(fnq),4));
    writeln('Extracting '+ fnq);
    rewrite(wfile);

    while toget>0 do
      begin
                   //form1.p1.StepBy(1);
                   If toget > 65536 Then
                   begin
                   writeh:=65536;
                   blockread(rfile,WriteData,writeh);
                   blockwrite(wfile,WriteData,writeh);
                   toget:= toget - 65536;
                   end;
                   if toget<=65536 then
                   begin
                   writeh:=toget;
                   blockread(rfile,WriteData,toget);
                   blockwrite(wfile,WriteData,writeh);
                   toget:= 0;
                   end;


     end;
     closefile(rfile);
     closefile(wfile);

     //form1.p1.Position :=0;
     end;
begin
if paramstr(1)='' then
begin
        writeln('Usage: pakextract <pakfilename>');
        halt;
end;
assignfile(levent,paramstr(1));
reset(levent);
try
BlockRead(levent,magic,4);
BlockRead(levent,DiroffSet,4);
BlockRead(levent,DirSize,4);
except
writeln('Reading Error!');
exit;
end;

seek(levent,diroffset);
total:= (dirsize div 64);
closefile(levent);
writeln('PAK file '+paramstr(1) +' total '+ inttostr(total) +' files.');

assignfile(tfile,'filenames.txt');
rewrite(tfile);
for kt:=1 to total do
begin
        extract(paramstr(1),kt);
end;
closefile(tfile);
  { TODO -oUser -cConsole Main : Insert code here }
end.

