unit UnitMyThread;

interface

uses Classes  ;

type TNotifyEvent = procedure(Sender: TObject) of object;

type TMYThread = class(TThread)
  private
    ExecFun:TNotifyEvent;
  protected
    procedure Execute; override;
  public
    constructor Create;
  end;

procedure MyThread(ExecuteEvent:TNotifyEvent;TerminateEvent:TNotifyEvent=nil);



implementation


{ TSortThread }

procedure MyThread(ExecuteEvent,TerminateEvent:TNotifyEvent);
begin
  with TMyThread.Create do
  begin
    ExecFun:=ExecuteEvent;
    if  @TerminateEvent<>nil then
      OnTerminate := TerminateEvent;
  end;
end;

constructor TMYThread.Create;
begin
  FreeOnTerminate := True;
  inherited Create(False);
end;

procedure TMyThread.Execute;
begin
   try
   ExecFun(self);
   except
   end;
end;

end.
