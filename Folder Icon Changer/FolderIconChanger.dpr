Program FolderIconChanger;

Uses
  Forms,
  windows,
  Unit_Base In 'Unit_Base\Unit_Base.pas' {Form1},
  Vcl.Themes,
  Vcl.Styles,
  GetBigIcon In 'Units\GetBigIcon\GetBigIcon.pas',
  WindowsDarkMode In 'Units\WindowsDarkMode\WindowsDarkMode.pas',
  Translation In 'Units\Translation\Translation.pas',
  Unit_Update In 'Unit_Update\Unit_Update.pas' {Form10},
  Unit_About In 'Unit_About\Unit_About.pas' {Form8},
  FileInfoUtils In 'Units\FileInfoUtils\FileInfoUtils.pas';

{$R *.res}

Var
  HM: THandle;

Function Check: boolean;
Begin
  HM := OpenMutex(MUTEX_ALL_ACCESS, false, 'Folder Icon Changer');
  Result := (HM <> 0);
  If HM = 0 Then
    HM := CreateMutex(Nil, false, 'Folder Icon Changer');
End;

Begin
  If Check Then
    exit;
  SetThreadLocale(1049);
  // ReportMemoryLeaksOnShutdown := True;
  Application.Initialize;
  Application.Title := 'Folder Icon Changer';
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TForm10, Form10);
  Application.CreateForm(TForm8, Form8);
  Form1.CleanTranslationsLikeGlobload;
  Form1.LoadLanguage;
  Form1.globload;
  If Not Form1.IsWindows10Or11 Then
  Begin
    MessageBox(0, Pchar(Form1.LangOnlyWindows.Caption), Pchar(Form1.LangError.Caption), MB_ICONERROR);
    exit;
  End;
  Application.Run;

End.

