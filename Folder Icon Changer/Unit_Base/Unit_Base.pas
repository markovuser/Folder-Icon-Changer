Unit Unit_Base;

Interface

Uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, ShellAPI, ShlObj, IniFiles, ActiveX, StrUtils, Spin,
  FileCtrl, Menus, SHFolder, Registry, ComCtrls, ImgList, PngImage, System.ImageList,
  TlHelp32, Vcl.ExtDlgs, GetBigIcon, WindowsDarkMode, Vcl.CheckLst, Translation,
  System.Threading, Vcl.WinXCtrls, System.IOUtils, FileInfoUtils, Math, ComObj;

Type
  TGRPICONDIRENTRY = Packed Record
    bWidth: Byte;
    bHeight: Byte;
    bColorCount: Byte;
    bReserved: Byte;
    wPlanes: Word;
    wBitCount: Word;
    dwBytesInRes: DWORD;
    nID: Word;
  End;

  PGRPICONDIRENTRY = ^TGRPICONDIRENTRY;

  // Структура заголовка группы иконок
  TGRPICONDIR = Packed Record
    idReserved: Word;
    idType: Word;
    idCount: Word;
  End;

  PGRPICONDIR = ^TGRPICONDIR;

Type
  TListView = Class(ComCtrls.TListView)
  Protected
    Procedure WMVScroll(Var Message: TWMVScroll); Message WM_VSCROLL;
  End;

Type
  TWmMoving = Record
    Msg: Cardinal;
    fwSide: Cardinal;
    lpRect: PRect;
    Result: Integer;
  End;

Function SHChangeIconDialog(hOwner: THandle; szFilename: PWideChar; Reserved: Integer; Var lpIconIndex: Integer): DWORD; Stdcall;

Type
  TForm1 = Class(TForm)
    StatusBar: TStatusBar;
    OpenDialog: TOpenDialog;
    ListViewList: TListView;
    TabControl2: TTabControl;
    GroupBox1: TGroupBox;
    EditFolder: TEdit;
    GroupBox2: TGroupBox;
    EditInfoTip: TEdit;
    GroupBox3: TGroupBox;
    EditIconFile: TEdit;
    ImageList3: TImageList;
    GroupBoxPuth: TGroupBox;
    TabControl7: TTabControl;
    ImageFolder: TImage;
    TabControl5: TTabControl;
    ImageIcon: TImage;
    ButtonUP: TButton;
    ButtonRefresh: TButton;
    SavePictureDialog: TSavePictureDialog;
    TabControl1: TTabControl;
    GroupBox4: TGroupBox;
    CheckBoxCopyIcon: TCheckBox;
    ButtonChangeIcon: TButton;
    GroupBox5: TGroupBox;
    TabControl6: TTabControl;
    GroupBox6: TGroupBox;
    MemoLog: TMemo;
    CheckBoxSys: TCheckBox;
    CheckBoxHide: TCheckBox;
    CheckBoxReadOnly: TCheckBox;
    SpinEditIcon: TSpinEdit;
    SpeedButton_GeneralMenu: TSpeedButton;
    PathFolder: TEdit;
    ComboBoxDrive: TComboBox;
    PopupMenuLanguage: TPopupMenu;
    LangElements: TMenuItem;
    LangError: TMenuItem;
    LangTarget: TMenuItem;
    LangIconInstall: TMenuItem;
    LangSelectFolderIcon: TMenuItem;
    LangAttention: TMenuItem;
    LangFolderWithIcon: TMenuItem;
    LangChangeIconFolder: TMenuItem;
    LangDelIconFolder: TMenuItem;
    LangCannotChangeIcon: TMenuItem;
    LangIconDelete: TMenuItem;
    StatusBar2: TStatusBar;
    PopupMenuFileList: TPopupMenu;
    IconDeleteMenu: TMenuItem;
    RefreshIconMenu: TMenuItem;
    N2: TMenuItem;
    PopupMenuIcon: TPopupMenu;
    AddIconMenu: TMenuItem;
    SaveIconMenu: TMenuItem;
    ButtonGetInfo: TButton;
    LangClearingCacheComplet: TMenuItem;
    ProgressBar1: TProgressBar;
    LangStartExplorer: TMenuItem;
    LangEndExplorer: TMenuItem;
    LangClearingCacheIcon: TMenuItem;
    LangClearingAddCacheIcon: TMenuItem;
    LangDeletingCacheFiles: TMenuItem;
    LangOnlyWindows: TMenuItem;
    LangFileNotFound: TMenuItem;
    LangFailedUploadFile: TMenuItem;
    LangNoIconGroups: TMenuItem;
    LangInvalidIconIndex: TMenuItem;
    LangIconExtractedSuccessfully: TMenuItem;
    LangErrorExtractingIcon: TMenuItem;
    ButtonSaveIcon: TButton;
    PopupMenuGeneral: TPopupMenu;
    ThemeMenu: TMenuItem;
    ThemeAuto: TMenuItem;
    ThemeLight: TMenuItem;
    ThemeDark: TMenuItem;
    LanguageMenu: TMenuItem;
    N3: TMenuItem;
    RefreshListMenu: TMenuItem;
    MenuItem2: TMenuItem;
    CheckUpdateMenu: TMenuItem;
    N1: TMenuItem;
    AboutMenu: TMenuItem;
    DeleteIconCache: TMenuItem;
    RestartExplorerMenu: TMenuItem;
    N4: TMenuItem;
    ViewMenu: TMenuItem;
    IconSmallMenu: TMenuItem;
    IconBigMenu: TMenuItem;
    ButtonBrowseIcon: TButton;
    ImageList1: TImageList;
    Procedure Button5Click(Sender: TObject);
    Procedure SpinEditIconChange(Sender: TObject);
    Procedure IconInstall;
    Procedure IconDelete;
    Procedure LoadNastr;
    Procedure SaveNastr;
    Procedure ExtractIconByIndex(Const ExeFile, OutputIcoFile: String; IconIndex: Integer);
    Procedure N43Click(Sender: TObject);
    Procedure SpinEditIconKeyPress(Sender: TObject; Var Key: Char);
    Procedure ButtonGetInfoClick(Sender: TObject);
    Procedure ComboBoxDriveChange(Sender: TObject);
    Procedure getIcon;
    Procedure getDrive;
    Procedure GetAttrFolder;
    Function Work(Str: String; Simvol: Char): String;
    Function KillTask(ExeFileName: String): Integer;
    Function GetDriveVolume(drive: String): String;
    Function TrimEX(Word, TrimSymbols: String): String;
    Function GetParentDir(StartDirectory: String): String;
    Function DeleteFileWithUndo(sFileName: String): Boolean;
    Procedure ButtonUPClick(Sender: TObject);
    Procedure ListViewListDblClick(Sender: TObject);
    Procedure ButtonRefreshClick(Sender: TObject);
    Procedure ButtonChangeIconClick(Sender: TObject);
    Procedure FormClose(Sender: TObject; Var Action: TCloseAction);
    Procedure IconBigMenuClick(Sender: TObject);
    Procedure IconSmallMenuClick(Sender: TObject);
    Procedure SpeedButton_GeneralMenuClick(Sender: TObject);
    Procedure RefreshListMenuClick(Sender: TObject);
    Procedure globload;
    Procedure LoadLanguage;
    Procedure UnCheckTheme;
    Procedure SetIconsForSize(UseLargeIcons: Boolean);
    Function GetIconCountSimple(Const ExeFile: String): Integer;
    Procedure CleanTranslationsLikeGlobload;
    Function RunAsAdmin(Const FileName, Parameters: String): Boolean;
    Procedure ClearIconCacheInThread;
    Procedure FormCreate(Sender: TObject);
    Procedure IconDeleteMenuClick(Sender: TObject);
    Procedure RefreshIconMenuClick(Sender: TObject);
    Procedure ImageIconClick(Sender: TObject);
    Procedure SaveIconMenuClick(Sender: TObject);
    Function IsWindows10Or11: Boolean;
    Function PortablePath: String;
    Procedure ButtonSaveIconClick(Sender: TObject);
    Procedure FormMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
    Procedure ThemeAutoClick(Sender: TObject);
    Procedure ThemeLightClick(Sender: TObject);
    Procedure ThemeDarkClick(Sender: TObject);
    Procedure CheckUpdateMenuClick(Sender: TObject);
    Procedure AboutMenuClick(Sender: TObject);
    Procedure DeleteIconCacheClick(Sender: TObject);
    Procedure RestartExplorerMenu1Click(Sender: TObject);
    Procedure MenuItem1Click(Sender: TObject);
    Procedure FormMouseWheel(Sender: TObject; Shift: TShiftState; WheelDelta: Integer; MousePos: TPoint; Var Handled: Boolean);
    Procedure SpinEditIconKeyDown(Sender: TObject; Var Key: Word; Shift: TShiftState);
    Procedure ButtonBrowseIconClick(Sender: TObject);
    Procedure FormDestroy(Sender: TObject);
    Procedure FormDragOver(Sender, Source: TObject; X, Y: Integer; State: TDragState; Var Accept: Boolean);
    Procedure UpdateFormState;
    Procedure ScanAndAddFoldersToListView(Const Path: String; ListView: TListView);
    Function IsSystemOrHiddenFolder(Const FolderName: String): Boolean;
    Procedure ListViewListClick(Sender: TObject);
    Procedure ListViewListSelectItem(Sender: TObject; Item: TListItem; Selected: Boolean);
    Procedure FormActivate(Sender: TObject);
    Procedure EditIconFileClick(Sender: TObject);
  Private
    Procedure WMMouseMove(Var Message: TWMMouseMove); Message WM_MOUSEMOVE;
    Procedure WMDropFiles(Var Msg: TWMDropFiles); Message WM_DROPFILEs;
    Procedure WMSettingChange(Var Message: TWMSettingChange); Message WM_SETTINGCHANGE;
    Procedure LanguageMenuItemClick(Sender: TObject);

    Var
      FListViewWndProc: TWndMethod;
    Procedure ListViewWndProc(Var Msg: TMessage);
  Public
    Procedure WMDeviceChange(Var Msg: TMessage); Message WM_DEVICECHANGE;

    Var
      FShowHoriz: Boolean;
      FShowVert: Boolean;
  End;

Var
  Form1: TForm1;
  drive: String;
  Sort: Integer = 1;
  Msg: Integer;
  Portable: boolean;
  LangCode: String;
  LangLocal: String;

Const
  ServerName = 'Folder-Icon-Changer';
  ApiGithub = 'https://api.github.com/repos/markovuser/' + ServerName + '/releases/latest';

Implementation

Uses
  Unit_About, Unit_Update;
{$R *.dfm}

//Function SHChangeIconDialog; External 'shell32.dll' Index 62;

Function SHChangeIconDialog; External 'shell32.dll' Name '#62';

Function TForm1.IsWindows10Or11: Boolean;
Begin
  Result := (TOSVersion.Major = 10) And (TOSVersion.Build >= 10240);
End;

Procedure TForm1.ListViewWndProc(Var Msg: TMessage);
Begin
  ShowScrollBar(ListViewList.Handle, SB_HORZ, FShowHoriz);
  ShowScrollBar(ListViewList.Handle, SB_VERT, FShowVert);
  FListViewWndProc(Msg); // process message
End;

Procedure TForm1.WMSettingChange(Var Message: TWMSettingChange);
Var
  Ini: TMemIniFile;
  ThemeAuto: boolean;
Begin
  Ini := TMemIniFile.Create(Form1.PortablePath);

  ThemeAuto := Ini.ReadBool('Option', 'ThemeAuto', false);
  If ThemeAuto = true Then
  Begin
    If DarkModeIsEnabled = true Then
    Begin
      SetSpecificThemeMode(true, 'Windows11 Modern Dark', 'Windows11 Modern Light');
      Application.ProcessMessages;
    End;

    If DarkModeIsEnabled = false Then
    Begin
      SetSpecificThemeMode(true, 'Windows11 Modern Light', 'Windows11 Modern Dark');
      Application.ProcessMessages;
    End;
  End;
  Ini.Free;
End;

// Удаление в корзину
Function TForm1.DeleteFileWithUndo(sFileName: String): Boolean;
Var
  fos: TSHFileOpStruct;
Begin
  sFileName := sFileName + #0;
  FillChar(fos, sizeof(fos), 0);
  With fos Do
  Begin
    wFunc := FO_DELETE;
    pFrom := pchar(sFileName);
    fFlags := FOF_ALLOWUNDO Or FOF_NOCONFIRMATION Or FOF_SILENT;
  End;
  Result := (0 = ShFileOperation(fos));
End;

Function TForm1.RunAsAdmin(Const FileName, Parameters: String): Boolean;
Var
  Info: TShellExecuteInfo;
Begin
  FillChar(Info, SizeOf(Info), 0);
  Info.cbSize := SizeOf(TShellExecuteInfo);
  Info.fMask := SEE_MASK_NOCLOSEPROCESS;
  Info.lpFile := pchar(FileName);
  Info.lpParameters := pchar(Parameters);
  Info.nShow := SW_HIDE;
  Info.lpVerb := 'runas'; // Запуск с правами администратора

  Result := ShellExecuteEx(@Info);
  If Result Then
  Begin
    WaitForSingleObject(Info.hProcess, INFINITE);
    CloseHandle(Info.hProcess);
  End;
End;

Function GetWindowsDirectory: String;
Var
  Buffer: Array[0..MAX_PATH] Of Char;
Begin
  Windows.GetWindowsDirectory(Buffer, MAX_PATH);
  Result := String(Buffer);
End;

Function StartExplorerLikeTaskManager: Boolean;
Var
  ProcessInfo: TProcessInformation;
  StartupInfo: TStartupInfo;
Begin
  FillChar(StartupInfo, SizeOf(StartupInfo), 0);
  StartupInfo.cb := SizeOf(StartupInfo);

  // Ключевой момент: запускаем explorer.exe из системной директории
  Result := CreateProcess(PChar(GetWindowsDirectory + '\explorer.exe'), // Полный путь
    Nil,                       // Без параметров
    Nil,                       // Без наследования дескриптора процесса
    Nil,                       // Без наследования дескриптора потока
    False,                     // Не наследует дескрипторы
    NORMAL_PRIORITY_CLASS,     // Обычный приоритет
    Nil,                       // Использует окружение родителя
    Nil,                       // Использует текущую директорию
    StartupInfo, ProcessInfo);

  If Result Then
  Begin
    CloseHandle(ProcessInfo.hThread);
    CloseHandle(ProcessInfo.hProcess);
    Sleep(3000); // Даем время на запуск
  End;
End;

Procedure TForm1.ClearIconCacheInThread;
Begin
  // Показываем анимацию ожидания перед запуском
  TThread.Synchronize(Nil,
    Procedure
    Begin
      ProgressBar1.Visible := True;
      ProgressBar1.Position := 0;
      Form1.StatusBar.Panels[0].Text := '';
    End);

  TTask.Run(
    Procedure
    Begin
      Try
      // Шаг 1: Закрытие проводника
        TThread.Synchronize(Nil,
          Procedure
          Begin
            ProgressBar1.Position := 10;
            Form1.StatusBar.Panels[0].Text := LangEndExplorer.Caption;
          End);

        ShellExecute(0, 'open', 'taskkill', '/f /im explorer.exe', Nil, SW_HIDE);
        Sleep(3000);

      // Шаг 2: Очистка основного кэша
        TThread.Synchronize(Nil,
          Procedure
          Begin
            ProgressBar1.Position := 30;
            Form1.StatusBar.Panels[0].Text := LangClearingCacheIcon.Caption;
          End);

        ShellExecute(0, 'open', 'cmd.exe', '/c attrib -h "%localappdata%\IconCache.db"', Nil, SW_HIDE);
        Sleep(500);
        ShellExecute(0, 'open', 'cmd.exe', '/c del /a /f /q "%localappdata%\IconCache.db"', Nil, SW_HIDE);
        Sleep(500);

      // Шаг 3: Очистка дополнительного кэша
        TThread.Synchronize(Nil,
          Procedure
          Begin
            ProgressBar1.Position := 50;
            Form1.StatusBar.Panels[0].Text := LangClearingAddCacheIcon.Caption;
          End);

        ShellExecute(0, 'open', 'cmd.exe', '/c attrib -h "%localappdata%\Microsoft\Windows\Explorer\iconcache*"', Nil, SW_HIDE);
        Sleep(500);
        ShellExecute(0, 'open', 'cmd.exe', '/c attrib -h "%localappdata%\Microsoft\Windows\Explorer\thumbcache*"', Nil, SW_HIDE);
        Sleep(500);

      // Шаг 4: Удаление файлов кэша
        TThread.Synchronize(Nil,
          Procedure
          Begin
            ProgressBar1.Position := 70;
            Form1.StatusBar.Panels[0].Text := LangDeletingCacheFiles.Caption;
          End);

        ShellExecute(0, 'open', 'cmd.exe', '/c del /a /f /q "%localappdata%\Microsoft\Windows\Explorer\iconcache*"', Nil, SW_HIDE);
        Sleep(500);
        ShellExecute(0, 'open', 'cmd.exe', '/c del /a /f /q "%localappdata%\Microsoft\Windows\Explorer\thumbcache*"', Nil, SW_HIDE);
        Sleep(500);

        Sleep(2000);

      // Шаг 5: Запуск проводника
        TThread.Synchronize(Nil,
          Procedure
          Begin
            ProgressBar1.Position := 90;
            Form1.StatusBar.Panels[0].Text := LangStartExplorer.Caption;
          End);

        StartExplorerLikeTaskManager;

      // Завершение
        TThread.Synchronize(Nil,
          Procedure
          Begin
            ProgressBar1.Position := 100;
            Form1.StatusBar.Panels[0].Text := LangClearingCacheComplet.caption;

        // Ждем немного перед скрытием прогресса
            Sleep(1000);

            ProgressBar1.Visible := False;
          End);

      Except
        On E: Exception Do
        Begin
          TThread.Synchronize(Nil,
            Procedure
            Begin
              ProgressBar1.Visible := False;
              Form1.StatusBar.Panels[0].Text := LangError.Caption + ' ' + E.Message;
            End);
        End;
      End;
    End);
End;

Procedure TListView.WMVScroll(Var Message: TWMVScroll);
Begin
  TWinControl(self).DefaultHandler(Message);
End;

Procedure TForm1.UnCheckTheme;
Var
  i: Integer;
Begin
  For i := 0 To ThemeMenu.Count - 1 Do
  Begin
    ThemeMenu.Items[i].Checked := False;
  End;
End;

Procedure TForm1.ThemeAutoClick(Sender: TObject);
Begin
  SpeedButton_GeneralMenu.AllowAllUp := true;
  SpeedButton_GeneralMenu.Down := False;
  UnCheckTheme;
  ThemeAuto.Checked := true;
  If DarkModeIsEnabled = true Then
  Begin
    SetSpecificThemeMode(true, 'Windows11 Modern Dark', 'Windows11 Modern Light');
    Application.ProcessMessages;
  End;

  If DarkModeIsEnabled = False Then
  Begin
    SetSpecificThemeMode(true, 'Windows11 Modern Light', 'Windows11 Modern Dark');
    Application.ProcessMessages;
  End;
End;

Procedure TForm1.ThemeDarkClick(Sender: TObject);
Begin
  SpeedButton_GeneralMenu.AllowAllUp := true;
  SpeedButton_GeneralMenu.Down := False;
  UnCheckTheme;
  ThemeDark.Checked := true;
  SetSpecificThemeMode(true, 'Windows11 Modern Dark', 'Windows11 Modern Light');
  Application.ProcessMessages;
End;

Procedure TForm1.ThemeLightClick(Sender: TObject);
Begin
  SpeedButton_GeneralMenu.AllowAllUp := true;
  SpeedButton_GeneralMenu.Down := False;
  UnCheckTheme;
  ThemeLight.Checked := true;
  SetSpecificThemeMode(true, 'Windows11 Modern Light', 'Windows11 Modern Dark');
  Application.ProcessMessages;
End;

Function TForm1.TrimEX(Word, TrimSymbols: String): String;
Var
  x, a, b: Integer;
Begin
  Result := Word;
  If TrimSymbols = '' Then
    Exit;
  Word := trim(Word);
  If Length(Word) = 0 Then
    Exit;
  x := 0;
  Repeat
    x := x + 1;
  Until (pos(AnsiUpperCase(Word[x]), AnsiUpperCase(TrimSymbols)) = 0) Or (x = Length(Word));
  a := x;
  x := Length(Word) + 1;
  Repeat
    x := x - 1;
  Until (pos(AnsiUpperCase(Word[x]), AnsiUpperCase(TrimSymbols)) = 0) Or (x = 1);
  b := x;
  Word := copy(Word, a, b - a + 1);
  Result := Word;
End;

Function TForm1.GetParentDir(StartDirectory: String): String;
Var
  x: Integer;
Begin
  Result := '';
  If SysUtils.DirectoryExists(StartDirectory) = false Then
    Exit;
  StartDirectory := TrimEX(StartDirectory, '\');
  If Length(StartDirectory) = 0 Then
    Exit;
  x := Length(StartDirectory) + 1;
  Repeat
    x := x - 1;
  Until (StartDirectory[x] = '\') Or (x = 1);
  Result := copy(StartDirectory, 1, x);
  If Result[Length(Result)] <> '\' Then
    Result := Result + '\';
  If SysUtils.DirectoryExists(Result) = false Then
  Begin
    Result := '';
    Exit;
  End;
End;

Function TForm1.GetDriveVolume(drive: String): String;
Var
  _VolumeName, _FileSystemName: Array[0..MAX_PATH - 1] Of Char;
  _VolumeSerialNo, _MaxComponentLength, _FileSystemFlags: LongWord;
Begin
  Result := '';
  If GetVolumeInformation(pchar(drive + ':\'), _VolumeName, MAX_PATH, @_VolumeSerialNo, _MaxComponentLength, _FileSystemFlags, _FileSystemName, MAX_PATH) Then
    Result := _VolumeName;
End;

Procedure TForm1.getDrive;
Var
  drive: Char;
Begin
  Form1.ComboBoxDrive.Clear;
  For drive := 'a' To 'z' Do
  Begin
    Case GetDriveType(pchar(drive + ':\')) Of
      DRIVE_REMOVABLE:
        Begin
          Form1.ComboBoxDrive.Items.Add(UpperCase(drive) + ':\ [' + GetDriveVolume(drive) + ' ]');
        End;
      DRIVE_FIXED:
        Begin
          Form1.ComboBoxDrive.Items.Add(UpperCase(drive) + ':\ [' + GetDriveVolume(drive) + ' ]');
        End;
      { DRIVE_CDROM:
        begin
        Form1.ComboBox1.Items.Add(AnsiUpperCase(Drive) + ':\ [' +
        GetDriveVolume(Drive) + ' ]');
        end; }
      { DRIVE_REMOTE:
        begin
        Form1.ComboBox1.Items.Add(UpperCase(drive) + ':\ [' +
        GetDriveVolume(drive) + ' ]');
        application.ProcessMessages;
        end; }
      { DRIVE_RAMDISK:
        begin
        Form1.ComboBox1.Items.Add(UpperCase(Drive) + ':\ [' +
        GetDriveVolume(Drive) + ' ]');
        end; }
    End;
  End;
  Form1.ComboBoxDrive.ItemIndex := 0;
End;

Procedure TForm1.WMDeviceChange(Var Msg: TMessage);
Const
  DBT_DEVICEARRIVAL = $8000; // system detected a new device
  DBT_DEVICEREMOVECOMPLETE = $8004; // device is gone
Var
  myMsg: String;
  r: LongWord;
  Drives: Array[0..128] Of Char;
  pDrive: pchar;
Begin
  Inherited;
  Case Msg.WParam Of
    DBT_DEVICEARRIVAL:
      myMsg := 'OK';
    DBT_DEVICEREMOVECOMPLETE:
      myMsg := 'NO';
  End;
  If myMsg = 'OK' Then
  Begin
    r := GetLogicalDriveStrings(sizeof(Drives), Drives);
    If r = 0 Then
      Exit;
    If r > sizeof(Drives) Then
      Raise Exception.Create(SysErrorMessage(ERROR_OUTOFMEMORY));
    pDrive := Drives;
    While pDrive^ <> #0 Do
    Begin
      If GetDriveType(pDrive) = DRIVE_REMOVABLE Then
        drive := pDrive;
      inc(pDrive, 4);
    End;
    getDrive;
  End;
  If myMsg = 'NO' Then
  Begin
    getDrive;
  End;
End;

Function SlashSep(Path, FName: String): String;
Begin
  If Path[Length(Path)] <> '\' Then
    // Если путь не заканчивается слешем '\'
    Result := Path + '\' + FName
    // то его нужно добавить, а затем добавить имя файла
  Else
    Result := Path + FName; // иначе не добавлять
End;

Function FileTimeToDateTimeStr(FileTime: TFileTime): String;
// Функция перевода сист. времени в строку
Var
  LocFTime: TFileTime;
  SysFTime: TSystemTime;
  DateStr: String;
  Date, Time: TDateTime;
Begin
  FileTimeToLocalFileTime(FileTime, LocFTime);
  FileTimeToSystemTime(LocFTime, SysFTime);
  Try
    With SysFTime Do
    Begin
      Date := EncodeDate(wYear, wMonth, wDay);
      DateStr := DateToStr(Date);
      Time := EncodeTime(wHour, wMinute, wSecond, wMilliseconds);
    End;
    Result := DateTimeToStr(Date + Time);
  Except
    Result := '';
  End;
End;

Function Sort1SubItemAsString1(Item1, Item2: TListItem; ParamSort: Integer): Integer; Stdcall;
Begin
  Result := 0; // Устанавливаем значение по умолчанию
  Try
    If AnsiUpperCase(Item1.SubItems[2]) > AnsiUpperCase(Item2.SubItems[2]) Then
      Result := ParamSort
    Else If AnsiUpperCase(Item1.SubItems[2]) < AnsiUpperCase(Item2.SubItems[2]) Then
      Result := -ParamSort;
    // Если строки равны, Result остается 0
  Except
    Result := 0; // В случае исключения возвращаем 0
  End;
End;

Function TForm1.IsSystemOrHiddenFolder(Const FolderName: String): Boolean;
Var
  SystemFolders: TArray<String>;
  i: Integer;
Begin
  Result := False;

  // Список системных/скрытых папок Windows
  SystemFolders := TArray<String>.Create('$RECYCLE.BIN', 'System Volume Information', '$Recycle.Bin', 'RECYCLER', 'RECYCLED', 'Thumbs.db', '.Trash',              // Для Linux/Unix
    '.Trash-1000',         // Для Linux/Unix
    '.local',              // Для Linux/Unix
    '.config',             // Для Linux/Unix
    '.cache',              // Для Linux/Unix
    '.git',                // Папка Git
    '.svn',                // Папка Subversion
    '.hg',                 // Папка Mercurial
    '__pycache__',         // Python кэш
    'node_modules'         // Node.js модули
  );

  // Проверяем, входит ли имя папки в список системных
  For i := 0 To High(SystemFolders) Do
  Begin
    If SameText(FolderName, SystemFolders[i]) Then
    Begin
      Result := True;
      Exit;
    End;
  End;
End;

Procedure TForm1.ScanAndAddFoldersToListView(Const Path: String; ListView: TListView);
Var
  sr: TSearchRec;
  ListItem: TListItem;
  ShInfo: TSHFileInfo;
  SearchPath: String;
Begin
  If Not Assigned(ListView) Then
    Exit;

  // Блокируем обновление для скорости
  ListView.Items.BeginUpdate;
  Try
    // Очищаем ListView
    ListView.Items.Clear;

    // Нормализуем путь
    SearchPath := IncludeTrailingPathDelimiter(Path);

    // Ищем все файлы и папки
    If FindFirst(SearchPath + '*.*', faAnyFile, sr) = 0 Then
    Begin
      Repeat
        // Пропускаем текущую и родительскую директории
        If (sr.Name = '.') Or (sr.Name = '..') Then
          Continue;

        // Проверяем, что это папка
        If (sr.Attr And faDirectory) = faDirectory Then
        Begin
          // Пропускаем системные и скрытые папки для Windows
          {$IFDEF MSWINDOWS}
          // Для Windows используем Windows-специфичные константы
          If (sr.Attr And (FILE_ATTRIBUTE_HIDDEN Or FILE_ATTRIBUTE_SYSTEM)) > 0 Then
            Continue;
          {$ELSE}
          // Для других платформ или кроссплатформенного кода
          // Можно использовать общие проверки или пропустить эту проверку
          {$ENDIF}

          // Дополнительная проверка для Windows системных папок
          // (например, "$RECYCLE.BIN", "System Volume Information", и т.д.)
          If IsSystemOrHiddenFolder(sr.Name) Then
            Continue;

          // Создаем новый элемент в ListView
          ListItem := ListView.Items.Add;
          ListItem.Caption := sr.Name; // Имя папки

          // Добавляем полный путь в SubItems
          ListItem.SubItems.Add(SearchPath + sr.Name);

          // Получаем иконку папки
          SHGetFileInfo(PChar(SearchPath + sr.Name), 0, ShInfo, SizeOf(ShInfo), SHGFI_SYSICONINDEX);

          // Устанавливаем иконку
          ListItem.ImageIndex := ShInfo.iIcon;
        End;
      Until FindNext(sr) <> 0;

      FindClose(sr);
    End;

  Finally
    ListView.Items.EndUpdate;
  End;

  // Ваш существующий код для выбора первого элемента
  If ListViewList.Items.Count > 0 Then
  Begin
    Try
      ListViewList.Items[0].Selected := True;
      ListViewList.ItemIndex := 0; // Альтернативный способ выделения
      ListViewList.Items[0].MakeVisible(True);
    Except
    End;
  End;

  Application.ProcessMessages;
  StatusBar2.Panels[0].Text := LangElements.Caption;
  StatusBar2.Panels[1].Text := inttostr(ListViewList.Items.Count);
End;

Procedure TForm1.ComboBoxDriveChange(Sender: TObject);
Begin
  Form1.StatusBar.Panels[0].Text := '';
  Form1.ImageFolder.Picture := Nil;
  Form1.ImageIcon.Picture := Nil;
  EditIconFile.Text := '';
  EditFolder.Text := '';
  PathFolder.Text := '';
  PathFolder.Text := ExtractFileDrive(Form1.ComboBoxDrive.Text) + '\';
  ScanAndAddFoldersToListView(PathFolder.Text, ListViewList);
  Application.ProcessMessages;
End;

Function GetTempDirectory: String;
Var
  Buffer: Array[0..MAX_PATH] Of Char;
Begin
  GetTempPath(MAX_PATH, Buffer);
  Result := IncludeTrailingPathDelimiter(StrPas(Buffer));
End;

// Удаление символов в строке после определенного символа
Function TForm1.Work(Str: String; Simvol: Char): String;
Var
  a: Integer;
Begin
  Try
    Result := Str;
    a := pos(Simvol, Result);
    If a > 0 Then
      Delete(Result, a, (Length(Result) - a) + 1);
  Except
  End;
End;

// Установка иконки
Procedure TForm1.IconInstall;
Var
  IconFile, IconResource, BaseName: String;
  iniFile: TMemIniFile;
  DesktopIniPath, TargetIconPath: String;
Begin
  DesktopIniPath := EditFolder.Text + 'Desktop.ini';
  iniFile := TMemIniFile.Create(DesktopIniPath);
  Try
    // Считываем текущие значения
    IconFile := iniFile.readString('.ShellClassInfo', 'IconFile', '');
    IconResource := iniFile.readString('.ShellClassInfo', 'IconResource', '');
    IconResource := Work(IconResource, ',');

    // Снимаем атрибуты с файлов
    SetFileAttributes(PChar(DesktopIniPath), FILE_ATTRIBUTE_NORMAL);
    If IconFile <> '' Then
      SetFileAttributes(PChar(EditFolder.Text + ExtractFileName(IconFile)), FILE_ATTRIBUTE_NORMAL);
    If IconResource <> '' Then
      SetFileAttributes(PChar(EditFolder.Text + ExtractFileName(IconResource)), FILE_ATTRIBUTE_NORMAL);

    If Not CheckBoxCopyIcon.Checked Then
    Begin
      // Не копировать иконку - используем исходный путь
      With iniFile Do
      Begin
        WriteString('.ShellClassInfo', 'IconFile', EditIconFile.Text);
        WriteString('.ShellClassInfo', 'IconIndex', IntToStr(SpinEditIcon.Value));
        WriteString('.ShellClassInfo', 'IconResource', EditIconFile.Text + ',' + IntToStr(SpinEditIcon.Value));
        WriteString('.ShellClassInfo', 'InfoTip', EditInfoTip.Text);
        WriteString('.ShellClassInfo', 'ConfirmFileOp', '0');
        UpdateFile;
      End;
    End
    Else
    Begin
      // Копировать иконку в папку ICO
      If LowerCase(ExtractFileExt(EditIconFile.Text)) = '.ico' Then
      Begin

        TargetIconPath := EditFolder.Text + ExtractFileName(EditIconFile.Text);

      // Удаляем старые файлы иконок если нужно
        If (IconFile <> '') And (LowerCase(ExtractFileExt(IconFile)) = '.ico') And (IconFile <> ExtractFileName(EditIconFile.Text)) Then
        Begin
          SetFileAttributes(PChar(EditFolder.Text + ExtractFileName(IconFile)), FILE_ATTRIBUTE_NORMAL);
          If FileExists(EditFolder.Text + ExtractFileName(IconFile)) Then
            DeleteFileWithUndo(EditFolder.Text + ExtractFileName(IconFile));
        End;

        If (IconResource <> '') And (LowerCase(ExtractFileExt(IconResource)) = '.ico') And (IconResource <> ExtractFileName(EditIconFile.Text)) Then
        Begin
          SetFileAttributes(PChar(EditFolder.Text + ExtractFileName(IconResource)), FILE_ATTRIBUTE_NORMAL);
          If FileExists(EditFolder.Text + ExtractFileName(IconResource)) Then
            DeleteFileWithUndo(EditFolder.Text + ExtractFileName(IconResource));
        End;

      // Копируем файл иконки если он в другом месте
        If Not SameText(EditIconFile.Text, TargetIconPath) Then
          CopyFile(PChar(EditIconFile.Text), PChar(TargetIconPath), False);

      // Настраиваем INI файл для скопированной иконки
        With iniFile Do
        Begin
          WriteString('.ShellClassInfo', 'IconFile', ExtractFileName(EditIconFile.Text));
          WriteString('.ShellClassInfo', 'IconIndex', IntToStr(SpinEditIcon.Value));
          WriteString('.ShellClassInfo', 'IconResource', ExtractFileName(EditIconFile.Text) + ',' + IntToStr(SpinEditIcon.Value));
          WriteString('.ShellClassInfo', 'InfoTip', EditInfoTip.Text);
          WriteString('.ShellClassInfo', 'ConfirmFileOp', '0');
          UpdateFile;
        End;
      End;
            // Копировать иконку в папку EXE DLL
      If (LowerCase(ExtractFileExt(EditIconFile.Text)) = '.exe') Or (LowerCase(ExtractFileExt(EditIconFile.Text)) = '.dll') Then
      Begin
        BaseName := ChangeFileExt(ExtractFileName(EditIconFile.Text), '');
        BaseName := BaseName + '.ico';
        ExtractIconByIndex(EditIconFile.Text, GetTempDirectory + BaseName, SpinEditIcon.Value);
        TargetIconPath := EditFolder.Text + ExtractFileName(GetTempDirectory + BaseName);
      // Удаляем старые файлы иконок если нужно
        If (IconFile <> '') And (LowerCase(ExtractFileExt(IconFile)) = '.ico') Then
        Begin
          SetFileAttributes(PChar(EditFolder.Text + ExtractFileName(IconFile)), FILE_ATTRIBUTE_NORMAL);
          If FileExists(EditFolder.Text + ExtractFileName(IconFile)) Then
            DeleteFileWithUndo(EditFolder.Text + ExtractFileName(IconFile));
        End;

        If (IconResource <> '') And (LowerCase(ExtractFileExt(IconResource)) = '.ico') Then
        Begin
          SetFileAttributes(PChar(EditFolder.Text + ExtractFileName(IconResource)), FILE_ATTRIBUTE_NORMAL);
          If FileExists(EditFolder.Text + ExtractFileName(IconResource)) Then
            DeleteFileWithUndo(EditFolder.Text + ExtractFileName(IconResource));
        End;

      // Копируем файл иконки если он в другом месте
        CopyFile(PChar(GetTempDirectory + BaseName), PChar(TargetIconPath), False);

      // Настраиваем INI файл для скопированной иконки
        With iniFile Do
        Begin
          WriteString('.ShellClassInfo', 'IconFile', ExtractFileName(BaseName));
          WriteString('.ShellClassInfo', 'IconIndex', IntToStr(SpinEditIcon.Value));
          WriteString('.ShellClassInfo', 'IconResource', ExtractFileName(BaseName));
          WriteString('.ShellClassInfo', 'InfoTip', EditInfoTip.Text);
          WriteString('.ShellClassInfo', 'ConfirmFileOp', '0');
          UpdateFile;
        End;
      End;
      // Устанавливаем атрибуты для скопированного файла иконки
      SetFileAttributes(PChar(TargetIconPath), FILE_ATTRIBUTE_SYSTEM Or FILE_ATTRIBUTE_HIDDEN);
    End;

    // Финальные настройки атрибутов
    SetFileAttributes(PChar(DesktopIniPath), FILE_ATTRIBUTE_SYSTEM Or FILE_ATTRIBUTE_HIDDEN);
    SetFileAttributes(PChar(EditFolder.Text), FILE_ATTRIBUTE_READONLY);

    // Обновляем интерфейс
    Application.ProcessMessages;
    StatusBar.Panels[0].Text := LangIconInstall.Caption;

  Finally
    iniFile.Free;
  End;
End;

// Функция для получения реального пути из ярлыка
Function GetFileNamefromLink(LinkFileName: String): String;
Var
  MyObject: IUnknown;
  MySLink: IShellLink;
  MyPFile: IPersistFile;
  FileInfo: TWin32FINDDATA;
  WidePath: Array[0..MAX_PATH] Of WideChar;
  Buff: Array[0..MAX_PATH] Of char;
Begin
  Result := '';
  If (fileexists(LinkFileName) = False) Then
    Exit;
  MyObject := CreateComObject(CLSID_ShellLink);
  MyPFile := MyObject As IPersistFile;
  MySLink := MyObject As IShellLink;
  StringToWideChar(LinkFileName, WidePath, SizeOf(WidePath));
  MyPFile.Load(WidePath, STGM_READ);
  MySLink.GetPath(Buff, MAX_PATH, FileInfo, SLGP_UNCPRIORITY);
  Result := Buff;
End;

Procedure TForm1.WMDropFiles(Var Msg: TWMDropFiles);
Var
  FileCount: Integer;
  FileName: Array[0..MAX_PATH] Of Char;
  I: Integer;
  icExeDll: TIcon;
  icIco: TIcon;
  FileExt: String;
  IconCount: Integer;
  FileNameStr: String;
  Success: Boolean;
  RealPath: String;
Begin
  Inherited;

  icExeDll := Nil;
  icIco := Nil;
  Success := False;

  Try
    FileCount := DragQueryFile(Msg.Drop, $FFFFFFFF, Nil, 0);

    If FileCount > 0 Then
    Begin
      For I := 0 To Min(FileCount, 1) - 1 Do
      Begin
        FillChar(FileName, SizeOf(FileName), 0);

        If DragQueryFile(Msg.Drop, I, @FileName[0], MAX_PATH) > 0 Then
        Begin
          FileNameStr := String(PChar(@FileName[0]));

          If Not FileExists(FileNameStr) Then
            Continue;

          FileExt := LowerCase(ExtractFileExt(FileNameStr));

          // Обработка ярлыков
          If FileExt = '.lnk' Then
          Begin
            // Используем вашу функцию GetFileNamefromLink
            RealPath := GetFileNamefromLink(FileNameStr);

            // Проверяем результат
            If (RealPath = '') Or (Not FileExists(RealPath)) Then
              Continue;

            // Обновляем путь и расширение
            FileNameStr := RealPath;
            FileExt := LowerCase(ExtractFileExt(RealPath));
          End;

          // Проверяем поддерживаемые расширения
          If (LowerCase(FileExt) = '.exe') Or (LowerCase(FileExt) = '.dll') Or (LowerCase(FileExt) = '.ico') Then
          Begin
            // Проверяем компоненты
            If Not Assigned(EditIconFile) Or Not Assigned(SpinEditIcon) Or Not Assigned(ImageIcon) Then
              Exit;

            // В EditIconFile показываем оригинальный путь (ярлыка или файла)
            EditIconFile.Text := FileNameStr;

            // Для EXE/DLL файлов
            If (LowerCase(FileExt) = '.exe') Or (LowerCase(FileExt) = '.dll') Then
            Begin
              SpinEditIcon.Value := 0;

              // Получаем количество иконок
              IconCount := GetIconCountSimple(FileNameStr);

              // Настройка SpinEdit
              If IconCount > 0 Then
              Begin
                SpinEditIcon.MaxValue := Max(0, IconCount - 1);
                SpinEditIcon.MinValue := 0;
                SpinEditIcon.Enabled := (IconCount > 1);
                CheckBoxCopyIcon.Enabled := true;
                // Загружаем иконку через отдельный объект
                icExeDll := TIcon.Create;
                Try
                  GetIconFromFile(PChar(FileNameStr), icExeDll, SHIL_EXTRALARGE);
                  ImageIcon.Picture.Icon := icExeDll;
                Except
                  // В случае ошибки очищаем изображение
                  ImageIcon.Picture := Nil;
                End;
              End
              Else
              Begin
                SpinEditIcon.Enabled := False;
                ImageIcon.Picture := Nil;
              End;

              If Assigned(ButtonSaveIcon) Then
                ButtonSaveIcon.Enabled := (IconCount > 0);

              Success := True;
            End
            // Для ICO файлов
            Else If FileExt = '.ico' Then
            Begin
              icIco := TIcon.Create;
              Try
                // Пытаемся загрузить иконку из файла
                GetIconFromFile(PChar(FileNameStr), icIco, SHIL_EXTRALARGE);
                ImageIcon.Picture.Icon := icIco;
                Success := True;
              Except
                // В случае ошибки очищаем
                ImageIcon.Picture := Nil;
                Success := False;
              End;

              SpinEditIcon.Enabled := False;
              SpinEditIcon.Value := 0;
              CheckBoxCopyIcon.Enabled := true;
              If Assigned(ButtonSaveIcon) Then
                ButtonSaveIcon.Enabled := False;
            End;

            If Success Then
            Begin
              If Assigned(ButtonChangeIcon) Then
              Begin
                ButtonChangeIcon.Enabled := True;
                ButtonChangeIcon.Default := True;
                CheckBoxCopyIcon.Enabled := true;
                If ButtonChangeIcon.CanFocus Then
                  ButtonChangeIcon.SetFocus;
              End;

              If Assigned(StatusBar) And (StatusBar.Panels.Count > 0) Then
                StatusBar.Panels[0].Text := '';

              Break;
            End;
          End;
        End;
      End;
    End;
  Finally
    // Освобождаем созданные объекты
    If Assigned(icExeDll) Then
      icExeDll.Free;

    If Assigned(icIco) Then
      icIco.Free;

    DragFinish(Msg.Drop);
  End;
End;

Function GetSpecialFolderPath(folder: Integer): String;
Const
  SHGFP_TYPE_CURRENT = 0;
Var
  Path: Array[0..MAX_PATH] Of Char;
Begin
  If SUCCEEDED(SHGetFolderPath(0, folder, 0, SHGFP_TYPE_CURRENT, @Path[0])) Then
    Result := Path
  Else
    Result := '';
End;

Function EraseParam(Path: String): String;
Var
  i: Integer;
Begin
  For i := 0 To Length(Path) Do
  Begin
    If (copy(Path, i, 2) = ' -') Then
    Begin
      Result := trim(copy(Path, i, Length(Path)));
      break;
    End;
  End;
End;

Function RunParam(Path: String): String;
Var
  a3, a4: String;
Begin
  a3 := StringReplace(Path, ',', ' -', [rfReplaceAll]);
  a4 := EraseParam(a3);
  Result := a4;
End;

Function ErasePath(Path: String): String;
Var
  i: Integer;
Begin
  Result := Path;
  For i := 0 To Length(Path) Do
  Begin
    If (copy(Path, i, 2) = ' -') Then
    Begin
      Result := trim(copy(Path, 1, i - 1));
      break;
    End;
  End;
End;

Function RunPath(Path: String): String;
Var
  a3, a4: String;
Begin
  a3 := StringReplace(Path, ',', ' -', [rfReplaceAll]);
  a4 := ErasePath(a3);
  Result := a4;
End;

Function TForm1.KillTask(ExeFileName: String): Integer;
Const
  PROCESS_TERMINATE = $0001;
Var
  ContinueLoop: BOOL;
  FSnapshotHandle: THandle;
  ProcessHandle: Cardinal;
  FProcessEntry32: TProcessEntry32;
Begin
  Result := 0;
  FSnapshotHandle := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);

  If FSnapshotHandle = INVALID_HANDLE_VALUE Then
    RaiseLastOSError;
  Try
    FProcessEntry32.dwSize := sizeof(FProcessEntry32);
    ContinueLoop := Process32First(FSnapshotHandle, FProcessEntry32);

    While Integer(ContinueLoop) <> 0 Do
    Begin
      If ((UpperCase(ExtractFileName(FProcessEntry32.szExeFile)) = UpperCase(ExeFileName)) Or (UpperCase(FProcessEntry32.szExeFile) = UpperCase(ExeFileName))) Then
      Begin
        ProcessHandle := OpenProcess(PROCESS_TERMINATE, BOOL(0), FProcessEntry32.th32ProcessID);
        If ProcessHandle > 0 Then
        Begin
          Try
            Result := Integer(TerminateProcess(ProcessHandle, 0));
          Finally
            CloseHandle(ProcessHandle);
          End;
        End
        Else
          RaiseLastOSError;
      End;
      ContinueLoop := Process32Next(FSnapshotHandle, FProcessEntry32);
    End;
  Finally
    CloseHandle(FSnapshotHandle);
  End;
End;

// Удаление иконки

Procedure TForm1.IconDelete;
Var
  IconFile, IconResource, FolderPath: String;
  iniFile: TMemIniFile;
Begin
  FolderPath := IncludeTrailingPathDelimiter(EditFolder.Text);

  Try
    // Читаем данные из desktop.ini
    If FileExists(FolderPath + 'Desktop.ini') Then
    Begin
      iniFile := TMemIniFile.Create(FolderPath + 'Desktop.ini');
      Try
        IconFile := iniFile.ReadString('.ShellClassInfo', 'IconFile', '');
        IconResource := iniFile.ReadString('.ShellClassInfo', 'IconResource', '');

        // Обрабатываем пути к файлам
        IconResource := StringReplace(IconResource, '.\', '', [rfReplaceAll]);
        IconResource := Work(IconResource, ',');
      Finally
        iniFile.Free;
      End;

      // Снимаем атрибуты и удаляем файлы
      SetFileAttributes(PChar(FolderPath + 'Desktop.ini'), FILE_ATTRIBUTE_NORMAL);

      // Удаляем файл иконки из IconFile
      If (IconFile <> '') And (LowerCase(ExtractFileExt(EditIconFile.Text)) = '.ico') And FileExists(FolderPath + ExtractFileName(IconFile)) Then
      Begin
        SetFileAttributes(PChar(FolderPath + ExtractFileName(IconFile)), FILE_ATTRIBUTE_NORMAL);
        DeleteFileWithUndo(FolderPath + ExtractFileName(IconFile));
      End;

      // Удаляем файл иконки из IconResource
      If (IconResource <> '') And (LowerCase(ExtractFileExt(EditIconFile.Text)) = '.ico') And FileExists(FolderPath + ExtractFileName(IconResource)) Then
      Begin
        SetFileAttributes(PChar(FolderPath + ExtractFileName(IconResource)), FILE_ATTRIBUTE_NORMAL);
        DeleteFileWithUndo(FolderPath + ExtractFileName(IconResource));
      End;

      // Удаляем desktop.ini
      DeleteFileWithUndo(FolderPath + 'Desktop.ini');
    End;

    Application.ProcessMessages;
    Form1.StatusBar.Panels[0].Text := LangIconDelete.Caption;

  Except
    On E: Exception Do
    Begin
      // Обработка ошибки (можно добавить логирование или сообщение)
    End;
  End;
End;

//Сохранение иконки
Function CreateICO_Simple(Const ExeFile, IcoFile: String): Boolean;
Var
  hModule: Cardinal;
  Stream: TMemoryStream;
  i, IconCount: Integer;
Begin
  Result := False;
  IconCount := 0;

  If Not FileExists(ExeFile) Then
    Exit;

  hModule := LoadLibraryEx(PChar(ExeFile), 0, LOAD_LIBRARY_AS_DATAFILE);
  If hModule = 0 Then
    Exit;

  // Считаем иконки
  For i := 1 To 50 Do
  Begin
    If FindResource(hModule, PChar(i), RT_ICON) <> 0 Then
      Inc(IconCount)
    Else If i > 10 Then
      Break;
  End;

  If IconCount = 0 Then
  Begin
    FreeLibrary(hModule);
    Exit;
  End;

  Stream := TMemoryStream.Create;
  Try
    // Заголовок
    Var Header: Packed Record
  Reserved: Word;
  ImageType: Word;
  ImageCount: Word;
End;

    Header.Reserved := 0;
    Header.ImageType := 1;
    Header.ImageCount := IconCount;
    Stream.Write(Header, SizeOf(Header));

    // Записи каталога
    Var DataOffset: DWORD := 6 + (IconCount * 16);

    For i := 1 To IconCount Do
    Begin
      Var hRes: Cardinal := FindResource(hModule, PChar(i), RT_ICON);
      If hRes = 0 Then
        Continue;

      Var hGlobal: Cardinal := LoadResource(hModule, hRes);
      If hGlobal = 0 Then
        Continue;

      Var IconSize: DWORD := SizeofResource(hModule, hRes);

      // Простая запись - все иконки как 32x32
      // В реальном ICO это исправится при открытии
      Var Entry: Packed Record
  Width: Byte;
  Height: Byte;
  ColorCount: Byte;
  Reserved: Byte;
  Planes: Word;
  BitCount: Word;
  SizeInBytes: DWORD;
  FileOffset: DWORD;
End;

      Entry.Width := 32;
      Entry.Height := 32;
      Entry.ColorCount := 0;
      Entry.Reserved := 0;
      Entry.Planes := 1;
      Entry.BitCount := 32;
      Entry.SizeInBytes := IconSize;
      Entry.FileOffset := DataOffset;

      Stream.Write(Entry, SizeOf(Entry));
      Inc(DataOffset, IconSize);

      UnlockResource(hGlobal);
      FreeResource(hGlobal);
    End;

    // Данные иконок
    For i := 1 To IconCount Do
    Begin
      Var hRes: Cardinal := FindResource(hModule, PChar(i), RT_ICON);
      If hRes = 0 Then
        Continue;

      Var hGlobal: Cardinal := LoadResource(hModule, hRes);
      If hGlobal = 0 Then
        Continue;

      Var pData: Pointer := LockResource(hGlobal);
      Var IconSize: DWORD := SizeofResource(hModule, hRes);

      If (pData <> Nil) And (IconSize > 0) Then
        Stream.Write(pData^, IconSize);

      If pData <> Nil Then
        UnlockResource(hGlobal);
      FreeResource(hGlobal);
    End;

    Stream.SaveToFile(IcoFile);
    Result := FileExists(IcoFile);

  Finally
    Stream.Free;
    FreeLibrary(hModule);
  End;
End;

// Используя логику CreateICO_Simple, но для главной иконки
Function ExtractMainIconUsingSimpleLogic(Const ExeFile, IcoFile: String): Boolean;
Var
  hModule: Cardinal;
  Stream: TMemoryStream;
  i: Integer;
  MainIconID: Integer;
Begin
  Result := False;
  MainIconID := 0;

  If Not FileExists(ExeFile) Then
    Exit;

  hModule := LoadLibraryEx(PChar(ExeFile), 0, LOAD_LIBRARY_AS_DATAFILE);
  If hModule = 0 Then
    Exit;

  Try
    // Определяем ID главной иконки
    // Обычно это первая иконка в EXE файле
    For i := 1 To 10 Do
    Begin
      If FindResource(hModule, PChar(i), RT_ICON) <> 0 Then
      Begin
        MainIconID := i;
        Break;
      End;
    End;

    If MainIconID = 0 Then
    Begin
      FreeLibrary(hModule);
      Exit;
    End;

    // Ищем все иконки, принадлежащие этой группе
    // Обычно главная иконка имеет несколько размеров с последовательными ID
    Var IconIDs: TStringList := TStringList.Create;
    Try
      i := MainIconID;
      While FindResource(hModule, PChar(i), RT_ICON) <> 0 Do
      Begin
        IconIDs.Add(IntToStr(i));
        Inc(i);
        If i > MainIconID + 10 Then // Максимум 10 размеров
          Break;
      End;

      If IconIDs.Count = 0 Then
        Exit;

      Stream := TMemoryStream.Create;
      Try
        // Заголовок ICO
        Var Header: Packed Record
  Reserved: Word;
  ImageType: Word;
  ImageCount: Word;
End;

        Header.Reserved := 0;
        Header.ImageType := 1;
        Header.ImageCount := IconIDs.Count;
        Stream.Write(Header, SizeOf(Header));

        // Записи каталога
        Var DataOffset: DWORD := 6 + (IconIDs.Count * 16);

        For i := 0 To IconIDs.Count - 1 Do
        Begin
          Var IconID: Integer := StrToInt(IconIDs[i]);
          Var hRes: Cardinal := FindResource(hModule, PChar(IconID), RT_ICON);
          If hRes = 0 Then
            Continue;

          Var hGlobal: Cardinal := LoadResource(hModule, hRes);
          If hGlobal = 0 Then
            Continue;

          Var IconSize: DWORD := SizeofResource(hModule, hRes);

          // Определяем размер иконки (примерно)
          Var Width, Height: Byte;
          If IconSize < 1000 Then
          Begin
            Width := 16;
            Height := 16;
          End
          Else If IconSize < 4000 Then
          Begin
            Width := 32;
            Height := 32;
          End
          Else If IconSize < 16000 Then
          Begin
            Width := 48;
            Height := 48;
          End
          Else
          Begin
            Width := 64;
            Height := 64;
          End;

          Var Entry: Packed Record
  Width: Byte;
  Height: Byte;
  ColorCount: Byte;
  Reserved: Byte;
  Planes: Word;
  BitCount: Word;
  SizeInBytes: DWORD;
  FileOffset: DWORD;
End;

          Entry.Width := Width;
          Entry.Height := Height;
          Entry.ColorCount := 0;
          Entry.Reserved := 0;
          Entry.Planes := 1;
          Entry.BitCount := 32;
          Entry.SizeInBytes := IconSize;
          Entry.FileOffset := DataOffset;

          Stream.Write(Entry, SizeOf(Entry));
          Inc(DataOffset, IconSize);

          UnlockResource(hGlobal);
          FreeResource(hGlobal);
        End;

        // Данные иконок
        For i := 0 To IconIDs.Count - 1 Do
        Begin
          Var IconID: Integer := StrToInt(IconIDs[i]);
          Var hRes: Cardinal := FindResource(hModule, PChar(IconID), RT_ICON);
          If hRes = 0 Then
            Continue;

          Var hGlobal: Cardinal := LoadResource(hModule, hRes);
          If hGlobal = 0 Then
            Continue;

          Var pData: Pointer := LockResource(hGlobal);
          Var IconSize: DWORD := SizeofResource(hModule, hRes);

          If (pData <> Nil) And (IconSize > 0) Then
            Stream.Write(pData^, IconSize);

          If pData <> Nil Then
            UnlockResource(hGlobal);
          FreeResource(hGlobal);
        End;

        Stream.SaveToFile(IcoFile);
        Result := FileExists(IcoFile);

      Finally
        Stream.Free;
      End;

    Finally
      IconIDs.Free;
    End;

  Finally
    FreeLibrary(hModule);
  End;
End;

Procedure TForm1.EditIconFileClick(Sender: TObject);
Begin
  ImageIconClick(Sender);
End;

Procedure TForm1.ExtractIconByIndex(Const ExeFile, OutputIcoFile: String; IconIndex: Integer);
Var
  hExe: THandle;
  IconNames: TStringList;
  SelectedIconName: String;

  Function EnumCallback(Module: hModule; ResType: LPCTSTR; ResName: LPTSTR; Data: LPARAM): BOOL; Stdcall;
  Var
    List: TStringList;
    Name: String;
  Begin
    List := TStringList(Data);

    If HiWord(DWORD(ResName)) = 0 Then
      Name := 'ID:' + IntToStr(LoWord(DWORD(ResName)))
    Else
      Name := WideCharToString(ResName);

    List.Add(Name);
    Result := True;
  End;

  // Функция извлечения иконки по имени
  Function ExtractIconByName(Const IconName: String): Boolean;
  Var
    GroupResInfo, IconResInfo: HRSRC;
    GroupResData, IconResData: hGlobal;
    GroupData: Pointer;
    GroupSize: DWORD;
    GroupDir: PGRPICONDIR;
    GroupEntry: PGRPICONDIRENTRY;
    i: Integer;
    Stream: TMemoryStream;
    ResID: PChar;
    IconID: Integer;
  Begin
    Result := False;

    // Определяем тип имени ресурса
    If Copy(IconName, 1, 3) = 'ID:' Then
    Begin
      // Числовой ID
      IconID := StrToInt(Copy(IconName, 4, Length(IconName)));
      ResID := MAKEINTRESOURCE(IconID);
    End
    Else
    Begin
      // Строковое имя
      ResID := PChar(IconName);
    End;

    // Находим группу иконок
    GroupResInfo := FindResource(hExe, ResID, RT_GROUP_ICON);
    If GroupResInfo = 0 Then
      Exit;

    // Загружаем данные группы
    GroupResData := LoadResource(hExe, GroupResInfo);
    If GroupResData = 0 Then
      Exit;

    GroupData := LockResource(GroupResData);
    GroupSize := SizeofResource(hExe, GroupResInfo);

    If (GroupData = Nil) Or (GroupSize < SizeOf(TGRPICONDIR)) Then
    Begin
      FreeResource(GroupResData);
      Exit;
    End;

    Stream := TMemoryStream.Create;
    Try
      GroupDir := PGRPICONDIR(GroupData);

      // Записываем заголовок ICO
      Var IcoHeader: Packed Record
  Reserved: Word;
  ImageType: Word;
  ImageCount: Word;
End;

      IcoHeader.Reserved := 0;
      IcoHeader.ImageType := 1; // ICO
      IcoHeader.ImageCount := GroupDir^.idCount;
      Stream.Write(IcoHeader, SizeOf(IcoHeader));

      // Записи каталога
      Var DataOffset: DWORD := SizeOf(IcoHeader) + (GroupDir^.idCount * 16);
      GroupEntry := PGRPICONDIRENTRY(DWORD(GroupData) + SizeOf(TGRPICONDIR));

      For i := 0 To GroupDir^.idCount - 1 Do
      Begin
        Var IcoEntry: Packed Record
  Width: Byte;
  Height: Byte;
  ColorCount: Byte;
  Reserved: Byte;
  Planes: Word;
  BitCount: Word;
  SizeInBytes: DWORD;
  FileOffset: DWORD;
End;

        IcoEntry.Width := GroupEntry^.bWidth;
        IcoEntry.Height := GroupEntry^.bHeight;
        IcoEntry.ColorCount := GroupEntry^.bColorCount;
        IcoEntry.Reserved := GroupEntry^.bReserved;
        IcoEntry.Planes := GroupEntry^.wPlanes;
        IcoEntry.BitCount := GroupEntry^.wBitCount;
        IcoEntry.SizeInBytes := GroupEntry^.dwBytesInRes;
        IcoEntry.FileOffset := DataOffset;

        Stream.Write(IcoEntry, SizeOf(IcoEntry));
        Inc(DataOffset, GroupEntry^.dwBytesInRes);
        Inc(GroupEntry);
      End;

      // Данные иконок
      GroupEntry := PGRPICONDIRENTRY(DWORD(GroupData) + SizeOf(TGRPICONDIR));

      For i := 0 To GroupDir^.idCount - 1 Do
      Begin
        IconResInfo := FindResource(hExe, MAKEINTRESOURCE(GroupEntry^.nID), RT_ICON);
        If IconResInfo <> 0 Then
        Begin
          IconResData := LoadResource(hExe, IconResInfo);
          If IconResData <> 0 Then
          Begin
            Var IconData: Pointer := LockResource(IconResData);
            Var IconSize: DWORD := SizeofResource(hExe, IconResInfo);

            If (IconData <> Nil) And (IconSize > 0) Then
              Stream.Write(IconData^, IconSize);

            If IconData <> Nil Then
              UnlockResource(IconResData);
            FreeResource(IconResData);
          End;
        End;
        Inc(GroupEntry);
      End;

      // Сохраняем файл
      Stream.SaveToFile(OutputIcoFile);
      Result := FileExists(OutputIcoFile);

    Finally
      Stream.Free;
    End;

    UnlockResource(GroupResData);
    FreeResource(GroupResData);
  End;

Begin
  If Not FileExists(ExeFile) Then
  Begin
    ShowMessage(LangFileNotFound.Caption + ' ' + ExeFile);
    Exit;
  End;

  // Загружаем EXE файл
  hExe := LoadLibraryEx(PChar(ExeFile), 0, LOAD_LIBRARY_AS_DATAFILE);
  If hExe = 0 Then
  Begin
    ShowMessage(LangFailedUploadFile.Caption);
    Exit;
  End;

  IconNames := TStringList.Create;
  Try
    // Получаем список всех иконок
    EnumResourceNames(hExe, RT_GROUP_ICON, @EnumCallback, LPARAM(IconNames));

    If IconNames.Count = 0 Then
    Begin
      ShowMessage(LangNoIconGroups.Caption);
      FreeLibrary(hExe);
      Exit;
    End;

    // Проверяем индекс
    If (IconIndex < 0) Or (IconIndex >= IconNames.Count) Then
    Begin
      ShowMessage(LangInvalidIconIndex.Caption);
      FreeLibrary(hExe);
      IconNames.Free;
      Exit;
    End;

    // Извлекаем иконку по указанному индексу
    SelectedIconName := IconNames[IconIndex];
    ExtractIconByName(SelectedIconName);
   { If ExtractIconByName(SelectedIconName) Then
    Begin
      ShowMessage(OutputIcoFile);
    End
    Else
    Begin
      ShowMessage(LangErrorExtractingIcon.Caption);
    End; }

  Finally
    IconNames.Free;
    FreeLibrary(hExe);
  End;
End;

Procedure TForm1.ButtonSaveIconClick(Sender: TObject);
Var
  BaseName: String;
Begin
  // Берем имя исходного файла без расширения
  BaseName := ChangeFileExt(ExtractFileName(EditIconFile.Text), '');
  If BaseName = '' Then
    BaseName := 'icon';
  SavePictureDialog.FileName := BaseName + '.ico';
  If SavePictureDialog.Execute Then
  Begin
    ExtractIconByIndex(EditIconFile.Text, SavePictureDialog.FileName, SpinEditIcon.Value);
  End;
End;

Procedure TForm1.ButtonBrowseIconClick(Sender: TObject);
Begin
  ImageIconClick(Sender);
End;

Procedure TForm1.Button5Click(Sender: TObject);
Begin
  IconDelete;
End;

Procedure TForm1.ButtonRefreshClick(Sender: TObject);
Begin
  Form1.StatusBar.Panels[0].Text := '';
  Form1.ImageFolder.Picture := Nil;
  Form1.ImageIcon.Picture := Nil;
  EditIconFile.Text := '';
  EditFolder.Text := '';
  EditInfoTip.Text := '';
  If Form1.PathFolder.Text = '' Then
    Exit;
  If SysUtils.DirectoryExists(StringReplace(PathFolder.Text, '\\', '\', [rfReplaceAll])) = true Then
  Begin
    ScanAndAddFoldersToListView(StringReplace(PathFolder.Text, '\\', '\', [rfReplaceAll]), ListViewList);
    Application.ProcessMessages;
  End;
End;

Procedure TForm1.ButtonUPClick(Sender: TObject);
Begin
  Form1.StatusBar.Panels[0].Text := '';
  Form1.ImageFolder.Picture := Nil;
  Form1.ImageIcon.Picture := Nil;
  EditIconFile.Text := '';
  EditFolder.Text := '';
  EditInfoTip.Text := '';
  If PathFolder.Text = ExtractFileDrive(PathFolder.Text) Then
    Exit;
  If PathFolder.Text = ExtractFileDrive(PathFolder.Text) + '\' Then
    Exit;
  If PathFolder.Text <> ExtractFileDrive(PathFolder.Text) Then
  Begin
    PathFolder.Text := GetParentDir(StringReplace(PathFolder.Text, '\\', '\', [rfReplaceAll]));
    // cbPathChange(Sender);
    ScanAndAddFoldersToListView(StringReplace(PathFolder.Text, '\\', '\', [rfReplaceAll]), ListViewList);
    Application.ProcessMessages;
  End;
End;

Procedure TForm1.WMMouseMove(Var Message: TWMMouseMove);
Begin
  SpeedButton_GeneralMenu.AllowAllUp := true;
  SpeedButton_GeneralMenu.Down := false;
End;

Procedure TForm1.SpeedButton_GeneralMenuClick(Sender: TObject);
Var
  ButtonRight: TPoint;
Begin
  If (SpeedButton_GeneralMenu.AllowAllUp) Then
  Begin
    SpeedButton_GeneralMenu.AllowAllUp := False;
    SpeedButton_GeneralMenu.Down := true;
    Application.ProcessMessages;
    If Sender Is TControl Then
    Begin
      ButtonRight.X := SpeedButton_GeneralMenu.Left;
      ButtonRight.Y := SpeedButton_GeneralMenu.Top;
      ButtonRight := ClientToScreen(ButtonRight);
      PopupMenuGeneral.Popup(ButtonRight.X + SpeedButton_GeneralMenu.Width, ButtonRight.Y + SpeedButton_GeneralMenu.Height);
    End;
  End
  Else
  Begin
    SpeedButton_GeneralMenu.AllowAllUp := true;
    SpeedButton_GeneralMenu.Down := False;
    PopupMenuGeneral.CloseMenu;
    Application.ProcessMessages;
  End;
End;

Procedure TForm1.SpinEditIconChange(Sender: TObject);
Var
  a: Array[0..MAX_PATH] Of Char;
  hIcon, hNewIcon: THandle;
Begin
  If Trim(EditIconFile.Text) = '' Then
  Begin
    ImageIcon.Picture := Nil;
    Exit;
  End;

  If Not FileExists(EditIconFile.Text) Then
  Begin
    ImageIcon.Picture := Nil;
    Exit;
  End;

  StrPCopy(a, EditIconFile.Text);

  // Освобождаем старую иконку
  If ImageIcon.Picture.Icon.Handle <> 0 Then
    ImageIcon.Picture.Icon.ReleaseHandle;

  // Извлекаем иконку
  hIcon := ExtractIcon(HInstance, a, SpinEditIcon.Value);

  If hIcon > 0 Then
  Begin
    Try
      // Масштабируем иконку до 48x48 с помощью CopyImage
      hNewIcon := CopyImage(hIcon, IMAGE_ICON, 48, 48, LR_COPYFROMRESOURCE);

      If hNewIcon <> 0 Then
      Begin
        ImageIcon.Picture.Icon.Handle := hNewIcon;

        // Настраиваем TImage
        ImageIcon.AutoSize := True;

        // Уничтожаем оригинальную иконку (скопированную)
        DestroyIcon(hIcon);
      End
      Else
      Begin
        // Если CopyImage не сработал, используем оригинальную
        ImageIcon.Picture.Icon.Handle := hIcon;
        ImageIcon.AutoSize := True;
      End;
    Except
      DestroyIcon(hIcon);
      ImageIcon.Picture := Nil;
    End;
  End
  Else
  Begin
    ImageIcon.Picture := Nil;
  End;
End;

Procedure TForm1.SpinEditIconKeyDown(Sender: TObject; Var Key: Word; Shift: TShiftState);
Begin
  If Not (Key In [VK_UP, VK_DOWN, VK_LEFT, VK_RIGHT, VK_HOME, VK_END, VK_BACK, VK_DELETE, VK_TAB, VK_RETURN, VK_ESCAPE]) Then
  Begin
    Key := 0;
  End;
End;

Procedure TForm1.SpinEditIconKeyPress(Sender: TObject; Var Key: Char);
Begin
  Key := #0;
End;

Procedure TForm1.FormActivate(Sender: TObject);
Begin
  DragAcceptFiles(Form1.Handle, true);
End;

Procedure TForm1.FormClose(Sender: TObject; Var Action: TCloseAction);
Begin
  Form1.SaveNastr;
  DragAcceptFiles(Handle, false);
  ListViewList.WindowProc := FListViewWndProc;
  FListViewWndProc := Nil;
End;

Procedure TForm1.CheckUpdateMenuClick(Sender: TObject);
Begin
  SpeedButton_GeneralMenu.AllowAllUp := true;
  SpeedButton_GeneralMenu.Down := False;
  Form10.ShowModal;
End;

Procedure TForm1.CleanTranslationsLikeGlobload;
Var
  i, j, k, m: Integer;
  Ini: TMemIniFile;
  Sections, Keys: TStringList;
  SearchRec: TSearchRec;
  FindResult: Integer;
  CompPath, FormName, CompName, PropName: String;
  FirstDot, SecondDot: Integer;
  FormExists, CompExists: Boolean;
  CurrentForm: TForm;
  CurrentComponent: TComponent;
  Modified: Boolean;
  IsDuplicate: Boolean;
  n: Integer;
  CompareKey, CompareFormName: String;
  CompareDotPos: Integer;
Begin
  // Создаем все формы проекта (если нужно)
  // CreateAllProjectForms;

  FindResult := FindFirst(ExtractFilePath(ParamStr(0)) + 'Language\*.ini', faAnyFile, SearchRec);
  If FindResult = 0 Then
  Begin
    Repeat
      If (SearchRec.Name <> '.') And (SearchRec.Name <> '..') Then
      Begin
        Ini := TMemIniFile.Create(ExtractFilePath(ParamStr(0)) + 'Language\' + SearchRec.Name);
        Sections := TStringList.Create;
        Keys := TStringList.Create;
        Modified := False;

        Try
          Ini.ReadSections(Sections);

          For i := 0 To Sections.Count - 1 Do
          Begin
            If SameText(Sections[i], Application.Title) Then
            Begin
            // ========== ИСКЛЮЧАЕМ ЭТИ СЕКЦИИ ИЗ ОБРАБОТКИ ==========
              If SameText(Sections[i], 'Language information') Or SameText(Sections[i], 'DestDir') Then
                Continue; // Пропускаем эти секции полностью

              Ini.ReadSection(Sections[i], Keys);

            // Проходим по всем ключам в обратном порядке
              For j := Keys.Count - 1 Downto 0 Do
              Begin
                CompPath := Keys[j];
                FirstDot := Pos('.', CompPath);

                If FirstDot > 0 Then
                Begin
                  FormName := Copy(CompPath, 1, FirstDot - 1);
                  FormExists := False;
                  CompExists := False;

                // ==================== ПРОВЕРКА СУЩЕСТВОВАНИЯ КОМПОНЕНТА ====================
                // Проверяем ВСЕ формы в Screen
                  For k := 0 To Screen.FormCount - 1 Do
                  Begin
                    If SameText(Screen.Forms[k].Name, FormName) Then
                    Begin
                      FormExists := True;
                      CurrentForm := Screen.Forms[k];

                    // Извлекаем остаток пути после имени формы
                      CompName := Copy(CompPath, FirstDot + 1, Length(CompPath));
                      SecondDot := Pos('.', CompName);

                      If SecondDot > 0 Then
                      Begin
                      // Есть вложенный компонент: Form1.TrayIcon1.Hint
                        PropName := Copy(CompName, SecondDot + 1, Length(CompName));
                        CompName := Copy(CompName, 1, SecondDot - 1);

                      // Ищем компонент на форме
                        CurrentComponent := CurrentForm.FindComponent(CompName);

                      // Если не нашли через FindComponent, ищем вручную
                        If CurrentComponent = Nil Then
                        Begin
                          For m := 0 To CurrentForm.ComponentCount - 1 Do
                          Begin
                            If SameText(CurrentForm.Components[m].Name, CompName) Then
                            Begin
                              CurrentComponent := CurrentForm.Components[m];
                              Break;
                            End;
                          End;
                        End;

                        CompExists := (CurrentComponent <> Nil);
                      End
                      Else
                      Begin
                      // Нет второй точки - это свойство формы (Form1.Caption)
                        CompExists := True;
                      End;

                      Break; // Форма найдена, выходим из цикла
                    End;
                  End;

                // ==================== ПРОВЕРКА ДУБЛИКАТОВ ====================
                  IsDuplicate := False;
                // Проверяем предыдущие ключи на дубликаты (только внутри той же формы)
                  For n := 0 To j - 1 Do
                  Begin
                    CompareKey := Keys[n];
                    CompareDotPos := Pos('.', CompareKey);

                    If CompareDotPos > 0 Then
                    Begin
                      CompareFormName := Copy(CompareKey, 1, CompareDotPos - 1);

                    // Дубликатом считаем только если:
                    // 1. Имя формы совпадает
                    // 2. Полный путь совпадает (регистронезависимо)
                      If (SameText(FormName, CompareFormName)) And (SameText(CompPath, CompareKey)) Then
                      Begin
                        IsDuplicate := True;
                        Break;
                      End;
                    End;
                  End;

                // ==================== УДАЛЕНИЕ КЛЮЧА ====================
                // Удаляем если:
                // 1. Форма или компонент не существуют ИЛИ
                // 2. Найден дубликат в той же форме
                  If (Not (FormExists And CompExists)) Or IsDuplicate Then
                  Begin
                    Ini.DeleteKey(Sections[i], Keys[j]);
                    Modified := True;
                  End;
                End
                Else
                Begin
                // Некорректный формат - удаляем
                  Ini.DeleteKey(Sections[i], Keys[j]);
                  Modified := True;
                End;
              End;

            // Проверяем, не пустая ли секция после удаления
            // (кроме исключенных секций)
              If Not (SameText(Sections[i], 'Language information') Or SameText(Sections[i], 'DestDir')) Then
              Begin
                Ini.ReadSection(Sections[i], Keys);
                If Keys.Count = 0 Then
                Begin
                  Ini.EraseSection(Sections[i]);
                  Modified := True;
                End;
              End;
            End;

            If Modified Then
              Ini.UpdateFile;
          End;

        Finally
          Keys.Free;
          Sections.Free;
          Ini.Free;
        End;
      End;
    Until FindNext(SearchRec) <> 0;
    FindClose(SearchRec);
  End;
End;

Procedure TForm1.LoadLanguage;
Var
  Ini: TMemIniFile;
  LangFiles: TStringList;
  i: Integer;
  MenuItem: TMenuItem;
  FileName, DisplayName, MenuCaption: String;
  SearchRec: TSearchRec;
Begin
  Ini := TMemIniFile.Create(PortablePath);
  While LanguageMenu.Count > 0 Do
    LanguageMenu.Items[0].Free;
  LangLocal := Ini.ReadString('Language', 'Language', '');
  Ini.Free;

  LangFiles := TStringList.Create;
  Try
    If FindFirst(ExtractFilePath(ParamStr(0)) + 'Language\*.ini', faAnyFile, SearchRec) = 0 Then
    Begin
      Repeat
        If (SearchRec.Name <> '.') And (SearchRec.Name <> '..') Then
          LangFiles.Add(SearchRec.Name);
      Until FindNext(SearchRec) <> 0;
      FindClose(SearchRec);
    End;
    LangFiles.Sort;
    For i := 0 To LangFiles.Count - 1 Do
    Begin
      FileName := LangFiles[i];
      LangCode := ChangeFileExt(FileName, '');
      DisplayName := LangCode;
      Try
        Ini := TMemIniFile.Create(ExtractFilePath(ParamStr(0)) + 'Language\' + FileName);
        Try
          DisplayName := Ini.ReadString('Language information', 'LANGNAME', LangCode);
        Finally
          Ini.Free;
        End;
      Except
      End;
      MenuCaption := LangCode + #9 + #9 + DisplayName;
      MenuItem := TMenuItem.Create(LanguageMenu);
      MenuItem.RadioItem := true;
      MenuItem.Caption := MenuCaption;
      MenuItem.AutoHotkeys := maManual;
      MenuItem.AutoCheck := True;
      If (LangCode = LangLocal) Or (SameText(LangCode, LangLocal)) Or (LangCode + '.ini' = LangLocal) Then
        MenuItem.Checked := True;
      MenuItem.OnClick := LanguageMenuItemClick;
      LanguageMenu.Add(MenuItem);
    End;

  Finally
    LangFiles.Free;
  End;
End;

Procedure TForm1.LanguageMenuItemClick(Sender: TObject);
Var
  MenuItem: TMenuItem;
  Ini: TMemIniFile;
  i: Integer;
Begin
  If Sender Is TMenuItem Then
  Begin
    MenuItem := TMenuItem(Sender);
    LangCode := Copy(MenuItem.Caption, 1, Pos(#9, MenuItem.Caption) - 1);
    LangLocal := LangCode;
    For i := 0 To LanguageMenu.Count - 1 Do
      LanguageMenu.Items[i].Checked := False;
    MenuItem.Checked := True;
    Ini := TMemIniFile.Create(PortablePath);
    Try
      Ini.WriteString('Language', 'Language', LangLocal);
      Ini.UpdateFile;
    Finally
      Ini.Free;
    End;
    LoadLanguage;
    Form1.Globload;
    RefreshListMenuClick(Sender);
  End;
End;

Function GetApplicationBitness: String;
Begin
  {$IFDEF WIN64}
  Result := '(64-bit)';
  {$ELSE}
  Result := '(32-bit)';
  {$ENDIF}
End;

Procedure TForm1.FormCreate(Sender: TObject);
Begin
  portable := fileExists(ExtractFilePath(ParamStr(0)) + 'portable.ini');
  If portable = True Then
    Form1.Caption := Form1.Caption + ' ' + GetFileVersion(ParamStr(0)) + ' ' + GetApplicationBitness + ' Portable'
  Else
    Form1.Caption := Form1.Caption + ' ' + GetFileVersion(ParamStr(0)) + ' ' + GetApplicationBitness;
  OnMouseWheel := FormMouseWheel;
  Form1.globload;

  DragAcceptFiles(Handle, true);
  createicon(ListViewList);
  Application.ProcessMessages;
  Form1.SpinEditIcon.MinValue := 0;
  Application.ProcessMessages;
  FShowHoriz := false;
  FShowVert := true;
  FListViewWndProc := ListViewList.WindowProc;
  ListViewList.WindowProc := ListViewWndProc;
  getDrive;
  Form1.ComboBoxDriveChange(self);
  Form1.StatusBar2.Panels[0].Text := LangElements.Caption;
  Form1.LoadNastr;
End;

Procedure TForm1.FormDestroy(Sender: TObject);
Begin
  DragAcceptFiles(Handle, False);
End;

Procedure TForm1.FormDragOver(Sender, Source: TObject; X, Y: Integer; State: TDragState; Var Accept: Boolean);
Begin
  Accept := True;
End;

Procedure TForm1.FormMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
Begin
  SpeedButton_GeneralMenu.AllowAllUp := true;
  SpeedButton_GeneralMenu.Down := False;
End;

Procedure TForm1.FormMouseWheel(Sender: TObject; Shift: TShiftState; WheelDelta: Integer; MousePos: TPoint; Var Handled: Boolean);
Var
  ComponentAtPosition: TControl;
Begin
  // Получаем компонент, находящийся под курсором мыши
  ComponentAtPosition := FindDragTarget(Mouse.CursorPos, False);

  // Проверяем, является ли компонент типом TSpinEdit
  If Assigned(ComponentAtPosition) And (ComponentAtPosition Is TSpinEdit) Then
  Begin
    With TComponent(ComponentAtPosition) As TSpinEdit Do
    Begin
      SetFocus();
      If WheelDelta > 0 Then
        Value := Value + 1
      Else
        Value := Value - 1;

      SelectAll(); // Выделение текста

      Handled := True;
    End;
  End;
End;

Procedure RefreshScreenIcons;
Begin
  SendMessage(HWND_BROADCAST, WM_SETTINGCHANGE, SPI_SETICONMETRICS, 0);
End;

Procedure TForm1.getIcon;
Var
  IconFile, IconResource, IconResourceIndex, IconIndex: String;
  ini: TMemIniFile;
  ic: TIcon;
  a: Integer;
  FolderPath, FullPath: String;
Begin
  If ListViewList.ItemIndex = -1 Then
    Exit;

  ic := TIcon.Create;
  Try
    FolderPath := StringReplace(ListViewList.Selected.SubItems[0], '\\', '\', [rfReplaceAll]);
    // Добавляем обратный слеш если его нет в конце
    If (FolderPath <> '') And (FolderPath[Length(FolderPath)] <> '\') Then
      FolderPath := FolderPath + '\';
    EditFolder.Text := FolderPath;

    Application.ProcessMessages;
    GetIconFromFile(FolderPath, ic, SHIL_EXTRALARGE);
    ImageFolder.Picture.Icon := ic;

    SpinEditIcon.Value := 0;
    EditInfoTip.Text := '';
    EditIconFile.Text := '';
    ImageIcon.Picture.Icon := Nil;

    If FileExists(FolderPath + 'Desktop.ini') Then
    Begin
      ini := TMemIniFile.Create(FolderPath + 'Desktop.ini');
      Try
        IconFile := ini.ReadString('.ShellClassInfo', 'IconFile', '');
        IconIndex := ini.ReadString('.ShellClassInfo', 'IconIndex', '');
        IconResource := ini.ReadString('.ShellClassInfo', 'IconResource', '');
        EditInfoTip.Text := ini.ReadString('.ShellClassInfo', 'InfoTip', '');
      Finally
        ini.Free;
      End;

      // Обработка путей
      IconFile := StringReplace(IconFile, '%SystemRoot%', GetSpecialFolderPath(CSIDL_WINDOWS), [rfReplaceAll]);
      IconResource := StringReplace(IconResource, '%SystemRoot%', GetSpecialFolderPath(CSIDL_WINDOWS), [rfReplaceAll]);
      IconResource := StringReplace(IconResource, '.\', '', [rfReplaceAll]);

      IconResourceIndex := IconResource;
      IconResource := Work(IconResource, ',');

      // Приоритетная проверка IconResource (локальный путь)
      If FileExists(FolderPath + IconResource) Then
      Begin
        FullPath := FolderPath + IconResource;
        EditIconFile.Text := FullPath;
        GetIconFromFile(FullPath, ic, SHIL_EXTRALARGE);
        ImageIcon.Picture.Icon := ic;

        If IconResourceIndex <> '' Then
        Begin
          Delete(IconResourceIndex, 1, Pos(',', IconResourceIndex));
          a := StrToIntDef(IconResourceIndex, 0);
          SpinEditIcon.Value := a;
        End;
      End
      // Проверка IconResource (абсолютный путь)
      Else If FileExists(IconResource) Then
      Begin
        FullPath := IconResource;
        EditIconFile.Text := FullPath;
        GetIconFromFile(FullPath, ic, SHIL_EXTRALARGE);
        ImageIcon.Picture.Icon := ic;

        If IconResourceIndex <> '' Then
        Begin
          Delete(IconResourceIndex, 1, Pos(',', IconResourceIndex));
          a := StrToIntDef(IconResourceIndex, 0);
          SpinEditIcon.Value := a;
        End;
      End
      // Проверка IconFile (локальный путь)
      Else If FileExists(FolderPath + IconFile) Then
      Begin
        FullPath := FolderPath + IconFile;
        EditIconFile.Text := FullPath;
        GetIconFromFile(FullPath, ic, SHIL_EXTRALARGE);
        ImageIcon.Picture.Icon := ic;

        If IconIndex <> '' Then
        Begin
          a := StrToIntDef(IconIndex, 0);
          SpinEditIcon.Value := a;
        End;
      End
      // Проверка IconFile (абсолютный путь)
      Else If FileExists(IconFile) Then
      Begin
        FullPath := IconFile;
        EditIconFile.Text := FullPath;
        GetIconFromFile(FullPath, ic, SHIL_EXTRALARGE);
        ImageIcon.Picture.Icon := ic;

        If IconIndex <> '' Then
        Begin
          a := StrToIntDef(IconIndex, 0);
          SpinEditIcon.Value := a;
        End;
      End;

      StatusBar.Panels[0].Text := '';
    End;

    Application.ProcessMessages;

  Finally
    ic.Free;
  End;

  // Обновление контролов файла иконки
 {
  If ExtractFileExt(EditIconFile.Text) = '.exe' Then
  Begin
    CheckBoxCopyIcon.Checked := False;
    CheckBoxCopyIcon.Enabled := False;
  End
  Else If ExtractFileExt(EditIconFile.Text) = '.dll' Then
  Begin
    CheckBoxCopyIcon.Checked := False;
    CheckBoxCopyIcon.Enabled := False;
  End
  Else If ExtractFileExt(EditIconFile.Text) = '.ico' Then
  Begin
    CheckBoxCopyIcon.Enabled := True;
    // CheckBoxCopyIcon.Checked := True;
  End
  Else
  Begin
    CheckBoxCopyIcon.Enabled := True;
  End;  }
End;

Procedure TForm1.ListViewListClick(Sender: TObject);
Begin
  UpdateFormState;
End;

Procedure TForm1.ListViewListDblClick(Sender: TObject);
Begin
  If ListViewList.ItemIndex = -1 Then
    exit;
  Form1.StatusBar.Panels[0].Text := '';
  Form1.StatusBar.Panels[0].Text := '';
  Form1.ImageFolder.Picture := Nil;
  Form1.ImageIcon.Picture := Nil;
  EditIconFile.Text := '';
  EditFolder.Text := '';
  EditInfoTip.Text := '';
  PathFolder.Text := PathFolder.Text + ListViewList.Selected.Caption + '\';
  ScanAndAddFoldersToListView(StringReplace(PathFolder.Text, '\\', '\', [rfReplaceAll]), ListViewList);
    // Ваш существующий код для выбора первого элемента
  If ListViewList.Items.Count > 0 Then
  Begin
    Try
      ListViewList.Items[0].Selected := True;
      ListViewList.ItemIndex := 0; // Альтернативный способ выделения
      ListViewList.Items[0].MakeVisible(True);
    Except
    End;
  End;
  Application.ProcessMessages;
End;

Procedure TForm1.ListViewListSelectItem(Sender: TObject; Item: TListItem; Selected: Boolean);
Begin
  UpdateFormState;
End;

Procedure TForm1.GetAttrFolder;
Var
  FileName: String;
  Attr: DWORD;
Begin
  Try
    FileName := StringReplace(Form1.PathFolder.Text + '\' + Form1.ListViewList.Selected.Caption, '\\', '\', [rfReplaceAll]);

    // Получение атрибутов файла через Windows API
    Attr := GetFileAttributes(PChar(FileName));

    If Attr = INVALID_FILE_ATTRIBUTES Then
      RaiseLastOSError; // Генерируем исключение при ошибке

    // Установка значений чекбоксов
    Form1.CheckBoxReadOnly.Checked := (Attr And FILE_ATTRIBUTE_READONLY) <> 0;
    Form1.CheckBoxHide.Checked := (Attr And FILE_ATTRIBUTE_HIDDEN) <> 0;
    Form1.CheckBoxSys.Checked := (Attr And FILE_ATTRIBUTE_SYSTEM) <> 0;
  Except
    On E: Exception Do
    Begin
      // Обработка ошибок
      Form1.CheckBoxReadOnly.Checked := False;
      Form1.CheckBoxHide.Checked := False;
      Form1.CheckBoxSys.Checked := False;

      // Можно показать сообщение об ошибке
      // ShowMessage('Ошибка при получении атрибутов: ' + E.Message);
    End;
  End;
End;

Procedure TForm1.UpdateFormState;
Var
  IconExt: String;
Begin
  If ListViewList.ItemIndex = -1 Then
  Begin
    EditFolder.Text := '';
    Form1.ImageFolder.Picture := Nil;
    Form1.ImageIcon.Picture := Nil;
    EditIconFile.Text := '';
    EditInfoTip.Text := '';
    MemoLog.Clear;
    Form1.CheckBoxSys.Checked := false;
    Form1.CheckBoxHide.Checked := false;
    Form1.CheckBoxReadOnly.Checked := false;
    ButtonBrowseIcon.Enabled := false;
    ButtonGetInfo.Enabled := false;
    ButtonChangeIcon.Enabled := false;
    CheckBoxCopyIcon.Enabled := false;
  End;
  If ListViewList.ItemIndex <> -1 Then
  Begin
    Form1.ImageFolder.Picture := Nil;
    Form1.ImageIcon.Picture := Nil;
    ButtonBrowseIcon.Enabled := true;
    ButtonGetInfo.Enabled := true;
    GetAttrFolder;
    getIcon;
    Application.ProcessMessages;
    MemoLog.Clear;
    If fileexists(EditFolder.Text + 'Desktop.ini') Then
    Begin
      MemoLog.Lines.LoadFromFile(EditFolder.Text + 'Desktop.ini');
      Application.ProcessMessages;
    End;
    If Trim(EditIconFile.Text) = '' Then
    Begin
      SpinEditIcon.Enabled := false;
      ButtonSaveIcon.Enabled := false;
      CheckBoxCopyIcon.Enabled := false;
    End;
    If Trim(EditIconFile.Text) <> '' Then
    Begin
      IconExt := LowerCase(ExtractFileExt(EditIconFile.Text));
      If (IconExt = '.exe') Or (IconExt = '.dll') Then
      Begin
        SpinEditIcon.Enabled := true;
        ButtonSaveIcon.Enabled := true;
        CheckBoxCopyIcon.Enabled := true;
      End;

      If IconExt = '.ico' Then
      Begin
        SpinEditIcon.Enabled := false;
        CheckBoxCopyIcon.Enabled := true;
        ButtonSaveIcon.Enabled := false;
      End;
    End;
  End;
End;

Procedure TForm1.Globload;
Var
  i: Integer;
  Internat: TTranslation;
  Ini: TMemIniFile;
  lang, lang_file: String;
Begin
  For i := 0 To Screen.FormCount - 1 Do
  Begin
    Ini := TMemIniFile.Create(Form1.PortablePath);
    lang := Ini.ReadString('Language', 'Language', '');
    lang_file := ExtractFilePath(ParamStr(0)) + 'Language\' + lang + '.ini';
    Ini.Free;
    Ini := TMemIniFile.Create(lang_file);
    If Ini.SectionExists(Application.Title) Then // Используем конкретную секцию
    Begin
      Internat.Execute(Screen.Forms[i], Application.Title); // Передаем имя секции
      Application.ProcessMessages;
    End;
    Ini.Free;
  End;
End;

Procedure TForm1.MenuItem1Click(Sender: TObject);
Begin
  SpeedButton_GeneralMenu.AllowAllUp := true;
  SpeedButton_GeneralMenu.Down := False;
  ButtonRefreshClick(Sender);
End;

Procedure TForm1.SetIconsForSize(UseLargeIcons: Boolean);
Var
  {$IFDEF WIN64}
  SysImageList: NativeUInt;
  {$ELSE}
  SysImageList: Cardinal;
  {$ENDIF}
  SFI: TSHFileInfo;
  Flags: DWORD;
Begin
  SpeedButton_GeneralMenu.AllowAllUp := true;
  SpeedButton_GeneralMenu.Down := False;

  If UseLargeIcons Then
    Flags := SHGFI_SYSICONINDEX Or SHGFI_LARGEICON
  Else
    Flags := SHGFI_SYSICONINDEX Or SHGFI_SMALLICON;

  // Получаем список иконок
  {$IFDEF WIN64}
  SysImageList := NativeUInt(SHGetFileInfo('', 0, SFI, sizeof(TSHFileInfo), Flags));
  {$ELSE}
  SysImageList := Cardinal(SHGetFileInfo('', 0, SFI, sizeof(TSHFileInfo), Flags));
  {$ENDIF}

  If SysImageList <> 0 Then
  Begin
    // Используем один и тот же список для обоих размеров
    Form1.ListViewList.LargeImages.Handle := SysImageList;
    Form1.ListViewList.LargeImages.ShareImages := true;

    Form1.ListViewList.SmallImages.Handle := SysImageList;
    Form1.ListViewList.SmallImages.ShareImages := true;
  End;

  // Обновляем состояние меню
  IconSmallMenu.Checked := Not UseLargeIcons;
  IconBigMenu.Checked := UseLargeIcons;
End;

Procedure TForm1.IconBigMenuClick(Sender: TObject);
Begin
  SetListViewSystemIcons(ListViewList, True);   // Большие иконки
End;

Procedure TForm1.IconSmallMenuClick(Sender: TObject);
Begin
  SetListViewSystemIcons(ListViewList, False);  // Маленькие иконки
End;

Function TForm1.GetIconCountSimple(Const ExeFile: String): Integer;
Var
  hExe: THandle;
  IconList: TStringList;

  // Локальная функция обратного вызова

  Function EnumProc(hModule: hModule; lpType: LPCTSTR; lpName: LPTSTR; lParam: lParam): BOOL; Stdcall;
  Begin
    // Просто добавляем пустую строку для подсчета
    TStringList(lParam).Add('');
    Result := True;
  End;

Begin
  Result := 0;

  If Not FileExists(ExeFile) Then
    Exit;

  hExe := LoadLibraryEx(PChar(ExeFile), 0, LOAD_LIBRARY_AS_DATAFILE);
  If hExe = 0 Then
    Exit;

  IconList := TStringList.Create;
  Try
    // Перечисляем иконки
    EnumResourceNames(hExe, RT_GROUP_ICON, @EnumProc, lParam(IconList));
    Result := IconList.Count;
  Finally
    IconList.Free;
    FreeLibrary(hExe);
  End;
End;

Procedure TForm1.ImageIconClick(Sender: TObject);
Var
  ic: ticon;
Begin
  If OpenDialog.Execute Then
  Begin
    If (LowerCase(ExtractFileExt(OpenDialog.FileName)) = '.exe') Or (LowerCase(ExtractFileExt(OpenDialog.FileName)) = '.dll') Then
    Begin
      ic := ticon.Create;
      Form1.EditIconFile.Text := OpenDialog.FileName;
      Form1.SpinEditIcon.Value := 0;
      SpinEditIcon.MaxValue := 0;
      SpinEditIcon.MinValue := 0;
      If (GetIconCountSimple(Form1.EditIconFile.Text) = 1) Or (GetIconCountSimple(Form1.EditIconFile.Text) = 0) Then
      Begin
        SpinEditIcon.MaxValue := GetIconCountSimple(Form1.EditIconFile.Text);
        SpinEditIcon.MinValue := 0;
      End;
      If (GetIconCountSimple(Form1.EditIconFile.Text) <> 1) And (GetIconCountSimple(Form1.EditIconFile.Text) <> 0) Then
      Begin
        SpinEditIcon.MaxValue := GetIconCountSimple(Form1.EditIconFile.Text) - 1;
        SpinEditIcon.MinValue := 0;
        SpinEditIconChange(Sender);
      End;
      GetIconFromFile(EditIconFile.Text, ic, SHIL_EXTRALARGE);
      ImageIcon.Picture.Icon := ic;
      ic.Free;
      Form1.StatusBar.Panels[0].Text := '';
      SpinEditIcon.Enabled := true;
      ButtonSaveIcon.Enabled := true;
      CheckBoxCopyIcon.Enabled:=true;
    End;

    If LowerCase(ExtractFileExt(OpenDialog.FileName)) = '.ico' Then
    Begin
      ic := ticon.Create;
      Form1.EditIconFile.Text := OpenDialog.FileName;
      GetIconFromFile(EditIconFile.Text, ic, SHIL_EXTRALARGE);
      ImageIcon.Picture.Icon := ic;
      ic.Free;
      SpinEditIcon.Enabled := false;
      ButtonSaveIcon.Enabled := false;
      CheckBoxCopyIcon.Enabled:=true;
    End;

  End;
  If EditIconFile.Text <> '' Then
  Begin
  CheckBoxCopyIcon.Enabled:=true;
    ButtonChangeIcon.Enabled := true;
    ButtonChangeIcon.Default := True;
    ButtonChangeIcon.SetFocus;
  End;
End;


Procedure TForm1.RefreshListMenuClick(Sender: TObject);
Begin
  SpeedButton_GeneralMenu.AllowAllUp := true;
  SpeedButton_GeneralMenu.Down := False;
  ButtonRefreshClick(Sender);
End;


Procedure TForm1.RestartExplorerMenu1Click(Sender: TObject);
Begin
  Try
    SpeedButton_GeneralMenu.AllowAllUp := true;
    SpeedButton_GeneralMenu.Down := False;
    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, Nil, Nil);
    SendMessage(FindWindow('Progman', 'Program Manager'), WM_COMMAND, $A065, 0);
    KillTask('explorer.exe');
    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, Nil, Nil);
    SendMessage(FindWindow('Progman', 'Program Manager'), WM_COMMAND, $A065, 0);
    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, Nil, Nil);
  Except
  End;
End;

Function GetAppDataRoamingPath: String;
Var
  Path: Array[0..MAX_PATH] Of Char;
Begin
  If SUCCEEDED(SHGetFolderPath(0, CSIDL_APPDATA, 0, 0, @Path[0])) Then
    Result := IncludeTrailingPathDelimiter(Path)
  Else
    Result := '';
End;

Function TForm1.PortablePath: String;
Begin
  If portable Then
    Result := ExtractFilePath(ParamStr(0)) + 'Config\Option.ini'
  Else
    Result := IncludeTrailingPathDelimiter(GetAppDataRoamingPath) + IncludeTrailingPathDelimiter(GetCompanyName(ParamStr(0))) + Application.Title + '\Config\Option.ini';
    //Result := ExtractFilePath(ParamStr(0)) + 'Config\Option.ini';
  // Создаем папку для конфига
  SysUtils.ForceDirectories(ExtractFilePath(Result));
End;

Procedure TForm1.LoadNastr;
Var
  ini: TMemIniFile;
  i: integer;
Begin

  ini := TMemIniFile.Create(Form1.PortablePath);

  Try
    For i := 0 To ThemeMenu.Count - 1 Do
    Begin
      ThemeMenu.Items[i].Checked := ini.ReadBool('Option', ThemeMenu.Items[i].Name, False);
      If ThemeMenu.Items[i].Checked Then
        ThemeMenu.Items[i].Click;
    End;

    Form1.CheckBoxCopyIcon.Checked := ini.ReadBool('Option', Form1.CheckBoxCopyIcon.Name, false);

    IconSmallMenu.Checked := ini.ReadBool('Option', IconSmallMenu.Name, false);
    If IconSmallMenu.Checked Then
    Begin
      IconSmallMenuClick(self);
    End;

    IconBigMenu.Checked := ini.ReadBool('Option', IconBigMenu.Name, false);
    If IconBigMenu.Checked Then
    Begin
      IconBigMenuClick(self);
    End;

    If ini.readString('Option', 'Window', '') = 'wsMaximized' Then
    Begin
      Form1.WindowState := wsMaximized;
    End;

    If ini.readString('Option', 'Window', '') <> 'wsMaximized' Then
    Begin
      Form1.top := ini.ReadInteger('Option', 'Top', Form1.top);
      Form1.left := ini.ReadInteger('Option', 'Left', Form1.left);
      Form1.Height := ini.ReadInteger('Option', 'Height', Form1.Height);
      Form1.Width := ini.ReadInteger('Option', 'Width', Form1.Width);
    End;
  Finally
    ini.Free;
  End;
End;

Procedure TForm1.SaveNastr;
Var
  ini: TMemIniFile;
  i: integer;
Begin
  ini := TMemIniFile.Create(Form1.PortablePath);

  For i := 0 To ThemeMenu.Count - 1 Do
  Begin
    ini.WriteBool('Option', ThemeMenu.Items[i].Name, ThemeMenu.Items[i].Checked);
  End;

  ini.WriteBool('Option', Form1.CheckBoxCopyIcon.Name, Form1.CheckBoxCopyIcon.Checked);

  ini.WriteBool('Option', IconSmallMenu.Name, IconSmallMenu.Checked);
  ini.WriteBool('Option', IconBigMenu.Name, IconBigMenu.Checked);

  If Form1.WindowState = wsMaximized Then
    ini.WriteString('Option', 'Window', 'wsMaximized');
  If Form1.WindowState <> wsMaximized Then
  Begin
    ini.WriteString('Option', 'Window', 'wsNormal');
    ini.WriteInteger('Option', 'Top', Form1.top);
    ini.WriteInteger('Option', 'Left', Form1.left);
    ini.WriteInteger('Option', 'Height', Form1.Height);
    ini.WriteInteger('Option', 'Width', Form1.Width);
  End;
  ini.UpdateFile;
  ini.Free;
End;

Procedure TForm1.IconDeleteMenuClick(Sender: TObject);
Begin
  SpeedButton_GeneralMenu.AllowAllUp := true;
  SpeedButton_GeneralMenu.Down := False;
  If EditFolder.Text = '' Then
  Begin
    Application.Messagebox(pchar(LangFolderWithIcon.Caption), pchar(LangError.Caption), mb_iconerror Or mb_ok);
    Exit;
  End;
  Msg := Messagebox(Application.MainForm.Handle, pchar(LangAttention.Caption), pchar(LangAttention.Caption), mb_iconquestion Or mb_yesno);

  If Msg = idyes Then
  Begin

    IconDelete;
    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, Nil, Nil);
    SendMessage(FindWindow('Progman', 'Program Manager'), WM_COMMAND, $A065, 0);

    Form1.ListViewList.Items[0].MakeVisible(true);
    Form1.ListViewList.Items[0].Selected := true;
    Form1.ListViewList.Items[0].Focused := true;
    Form1.ListViewList.SetFocus;
    getIcon;
    Application.ProcessMessages;
    MemoLog.Clear;
    Try
      If fileexists(EditFolder.Text + 'Desktop.ini') Then
      Begin
        MemoLog.Lines.LoadFromFile(EditFolder.Text + 'Desktop.ini');
      End;
    Except
    End;
    Application.ProcessMessages;
    ButtonRefreshClick(Sender);
    Application.ProcessMessages;
    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, Nil, Nil);
  End;
  If Msg = IDNO Then
    Exit;
End;

Procedure TForm1.RefreshIconMenuClick(Sender: TObject);
Begin
  SpeedButton_GeneralMenu.AllowAllUp := true;
  SpeedButton_GeneralMenu.Down := False;
  Form1.ImageFolder.Picture := Nil;
  Form1.ImageIcon.Picture := Nil;
  EditIconFile.Text := '';
  EditFolder.Text := '';
  EditInfoTip.Text := '';
  If Form1.PathFolder.Text = '' Then
    Exit;
  If SysUtils.DirectoryExists(StringReplace(PathFolder.Text, '\\', '\', [rfReplaceAll])) = true Then
  Begin
    ScanAndAddFoldersToListView(StringReplace(PathFolder.Text, '\\', '\', [rfReplaceAll]), ListViewList);
    Application.ProcessMessages;
  End;
End;

Procedure TForm1.AboutMenuClick(Sender: TObject);
Begin
  SpeedButton_GeneralMenu.AllowAllUp := true;
  SpeedButton_GeneralMenu.Down := False;
  Form8.ShowModal;
End;

Procedure TForm1.SaveIconMenuClick(Sender: TObject);
Begin
  If EditIconFile.Text = '' Then
  Begin
    Application.Messagebox(pchar(LangTarget.Caption), pchar(LangError.Caption), mb_iconerror Or mb_ok);
    Exit;
  End;
  If LowerCase(ExtractFileExt(Form1.EditIconFile.Text)) <> '.ico' Then
    Exit;
  SavePictureDialog.FileName := 'Image_' + DateToStr(Now) + FormatDateTime('_hh-mm-ss', Now);
  Application.ProcessMessages;
  If SavePictureDialog.Execute Then
  Begin
    CopyFile(PWideChar(EditIconFile.Text), PWideChar(SavePictureDialog.FileName), false);
    SetFileAttributes(PWideChar(SavePictureDialog.FileName), FILE_ATTRIBUTE_NORMAL);
    FileSetDate(SavePictureDialog.FileName, DateTimeToFileDate(Now));
    Application.ProcessMessages;
  End;
End;

Procedure TForm1.N43Click(Sender: TObject);
Begin
  close;
End;

Procedure TForm1.DeleteIconCacheClick(Sender: TObject);
Begin
  SpeedButton_GeneralMenu.AllowAllUp := true;
  SpeedButton_GeneralMenu.Down := False;
  TabControl1.Enabled := false;
  TabControl2.Enabled := false;
  TabControl6.Enabled := false;
  ClearIconCacheInThread;
  TabControl1.Enabled := true;
  TabControl2.Enabled := true;
  TabControl6.Enabled := true;
  SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, Nil, Nil);
End;

Procedure TForm1.ButtonGetInfoClick(Sender: TObject);
Begin
  If EditFolder.Text = '' Then
    Exit;
  EditInfoTip.Text := ExtractFileName(ExcludeTrailingPathDelimiter(EditFolder.Text));
  Application.ProcessMessages;
  If EditIconFile.Text <> '' Then
  Begin
    ButtonChangeIcon.Enabled := true;
    ButtonChangeIcon.Default := True;
    ButtonChangeIcon.SetFocus;
  End;
End;

Procedure TForm1.ButtonChangeIconClick(Sender: TObject);
Begin
  Form1.StatusBar.Panels[0].Text := '';
  If ListViewList.ItemIndex = -1 Then
    Exit;

  If AnsiUpperCase(ListViewList.Selected.Caption) = AnsiUpperCase('Documents and Settings') Then
    Exit;
  If AnsiUpperCase(ListViewList.Selected.Caption) = AnsiUpperCase('System Volume Information') Then
    Exit;
  If AnsiUpperCase(ListViewList.Selected.Caption) = AnsiUpperCase('$RECYCLE.BIN') Then
    Exit;
  If AnsiUpperCase(ListViewList.Selected.Caption) = AnsiUpperCase('Boot') Then
    Exit;
  If AnsiUpperCase(ListViewList.Selected.Caption) = AnsiUpperCase('Windows') Then
    Exit;
  If ListViewList.Selected.Caption = '' Then
    Exit;
  If EditFolder.Text = '' Then
    Exit;
  If ListViewList.ItemIndex = -1 Then
    Exit;

  If EditFolder.Text = '' Then
  Begin
    Application.Messagebox(pchar(LangSelectFolderIcon.Caption), pchar(LangError.Caption), mb_iconerror Or mb_ok);
    Exit;
  End;
  If EditIconFile.Text = '' Then
  Begin
    Application.Messagebox(pchar(LangTarget.Caption), pchar(LangError.Caption), mb_iconerror Or mb_ok);
    Exit;
  End;
  If EditFolder.Text = ExtractFileDrive(Form1.EditFolder.Text) + '\' Then
  Begin
    Application.Messagebox(pchar(LangCannotChangeIcon.Caption), pchar(LangError.Caption), mb_iconerror Or mb_ok);
    Exit;
  End;
  Msg := Messagebox(Application.MainForm.Handle, pchar(LangChangeIconFolder.Caption), pchar(LangAttention.Caption), mb_iconquestion Or mb_yesno);

  If Msg = idyes Then
  Begin
    IconInstall;
    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, Nil, Nil);
    SendMessage(FindWindow('Progman', 'Program Manager'), WM_COMMAND, $A065, 0);
  End;

  Form1.ListViewList.Items[0].MakeVisible(true);
  Form1.ListViewList.Items[0].Selected := true;
  Form1.ListViewList.Items[0].Focused := true;
  Form1.ListViewList.SetFocus;
  getIcon;
  Application.ProcessMessages;
  MemoLog.Clear;
  Try
    If fileexists(EditFolder.Text + 'Desktop.ini') Then
    Begin
      MemoLog.Lines.LoadFromFile(EditFolder.Text + 'Desktop.ini');
    End;
  Except
  End;
  Application.ProcessMessages;
  ButtonRefreshClick(Sender);
  Application.ProcessMessages;

  If Msg = IDNO Then
    Exit;
End;

End.

