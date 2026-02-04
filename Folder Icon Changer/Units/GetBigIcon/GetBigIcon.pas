unit GetBigIcon;

interface

uses ShellApi, Commctrl, ShlObj, Windows, Sysutils, Graphics;

const
  SHIL_LARGE = $00;
  SHIL_SMALL = $01;
  SHIL_EXTRALARGE = $02;
  SHIL_SYSSMALL = $03;
  SHIL_JUMBO = $04;
  IID_IImageList: TGUID = '{46EB5926-582E-4017-9FDF-E8998DAA0950}';

Procedure GetIconFromFile(aFile: String; var aIcon: TIcon; SHIL_FLAG: Cardinal); overload;


implementation

function GetImageListSH(SHIL_FLAG: Cardinal): HIMAGELIST;
type
  _SHGetImageList = function(iImageList: integer; const riid: TGUID;
    var ppv: Pointer): hResult; stdcall;
var
  Handle: THandle;
  SHGetImageList: _SHGetImageList;
begin
  Result := 0;
  Handle := LoadLibrary('Shell32.dll');
  if Handle <> 0 then
    try
      SHGetImageList := GetProcAddress(Handle, PChar(727));
      if Assigned(SHGetImageList) and (Win32Platform = VER_PLATFORM_WIN32_NT) then
        SHGetImageList(SHIL_FLAG, IID_IImageList, Pointer(Result));
    finally
      FreeLibrary(Handle);
    end;
end;

// Оригинальная процедура (без изменений)
Procedure GetIconFromFile(aFile: String; var aIcon: TIcon; SHIL_FLAG: Cardinal);
var
  aImgList: HIMAGELIST;
  SFI: TSHFileInfo;
  aIndex: integer;
Begin
  FillChar(SFI, Sizeof(SFI), #0);

  // Получаем информацию об иконке (работает и для файлов, и для папок)
  SHGetFileInfo(PChar(aFile), 0, SFI, Sizeof(TSHFileInfo),
    SHGFI_ICON or SHGFI_LARGEICON or SHGFI_SHELLICONSIZE or
    SHGFI_SYSICONINDEX or SHGFI_TYPENAME or SHGFI_DISPLAYNAME);

  // Освобождаем иконку, полученную SHGetFileInfo, так как будем использовать другую
  if SFI.hIcon <> 0 then
    DestroyIcon(SFI.hIcon);

  if not Assigned(aIcon) then
    aIcon := TIcon.Create
  else
    aIcon.Handle := 0;  // Освобождаем предыдущую иконку

  aImgList := GetImageListSH(SHIL_FLAG);
  if aImgList <> 0 then
  begin
    aIndex := SFI.iIcon;
    aIcon.Handle := ImageList_GetIcon(aImgList, aIndex, ILD_NORMAL);
  end;

  // Если не удалось получить ImageList, используем резервный вариант
  if (aIcon.Handle = 0) then
  begin
    SHGetFileInfo(PChar(aFile), 0, SFI, Sizeof(TSHFileInfo),
      SHGFI_ICON or SHGFI_LARGEICON);
    aIcon.Handle := SFI.hIcon;
  end;
End;


end.
