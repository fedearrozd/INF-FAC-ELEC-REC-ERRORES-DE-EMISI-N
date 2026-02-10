LPARAMETERS LPnTiempo

DECLARE INTEGER SetWindowLong IN Win32Api ;
   INTEGER HWND, INTEGER INDEX, INTEGER NewVal

DECLARE INTEGER SetLayeredWindowAttributes IN Win32Api ;
   INTEGER HWND, STRING ColorKey, ;
   INTEGER Opacity, INTEGER Flags

DECLARE Sleep IN WIN32API INTEGER Duration

Sleep(LPnTiempo)
