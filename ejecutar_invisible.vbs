Set WshShell = CreateObject("WScript.Shell")
WshShell.Run chr(34) & "C:\Users\tamar\OneDrive\Escritorio\StockControl\ejecutar_sync.bat" & chr(34), 0, False
Set WshShell = Nothing