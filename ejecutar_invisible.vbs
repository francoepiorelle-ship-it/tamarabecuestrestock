Set WShell = CreateObject("WScript.Shell")
' El 0 al final indica que la ventana se ejecuta de forma oculta (invisible)
WShell.Run "cmd /c ""C:\Users\tamar\OneDrive\Escritorio\StockControl\ejecutar.bat""", 0, False