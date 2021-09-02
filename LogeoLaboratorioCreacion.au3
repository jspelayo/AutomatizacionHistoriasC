main3()
#include <MsgBoxConstants.au3>
#include <Excel.au3>



func CicloExcel3()
	; Ciclo de Excel respecto a cada numero de Cedula
    WinActivate("Microsoft Excel - ValidacionLaboratorioCreacion") ; HABILITAR EXCEL
	Sleep(700) ;tiempo
	Send("{HOME}") ; casilla principal de excel
	Sleep(5000) ;tiempo  
    Send("{HOME}") ; casilla principal de excel
    Sleep(3000) ;tiempo
    Send("{CTRLDOWN}{HOME}{CTRLUP}") ;Ubicacion en la primera (Que se muestra visual)
    ;------------------------------------------------
	send("{TAB 4}")
	Sleep(3000) ;tiempo
	Send("laboratorio Exportado")
	Sleep(3000)
	Send("{ENTER}") 
	
	Send("{CTRLDOWN}{HOME}{CTRLUP}")
	Sleep(1000)
	Send("{CTRLDOWN}{SHIFTDOWN}{RIGHT}{CTRLUP}{SHIFTUP}")
	Sleep(1000)
	Send("{CTRLDOWN}c{CTRLUP}")
	Sleep(1000)
	Send("{CTRLDOWN}{DOWN}{CTRLUP}{DOWN}")
	Sleep(1000)
	Send("{CTRLDOWN}v{CTRLUP}")
	Sleep(3000)
	Send("{CTRLDOWN}{HOME}{CTRLUP}")
	Sleep(3000)	
	
	;------------------------------------------------
	Sleep(3000) ;tiempo
    Send("{SHIFTDOWN}{SPACE}{SHIFTUP}{CTRLDOWN}-{CTRLUP}") ; Señala la casilla completa y la elimina
	Sleep(5000) ;tiempo
	send("{LEFT}"); mover a un lado despues de ser señalada la celda de forma horizontal
	Sleep(5000) ;tiempo
	Send("{CTRLDOWN}g{CTRLUP}") ; Guarda y luego Cierra el Excel
	Sleep(5000) ;tiempo
	Send("{ALTDOWN}{F4}{ALTUP}")
	Sleep(3000)	
	WinActivate("Acceso Centralizado")
	Sleep(6000)
	laboratorio3()
endfunc


func NohayInformacion3()
	; Ciclo de Excel respecto a cada numero de Cedula
    WinActivate("Microsoft Excel - ValidacionLaboratorioCreacion") ; HABILITAR EXCEL
	Sleep(700) ;tiempo
	WinActivate("Microsoft Excel - ValidacionLaboratorioCreacion")
	Sleep(700) ;tiempo
	Send("{HOME}") ; casilla principal de excel
	Sleep(5000) ;tiempo  
    Send("{HOME}") ; casilla principal de excel
    Sleep(3000) ;tiempo
    Send("{CTRLDOWN}{HOME}{CTRLUP}") ;Ubicacion en la primera (Que se muestra visual)
    ;------------------------------------------------
	send("{TAB 4}")
	Sleep(3000) ;tiempo
	Send("No hay Informacion")
	Sleep(3000)
	Send("{ENTER}") 
	
	Send("{CTRLDOWN}{HOME}{CTRLUP}")
	Sleep(1000)
	Send("{CTRLDOWN}{SHIFTDOWN}{RIGHT}{CTRLUP}{SHIFTUP}")
	Sleep(1000)
	Send("{CTRLDOWN}c{CTRLUP}")
	Sleep(1000)
	Send("{CTRLDOWN}{DOWN}{CTRLUP}{DOWN}")
	Sleep(1000)
	Send("{CTRLDOWN}v{CTRLUP}")
	Sleep(3000)
	Send("{CTRLDOWN}{HOME}{CTRLUP}")
	Sleep(3000)	
	
	;------------------------------------------------
	Sleep(3000) ;tiempo
    Send("{SHIFTDOWN}{SPACE}{SHIFTUP}{CTRLDOWN}-{CTRLUP}") ; Señala la casilla completa y la elimina
	Sleep(5000) ;tiempo
	send("{LEFT}"); mover a un lado despues de ser señalada la celda de forma horizontal
	Sleep(5000) ;tiempo
	Send("{CTRLDOWN}g{CTRLUP}") ; Guarda y luego Cierra el Excel
	Sleep(5000) ;tiempo
	Send("{ALTDOWN}{F4}{ALTUP}")
	Sleep(3000)	
	WinActivate("Acceso Centralizado")
	Sleep(6000)
	laboratorio3()
endfunc

func main3()


Run("C:\Servinte\Acceso\Servinte.AccesCentralized.exe")

Sleep(6000)
MouseClick("right",678,477,3)
Send("{ENTER}")
Send("jeffpeli")
send("{TAB}")
send("1052403207")
Send("{ENTER}}")
Sleep(6000)
Send("Laboratorio")
Sleep(6000)
laboratorio3()
endfunc



Func Correo3()
    ;Codigo para enviar Correo
	
Run("C:\Program Files\Google\Chrome\Application\chrome.exe"); Ruta
    Sleep(9000) ;tiempo
    WinWaitActive("Google Chrome","",@SW_MAXIMIZE) ;Abrir enlace en pantalla completa
	Sleep(9000) ;tiempo
	;Send("{F6}"); habilita el buscador
	Sleep(9000);tiempo
    Send("https://outlook.office.com/mail/deeplink/compose?popoutv2=1&version=20210823004.06", 1) ; link que se envia
	Sleep(9000) ;tiempo
    Send("{ENTER}") 
    
	
	Sleep(10000) ;tiempo
	Send("{ENTER}")
	
	send("LIZARAZO333@GMAIL.COM") ; Correo al que se envia
	Sleep(6000) ;tiempo
	Send("{ENTER}")
	Sleep(1000) ;tiempo
	send("{TAB 4}") ; Asunto
	Sleep(6000) ;tiempo
	send("Laboratorios Generados")
	Sleep(6000) ;tiempo
	send("{TAB}") ; Mensaje
	Sleep(6000) ;tiempo
	send("SE HA REALIZADO EXITOSAMENTE, LA GENERACION DE LABORATORIOS GENERADOS")
	Sleep(6000) ;tiempo
	send("{TAB 2}") ;ENVIAR CORREO
	Sleep(6000) ;tiempo
	Send ("{ENTER}")	  

endfunc

func laboratorio3()
MouseClick("left",191,230,1)
Sleep(3000)
send("{DOWN 3}}")
Sleep(3000)
send("{RIGHT}")
Sleep(3000)
send("{DOWN}}")
Sleep(1000)
send("{RIGHT}")
Sleep(1000)
send("{DOWN}}")
Sleep(1000)
send("{ENTER}")
Sleep(6000)

send("{TAB 21}")
	Sleep(2000)
	send("{RIGHT}")
	Sleep(2000)
	send("{TAB}")

Global $oExcel = _Excel_Open()
Global $oWorkbook = _Excel_BookOpen($oExcel, "C:\Users\jpelayo\Documents\HistoriasClinicas\ValidacionLaboratorioCreacion.xlsx" )



Sleep(5000)
; Toma la casilla respectiva para copiar y guarddar en una variable, para despues pegar
Local $sResult = _Excel_RangeRead($oWorkbook, Default, "B2")

ClipPut($sResult)
$sData = ClipGet()
    Sleep(1000)
    WinActivate(" SERVINTE CLINICAL SUITE - LABORATORIO Versión 6.0 - HOSPITAL MANUEL URIBE ANGEL - Consulta de Resultados por Cargo  -  6.0.1  -  (rlaconres)")
	
	Sleep(3000)
	Send("^v")
	
	Sleep(1000)
	Send ("{ENTER}")
	Sleep(4000)
    if WinActivate("Consulta de Resultados por Cargo  -  6.0.1  -  (rlaconres)") then	
    send("{SPACE}") 
	Sleep(6000)
	Send("{ALTDOWN}{F4}{ALTUP}")
	NohayInformacion3()
	EndIf
MouseClick("left",91,42,1)
send("{RIGHT 2}")
send("{TAB}")
send("{SPACE}")
Sleep(10000)
Sleep(3000)
if WinActivate("Reporte de Resultados 6.0.1") then	
send("{SPACE}") 
EndIf
WinActivate("SERVINTE CLINICAL SUITE - LABORATORIO Versión 6.0 - HOSPITAL MANUEL URIBE ANGEL - Reporte de Resultados 6.0.1")
Sleep(6000);CTRL+SHIFT+A
if WinActivate("Consulta de Resultados") then
send("{SPACE}") 
EndIf
if WinActivate("Validar Impresión") then	
send("{SPACE}") 
EndIf
Sleep(3000)
MouseClick("left",209,43,1)
WinActivate("Opciones de Impresion")
Sleep(3000)
Send("{SHIFTDOWN}{TAB 2}{SHIFTUP}")
Sleep(3000)
Send ("{ENTER}")
Sleep(3000)
send("{DOWN}")
Sleep(3000)
Send ("{ENTER}")
Sleep(3000)
send("{TAB}")
Sleep(3000)
Send ("{ENTER}")
Sleep(30000)
WinWaitActive("PDF24 Assistant")
MouseClick("left",572,279,1)
;-----------------------------------------------
    ;Sleep(3000)
	;send("{TAB 5}") ; Se abre un cuadro de Dialogo para ubicarse en el boton 
	;Sleep(3000)
	;send ("{SPACE}") ; Aceptar
	;Sleep(6000)	; Tiempo
;-----------------------------------------------
Local $sResult2 = _Excel_RangeRead($oWorkbook, Default, "C2")	
    Local $sData2 = ClipGet() ; Guarda Variable
    ClipPut($sResult2)
    $sData2 = ClipGet() ; Se instancia
    WinActivate("Guardar un archivo PDF") ; Activa la ventana
	Sleep(8000) ; Tiempo
	send("{TAB 5}") ; Se ubica en la parte de ruta en el buscador
	Sleep(380); Tiempo
	Send ("{ENTER}") ; Enter
	Sleep(1000) ; Tiempo
    Send("^v") ; Pega la ruta desde el excel (Parametrizado)
	Send ("{ENTER}") ; Enter
;-------------------------------------------------
;Sleep(90000) ; Tiempo
;-----------------------------------------------------
Local $sResult5 = _Excel_RangeRead($oWorkbook, Default, "D2")	
    Local $sData5 = ClipGet() ; Guarda la variable
    ClipPut($sResult5) 
    $sData5 = ClipGet() ; Se instancia
    WinActivate("Guardar un archivo PDF") ; Se habilita la ventana
	Sleep(1000) ; Tiempo
    send("{TAB 2}") ; Ubicacion donde esta el boton de crear carpeta
	send("{RIGHT}") ; Flecha hacia el boton de crear carpeta
	Sleep(500) ; Tiempo
	send ("{SPACE}") ; Enter
	Sleep(1000) ; Tiempo
	Send("^v") ; Se pega el valor desde el excel, como se debe guardar
	Sleep(1000) ; Tiempo
	Send ("{ENTER}") ; Se crea la carpeta con el nombre respectivo
	Sleep(1000) ; Tiempo
	
	
	If WinActivate("Confirmar el reemplazo de carpetas") then
	Sleep(3000)
    send("{RIGHT}")
	Sleep(3000)
	send ("{SPACE}") ; Enter
	Sleep(3000) ; tiempo	
	send("{TAB 6}")
	Sleep(3000)
	Send ("{ENTER}")
	Sleep(6000)	
	Send("{ALTDOWN}{F4}{ALTUP}")
	Sleep(3000)  
	WinActivate("PDF24 Assistant")
	Sleep(3000) 	
	Send("{ALTDOWN}{F4}{ALTUP}")
	WinActivate("SERVINTE CLINICAL SUITE - LABORATORIO Versión 6.0 - HOSPITAL MANUEL URIBE ANGEL - [Reporte de Resultados 6.0.1]")
	Send("{ALTDOWN}{F4}{ALTUP}")
	Sleep(3000)
	WinActivate("Acceso Centralizado")
	Sleep(3000)
	MouseClick("left",1232,85,1)
	Sleep(3000)
	MouseClick("left",1238,219,1)
	Sleep(3000)
	Correo3() ; Hbailita la funcion de Correo para enviar el mensaje de finalizado
	Exit	
	EndIf
	
	
	if WinActivate("Confirmar Guardar como") then	
    Sleep(3000)
	send("{SPACE}") 
	Sleep(6000)	
	Send("{ALTDOWN}{F4}{ALTUP}")
	Sleep(3000)  
	WinActivate("PDF24 Assistant")
	Sleep(3000) 	
	Send("{ALTDOWN}{F4}{ALTUP}")
	WinActivate("SERVINTE CLINICAL SUITE - LABORATORIO Versión 6.0 - HOSPITAL MANUEL URIBE ANGEL - [Reporte de Resultados 6.0.1]")
	Send("{ALTDOWN}{F4}{ALTUP}")
	Sleep(3000)
	WinActivate("Acceso Centralizado")
	Sleep(3000)
	MouseClick("left",1232,85,1)
	Sleep(3000)
	MouseClick("left",1238,219,1)
	Sleep(3000)
	Correo3()
	exit
	EndIf
	Send ("{ENTER}") ; Se crea la carpeta con el nombre respectivo
;------------------------------------------------------
   Local $sResult5 = _Excel_RangeRead($oWorkbook, Default, "E2")	
    Local $sData5 = ClipGet() ; Guarda Variable
    ClipPut($sResult5)
    $sData5 = ClipGet() ; Se instancia
    WinActivate("Guardar un archivo PDF") ;Se habilita la ventana
	Sleep(1000) ; Tiempo
    send("{TAB 2}") ; Se ubica en la parte de la pestaña donde se escribe el nombre de como se debe guardar.
	Sleep(1000) ; Tiempo									
	Send("^v") ; Se pega desde el excel
	Sleep(2000) ; Tiempo
	Send ("{ENTER}")
	Sleep(4000)   
;-------------------------------------------------------------	
	
	WinActivate(" SERVINTE CLINICAL SUITE - LABORATORIO Versión 6.0 - HOSPITAL MANUEL URIBE ANGEL - Reporte de Resultados 6.0.1")
    Sleep(3000)
	Send("{ALTDOWN}{F4}{ALTUP}") ;Se cierra la ventana
	Sleep(6000) ; Tiempo
	if WinActivate("PowerBuilder application execution error (R0002)") then	
    send("{SPACE}") 
	Sleep(6000)
	EndIf
	CicloExcel3()
	;WinActivate("Acceso Centralizado")


endfunc
