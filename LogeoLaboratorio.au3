principalmain()
#include <MsgBoxConstants.au3>
#include <Excel.au3>



func CicloExcel2()
	; Ciclo de Excel respecto a cada numero de Cedula
    WinActivate("Microsoft Excel - ValidacionLaboratorio") ; HABILITAR EXCEL
	Sleep(700) ;tiempo
	Send("{HOME}") ; casilla principal de excel
	Sleep(5000) ;tiempo  
    Send("{HOME}") ; casilla principal de excel
    Sleep(3000) ;tiempo
    Send("{CTRLDOWN}{HOME}{CTRLUP}") ;Ubicacion en la primera (Que se muestra visual)
    ;------------------------------------------------
	send("{TAB 3}")
	Sleep(3000) ;tiempo
	Send("laboratorio Exportado") ; se envia el exportado a la casilla excel
	Sleep(3000)
	Send("{ENTER}") ; enter
	
	Send("{CTRLDOWN}{HOME}{CTRLUP}") ;casilla principal
	Sleep(1000)
	Send("{CTRLDOWN}{SHIFTDOWN}{RIGHT}{CTRLUP}{SHIFTUP}") ;selecciona la fila consultada 
	Sleep(1000)
	Send("{CTRLDOWN}c{CTRLUP}"); copia la fila excel
	Sleep(1000)
	Send("{CTRLDOWN}{DOWN}{CTRLUP}{DOWN}") ; se pasa la ultima fila, un cuadro mas abajo
	Sleep(1000)
	Send("{CTRLDOWN}v{CTRLUP}"); se pega
	Sleep(3000)
	Send("{CTRLDOWN}{HOME}{CTRLUP}"); regresa a la casilla principal
	Sleep(3000)	
	
	;------------------------------------------------
	Sleep(3000) ;tiempo
    Send("{SHIFTDOWN}{SPACE}{SHIFTUP}{CTRLDOWN}-{CTRLUP}") ; Señala la casilla completa y la elimina
	Sleep(5000) ;tiempo
	send("{LEFT}"); mover a un lado despues de ser señalada la celda de forma horizontal
	Sleep(5000) ;tiempo
	Send("{CTRLDOWN}g{CTRLUP}") ; Guarda y luego Cierra el Excel
	Sleep(5000) ;tiempo
	Send("{ALTDOWN}{F4}{ALTUP}") ; se guarda y cierra el excel
	Sleep(3000)	
	WinActivate("Acceso Centralizado"); se activa la ventana de acceso centralizado
	Sleep(6000)
	laboratorio() ; retorna al laboratorio
endfunc


func NohayInformacion2()
	; Ciclo de Excel respecto a cada numero de Cedula
    WinActivate("Microsoft Excel - ValidacionLaboratorio") ; HABILITAR EXCEL
	Sleep(700) ;tiempo
	WinActivate("Microsoft Excel - ValidacionLaboratorio")
	Sleep(700) ;tiempo
	Send("{HOME}") ; casilla principal de excel
	Sleep(5000) ;tiempo  
    Send("{HOME}") ; casilla principal de excel
    Sleep(3000) ;tiempo
    Send("{CTRLDOWN}{HOME}{CTRLUP}") ;Ubicacion en la primera (Que se muestra visual)
    ;------------------------------------------------
	send("{TAB 3}") ; se ubica en la casilla 
	Sleep(3000) ;tiempo
	Send("No hay Informacion"); se envia ese parametro
	Sleep(3000)
	Send("{ENTER}") ; enter
	
	Send("{CTRLDOWN}{HOME}{CTRLUP}") ; regresa al principal
	Sleep(1000)
	Send("{CTRLDOWN}{SHIFTDOWN}{RIGHT}{CTRLUP}{SHIFTUP}") ; selecciona la fila
	Sleep(1000)
	Send("{CTRLDOWN}c{CTRLUP}"); copia la fila
	Sleep(1000)
	Send("{CTRLDOWN}{DOWN}{CTRLUP}{DOWN}") ; pasa a la ultima casilla
	Sleep(1000)
	Send("{CTRLDOWN}v{CTRLUP}") ; pega la fila seleccionada
	Sleep(3000)
	Send("{CTRLDOWN}{HOME}{CTRLUP}"); regresa a la casilla principal
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
	laboratorio()
endfunc

func principalmain() ; funcion principal


Run("C:\Servinte\Acceso\Servinte.AccesCentralized.exe") ; se ejecuta servinte

Sleep(6000)
MouseClick("right",678,477,3) ;click en servinte
Send("{ENTER}"); enter
Send("jeffpeli") ; usuario
send("{TAB}") ; pasa a la siguiente casilla
send("1052403207") ; contraseña
Send("{ENTER}}") ; enter
Sleep(6000)
Send("Laboratorio") ; una vez abierta la aplicacion se esccribe laboratorio
Sleep(6000)
laboratorio() ; pasa a la funcion de laboratotio
endfunc



Func Correo2()
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
	send("Laboratorio Generado")
	Sleep(6000) ;tiempo
	send("{TAB}") ; Mensaje
	Sleep(6000) ;tiempo
	send("SE HA REALIZADO EXITOSAMENTE, LA GENERACION DE LABORATORIOS")
	Sleep(6000) ;tiempo
	send("{TAB 2}") ;ENVIAR CORREO
	Sleep(6000) ;tiempo
	Send ("{ENTER}")	   

endfunc

func laboratorio()
MouseClick("left",191,230,1) ; click en el boton laboratorio
Sleep(3000)
send("{DOWN 3}}") ; baja 3 veces hasta llegar a consultas
Sleep(3000)
send("{RIGHT}") ; derecha (abrir pestaña)
Sleep(3000)
send("{DOWN}}") ; abajo (consuta de resultados)
Sleep(1000)
send("{RIGHT}") ; derecha (abre pestaña)
Sleep(1000)
send("{DOWN}}"); abajo por cargos
Sleep(1000)
send("{ENTER}"); enter
Sleep(6000)

send("{TAB 21}") ; tabular hasta llegar a paciente
	Sleep(2000)
	send("{RIGHT}"); boton identificacion
	Sleep(2000)
	send("{TAB}"); identificacion

Global $oExcel = _Excel_Open()
Global $oWorkbook = _Excel_BookOpen($oExcel, "C:\Users\jpelayo\Documents\HistoriasClinicas\ValidacionLaboratorio.xlsx" ); abre excel

Sleep(5000)
; Toma la casilla respectiva para copiar y guarddar en una variable, para despues pegar
Local $sResult = _Excel_RangeRead($oWorkbook, Default, "B2") ; selecciona la casilla B2 la cual tiene el numero de identificacion

ClipPut($sResult) ;se envia la variable
$sData = ClipGet() ; se obtiene la variable
    Sleep(1000)
    WinActivate(" SERVINTE CLINICAL SUITE - LABORATORIO Versión 6.0 - HOSPITAL MANUEL URIBE ANGEL - Consulta de Resultados por Cargo  -  6.0.1  -  (rlaconres)")
	
	Sleep(3000)
	Send("^v") ; se pega la identificacion
	
	Sleep(1000)
	Send ("{ENTER}");enter para generar historia
	Sleep(4000)
    if WinActivate("Consulta de Resultados por Cargo  -  6.0.1  -  (rlaconres)") then	; si se activa esta ventana dar enter para continuar
    send("{SPACE}") 
	Sleep(6000)
	Send("{ALTDOWN}{F4}{ALTUP}"); ; se cierra
	NohayInformacion2() ; pasa la funcion de no hay informacion para que continue con el siguiente en la lista
	EndIf
MouseClick("left",91,42,1) ;click en el boton de impresion  
send("{RIGHT 2}") ;imprimir toda la consulta
send("{TAB}")
send("{SPACE}") ; enter
Sleep(10000)
Sleep(3000)
if WinActivate("Reporte de Resultados 6.0.1") then	; si se activa esta ventana dar enter para continuar
send("{SPACE}") ; enter para continuar
EndIf
WinActivate("SERVINTE CLINICAL SUITE - LABORATORIO Versión 6.0 - HOSPITAL MANUEL URIBE ANGEL - Reporte de Resultados 6.0.1") ; si se activa esta ventana dar enter para continuar
Sleep(6000);CTRL+SHIFT+A
if WinActivate("Consulta de Resultados") then ; si se activa esta ventana dar enter para continuar
send("{SPACE}") ; enter
EndIf
if WinActivate("Validar Impresión") then ; si se activa esta ventana dar enter para continuar	
send("{SPACE}") ; enter
EndIf
Sleep(3000)
MouseClick("left",209,43,1) ; click en imprimir
WinActivate("Opciones de Impresion") ;se activa ventana
Sleep(3000)
Send("{SHIFTDOWN}{TAB 2}{SHIFTUP}") ; ubicacion para llegara impresoras
Sleep(3000)
Send ("{ENTER}"); enter
Sleep(3000)
send("{DOWN}") ; Selecciona PDF24 
Sleep(3000)
Send ("{ENTER}") ; Enter
Sleep(3000)
send("{TAB}") ; TABULACION OK
Sleep(3000)
Send ("{ENTER}") ; Enter
Sleep(30000)
WinWaitActive("PDF24 Assistant")
MouseClick("left",572,279,1) ; Click en guardar PDF24 
;-----------------------------------------------
    send("{TAB 5}") ; Se abre un cuadro de Dialogo para ubicarse en el boton 
	send ("{SPACE}") ; Aceptar
	Sleep(6000)	; Tiempo
;-----------------------------------------------
Local $sResult2 = _Excel_RangeRead($oWorkbook, Default, "C2")	
    Local $sData2 = ClipGet() ; Guarda Variable
    ClipPut($sResult2)
    $sData2 = ClipGet() ; Se instancia
    WinActivate("Guardar un archivo PDF") ; Activa la ventana
	Sleep(5000) ; Tiempo
	send("{TAB 5}") ; Se ubica en la parte de ruta en el buscador
	Sleep(380); Tiempo
	Send ("{ENTER}") ; Enter
	Sleep(1000) ; Tiempo
    Send("^v") ; Pega la ruta desde el excel (Parametrizado)
	Send ("{ENTER}") ; Enter
;-------------------------------------------------
	Sleep(3000)
	send("{TAB}")
Local $sResult3 = _Excel_RangeRead($oWorkbook, Default, "B2")
Local $sData3 = ClipGet()
ClipPut($sResult3)
$sData3 = ClipGet()
   
	Sleep(3000)
	Send("^v") ;Se pega la cedula en caso de que esta ya exista
	Sleep(1000)
	Send ("{ENTER}"); buscar
	Sleep(3000)
	send("{DOWN}") ; baja  para seleccionar la carpeta ya creada de la historia clinica
	Sleep(3000)
	send("{SPACE}") ; entra la carpeta
	Sleep(3000)
	Send ("{ENTER}"); enter
	
;-----------------------------------------------------
Local $sResult4 = _Excel_RangeRead($oWorkbook, Default, "D2")	
    Local $sData4 = ClipGet() ; Guarda Variable
    ClipPut($sResult4)
    $sData4 = ClipGet() ; Se instancia
    WinActivate("Guardar como") ;Se habilita la ventana
	Sleep(1000) ; Tiempo
    send("{TAB 2}") ; Se ubica en la parte de la pestaña donde se escribe el nombre de como se debe guardar.
	Sleep(1000) ; Tiempo									
	Send("^v") ; Se pega desde el excel
	Sleep(2000) ; Tiempo
	Send ("{ENTER}")
	Sleep(6000)
	
	if WinActivate("Confirmar Guardar como") then	; si se activa esta ventana dar enter para continuar y de esta manera terminar el cilclo
    Sleep(3000)
	send("{SPACE}") ; entrar
	Sleep(6000)	
	Send("{ALTDOWN}{F4}{ALTUP}") ; cierra la ventana
	Sleep(3000)  
	WinActivate("PDF24 Assistant") ; si se activa esta ventana se cierra 
	Sleep(3000) 	
	Send("{ALTDOWN}{F4}{ALTUP}") ; cerrar ventana PDF24
	WinActivate("SERVINTE CLINICAL SUITE - LABORATORIO Versión 6.0 - HOSPITAL MANUEL URIBE ANGEL - [Reporte de Resultados 6.0.1]")
	Send("{ALTDOWN}{F4}{ALTUP}"); CERRAR VENTANA 
	Sleep(3000)
	if WinActivate("PowerBuilder application execution error (R0002)") then	; si se activa esta ventana de error dar enter para continuar
    send("{SPACE}") ; enter 
	Sleep(6000)
	EndIf
	WinActivate("Acceso Centralizado")
	Sleep(3000)
	MouseClick("left",1232,85,1)
	Sleep(3000)
	MouseClick("left",1238,219,1)
	Sleep(3000)
	Correo2() ; Enviar correo
	exit
	EndIf
;------------------------------------------------------
WinActivate(" SERVINTE CLINICAL SUITE - LABORATORIO Versión 6.0 - HOSPITAL MANUEL URIBE ANGEL - Reporte de Resultados 6.0.1")
    Sleep(3000) 
	Send("{ALTDOWN}{F4}{ALTUP}") ;Se cierra la ventana
	Sleep(6000) ; Tiempo
	if WinActivate("PowerBuilder application execution error (R0002)") then	; si se activa esta ventana de error dar enter para continuar
    send("{SPACE}") ; enter 
	Sleep(6000)
	EndIf
	CicloExcel2() ; recorre el ciclo
	;WinActivate("Acceso Centralizado")


endfunc