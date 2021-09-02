
HCmain()
#include <MsgBoxConstants.au3>
#include <Excel.au3>

; Funcion de Win activate para el instanciamiento de cualquier Ventana
Func _Au3RecordSetup()
Opt('WinWaitDelay',100)
Opt('WinDetectHiddenText',1)
Opt('MouseCoordMode',0)
Local $aResult = DllCall('User32.dll', 'int', 'GetKeyboardLayoutNameW', 'wstr', '')
If $aResult[1] <> '0000080A' Then
  MsgBox(64, 'Warning', 'Recording has been done under a different Keyboard layout' & @CRLF & '(0000080A->' & $aResult[1] & ')')
EndIf

EndFunc

Func _WinWaitActivate($title,$text,$timeout=0)
	WinWait($title,$text,$timeout)
	If Not WinActive($title,$text) Then WinActivate($title,$text)
	WinWaitActive($title,$text,$timeout)
EndFunc

func HCmain() ; funcion principal


Run("C:\Servinte\Acceso\Servinte.AccesCentralized.exe") ; se ejecuta servinte

Sleep(6000)
MouseClick("right",678,477,3) ;click en servinte
Send("{ENTER}"); enter
Send("jeffpeli") ; usuario
send("{TAB}") ; pasa a la siguiente casilla
send("1052403207") ; contraseña
Send("{ENTER}}") ; enter
Sleep(7000)
Send("Tablero de hospitalización")
Sleep(150)
MouseClick("left",191,230,3)
Sleep(300)
send("{TAB}")
Sleep(390)
MouseClick("left",592,411,3)
Sleep(5000)
main() ; pasa a la funcion de laboratotio
endfunc


func CicloExcel()
	; Ciclo de Excel respecto a cada numero de Cedula
    WinActivate("Microsoft Excel - robotvalidacion2") ; HABILITAR EXCEL
	Sleep(700) ;tiempo
	Send("{HOME}") ; casilla principal de excel
	Sleep(5000) ;tiempo  
    Send("{HOME}") ; casilla principal de excel
    Sleep(3000) ;tiempo
    Send("{CTRLDOWN}{HOME}{CTRLUP}") ;Ubicacion en la primera (Que se muestra visual)
    ;------------------------------------------------
	send("{TAB 6}")
	Sleep(3000) ;tiempo
	Send("Exportado")
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
	WinActivate("SERVINTE CLINICAL SUITE  - HOSPITAL MANUEL URIBE ANGEL -  Tablero de Pacientes de HCE 4.3.5 (chiepiact)")
	Sleep(3000)
	main()
	
endfunc


func Mensaje961()

    ; Ciclo de Excel respecto a cada numero de Cedula
    WinActivate("Microsoft Excel - robotvalidacion2") ; HABILITAR EXCEL
	Sleep(700) ;tiempo
	Send("{HOME}") ; casilla principal de excel
	Sleep(5000) ;tiempo  
    Send("{HOME}") ; casilla principal de excel
    Sleep(3000) ;tiempo
    Send("{CTRLDOWN}{HOME}{CTRLUP}") ;Ubicacion en la primera (Que se muestra visual)
    ;------------------------------------------------
	send("{TAB 6}")
	Sleep(3000) ;tiempo
	Send("No hay informacion")
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
	WinActivate("SERVINTE CLINICAL SUITE  - HOSPITAL MANUEL URIBE ANGEL -  Tablero de Pacientes de HCE 4.3.5 (chiepiact)")
	Sleep(3000)
	;main()	
endfunc


Func Correo()
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
	send("Historias Generadas")
	Sleep(6000) ;tiempo
	send("{TAB}") ; Mensaje
	Sleep(6000) ;tiempo
	send("SE HA REALIZADO EXITOSAMENTE, LA GENERACION DE HISTTORIAS CLINICAS")
	Sleep(6000) ;tiempo
	send("{TAB 2}") ;ENVIAR CORREO
	Sleep(6000) ;tiempo
	Send ("{ENTER}")	  

endfunc


func main()

;LogeoTablero() ;Instancia de logeo

WinActivate("SERVINTE CLINICAL SUITE  - HOSPITAL MANUEL URIBE ANGEL -  Tablero de Pacientes de HCE 4.3.5 (chiepiact)")
Sleep(6000)
MouseClick("left",230,59,1) ; Click Gestión de HISTORIAS
Sleep(4000) ; TIempo
MouseClick("left",1015,344,1) ; Click Botom Limpiar

Sleep(1000)
send("{TAB 9}") ; Ubicacion de casilla por busqueda por cedula
Sleep(4000)

; Abrir Excel

Global $oExcel = _Excel_Open()
Global $oWorkbook = _Excel_BookOpen($oExcel, "C:\Users\jpelayo\Documents\HistoriasClinicas\robotvalidacion2.xlsx" )
Sleep(5000)

;WinActivate("Gestión de historias 4.3.3 (chigeshis)")	; Abrir la ventana de Gestion de historias

; Toma la casilla respectiva para copiar y guarddar en una variable, para despues pegar
Local $sResult = _Excel_RangeRead($oWorkbook, Default, "B2")

Local $sData = ClipGet() ; Cluck Obtiene la variable 

;If $sData <> '' Then
ClipPut($sResult)
$sData = ClipGet() ; Envia La variable obtenida
    WinActivate("Gestión de historias 4.3.3 (chigeshis)")	; Abrir la ventana de Gestion de historias
	Sleep(6000) ; tiempo
	MouseClick("left",1007,342,1) ; Limpiar
	Sleep(3000)
    send("{TAB 9}") ; Ubicacion de Cedula
	Sleep(1000)
    Send("^v") ; Pega la variable
	
	
    send("{TAB 6}")	; buscar
	Send ("{ENTER}")
    send ("{SPACE}")  
	Sleep(5000)
	
	; Condicional en caso de no encontrar La cedula en el momento
	if WinActivate("Mensaje 961") then              
	Sleep(1000)
	send ("{SPACE}") ; Darle enter
	Sleep(1000)
	Mensaje961() ; Instancia el Ciclo de Excel para que pase a enviar la otra cedula
	Sleep(3000)
	WinActivate("Gestión de historias 4.3.3 (chigeshis)")
	Sleep(3000)
    Send("{ALTDOWN}{F4}{ALTUP}")
	Sleep(3000)
	return main() ; Retorna para que tome la siguiente Cedula en la lista
	EndIf
	; Condicional En Caso de pasar una casilla en blanco, que en general sera cuando finalice el recorrido.
	If WinActivate("Mensaje 1751") then
	Sleep(2000) ; tiempo
	send ("{SPACE}") ; Enter
	Sleep(1000) ; tiempo	
	;Correo() ; Hbailita la funcion de Correo para enviar el mensaje de finalizado
	Exit	
	EndIf
	; En caso de que no cumpla ninguna de esas condiciones  se siguel el proceso de exportacion Historias
	Sleep(3000)
	MouseClick("left",695,463,1) ; Darle clip a la primer historia por fecha para generar 
	Sleep(3000)
	MouseClick("left",695,463,2) ; Misma Coordenada
	Sleep(9000) ; tiempo
    Sleep(9000)
	Sleep(3000)
	
    ; Envio de Fecha de inicio
Local $sResult2 = _Excel_RangeRead($oWorkbook, Default, "C2")	
Local $sData2 = ClipGet() ; Guarda la variable
ClipPut($sResult2) ; envia
$sData2 = ClipGet() ; instancia
    WinActivate("SERVINTE CLINICAL SUITE - HCECLI 1.3 -  HOSPITAL MANUEL URIBE ANGEL") ; Activar pantalla
	Sleep(6000) ; tiempo
	Send("{SHIFTDOWN}{TAB 9}{SHIFTUP}")
	Sleep(2000)
	send("{UP 8}")
	Sleep(2000)
	Send("{SHIFTDOWN}{TAB}{SHIFTUP}")
	
    Send("^v") ; pega la fecha
	Sleep(5000) ; tiempo
	; Envio de Fecha Final
Local $sResult3 = _Excel_RangeRead($oWorkbook, Default, "D2")	
Local $sData3 = ClipGet() ; Guarda la variable
ClipPut($sResult3) ; envia
$sData3 = ClipGet() ; instancia
    WinActivate("SERVINTE CLINICAL SUITE - HCECLI 1.3 -  HOSPITAL MANUEL URIBE ANGEL") ; Hailita pantalla
	Sleep(6000) ; Tiempo
	send("{TAB}") ; Se pasa a la otra casilla donde se ubica la fecha final
	Sleep(6000) ; Tiempo
    Send("^v") ; Pega la fecha correspondiente
	Sleep(6000) ; Tiempo
	
	send("{TAB}")	; Se Ubica boton Buscar
	Sleep(3000)	; Tiempo
	send("{TAB}")
	Sleep(3000)
	send("{TAB}")
	Sleep(3000)
	send("{TAB}")
	Sleep(3000)
	send ("{SPACE}") ; Da click
	Sleep(9000)
	Send("{SHIFTDOWN}{TAB}{SHIFTUP}")
	Sleep(3000)
	send("{UP}")
	Sleep(3000)
	send ("{SPACE}") ; Da click
	
	Sleep(3000) ; Tiempo
	
	;Seccion de Generacion de Historias por pasiente
	send("{TAB 3}") ; IMPORTANTE abajo para seleccionar los parametros del paciente
	send ("{SPACE}") ; Selecciona todo
	Sleep(400) ; tiempo
	send ("{SPACE}") ; ; quita seleccion
	send("{DOWN 2}")  ; baja dos veces
	send ("{SPACE}")  ; seleccionar
	send("{DOWN 2}") ; baja dos veces
	send ("{SPACE}") ; seleccionar
	send("{DOWN 1}") ; baja
	send ("{SPACE}")  ; seleccionar
	send("{DOWN 1}")  ; baja
	send ("{SPACE}")  ; seleccionar
	send("{DOWN 3}")  ; baja tres veces
	send ("{SPACE}")  ; seleccionar
	send("{DOWN 20}")  ; baja
	send("{TAB}")     ; abajo
	Send ("{ENTER}") ; Visualizar
	;----------------------------------------------
	;Sleep(9000) ; Tiempo de Generacion de Historias
    
	;----------------------------------------------
	Sleep(800000) ; Tiempo de Generacion de Historias
	
	MouseClick("left",739,649,3) ; Coordenada de Boton de exportacion PDF 
	MouseClick("left",739,649,3) ; Se asegura la misma coordenada
	Sleep(2000)
	WinActivate("Opciones de exportación PDF")
	;-------PROBAR  TIEMPO----------------------------
	If not WinActivate("Opciones de exportación PDF") then
	Sleep(200000)
	MouseClick("left",739,649,3) ; Coordenada de Boton de exportacion PDF 
	MouseClick("left",739,649,3) ; Se asegura la misma coordenada
	EndIf
		
	Sleep(5000) ; Tiempo
	send("{TAB 11}") ; Se abre un cuadro de Dialogo para ubicarse en el boton 
	send ("{SPACE}") ; Aceptar
	Sleep(2000)	; Tiempo				
		
		
	; Se envia la ruta de Excel donde se guardara( Teniendo en cuenta el estandar del excel)
	Local $sResult4 = _Excel_RangeRead($oWorkbook, Default, "E2")	
    Local $sData4 = ClipGet() ; Guarda Variable
    ClipPut($sResult4)
    $sData4 = ClipGet() ; Se instancia
    WinActivate("Guardar un archivo PDF") ; Activa la ventana
	Sleep(5000) ; Tiempo
	send("{TAB 5}") ; Se ubica en la parte de ruta en el buscador
	Sleep(380); Tiempo
	Send ("{ENTER}") ; Enter
	Sleep(1000) ; Tiempo
    Send("^v") ; Pega la ruta desde el excel (Parametrizado)
	Send ("{ENTER}") ; Enter
	; Condicional En Caso de pasar una casilla en blanco, que en general sera cuando finalice el recorrido.
	
	
	; Se Crea la carperta con el Nombre con el que se Guarda. ( Parametrizacion de Excel)
	Local $sResult5 = _Excel_RangeRead($oWorkbook, Default, "F2")	
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
	
	If WinActivate("Confirmar el reemplazo de carpetas") then
	Sleep(3000)
    send("{RIGHT}")
	Sleep(3000)
	send ("{SPACE}") ; Enter
	Sleep(3000) ; tiempo	
	send("{TAB 6}")
	Sleep(3000)
	Send ("{ENTER}")
	Sleep(3000)
	Send("{ALTDOWN}{F4}{ALTUP}")
	Sleep(3000)
	WinActivate("Gestión de historias 4.3.3 (chigeshis)")
	Sleep(3000)
	Send("{ALTDOWN}{F4}{ALTUP}")
	Sleep(3000)
	WinActivate("SERVINTE CLINICAL SUITE  - HOSPITAL MANUEL URIBE ANGEL -  Tablero de Pacientes de HCE 4.3.5 (chiepiact)")
	Sleep(3000)	
	Send("{ALTDOWN}{F4}{ALTUP}")
	Sleep(3000)
	WinActivate("Acceso Centralizado")
	Sleep(3000)
	MouseClick("left",1232,85,1)
	Sleep(3000)
	MouseClick("left",1238,219,1)
	Sleep(3000)
	Correo() ; Hbailita la funcion de Correo para enviar el mensaje de finalizado
	Exit	
	EndIf
	
	Send ("{ENTER}") ; se entra a la carpeta
	Sleep(1000) ; Tiempo
	
	; Se escribe el nombre con el que se debe guardar desde el excel. ( Parametrizacion de Excel)
	Local $sResult5 = _Excel_RangeRead($oWorkbook, Default, "G2")	
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
	;-----------------------
	Sleep(4000)
Send("{SHIFTDOWN}{TAB 3}{SHIFTUP}")
Sleep(3000)
send ("{SPACE}") ; Enter
Sleep(4000)
Send("{CTRLDOWN}c{CTRLUP}")
Sleep(4000)
	
    $Copy = ClipGet() ;
	If StringInstr($Copy,"Reporte de cirugía",0) then ;si el reporte de cirujia se ubica en la ultima casilla, pasa a realizar el siguiente proceso
	
	send ("{SPACE}")
	Sleep(400) ; tiempo
	send("{UP 20}")	
	Sleep(400) ; tiempo
	send ("{SPACE}") ; Selecciona todo
	Sleep(400) ; tiempo
	send ("{SPACE}") ; ; quita seleccion
	send("{DOWN 20}")  ; baja dos veces
	send ("{SPACE}")
	send("{TAB}")     ; abajo
	Send ("{ENTER}") ; Visualizar
	Sleep(4000)
	Sleep(9000)
	MouseClick("left",739,649,3) ; Coordenada de Boton de exportacion PDF 
	MouseClick("left",739,649,3) ; Se asegura la misma coordenada
	Sleep(5000) ; Tiempo
	Sleep(5000) ; Tiempo
	send("{TAB 11}") ; Se abre un cuadro de Dialogo para ubicarse en el boton 
	send ("{SPACE}") ; Aceptar
	Sleep(3000)	; Tiempo				
	send("{TAB 9}") ;vuelve y toma el proceso de guardado
	Sleep(3000)
	send("{DOWN}")
	Sleep(3000)
	send("{TAB 2}")
	Sleep(1000)
    send("{RIGHT}")
	Sleep(1000)
    send("{LEFT 11}")
	Sleep(1000)
	Send("{BACKSPACE}{BACKSPACE}")
	Sleep(3000)
	Send("CX") ; reemplaza el valor de HC POR CX
	Sleep(3000)
	send("{TAB 3}")
	Sleep(3000)
	Send ("{ENTER}")
	Sleep(4000)
	EndIf
   
    	
	;--------------------------
	
	
	Sleep(3000) ; Tiempo
	WinActivate("SERVINTE CLINICAL SUITE - HCECLI 1.3 -  HOSPITAL MANUEL URIBE ANGEL")
	Sleep(3000)
	Send("{ALTDOWN}{F4}{ALTUP}") ;Se cierra la ventana
	Sleep(5000) ; Tiempo
	WinActivate("Tablero de Pacientes") ;Habilitacion de la ventana
	Send ("{ENTER}") ; Enter
    Sleep(5000) ; Tiempo
	WinActivate("Gestión de historias 4.3.3 (chigeshis)") ; Habilitacion de la ventana
	Send("{ALTDOWN}{F4}{ALTUP}") ; Se cierra 
	Sleep(5000) ; Tiempo	
	
	CicloExcel() ; Se Instancia la funcion de excel para la siguiente cedula en la lista
	Sleep(5000) ; Tiempo
	WinActivate("SERVINTE CLINICAL SUITE  - HOSPITAL MANUEL URIBE ANGEL -  Tablero de Pacientes de HCE 4.3.5 (chiepiact)")
	

endfunc