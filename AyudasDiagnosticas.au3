#include <Excel.au3>

AyudasMain2()
func AyudasMain2()
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

main2()
endfunc



func main2()

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
Global $oWorkbook = _Excel_BookOpen($oExcel, "C:\Users\jpelayo\Documents\HistoriasClinicas\AyudasDiagnosticas.xlsx" )
Sleep(5000)
Local $sResult = _Excel_RangeRead($oWorkbook, Default, "B2")
Local $sData = ClipGet() ; Cluck Obtiene la variable 1
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
	Mensaje9611() ; Instancia el Ciclo de Excel para que pase a enviar la otra cedula
	Sleep(3000)
	WinActivate("Gestión de historias 4.3.3 (chigeshis)")
	Sleep(3000)
    Send("{ALTDOWN}{F4}{ALTUP}")
	Sleep(3000)
	return main2() ; Retorna para que tome la siguiente Cedula en la lista
	EndIf
	; Condicional En Caso de pasar una casilla en blanco, que en general sera cuando finalice el recorrido.
	If WinActivate("Mensaje 1751") then
	Sleep(2000) ; tiempo
	send ("{SPACE}") ; Enter
	Sleep(1000) ; tiempo	
	Correo5() ; Hbailita la funcion de Correo para enviar el mensaje de finalizado
	Exit	
	EndIf
	; En caso de que no cumpla ninguna de esas condiciones  se siguel el proceso de exportacion Historias
	Sleep(3000)
	MouseClick("left",643,161,2) ; Darle clip a la primer historia por fecha para generar 
	Sleep(7000)
	WinActivate("SERVINTE CLINICAL SUITE - Resultados Clínicos - HOSPITAL MANUEL URIBE ANGEL     -  2.2.1 - chiconres")

;$cont=0
;global $num_elem = 100
;While $cont < $num_elem
    
	Sleep(6000)
	Send("{LEFT}")
	Sleep(6000)
	Send("{ENTER}")
	Sleep(2000)
	MouseClick("left",118,163,1)
	Sleep(2000)
	MouseClick("left",118,163,1)
	Sleep(2000)
	Send("{DOWN 100}")
	Sleep(2000)
	Sleep(4000)
	If WinActivate("Servinte - Visor de Imágenes") Then
	Sleep(7000)
	send("{TAB 6}")
	Sleep(2000)
	Send("{ENTER}")
	Sleep(2000)
	Send("{UP}")
	Sleep(2000)
	Send("{ENTER}")
	Sleep(2000)
	WinActivate("PDF24 Assistant")
	Sleep(2000)
	MouseClick("left",565,273,2)
	guardar()	
	endif
	
	If not WinActivate("Servinte - Visor de Imágenes") Then
	Sleep(2000)
	WinActivate("SERVINTE CLINICAL SUITE - Resultados Clínicos - HOSPITAL MANUEL URIBE ANGEL     -  2.2.1 - chiconres")
	Sleep(2000)
	Send("{ALTDOWN}{F4}{ALTUP}")
	Sleep(2000)
	WinActivate("Gestión de historias 4.3.3 (chigeshis)")
	Sleep(3000)
	Send("{ALTDOWN}{F4}{ALTUP}")
	Sleep(3000)
	Mensaje9611()
	;return main2()
	EndIf
	;-----------------------
    Send("{DOWN 100}")
	Sleep(2000)
	WinActivate("Servinte - Visor de Imágenes")
	If WinActivate("Servinte - Visor de Imágenes") Then
	Sleep(7000)
	send("{TAB 6}")
	Sleep(2000)
	Send("{ENTER}")
	Sleep(2000)
	Send("{UP}")
	Sleep(2000)
	Send("{ENTER}")
	Sleep(2000)
	WinActivate("PDF24 Assistant")
	Sleep(2000)
	MouseClick("left",565,273,2)
	guardar2()	
	endif
	Send("{DOWN 5}")
	Sleep(2000)
	WinActivate("Servinte - Visor de Imágenes")
	If WinActivate("Servinte - Visor de Imágenes") Then
	Sleep(7000)
	send("{TAB 6}")
	Sleep(2000)
	Send("{ENTER}")
	Sleep(2000)
	Send("{UP}")
	Sleep(2000)
	Send("{ENTER}")
	Sleep(2000)
	WinActivate("PDF24 Assistant")
	Sleep(2000)
	MouseClick("left",565,273,2)
	guardar3()	
	endif
	Send("{DOWN 5}")
	Sleep(2000)
	WinActivate("Servinte - Visor de Imágenes")
	If WinActivate("Servinte - Visor de Imágenes") Then
	Sleep(7000)
	send("{TAB 6}")
	Sleep(2000)
	Send("{ENTER}")
	Sleep(2000)
	Send("{UP}")
	Sleep(2000)
	Send("{ENTER}")
	Sleep(2000)
	WinActivate("PDF24 Assistant")
	Sleep(2000)
	MouseClick("left",565,273,2)
	guardar4()	
	endif
	Send("{DOWN 5}")
	Sleep(2000)
	WinActivate("Servinte - Visor de Imágenes")
	If WinActivate("Servinte - Visor de Imágenes") Then
	Sleep(7000)
	send("{TAB 6}")
	Sleep(2000)
	Send("{ENTER}")
	Sleep(2000)
	Send("{UP}")
	Sleep(2000)
	Send("{ENTER}")
	Sleep(2000)
	WinActivate("PDF24 Assistant")
	Sleep(2000)
	MouseClick("left",565,273,2)
	guardar5()	
	endif
	
	Send("{DOWN 5}")
	Sleep(2000)
	WinActivate("Servinte - Visor de Imágenes")
	If WinActivate("Servinte - Visor de Imágenes") Then
	Sleep(7000)
	send("{TAB 6}")
	Sleep(2000)
	Send("{ENTER}")
	Sleep(2000)
	Send("{UP}")
	Sleep(2000)
	Send("{ENTER}")
	Sleep(2000)
	WinActivate("PDF24 Assistant")
	Sleep(2000)
	MouseClick("left",565,273,2)
	guardar6()	
	endif
	Send("{DOWN 5}")
	Sleep(2000)
	WinActivate("Servinte - Visor de Imágenes")
	If WinActivate("Servinte - Visor de Imágenes") Then
	Sleep(7000)
	send("{TAB 6}")
	Sleep(2000)
	Send("{ENTER}")
	Sleep(2000)
	Send("{UP}")
	Sleep(2000)
	Send("{ENTER}")
	Sleep(2000)
	WinActivate("PDF24 Assistant")
	Sleep(2000)
	MouseClick("left",565,273,2)
	guardar7()	
	endif
	Send("{DOWN 5}")
	Sleep(2000)
	WinActivate("Servinte - Visor de Imágenes")
	If WinActivate("Servinte - Visor de Imágenes") Then
	Sleep(7000)
	send("{TAB 6}")
	Sleep(2000)
	Send("{ENTER}")
	Sleep(2000)
	Send("{UP}")
	Sleep(2000)
	Send("{ENTER}")
	Sleep(2000)
	WinActivate("PDF24 Assistant")
	Sleep(2000)
	MouseClick("left",565,273,2)
	guardar8()	
	endif
	Send("{DOWN 5}")
	Sleep(2000)
	WinActivate("SERVINTE CLINICAL SUITE - Resultados Clínicos - HOSPITAL MANUEL URIBE ANGEL     -  2.2.1 - chiconres - ")
	Sleep(3000)
	Send("{ALTDOWN}{F4}{ALTUP}")
	Sleep(3000)	
	WinActivate("Gestión de historias 4.3.3 (chigeshis)")
	Sleep(3000)
	Send("{ALTDOWN}{F4}{ALTUP}")
	Sleep(3000)
	CicloExcel4()

endfunc
;-------------PRINCIPAL AYUDAS DIAGNOSTICAS ------------------------

func Mensaje9611()
	; Ciclo de Excel respecto a cada numero de Cedula
    WinActivate("Microsoft Excel - AyudasDiagnosticas") ; HABILITAR EXCEL
	Sleep(700) ;tiempo
	WinActivate("Microsoft Excel - AyudasDiagnosticas")
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
	main2()
endfunc

func CicloExcel4()
	; Ciclo de Excel respecto a cada numero de Cedula
    WinActivate("Microsoft Excel - AyudasDiagnosticas") ; HABILITAR EXCEL
	Sleep(700) ;tiempo
	Send("{HOME}") ; casilla principal de excel
	Sleep(5000) ;tiempo  
    Send("{HOME}") ; casilla principal de excel
    Sleep(3000) ;tiempo
    Send("{CTRLDOWN}{HOME}{CTRLUP}") ;Ubicacion en la primera (Que se muestra visual)
    ;------------------------------------------------ 
	send("{TAB 4}")
	Sleep(3000) ;tiempo
	Send("Ayudas Diagnosticas Exportado")
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
	main2()
endfunc

func guardar()
   Sleep(9000)
Local $sResult2 = _Excel_RangeRead($oWorkbook, Default, "C2")	
    Local $sData2 = ClipGet() ; Guarda Variable
    ClipPut($sResult2)
    $sData2 = ClipGet() ; Se instancia
    send("{TAB 5}") ; Se ubica en la parte de ruta en el buscador
	Sleep(2000); Tiempo
	Send ("{ENTER}") ; Enter
	Sleep(1000) ; Tiempo
    Send("^v") ; Pega la ruta desde el excel (Parametrizado)
	Send ("{ENTER}") ; Enter
;-----------------------------------------------------
Local $sResult5 = _Excel_RangeRead($oWorkbook, Default, "D2")	
    Local $sData5 = ClipGet() ; Guarda la variable
    ClipPut($sResult5) 
    $sData5 = ClipGet() ; Se instancia
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
	WinActivate("Servinte - Visor de Imágenes")
	Sleep(3000)
	Send("{ALTDOWN}{F4}{ALTUP}")
	Sleep(3000)
	WinActivate("SERVINTE CLINICAL SUITE - Resultados Clínicos - HOSPITAL MANUEL URIBE ANGEL     -  2.2.1 - chiconres - ")
	Sleep(3000)
	Send("{ALTDOWN}{F4}{ALTUP}")
	Sleep(3000)
	WinActivate("Servinte - Visor de Imágenes")
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
	Correo5() ; Hbailita la funcion de Correo para enviar el mensaje de finalizado
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
	Sleep(3000)
	WinActivate("Servinte - Visor de Imágenes")
	Send("{ALTDOWN}{F4}{ALTUP}")
	Correo5()
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
    WinActivate("Servinte - Visor de Imágenes")
	Send("{ALTDOWN}{F4}{ALTUP}")
	Sleep(4000)	
;-------------------------------------------------------	

endfunc

func guardar2()
   Sleep(9000)
   send("{TAB 9}")
   Sleep(1000)
   send("{DOWN}") 
   Sleep(1000)
   send("{TAB 2}")
   Sleep(2000)
   send("{RIGHT}")
    Sleep(1000)
    send("{LEFT 11}")
	Sleep(1000)
	Send("2") ; reemplaza el valor de HC POR CX
	Sleep(3000)
	send("{TAB 3}")
	Sleep(3000)
	Send ("{ENTER}")
	Sleep(4000)
	
    WinActivate("Servinte - Visor de Imágenes")
	Send("{ALTDOWN}{F4}{ALTUP}")
	Sleep(4000)	
endfunc

func guardar3()
   Sleep(9000)
   send("{TAB 9}")
   Sleep(1000)
   send("{DOWN}") 
   Sleep(1000)
   send("{TAB 2}")
   Sleep(2000)
   send("{RIGHT}")
    Sleep(1000)
    send("{LEFT 11}")
	Sleep(1000)
	Send("{BACKSPACE}")
	Sleep(1000)
	Send("3") ; reemplaza el valor de HC POR CX
	Sleep(3000)
	send("{TAB 3}")
	Sleep(3000)
	Send ("{ENTER}")
	
	Sleep(4000)
    WinActivate("Servinte - Visor de Imágenes")
	Send("{ALTDOWN}{F4}{ALTUP}")
	Sleep(4000)	
endfunc
func guardar4()
   Sleep(9000)
   send("{TAB 9}")
   Sleep(1000)
   send("{DOWN}") 
   Sleep(1000)
   send("{TAB 2}")
   Sleep(2000)
   send("{RIGHT}")
    Sleep(1000)
    send("{LEFT 11}")
	Sleep(1000)
	Send("{BACKSPACE}")
	Sleep(1000)
	Send("4") ; reemplaza el valor de HC POR CX
	Sleep(3000)
	send("{TAB 3}")
	Sleep(3000)
	Send ("{ENTER}")
	
    WinActivate("Servinte - Visor de Imágenes")
	Send("{ALTDOWN}{F4}{ALTUP}")
	Sleep(4000)	
endfunc
func guardar5()
   Sleep(9000)
   send("{TAB 9}")
   Sleep(1000)
   send("{DOWN}") 
   Sleep(1000)
   send("{TAB 2}")
   Sleep(2000)
   send("{RIGHT}")
    Sleep(1000)
    send("{LEFT 11}")
	Sleep(1000)
	Send("{BACKSPACE}")
	Sleep(1000)
	Send("5") ; reemplaza el valor de HC POR CX
	Sleep(3000)
	send("{TAB 3}")
	Sleep(3000)
	Send ("{ENTER}")
	Sleep(4000)

    WinActivate("Servinte - Visor de Imágenes")
	Send("{ALTDOWN}{F4}{ALTUP}")
	Sleep(4000)	
endfunc
func guardar6()
   Sleep(9000)
   send("{TAB 9}")
   Sleep(1000)
   send("{DOWN}") 
   Sleep(1000)
   send("{TAB 2}")
   Sleep(2000)
   send("{RIGHT}")
    Sleep(1000)
    send("{LEFT 11}")
	Sleep(1000)
	Send("{BACKSPACE}")
	Sleep(1000)
	Send("6") ; reemplaza el valor de HC POR CX
	Sleep(3000)	
	send("{TAB 3}")
	Sleep(3000)
	Send ("{ENTER}")
	Sleep(4000)
    WinActivate("Servinte - Visor de Imágenes")
	Send("{ALTDOWN}{F4}{ALTUP}")
	Sleep(4000)	
endfunc
func guardar7()
   Sleep(9000)
   send("{TAB 9}")
   Sleep(1000)
   send("{DOWN}") 
   Sleep(1000)
   send("{TAB 2}")
   Sleep(2000)
   send("{RIGHT}")
    Sleep(1000)
    send("{LEFT 11}")
	Sleep(1000)
	Send("{BACKSPACE}")
	Sleep(1000)
	Send("7") ; reemplaza el valor de HC POR CX
	Sleep(3000)
	send("{TAB 3}")
	Sleep(3000)
	Send ("{ENTER}")
	Sleep(4000)

    WinActivate("Servinte - Visor de Imágenes")
	Send("{ALTDOWN}{F4}{ALTUP}")
	Sleep(4000)	
endfunc
func guardar8()
   Sleep(9000)
   send("{TAB 9}")
   Sleep(1000)
   send("{DOWN}") 
   Sleep(1000)
   send("{TAB 2}")
   Sleep(2000)
   send("{RIGHT}")
    Sleep(1000)
    send("{LEFT 11}")
	Sleep(1000)
	Send("{BACKSPACE}")
	Sleep(1000)
	Send("8") ; reemplaza el valor de HC POR CX
	Sleep(3000)
	send("{TAB 3}")
	Sleep(3000)
	Send ("{ENTER}")
	Sleep(4000)

    WinActivate("Servinte - Visor de Imágenes")
	Send("{ALTDOWN}{F4}{ALTUP}")
	Sleep(4000)	
endfunc

Func Correo5()
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
	send("Ayudas Diagnosticas")
	Sleep(6000) ;tiempo
	send("{TAB}") ; Mensaje
	Sleep(6000) ;tiempo
	send("SE HA REALIZADO EXITOSAMENTE, LA GENERACION DE AYUDAS DIAGNOSTICAS")
	Sleep(6000) ;tiempo
	send("{TAB 2}") ;ENVIAR CORREO
	Sleep(6000) ;tiempo
	Send ("{ENTER}")	   

endfunc