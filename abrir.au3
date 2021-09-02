#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>


$Form1 = GUICreate("Hopital Manuel Uribe Angel - Jefferson Pelayo", 615, 438, 192, 124)
$Button1 = GUICtrlCreateButton("Generar Historias Clinicas 1", 120, 40, 353, 57)
$Button2 = GUICtrlCreateButton("Generar Historias Clinicas 2 (Laboratorio)", 120, 112, 353, 57)
$Button3 = GUICtrlCreateButton("Generar Laboratorio (Crear)", 120, 192, 353, 65)
$Button4 = GUICtrlCreateButton("Generar Ayudas Diagnosticas", 120, 280, 353, 57)
$Button5 = GUICtrlCreateButton("Cancelar", 232, 373, 121, 49)
GUISetState(@SW_SHOW)

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

$x = 1


While 1
$msg = GuiGetMsg()
Select
Case $msg = $GUI_EVENT_CLOSE
ExitLoop


   Case $msg = $Button1      ; Presiona Botón1
   if $x = 1 Then
   #include <ExportarHC.au3>  ; Historia Clinica
   Else
    GUICtrlSetState($Button1, $GUI_DISABLE)  ; Deshabilita el Botón1
   EndIf

   Case $msg = $Button2      ; Presiona Botón2
    if $x = 1 Then
	#include <LogeoLaboratorio.au3>  ; Logeo Laboratorio
   Else
    GUICtrlSetState($Button2, $GUI_DISABLE)  ; Deshabilita el Botón2
   EndIf

   Case $msg = $Button3      ; Presiona Botón3
    if $x = 1 Then
    #include <LogeoLaboratorioCreacion.au3>   ; Laboratorio Crear Carpeta
    Else
    GUICtrlSetState($Button3, $GUI_DISABLE)  ; Deshabilita el Botón3
   EndIf   
   
   
    Case $msg = $Button4      ; Presiona Botón4
    if $x = 1 Then
    #include <AyudasDiagnosticas.au3>   ; Ayudas Diagnosticas
    Else
    GUICtrlSetState($Button4, $GUI_DISABLE)  ; Deshabilita el Botón3
   EndIf   

    Case $msg = $Button5      ; Presiona Botón4
    if $x = 1 Then
    exit  ; Cancelar
    Else
    GUICtrlSetState($Button4, $GUI_DISABLE)  ; Deshabilita el Botón3
   EndIf   
   

EndSelect
Wend
Exit


