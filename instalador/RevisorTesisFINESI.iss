; Instalador de Windows del Revisor de Tesis FINESI (Inno Setup 6).
;
; Se genera con construir_instalador.bat, que primero arma dist\RevisorTesisFINESI.exe
; con PyInstaller y después compila este script. El resultado queda en
; instalador\salida\RevisorTesisFINESI-Setup-<versión>.exe
;
; Los datos del revisor (reglas, registro.csv, reportes, configuración) NO van en la
; carpeta de instalación, donde Windows no deja escribir: el .exe ve 'instalado.txt' a
; su lado y los guarda en Documentos\Revisor de Tesis FINESI. Desinstalar no los borra.

#define Nombre "Revisor de Tesis FINESI"
#define Version "1.0.0"
#define Exe "RevisorTesisFINESI.exe"
; carpeta del .exe ya construido; se puede cambiar al compilar: ISCC /DOrigen=otra\carpeta
#ifndef Origen
  #define Origen "..\dist"
#endif

[Setup]
; El AppId identifica al programa entre versiones: no cambiarlo, o una actualización
; se instalaría al lado de la anterior en vez de reemplazarla.
AppId={{6B684492-BBDB-4E68-9FD9-1B8D0A087B56}
AppName={#Nombre}
AppVersion={#Version}
AppVerName={#Nombre} {#Version}
AppPublisher=FINESI · Universidad Nacional del Altiplano
DefaultDirName={autopf}\{#Nombre}
DefaultGroupName={#Nombre}
DisableProgramGroupPage=yes
; Sin permisos de administrador se instala solo para el usuario actual; el asistente
; ofrece instalar para todos los usuarios a quien sí los tenga.
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=salida
OutputBaseFilename=RevisorTesisFINESI-Setup-{#Version}
SetupIconFile=..\icono.ico
UninstallDisplayIcon={app}\{#Exe}
UninstallDisplayName={#Nombre}
WizardStyle=modern
Compression=lzma2/max
SolidCompression=yes
; si el programa está abierto al actualizar, se ofrece cerrarlo
CloseApplications=yes

[Languages]
Name: "es"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "escritorio"; Description: "Crear un acceso directo en el escritorio"; GroupDescription: "Accesos directos:"

[Files]
Source: "{#Origen}\{#Exe}"; DestDir: "{app}"; Flags: ignoreversion
Source: "instalado.txt"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#Nombre}"; Filename: "{app}\{#Exe}"
Name: "{autodesktop}\{#Nombre}"; Filename: "{app}\{#Exe}"; Tasks: escritorio

[Run]
Filename: "{app}\{#Exe}"; Description: "Abrir el {#Nombre}"; Flags: nowait postinstall skipifsilent

