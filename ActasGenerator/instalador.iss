; instalador.iss - Script de Inno Setup para el Generador de Actas INTi
;
; Genera un instalador profesional de Windows que:
;   - Instala la app en C:\Program Files\Generador de Actas INTi\
;   - Crea acceso directo en el Escritorio con el logo del INTi
;   - Crea acceso directo en el Menú Inicio
;   - Crea la carpeta Documentos\Actas de ATASIS
;   - Aparece en "Agregar o quitar programas" con opción de desinstalar
;
; Para compilar: abrir este archivo con Inno Setup Compiler

#define AppName "ATASIS-INTi"
#define AppVersion "1.0.0"
#define AppPublisher "Instituto Nacional de Tierras (INTi)"
#define AppExeName "GeneradorActasINTi.exe"
#define SourceDir "dist\GeneradorActasINTi"

[Setup]
; Información de la aplicación
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL=https://www.inti.gob.ve
AppSupportURL=https://www.inti.gob.ve
AppUpdatesURL=https://www.inti.gob.ve

; Carpeta de instalación
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes

; Icono del instalador y la app (comentado temporalmente por error de formato)
; SetupIconFile=assets\logo.ico
UninstallDisplayIcon={app}\{#AppExeName}

; Permisos: lowest permite instalación sin admin, pero usaremos admin para Program Files
; Para evitar conflictos con userdocs, usaremos {commondocs} en lugar de {userdocs}
PrivilegesRequired=admin

; Archivo de salida del instalador
OutputDir=instalador
OutputBaseFilename=Instalador_GeneradorActasINTi_v{#AppVersion}

; Compresión
Compression=lzma2/ultra64
SolidCompression=yes

; Archivo de licencia que el usuario debe aceptar
LicenseFile=license.txt

; Información de la licencia y bienvenida
WizardStyle=modern
; WizardSmallImageFile=assets\logo.ico  ; Comentado por error de formato de imagen

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
; Opciones que el usuario puede elegir durante la instalación
Name: "desktopicon"; Description: "Crear acceso directo en el Escritorio"; GroupDescription: "Accesos directos:"
Name: "startmenuicon"; Description: "Crear acceso directo en el Menú Inicio"; GroupDescription: "Accesos directos:"

[Files]
; Copiar todos los archivos del ejecutable generado por PyInstaller
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; Iconos para el instalador (necesarios para SetupIconFile y WizardSmallImageFile)
Source: "assets\logo.ico"; DestDir: "{app}\assets"; Flags: ignoreversion

[Dirs]
; Crear la carpeta de actas en Documentos comunes (todos los usuarios)
; Usamos {commondocs} en lugar de {userdocs} para evitar conflicto con PrivilegesRequired=admin
Name: "{commondocs}\Actas de ATASIS"

[Icons]
; Acceso directo en el Escritorio
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\assets\logo.ico"; Tasks: desktopicon; Comment: "Generador de Actas del Instituto Nacional de Tierras"

; Acceso directo en el Menú Inicio
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\assets\logo.ico"; Tasks: startmenuicon; Comment: "Generador de Actas del Instituto Nacional de Tierras"

; Acceso directo para desinstalar en el Menú Inicio
Name: "{group}\Desinstalar {#AppName}"; Filename: "{uninstallexe}"

[Run]
; Ofrecer ejecutar la app al terminar la instalación
Filename: "{app}\{#AppExeName}"; Description: "Iniciar {#AppName}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Limpiar archivos generados durante el uso (logs, backups)
; NO elimina la carpeta Actas de ATASIS para no perder los documentos
Type: filesandordirs; Name: "{app}\historial\backup"
Type: files; Name: "{app}\historial\actas.log"
