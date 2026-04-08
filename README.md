# AutoCPV

Aplicació visual per a Windows per revisar activitats culturals des d'un Excel i enviar-les a un Google Form de manera ràpida.

## Què fa

- carrega un Excel amb activitats
- deixa editar cada fila amb desplegables i validacions
- obri cerques ràpides a Google i fonts
- envia una fila o totes al formulari
- guarda i reobri sessions de treball
- genera un instal·lador per a distribuir l'app

## Fitxers principals

- `app.py`: codi principal de l'aplicació
- `AutoCPV.spec`: empaquetat de l'executable amb PyInstaller
- `installer.iss`: script de l'instal·lador amb Inno Setup
- `assets/`: icones i imatges de l'instal·lador

## Compilar l'executable

```powershell
py -m PyInstaller --noconfirm AutoCPV.spec
```

L'executable es genera en `dist\AutoCPV.exe`.

## Compilar l'instal·lador

```powershell
& "C:\Users\solso\AppData\Local\Programs\Inno Setup 6\ISCC.exe" "C:\Users\solso\Documents\New project\installer.iss"
```

L'instal·lador es genera en `installer-dist\AutoCPV-Setup.exe`.

També es pot fer tot de colp amb:

```powershell
build-release.bat
```

## Publicació recomanada

La manera més simple de distribuir noves versions és pujar `AutoCPV-Setup.exe` a GitHub Releases.
