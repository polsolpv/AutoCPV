# AutoCPV

Aplicacio visual per a Windows per revisar activitats culturals des d'un Excel, treballar PDFs amb OCR i enviar dades a un Google Form.

## Que fa

- carrega un Excel amb activitats
- deixa editar cada fila amb desplegables i validacions
- obri cerques rapides a Google i fonts
- envia una fila o totes al formulari
- guarda i reobri sessions de treball
- processa PDFs amb OCR local de PDF24, permet cancel.lar i exporta a DOCX/PDF net
- genera un Excel AutoCPV des del PDF OCR amb ChatGPT i el prompt del projecte

## Entrada unica

- executable principal: `dist\AutoCPV.exe`
- instal.lador Windows: `installer-dist\AutoCPV-Setup.exe`
- acces directe recomanat: `C:\Users\solso\Desktop\AutoCPV.lnk`

No hi ha variants separades ni instal.ladors antics dins del projecte.

## Fitxers principals

- `app.py`: codi principal de l'aplicacio
- `AutoCPV.spec`: empaquetat de l'executable amb PyInstaller
- `installer.iss`: instal.lador Windows amb Inno Setup
- `build-release.ps1`: build local de l'executable i l'instal.lador
- `.github/workflows/build-windows.yml`: build automatic en GitHub Actions
- `requirements.txt`: dependencies de Python per a la build
- `PROMPT AutoCPV.txt`: prompt que s'envia a ChatGPT per crear l'Excel
- `assets/`: icona i imatges de marca
- `sample_scan.pdf`: PDF de prova

## OCR local

AutoCPV usa la instal.lacio local de PDF24:

- `C:\Program Files\PDF24\pdf24-Ocr.exe`
- o `C:\Program Files (x86)\PDF24\pdf24-Ocr.exe`
- idiomes OCR: catala + castella (`cat+spa`)

No usa PDF24 web ni navegador.

L'instal.lador comprova si PDF24 Creator esta instal.lat. Si no el troba, intenta instal.lar-lo amb:

```powershell
winget install --id geeksoftwareGmbH.PDF24Creator --source winget
```

Si `winget` no esta disponible, obri la web de PDF24 per a instal.lar-lo manualment.

## Excel amb ChatGPT via NVIDIA

En la pestanya OCR:

1. obri el PDF en PDF24 OCR
2. guarda el PDF OCR resultant
3. carrega el PDF OCR en AutoCPV
4. prem `Generar Excel amb ChatGPT`

La generacio usa NVIDIA Integrate amb el model `openai/gpt-oss-120b`.

La forma recomanada es fer-ho des de la mateixa app:

1. obri AutoCPV
2. entra en la pestanya OCR
3. prem `Configurar IA`
4. enganxa la clau `nvapi-...`
5. guarda

La configuracio queda en `%APPDATA%\AutoCPV\settings.json` en cada ordinador.

També es pot configurar per PowerShell:

```powershell
[Environment]::SetEnvironmentVariable("NVIDIA_API_KEY", "nvapi-...", "User")
```

Despres tanca i torna a obrir AutoCPV. Opcionalment pots canviar el model amb `NVIDIA_MODEL`; si no, s'usa `openai/gpt-oss-120b`.

## Recompilar

```powershell
.\build-release.ps1
```

Resultats:

- `dist\AutoCPV.exe`
- `installer-dist\AutoCPV-Setup.exe`

En GitHub, el workflow `Build AutoCPV Windows` genera els mateixos artefactes.
