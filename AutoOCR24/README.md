# AutoOCR24

Aplicació visual per a Windows que:

- carrega un PDF
- executa OCR amb PDF24 Tools en segon pla
- detecta català i castellà
- deixa revisar i editar el text per pàgines
- exporta a `DOCX` i a `PDF` net

## Dependències

```powershell
py -m pip install selenium pdfplumber python-docx reportlab openpyxl pillow
```

També necessita Microsoft Edge instal·lat i connexió a internet.

## Ús

1. Obri `launch.bat`.
2. Selecciona un PDF.
3. Prem `Processar OCR`.
4. Revisa el text per pàgines.
5. Prem `Aplicar canvis de pàgina`.
6. Exporta a `DOCX` o a `PDF net`.

## Notes

- PDF24 Tools es fa servir via web en segon pla.
- El PDF OCR temporal es guarda en una carpeta temporal del sistema.
- El PDF net es genera a partir del text editat, no és el PDF OCR original.
