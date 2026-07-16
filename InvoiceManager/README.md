# Finish Invoice Manager

Offline Windows invoice application with machine-bound, digitally signed licensing.

## License workflow

1. Client opens `FinishInvoiceManager.exe` and sends the displayed Machine Code.
2. Owner opens `FinishLicenseAdmin.exe`, enters the customer, Machine Code, and optional expiry.
3. Owner sends the generated `.lic` file to the client.
4. Client imports the file from the activation window.

`FinishLicenseAdmin.key` is the private signing key. It is intentionally ignored by Git and must remain beside the Admin EXE only. Never send it to a client or commit it to a repository.

Modern Windows desktop invoice entry application. Data is stored locally in
`invoice_data/invoices.db`; PDFs are created in `invoice_data/pdfs`.

## Run from source

```powershell
python -m pip install -r requirements.txt
python app.py
```

## Build Windows EXE

Double-click `build_exe.bat`. The portable executable will be created at
`dist\FinishInvoiceManager.exe`.
