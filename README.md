# OneDrive Excel Row App

Simple app: submit form -> append one row to an Excel table stored in OneDrive.

## 1) Setup

```bash
cd onedrive-excel-row-app
npm install
```

Copy env file:

```bash
# PowerShell
Copy-Item .env.example .env
```

Fill `.env` values:
- `MS_TENANT_ID`
- `MS_CLIENT_ID`
- `MS_CLIENT_SECRET`
- `MS_DRIVE_ID`
- `MS_FILE_ITEM_ID`
- `MS_TABLE_NAME`

## 2) Run

```bash
npm run dev
```

Open:
- `http://localhost:4000`

## 3) Excel requirements

1. In your OneDrive Excel file, convert your range into a table (`Insert -> Table`).
2. Set table name (e.g. `Table1`) and use it as `MS_TABLE_NAME`.
3. In Azure app registration, grant Graph app permission to write files (typically `Files.ReadWrite.All`) and grant admin consent.

## Inserted columns

The app writes this row:
1. date
2. expense
3. subexpense
4. amount
5. note
