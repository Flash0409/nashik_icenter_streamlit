# Nashik iCenter Streamlit App

This app reads your orderbook `.xlsm` file and provides:

- Project-wise overall details with only:
  - Assembly Start Date / Job Start Date
  - ITEM
  - Item Description
  - Components
  - Component Description
- ITEM/Work Order picker by project
- Editable Job Start Date for selected items
- Stock availability view (if stock columns are mapped)

## 1) Install dependencies

```powershell
cd c:\Users\E1547548\Desktop\nashik_icenter_streamlit
pip install -r requirements.txt
```

## 2) Run app

```powershell
streamlit run app.py
```

## Notes

- Default file path is pre-filled as:
  `c:\Users\E1547548\Desktop\Nashik iCenter Orderbook_NSK_Feb 242026.xlsm`
- If auto-column mapping is not correct, update mappings in the app UI.
- Use "Save as a new file" to keep original workbook unchanged.
