# Sample Excel File Structure

This shows the expected structure for the `lmunyc.xlsx` file:

## Required Column: "Client" and "Numéro"

The Excel file should have columns named **"Client"** (containing client names) and **"Numéro"** (containing phone numbers to extract RIO codes for).

### Example structure:

| Client | Numéro | RIO |
|--------|--------|-----|
| ABC DELIGHT DENTAIRE | 0607371063 | (will be filled automatically) |
| ABC DELIGHT DENTAIRE | 0698741258 | (will be filled automatically) |
| ANOTHER CLIENT NAME | 0612345678 | (will be filled automatically) |
| THIRD CLIENT | 0698765432 | (will be filled automatically) |

### Important Notes:

1. **Column Names**: The columns MUST be named exactly "Client" and "Numéro" (case-sensitive)
2. **File Name**: The file MUST be named "lmunyc.xlsx" and placed in the same directory as the script
3. **Sheet Name**: Must use the "Feuil1" sheet
4. **Format**: Must be an Excel file (.xlsx format)
5. **Multiple Numbers**: Each client can have multiple rows for different phone numbers
6. **RIO Column**: Will be created automatically to store extracted RIO codes
7. **Empty Rows**: Empty or blank entries will be automatically skipped

### How it works:

- The script will only look for phone numbers that are listed in the "Numéro" column
- It will NOT extract RIO codes for phone numbers found on the client page that are not in your Excel list
- This gives you precise control over which phone numbers to process

### Example Content:

```
Client Names to include in your Excel file:
- ABC DELIGHT DENTAIRE
- CABINET MEDICAL EXAMPLE
- SOCIETE EXEMPLE
- etc...
```

The script will:
1. Read all client names from the "Client" column
2. Search for each client on the UNYC platform
3. Click on the client link when found
4. Navigate to their detail page
5. Return to the client list for the next search
6. Provide a summary report at the end