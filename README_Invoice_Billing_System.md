# Invoice Billing and Payment Analytics System

A Microsoft Excel VBA UserForm project for creating invoices, managing product line items, tracking payment status, saving invoice data, searching existing invoices, updating/deleting invoices, and generating PDF invoices with a QR code.

This project is built completely inside Excel VBA. It uses a dynamic UserForm interface and stores invoice data in Excel worksheets.

---

## Features

- Dynamic VBA UserForm interface
- Invoice number auto-generation
- Customer and contact details entry
- Product dropdown list
- Auto rate fill based on selected product
- GST dropdown support
- Multiple product line items
- Automatic amount, GST amount, total, paid amount, and balance calculation
- Duplicate product prevention inside the same invoice
- Add and remove product items
- S.No re-numbering after removing items
- Save invoice data to Excel sheets
- Search invoice by invoice number
- Update existing invoice
- Delete invoice
- Export invoice as PDF
- Generate invoice QR code
- Phone, quantity, and rate validation
- Indian Rupee symbol formatting

---

## Technologies Used

- Microsoft Excel
- VBA
- MSForms UserForm controls
- Excel Worksheets as local database
- PDF export using `ExportAsFixedFormat`
- QR image download using `URLDownloadToFile`

---

## Workbook Sheets Used

The VBA code creates and uses these sheets automatically:

| Sheet Name | Purpose |
|---|---|
| `Invoice_Data` | Stores one row per invoice with customer details, payment status, totals, and balance |
| `Invoice_Items` | Stores product line items for each invoice |
| `Invoice_Print` | Temporary formatted invoice sheet used for PDF generation |

---

## Basic Workflow

1. Open the Excel macro-enabled workbook.
2. Open the invoice UserForm.
3. Enter customer details.
4. Select product from dropdown.
5. Enter quantity.
6. GST and rate are used to calculate totals.
7. Click **ADD ITEM** to add product into the item table.
8. Choose payment status and paid amount.
9. Click **SAVE** to store invoice.
10. Click **PDF** to generate a printable invoice PDF.
11. Use invoice number and click **SEARCH** to load old invoices.

---

## Code Explanation: Section by Section

### 1. `Option Explicit`

Forces every variable to be declared before use. This helps avoid spelling mistakes and runtime errors in VBA.

---

### 2. Company and Payment Constants

```vb
Private Const UPI_ID As String
Private Const PAYEE_NAME As String
Private Const COMPANY_ADDRESS As String
Private Const COMPANY_GSTIN As String
Private Const COMPANY_PHONE As String
Private Const COMPANY_EMAIL As String
```

These constants store fixed company and payment details used in the invoice PDF and QR section.

---

### 3. `URLDownloadToFile` Declaration

This Windows API function downloads the QR code image from an online QR service and saves it temporarily on the computer.

The code supports both:

- 64-bit Office using `PtrSafe`
- 32-bit Office using normal `Declare`

---

### 4. `WithEvents` Control Variables

```vb
Private WithEvents txtInvoiceNoEvt As MSForms.TextBox
Private WithEvents cmbStatusEvt As MSForms.ComboBox
Private WithEvents cmdSaveEvt As MSForms.CommandButton
```

These variables allow dynamically created UserForm controls to trigger events such as button clicks, textbox changes, and combobox changes.

This is important because the controls are created through VBA code instead of manually placed on the form.

---

### 5. `UserForm_Initialize`

Runs automatically when the UserForm opens.

It calls:

- `BuildInvoiceForm` to create the UI
- `CreateInvoiceSheets` to create required sheets
- `SetDefaultValues` to fill default invoice number, date, GST, and payment mode

---

### 6. `OnlyInteger`

Allows only digits from `0` to `9` and Backspace while typing.

Used for fields like:

- Phone
- Quantity
- Rate

---

### 7. `KeepOnlyDigits`

Cleans pasted text and keeps only numbers.

Example:

```text
abc123xyz
```

becomes:

```text
123
```

It also supports maximum length, such as limiting phone number to 10 digits.

---

### 8. TextBox KeyPress Events

```vb
txtPhoneEvt_KeyPress
txtQtyEvt_KeyPress
txtRateEvt_KeyPress
```

These events call `OnlyInteger`, so users cannot type letters into number-only fields.

---

### 9. `txtPhoneEvt_Change`

Calls `KeepOnlyDigits` and limits phone number to 10 digits.

---

### 10. `BuildInvoiceForm`

Creates the full UserForm design dynamically.

It adds:

- Header
- Invoice details section
- Customer contact section
- Product and amount section
- Product dropdown
- GST dropdown
- Add / Remove item buttons
- Product item list
- Save, Search, Update, Delete, Clear, PDF, and Close buttons

This procedure controls the main screen layout.

---

### 11. UI Helper Procedures

These procedures create reusable controls:

| Procedure | Purpose |
|---|---|
| `AddFrame` | Creates section heading labels |
| `AddProductDetailsBox` | Creates outer border/background for product table |
| `AddItemHeader` | Creates product table header labels |
| `AddLabel` | Creates normal labels |
| `AddTextBox` | Creates textboxes |
| `AddComboBox` | Creates dropdown boxes |
| `AddListBox` | Creates product list table |
| `AddButton` | Creates buttons |

These helper procedures reduce repeated code and keep the UI creation organized.

---

### 12. `GetProductPrice`

Returns the default price for a selected product.

Example:

- `Laptop` returns `45000`
- `Power Bank` returns `1200`
- `USB Cable` returns `250`

This is used to automatically fill the rate textbox when a product is selected.

---

### 13. `IsProductInList`

Checks whether the selected product is already added to the invoice item list.

This prevents duplicate products from being added to the same invoice.

---

### 14. `LoadProductDropdown`

Loads the product dropdown.

It also hides products that are already added to the invoice list. If a product is removed, it becomes available again in the dropdown.

---

### 15. `cmbProductEvt_Change`

Runs when product dropdown value changes.

It:

- Gets the product price using `GetProductPrice`
- Fills the rate textbox automatically
- Recalculates the invoice amount

---

### 16. `cmbGSTEvt_Change`

Runs when GST dropdown value changes.

It recalculates the invoice amount.

---

### 17. `RupeeSymbol`

Returns the Indian Rupee symbol:

```text
₹
```

Used in PDF and worksheet number formatting.

---

### 18. `CreateInvoiceSheets`

Creates required Excel sheets if they do not already exist.

It creates:

- `Invoice_Data`
- `Invoice_Items`

It also adds header columns for storing invoice and product item data.

---

### 19. `GetNextInvoiceNo`

Generates the next invoice number automatically.

Example:

```text
INV-0001
INV-0002
INV-0003
```

---

### 20. `SetDefaultValues`

Sets default values when the form opens or clears.

It fills:

- Next invoice number
- Current date and time
- GST default value
- Payment status
- Payment mode

---

### 21. `ItemListCount`

Returns how many product items are currently added in the invoice list.

---

### 22. `CalculateSingleLineAmount`

Calculates totals when only one product entry is being prepared before adding it to the item list.

It calculates:

- Amount before GST
- GST amount
- Total amount
- Paid amount
- Balance amount

---

### 23. `CalculateInvoiceAmount`

Decides which calculation method to use.

If product list has items, it calls:

```vb
RecalculateMultipleItemsTotal
```

Otherwise, it calls:

```vb
CalculateSingleLineAmount
```

---

### 24. `AddProductItem`

Adds a selected product into the invoice item list.

It validates:

- Product selected
- Quantity entered
- Rate available
- Duplicate product not already added

Then it calculates:

- Amount
- GST amount
- Total

Finally, it inserts the row into the ListBox.

---

### 25. `RemoveProductItem`

Removes the selected item from the product list.

After removing, it:

- Renumbers S.No
- Recalculates total
- Reloads product dropdown

---

### 26. `ReNumberItems`

Reorders the S.No column after an item is deleted.

Example:

Before deleting:

```text
1, 2, 4
```

After renumbering:

```text
1, 2, 3
```

---

### 27. `RecalculateMultipleItemsTotal`

Calculates grand total from all added product items.

It updates:

- Total amount
- Paid amount
- Balance amount

Paid amount is kept editable according to the latest code behavior.

---

### 28. `EnsureItemsAvailable`

Checks whether at least one product item exists before saving, updating, or generating PDF.

If the user entered product details but did not click **ADD ITEM**, this function can add the item before continuing.

---

### 29. `GetItemSummary`

Creates a summary text of products.

Example:

```text
Laptop x 1; Mouse x 2
```

This summary is saved in the main invoice data sheet.

---

### 30. `ClearInvoiceForm`

Clears customer details, product fields, totals, and product list.

It also reloads the product dropdown.

---

### 31. `FindInvoiceRow`

Searches the `Invoice_Data` sheet and returns the row number for the given invoice number.

Used by:

- Search
- Update
- Delete
- Duplicate invoice check

---

### 32. `DeleteInvoiceItems`

Deletes all product item rows from `Invoice_Items` for a given invoice number.

Used before updating or deleting invoice data.

---

### 33. `WriteInvoiceItems`

Writes each product item into the `Invoice_Items` sheet.

Each row stores:

- Invoice number
- Line number
- Product name
- Quantity
- Rate
- Amount
- GST %
- GST amount
- Line total

---

### 34. `GetProductNameFromSummary`

Extracts product name from an old summary format.

Used as fallback when old invoice records do not have item details in `Invoice_Items`.

---

### 35. `GetQtyFromSummary`

Extracts quantity from an old summary format.

Used for old saved invoices.

---

### 36. `LoadInvoiceItems`

Loads product items for a searched invoice.

It first checks `Invoice_Items`.

If no item rows are found, it uses fallback logic from `Invoice_Data`.

---

### 37. `SaveInvoiceData`

Saves a new invoice.

It:

- Checks duplicate invoice number
- Validates customer name
- Ensures items are available
- Saves invoice master data into `Invoice_Data`
- Saves item rows into `Invoice_Items`
- Shows success message
- Clears the form

---

### 38. `SearchInvoiceData`

Searches invoice by invoice number.

It loads:

- Customer details
- Payment status
- Payment mode
- Total amount
- Paid amount
- Balance amount
- Product line items

---

### 39. `UpdateInvoiceData`

Updates an existing invoice.

It requires searching invoice first, then updates:

- Invoice master data
- Product item rows
- Totals and balance

---

### 40. `DeleteInvoiceData`

Deletes the searched invoice.

It deletes:

- Invoice row from `Invoice_Data`
- Related item rows from `Invoice_Items`

---

### 41. Invoice Number Auto Search Event

```vb
txtInvoiceNoEvt_AfterUpdate
```

When the invoice number textbox is updated, the code checks if that invoice exists and loads it automatically.

---

### 42. Payment Status Event

```vb
cmbStatusEvt_Change
```

Runs when the payment status changes.

It adjusts paid amount behavior and recalculates totals.

---

### 43. Quantity, Rate, and Paid Amount Events

```vb
txtQtyEvt_Change
txtRateEvt_Change
txtPaidEvt_Change
```

These events recalculate the invoice whenever quantity, rate, or paid amount changes.

---

### 44. Button Click Events

| Event | Purpose |
|---|---|
| `cmdAddItemEvt_Click` | Adds product item |
| `cmdRemoveItemEvt_Click` | Removes product item |
| `cmdSaveEvt_Click` | Saves invoice |
| `cmdSearchEvt_Click` | Searches invoice |
| `cmdUpdateEvt_Click` | Updates invoice |
| `cmdDeleteEvt_Click` | Deletes invoice |
| `cmdClearEvt_Click` | Clears form |
| `cmdPDFEvt_Click` | Generates PDF |
| `cmdCloseEvt_Click` | Closes form |

---

### 45. `GenerateInvoicePDF`

Creates the invoice PDF.

It:

- Validates invoice number and customer name
- Ensures product items exist
- Prepares print sheet
- Creates PDF file
- Opens the generated PDF

---

### 46. `GetInvoicePrintSheet`

Gets or creates the `Invoice_Print` worksheet.

This sheet is used only for preparing the formatted PDF invoice.

---

### 47. `PrepareInvoicePrintSheet`

Formats the printable invoice.

It creates:

- Company header
- Invoice details
- Customer details
- Product table
- Grand total section
- Payment status and mode
- QR code section
- Page setup for PDF export

---

### 48. `CleanFileName`

Removes invalid characters from file names.

This prevents errors when saving the PDF file.

---

### 49. `BuildInvoiceDetailsQRText`

Builds short invoice text for the QR code.

QR text includes:

- Invoice number
- Customer name
- Phone
- Product details
- Total
- Paid amount
- Balance
- Status

---

### 50. `UrlEncodeText`

Converts QR text into URL-safe format.

This is useful when sending invoice text to the QR code generator URL.

---

### 51. `CleanQRFileName`

Cleans QR image file name before saving it temporarily.

---

### 52. `DownloadQRCodeImage`

Downloads QR code image using invoice QR text.

It saves the QR image as a temporary PNG file.

---

### 53. `AddUPIQRCodeToInvoice`

Adds the QR code section into the PDF invoice sheet.

It creates:

- Blue header: `SCAN INVOICE DETAILS`
- QR image
- UPI ID
- Balance amount
- Outer border

---

### 54. `FormatPhoneNumber`

Formats phone number in readable form before displaying it in PDF or QR text.

---

## Data Storage Design

### `Invoice_Data`

Stores one row per invoice.

Main fields:

- Invoice No
- Invoice Date Time
- Customer Name
- Phone
- Email
- City
- Payment Status
- Payment Mode
- Product Summary
- Item Count
- Total Amount
- Paid Amount
- Balance Amount

### `Invoice_Items`

Stores multiple rows per invoice.

Each product item is linked to the invoice number.

---

## Payment Logic

The system tracks:

```text
Balance = Total Amount - Paid Amount
```

Payment status options:

- Paid
- Pending
- Partial

Paid amount can be typed manually based on the latest version of the project.

---

## QR Code Logic

The system builds invoice details as text and converts that text into a QR code image.

The QR code can be scanned to view invoice details such as:

- Invoice number
- Customer name
- Phone number
- Product details
- Total
- Paid amount
- Balance
- Status

---

## How to Use

1. Open the Excel workbook.
2. Enable macros.
3. Open the invoice UserForm.
4. Fill customer details.
5. Select product and GST.
6. Enter quantity.
7. Click **ADD ITEM**.
8. Repeat for multiple products.
9. Enter paid amount.
10. Click **SAVE**.
11. Click **PDF** to export invoice.
12. Use **SEARCH** to find old invoices.
13. Use **UPDATE** or **DELETE** if needed.

---

## Notes

- This is an offline Excel-based invoice system.
- Internet is required only when generating QR code images.
- Invoice data is stored inside the workbook.
- The workbook should be saved as `.xlsm` to keep VBA macros.
- The QR code generator depends on an external QR image URL service.

---

## Future Improvements

- Add product master sheet instead of hardcoded products
- Add customer master sheet
- Add dashboard/report sheet
- Add tax summary report
- Add stock/inventory tracking
- Add logo in PDF invoice
- Add backup/export to CSV
- Add password protection
- Add cloud database support
- Add Power BI dashboard integration

---

## Author

**Loganathan K**

Project: Excel VBA Invoice Billing and Payment Analytics System
