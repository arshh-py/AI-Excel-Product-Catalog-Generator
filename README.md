# Local Inventory Catalog

Retail return stock ke liye local catalog app. Excel upload karo, product catalog banao, images cache karo, prices edit karo, aur catalog export karo.

## Easy Start

Windows par double-click:

```text
start_windows.bat
```

Browser khul jayega. Agar nahi khulta, manually open karo:

```text
http://127.0.0.1:5000
```

Command window open rakhni hai. Window band karoge to app band ho jayega.

## Excel Format

Excel file `.xlsx` honi chahiye.

Required columns:

```text
Product Name
Price
```

Important:

- Image search sirf `Product Name` se hoti hai.
- Price image matching me kabhi use nahi hota.
- Full product title pehle search hota hai.
- App ab Amazon, Amazon image CDN, image search, web search, manufacturer sites, Indian retailers, global marketplaces, aur trusted shopping sites ko priority order me try karta hai.
- `Image not found` sirf tab dikhega jab saare source, query fallback, webpage extraction, aur candidate download attempts fail ho jayein.

## New Features

- Smarter image search with multiple candidates and title matching.
- Exhaustive image fallback search with source-by-source logs.
- Background import progress page.
- Admin page me product price edit.
- Admin page me image upload ya image URL se replace.
- Refresh image search for a product.
- Export catalog to Excel.
- Windows launcher file.
- Windows package builder for sharing with non-technical users.

## Admin

Open:

```text
http://127.0.0.1:5000/admin
```

Admin me:

- `Edit Product` se price, product name, image upload, ya image URL update karo.
- `Refresh Search` se image dobara internet se search karo.
- `Delete Cached Images` se image cache clear karo.
- `Export Catalog` se Excel download karo.

## Export

Catalog page ya Admin page par `Export Catalog` button use karo.

Export file me columns:

- Product Name
- Price
- Image File
- Image Source
- Created At
- Updated At

## Share With Another Windows PC

Apne PC par run karo:

```text
package_windows.bat
```

Ye folder banayega:

```text
dist\InventoryCatalog
```

Is poore `InventoryCatalog` folder ko doosre Windows PC me copy karo. Doosre PC par:

```text
InventoryCatalog.exe
```

double-click karo. Browser me open karo:

```text
http://127.0.0.1:5000
```

Note: `package_windows.bat` chalane ke liye internet chahiye kyunki ye PyInstaller install karta hai. Package banne ke baad doosre PC par Python/VS Code ki zarurat nahi honi chahiye.

## Local Data

Data yahin store hota hai:

```text
data/uploads/          Uploaded Excel files
data/image_cache/      Cached product images
data/inventory.sqlite3 Product database
data/image_search.log  Detailed image search logs
```

Fresh start ke liye app band karo aur `data` folder delete kar do.

## Manual Start

```bash
pip install -r requirements.txt
python app.py
```
