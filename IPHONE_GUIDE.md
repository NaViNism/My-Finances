# 📱 FinanceFlow - iPhone Guide & Local Hosting

Welcome to **FinanceFlow**, an offline-first personal finance tracking web application built specifically for your iPhone.

---

## 🚀 Part 1: All Requested Updates Applied

### 1. Compact Buttons & Order
- **Record Transaction** is now placed **first**, followed by **Export Excel**.
- The button dimensions and padding have been compacted so that both buttons fit side-by-side on any iPhone screen size (including iPhone SE and standard models) without overflowing or clipping.

### 2. Title Changed
- Changed the header title from **"Total Net Worth"** to **"Summary"**.

### 3. Record Modal Centered & Fully Scrollable
- The modal is now **anchored in the center of the screen** (`max-height: 84dvh`).
- Smooth internal scrolling ensures you can comfortably scroll through all inputs.
- The **Submit / Save Transaction** button is placed with generous bottom margin and is **never obscured by Safari's bottom search or navigation bar**.

### 4. Excel Data Import (Setup Page)
- Added **"Import Data from Excel"** in the **Setup / Settings** tab.
- Supports `.xlsx`, `.xls`, and `.csv`.
- Maps columns according to your exact structure:
  - **`Period`**: Extracted as Date, and time is automatically set to **`09:00:00` (9 AM)**.
  - **`Accounts`**: Mapped to your account (e.g. SBI, Bank 2, Cash). If a new account name is encountered, it is automatically created!
  - **`Category`**: Mapped to Category. Any new category is automatically added to your persistent categories.
  - **`INR`**: Mapped to Amount (with fallback to `Amount`).
  - **`Income/Expense`**: Mapped to Income, Expense, or Transfer.
  - **`Note`**: Mapped to Note (with fallback to `Description`).
- **Real-time Balance Reflection**: When the sheet is imported, every transaction is added to your records and your account balances are calculated and updated immediately!

### 5. Zero Sample Values
- Removed all dummy transactions.
- Clean zero starting state for accounts (`₹0.00`) and transaction list (`[]`).

---

## 📲 Part 2: Can You Host It Locally from Your iPhone Itself?

**YES! Absolutely.** You have two great ways to run it 100% locally on your iPhone without paying for any domain or hosting:

---

### Option A (Recommended): Add to Home Screen as an Offline PWA (Zero Server Needed!)
Because FinanceFlow is an **offline-first Progressive Web App** with a local Service Worker and LocalStorage:
1. Make sure your PC server is running (`python serve.py` or double-click `start.bat`).
2. On your iPhone connected to the same Wi-Fi/Hotspot, open Safari and visit:
   ```
   http://172.20.10.2:8080
   ```
3. Tap Safari's **Share button** (square with arrow up at the bottom).
4. Tap **"Add to Home Screen"** and tap **"Add"**.
5. **From now on, you do NOT need your PC or any server at all!**
   - The Service Worker caches all HTML, CSS, JavaScript, icons, and libraries directly in your iPhone's internal storage.
   - When you tap the FinanceFlow app icon on your iPhone home screen, it opens and runs **100% locally on your iPhone hardware**—even in Airplane Mode or without Wi-Fi! All data is saved directly on your phone.

---

### Option B: Run a Genuine Local Web Server Directly Inside Your iPhone (Using a-Shell)

If you want an actual Python web server running inside iOS on `localhost`:

1. Download the free app **[a-Shell](https://apps.apple.com/app/a-shell/id1473805438)** or **a-Shell mini** from the Apple App Store (it is a full local terminal environment for iOS).
2. Transfer or copy the `finance-tracker` folder to your iPhone:
   - On your iPhone, open the **Files** app.
   - Go to **On My iPhone** → **a-Shell**.
   - Copy the `finance-tracker` folder there (via AirDrop, iCloud Drive, or USB).
3. Open the **a-Shell** app and type:
   ```bash
   cd finance-tracker
   python -m http.server 8080
   ```
4. Now open Safari on your iPhone and visit:
   ```
   http://localhost:8080
   ```
5. It will run completely self-contained on your iPhone with zero external devices or domains!
