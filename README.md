# Complaint Handling App — Nanobank Syariah
## Google Apps Script Project

### Script ID & Deployment
- **Script ID**: `1xRn_z0JTnzNK78sTTx0iv2kiypJPXR-LySJ3T2OINK_Stw84OmkgQFhZ`
- **Web App URL**: `https://script.google.com/macros/s/AKfycbwdCr454KpSMuO3oeBy4GLqysrUyq3z75mKzGlHLGuOMC1Gdq630QuLfP4p17EJfOJowA/exec`
- **GAS Editor**: `https://script.google.com/d/1xRn_z0JTnzNK78sTTx0iv2kiypJPXR-LySJ3T2OINK_Stw84OmkgQFhZ/edit`

3. **Push file ke GAS**
   ```bash
   cd d:\Project\complaint-handling
   clasp push
   ```

4. **Deploy sebagai Web App**
   - Di GAS editor: **Deploy → New Deployment**
   - Type: **Web App**
   - Execute as: `User accessing the web app`
   - Who has access: `Anyone` (atau `Anyone within your organization`)
   - Klik **Deploy** → Salin URL Web App

### Struktur File
| File | Keterangan |
|------|-----------|
| `Code.gs` | Server-side: Auth, CRUD users, submit complaint, email, reporting |
| `Index.html` | Main shell: Login screen + App layout |
| `Stylesheet.html` | CSS global styles |
| `ViewComplaint.html` | View: Input Complaint |
| `ViewReporting.html` | View: Reporting |
| `ViewUsers.html` | View: User Management |
| `JavaScript.html` | Client-side JavaScript |
| `appsscript.json` | GAS manifest |

### User Matrix (Hak Akses)
| Role | Input Complaint | Reporting | Manajemen User |
|------|:-:|:-:|:-:|
| admin | ✓ | ✓ | ✓ |
| cs-cc | ✓ | ✓ | — |
| pk | — | ✓ | — |

### Google Spreadsheet
- ID: `1iJfE0AdB0kbHAnQv2OCJRip7jNxpsgB1q8obn8dQzL4`
- Sheets: `users`, `complaints`, `mapping`, `product`, `report`, `copy-writing`, `user-matrix`

### Catatan
- Autentikasi otomatis menggunakan **Profil Akun Google Chrome (SSO)** yang dicocokkan dengan sheet `users`
- Ticket ID berformat `SC-00000` auto-increment
- Priority dipilih secara manual (dropdown Low / Normal / High)
- Setelah complaint disubmit, email otomatis terkirim ke nasabah & PIC
- Code ID, SLA, Classification, Reason, PIC → **auto-filled** berdasarkan mapping

