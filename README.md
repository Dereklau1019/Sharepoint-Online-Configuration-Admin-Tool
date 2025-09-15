# SharePoint Online Configuration Admin Tool

**SharePoint Online Configuration Admin Tool** 是一個用於管理 SharePoint Online 的管理工具，可透過 Graph API 或 PnPjs 對網站、列表、項目及 Web Part 屬性進行測試與操作。

---

## 功能

### 1. Graph API 測試

- 支援對 SharePoint Online 的各類資源進行 **CRUD 操作**：
  - **Create**: 建立網站、列表或項目
  - **Read**: 查詢網站資訊、列表、項目、使用者或群組
  - **Update**: 更新列表設定、項目內容、網站屬性
  - **Delete**: 刪除列表、項目或網站資源
- 可快速測試 Graph API 並回傳 JSON 結果。

### 2. PnPjs @3

- 使用 **PnPjs v3** 操作 SharePoint 資源，支持：
  - 網站 (Web)
  - 列表 (Lists)
  - 項目 (Items)
  - 欄位 (Fields)
  - 資料夾 (Folders)
  - 檔案 (Files)
- 提供完整 **CRUD 功能**：
  - Create / Add
  - Read / Get
  - Update
  - Delete

### 3. 修改 Site Page Web Part 屬性

- 可修改現有 Site Page 上 Web Part 的 **屬性設定**。
- 支援動態更新 Web Part 的內容或設定，方便測試與管理。

---

## 安裝與設定

### 前置條件

1. Node.js >= 18.x
2. npm
3. SharePoint Online Site Collection Admin 權限
4. 可存取 Microsoft Graph API 的帳號
5. 可存取 PnpJs 的帳號

### 安裝

```bash
git clone https://github.com/your-repo/SharePoint-Config-AdminTool.git
cd SharePoint-Config-AdminTool
npm install
npm run serve
```

### 授權

- MIT License
