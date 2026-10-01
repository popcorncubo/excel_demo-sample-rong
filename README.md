# Excel Inventory Manager

Upload Excel files and instantly manage products, inventory, stock records, and reports directly in your browser.

一個基於瀏覽器運作的 Excel 商品／庫存管理工具。
使用者只需要上傳 Excel，即可快速查看商品資料、庫存狀態、低庫存提醒與基本統計報表。

Tech Stack

Frontend

HTML5

CSS3

JavaScript

Excel Processing

SheetJS / XLSX

Data Storage

LocalStorage

Database

無後端資料庫

所有資料預設儲存在使用者瀏覽器本機

Excel 檔案不會上傳至伺服器

Features

Excel 商品資料匯入

商品搜尋與篩選

庫存數量統計

低庫存提醒

入庫／出庫紀錄

商品分類統計

Excel 報表匯出

管理者／檢視者模式

本機資料保存

Responsive Web Design

Custom Development

此專案為基礎展示版。

如有以下需求，可聯絡進行客製開發：

POS 系統

ERP 系統

SQL 資料庫

MySQL / PostgreSQL

PHP / C# 後端

REST API

多人帳號與權限

雲端資料同步

多倉庫管理

金流串接

AI 智能客服

WordPress 系統

自動化報表

公司內部管理系統

Contact

For custom development or business inquiries:

Email: chinsun1345@gmail.com

This project is intended as a lightweight frontend demo and starting point for customized business systems.

# StockLite — Excel 庫存管理 Demo

這是一套可直接放到 **GitHub Pages** 的純前端展示模板。

## 已完成

- 上傳 `.xlsx` / `.xls`，讀取第一個工作表
- 商品搜尋、類別篩選、庫存狀態篩選
- KPI：品項、總庫存、庫存成本、低庫存
- 本機出入庫 Demo（管理者模式）
- 檢視者 / 管理者的前端權限展示
- 出入庫紀錄
- 類別庫存報表、低庫存報表
- 匯出商品 Excel、紀錄 Excel、完整管理報表
- LocalStorage：同一瀏覽器重新整理後可保留資料
- Excel 不上傳伺服器，全程在瀏覽器本機處理
- 客製開發聯絡按鈕

> 注意：這個「權限」只是 UI Demo，不具真正資安隔離能力。正式多人帳號、權限、雲端同步、資料庫等功能需要後端。

## GitHub Pages 部署

1. 建立 GitHub repository。
2. 把本資料夾內所有檔案放在 repo 根目錄。
3. 到 **Settings → Pages**。
4. Source 選 **Deploy from a branch**。
5. Branch 選 `main` / `/ (root)`。
6. 儲存後即可取得網站網址。

## 修改聯絡信箱

打開 `app.js`，修改最上方：

```js
const CONFIG = {
  contactEmail: "yourname.service@gmail.com",
  storageKey: "stocklite-demo-v1"
};
```

建議使用工作專用 Gmail，不要放私人主要信箱。

## Excel 欄位

最低必須有：

- 商品編號（也接受 SKU / 品號 / 產品編號）
- 商品名稱（也接受 產品名稱 / 名稱 / 品名）

建議完整欄位：

`商品編號 | 商品名稱 | 類別 | 售價 | 成本 | 庫存 | 安全庫存 | 供應商`

可直接使用附帶的 `excel-inventory-demo-sample.xlsx` 測試。
## 可客製

- 真正登入 / RBAC 權限
- MySQL / PostgreSQL / SQL Server
- 多人、多裝置同步
- POS / ERP
- 多倉庫與庫存批次
- API 串接
- 金流
- AI 客服
- 自動備份 / 排程
- 管理後台

## 使用到的前端套件

- SheetJS CE（透過 jsDelivr CDN 載入）

若未連上網路，Excel 匯入與匯出功能可能無法使用；其餘介面仍可顯示。
