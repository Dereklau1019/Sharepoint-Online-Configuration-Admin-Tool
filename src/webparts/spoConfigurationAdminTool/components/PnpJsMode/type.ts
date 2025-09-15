import { IDropdownOption } from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ISpoConfigurationAdminToolWebPartProps } from "../../SpoConfigurationAdminToolWebPart";

export interface IPnpJsTesterProps {
  context: WebPartContext;
  properties: ISpoConfigurationAdminToolWebPartProps;
}

// 定義不同類別的 PnP JS 功能
export const pnpCategories: IDropdownOption[] = [
  { key: "web", text: "網站 (Web)" },
  { key: "lists", text: "列表 (Lists)" },
  { key: "items", text: "項目 (Items)" },
  { key: "fields", text: "欄位 (Fields)" },
  { key: "folders", text: "資料夾 (Folders)" },
  { key: "files", text: "檔案 (Files)" },
  { key: "users", text: "使用者 (Users)" },
  { key: "groups", text: "群組 (Groups)" },
  { key: "search", text: "搜尋 (Search)" },
  { key: "profiles", text: "使用者設定檔 (Profiles)" },
];

// 每個類別的具體功能選項
export const pnpFunctions: Record<string, IDropdownOption[]> = {
  web: [
    { key: "web.get", text: "取得網站資訊 (sp.web.get())" },
    { key: "web.webs.get", text: "取得子網站 (sp.web.webs.get())" },
    { key: "web.lists.get", text: "取得所有列表 (sp.web.lists.get())" },
    {
      key: "web.currentUser.get",
      text: "取得當前使用者 (sp.web.currentUser.get())",
    },
    {
      key: "web.siteUsers.get",
      text: "取得網站使用者 (sp.web.siteUsers.get())",
    },
    {
      key: "web.roleDefinitions.get",
      text: "取得權限等級 (sp.web.roleDefinitions.get())",
    },
  ],
  lists: [
    {
      key: "lists.getByTitle",
      text: "依標題取得列表 (sp.web.lists.getByTitle(title))",
    },
    { key: "lists.getById", text: "依 ID 取得列表 (sp.web.lists.getById(id))" },
    {
      key: "lists.ensureList",
      text: "確保列表存在 (sp.web.lists.ensure(title, template))",
    },
    {
      key: "list.fields.get",
      text: "取得列表欄位 (sp.web.lists.getByTitle(title).fields.get())",
    },
    {
      key: "list.items.get",
      text: "取得列表項目 (sp.web.lists.getByTitle(title).items.get())",
    },
    {
      key: "list.views.get",
      text: "取得列表檢視 (sp.web.lists.getByTitle(title).views.get())",
    },
  ],
  items: [
    {
      key: "items.getById",
      text: "依 ID 取得項目 (sp.web.lists.getByTitle(title).items.getById(id))",
    },
    {
      key: "items.add",
      text: "新增項目 (sp.web.lists.getByTitle(title).items.add(data))",
    },
    {
      key: "items.update",
      text: "更新項目 (sp.web.lists.getByTitle(title).items.getById(id).update(data))",
    },
    {
      key: "items.delete",
      text: "刪除項目 (sp.web.lists.getByTitle(title).items.getById(id).delete())",
    },
    {
      key: "items.filter",
      text: "篩選項目 (sp.web.lists.getByTitle(title).items.filter(filter).get())",
    },
    {
      key: "items.orderBy",
      text: "排序項目 (sp.web.lists.getByTitle(title).items.orderBy(field).get())",
    },
  ],
  fields: [
    {
      key: "fields.get",
      text: "取得所有欄位 (sp.web.lists.getByTitle(title).fields.get())",
    },
    {
      key: "fields.getByTitle",
      text: "依標題取得欄位 (sp.web.lists.getByTitle(title).fields.getByTitle(fieldTitle))",
    },
    {
      key: "fields.getById",
      text: "依 ID 取得欄位 (sp.web.lists.getByTitle(title).fields.getById(id))",
    },
    {
      key: "fields.add",
      text: "新增欄位 (sp.web.lists.getByTitle(title).fields.add(fieldInfo))",
    },
    {
      key: "fields.update",
      text: "更新欄位 (sp.web.lists.getByTitle(title).fields.getByTitle(fieldTitle).update(data))",
    },
  ],
  folders: [
    {
      key: "folders.get",
      text: "取得資料夾 (sp.web.getFolderByServerRelativeUrl(url))",
    },
    { key: "folders.add", text: "新增資料夾 (sp.web.folders.add(name))" },
    {
      key: "folders.getByServerRelativeUrl",
      text: "依 URL 取得資料夾 (sp.web.getFolderByServerRelativeUrl(url))",
    },
    {
      key: "folder.files.get",
      text: "取得資料夾檔案 (sp.web.getFolderByServerRelativeUrl(url).files.get())",
    },
    {
      key: "folder.folders.get",
      text: "取得子資料夾 (sp.web.getFolderByServerRelativeUrl(url).folders.get())",
    },
  ],
  files: [
    {
      key: "files.get",
      text: "取得檔案 (sp.web.getFileByServerRelativeUrl(url))",
    },
    {
      key: "files.getByServerRelativeUrl",
      text: "依 URL 取得檔案 (sp.web.getFileByServerRelativeUrl(url))",
    },
    {
      key: "file.getText",
      text: "取得檔案文字內容 (sp.web.getFileByServerRelativeUrl(url).getText())",
    },
    {
      key: "file.getBuffer",
      text: "取得檔案緩衝區 (sp.web.getFileByServerRelativeUrl(url).getBuffer())",
    },
    {
      key: "files.add",
      text: "上傳檔案 (sp.web.getFolderByServerRelativeUrl(folderUrl).files.add(name, content))",
    },
  ],
  users: [
    { key: "users.get", text: "取得所有使用者 (sp.web.siteUsers.get())" },
    {
      key: "users.getById",
      text: "依 ID 取得使用者 (sp.web.siteUsers.getById(id))",
    },
    {
      key: "users.getByEmail",
      text: "依電子郵件取得使用者 (sp.web.siteUsers.getByEmail(email))",
    },
    {
      key: "users.getByLoginName",
      text: "依登入名稱取得使用者 (sp.web.siteUsers.getByLoginName(loginName))",
    },
    {
      key: "users.add",
      text: "新增使用者 (sp.web.siteUsers.add(loginName, email, title))",
    },
  ],
  groups: [
    { key: "groups.get", text: "取得所有群組 (sp.web.siteGroups.get())" },
    {
      key: "groups.getById",
      text: "依 ID 取得群組 (sp.web.siteGroups.getById(id))",
    },
    {
      key: "groups.getByName",
      text: "依名稱取得群組 (sp.web.siteGroups.getByName(name))",
    },
    {
      key: "group.users.get",
      text: "取得群組使用者 (sp.web.siteGroups.getById(id).users.get())",
    },
    { key: "groups.add", text: "新增群組 (sp.web.siteGroups.add(title))" },
  ],
  search: [
    { key: "search.query", text: "執行搜尋查詢 (sp.search(query))" },
    { key: "search.suggest", text: "取得搜尋建議 (sp.searchSuggest(query))" },
    { key: "search.peopleQuery", text: "搜尋人員 (sp.searchPeople(query))" },
  ],
  profiles: [
    {
      key: "profiles.myProperties",
      text: "取得我的屬性 (sp.profiles.myProperties.get())",
    },
    {
      key: "profiles.getPropertiesFor",
      text: "取得指定使用者屬性 (sp.profiles.getPropertiesFor(accountName))",
    },
    {
      key: "profiles.editProfile",
      text: "編輯使用者屬性 (sp.profiles.editProfile(properties))",
    },
  ],
};
