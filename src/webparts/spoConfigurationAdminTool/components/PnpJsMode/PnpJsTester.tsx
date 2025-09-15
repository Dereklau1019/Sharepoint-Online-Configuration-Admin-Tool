import * as React from "react";
import { useState, useEffect } from "react";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/site-users";
import "@pnp/sp/site-groups";
import "@pnp/sp/search";
import "@pnp/sp/profiles";
import { IListUpdateResult } from "@pnp/sp/lists";
import {
  Text,
  TextField,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  DefaultButton,
  IconButton,
  Stack,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Pivot,
  PivotItem,
  Label,
} from "@fluentui/react";
import { IPnpJsTesterProps, pnpCategories, pnpFunctions } from "./type";
import EditableJsonList from "./EditableJsonList";

const PnpJsTester: React.FC<IPnpJsTesterProps> = ({ context, properties }) => {
  const [category, setCategory] = useState<
    | "web"
    | "lists"
    | "items"
    | "fields"
    | "folders"
    | "files"
    | "users"
    | "groups"
    | "search"
    | "profiles"
  >("web");
  const [selectedFunction, setSelectedFunction] = useState<string>("");
  const [parameters, setParameters] = useState<
    { key: string; value: string }[]
  >([{ key: "", value: "" }]);
  const [queryOptions, setQueryOptions] = useState("");
  const [response, setResponse] = useState<any>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [responseMode, setResponseMode] = useState<"json" | "raw" | "editJson">(
    "json"
  );
  const [exportSuccess, setExportSuccess] = useState(false);
  const [sp, setSp] = useState<any>(null);

  // 新增：用於管理編輯狀態
  const [editedData, setEditedData] = useState<string>("");
  const [hasChanges, setHasChanges] = useState(false);
  const [saveMessage, setSaveMessage] = useState<{
    type: MessageBarType;
    text: string;
  } | null>(null);

  // 初始化 PnP JS
  useEffect(() => {
    const spInstance = spfi().using(SPFx(context));
    setSp(spInstance);
  }, [context]);
  // 保存列表項目
  const saveListItems = async (updatedData: any) => {
    const listTitle = parameters.find((p) => p.key === "listTitle")?.value;
    if (!listTitle) throw new Error("找不到列表標題");

    if (Array.isArray(updatedData)) {
      // 批量更新多個項目
      for (const item of updatedData) {
        if (item.Id) {
          const { Id, ...updateData } = item;
          await sp.web.lists
            .getByTitle(listTitle)
            .items.getById(Id)
            .update(updateData);
        }
      }
    } else if (updatedData.Id) {
      // 單個項目更新
      const { Id, ...updateData } = updatedData;
      await sp.web.lists
        .getByTitle(listTitle)
        .items.getById(Id)
        .update(updateData);
    }
  };

  // 保存列表資料（如列表設定）所有資料
  const saveListData = async (updatedData: any) => {
    const listTitle = parameters.find(
      (p) => p.key === "title" || p.key === "listTitle"
    )?.value;
    if (!listTitle) throw new Error("找不到列表標題");

    // 更新列表本身的設定
    if (updatedData.Title) {
      const listTitle = updatedData.Title;
      const updateData = Object.fromEntries(
        properties.updatableListProperties.map((str) => [str, updatedData[str]])
      );
      await sp.web.lists.getByTitle(listTitle).update(updateData);
    }
  };

  // 保存使用者資料
  const saveUserData = async (updatedData: any) => {
    // SharePoint 使用者資料通常是唯讀的
    // 這裡主要用於記錄變更
    console.log("使用者資料變更 (唯讀):", updatedData);
  };

  // 保存群組資料
  const saveGroupData = async (updatedData: any) => {
    if (updatedData.Id) {
      const { Id, ...updateData } = updatedData;
      if (updateData.Title || updateData.Description) {
        await sp.web.siteGroups.getById(Id).update(updateData);
      }
    }
  };

  const handleParameterChange = (
    index: number,
    field: "key" | "value",
    value: string
  ) => {
    const updated = [...parameters];
    updated[index][field] = value;
    setParameters(updated);
  };

  const addParameter = () => {
    setParameters([...parameters, { key: "", value: "" }]);
  };

  const removeParameter = (index: number) => {
    if (parameters.length > 1) {
      setParameters(parameters.filter((_, idx) => idx !== index));
    }
  };

  const buildPnpQuery = () => {
    if (!sp || !selectedFunction) return null;

    const params = parameters.filter((p) => p.key.trim() && p.value.trim());
    const options = queryOptions.trim();

    try {
      let query = sp.web;
      const parts = selectedFunction.split(".");

      // 根據選擇的功能建構查詢
      switch (category) {
        case "web":
          if (selectedFunction === "web.get") {
            query = sp.web();
          } else if (selectedFunction === "web.webs.get") {
            query = sp.web.webs();
          } else if (selectedFunction === "web.lists.get") {
            query = sp.web.lists();
          } else if (selectedFunction === "web.currentUser.get") {
            query = sp.web.currentUser();
          } else if (selectedFunction === "web.siteUsers.get") {
            query = sp.web.siteUsers();
          } else if (selectedFunction === "web.roleDefinitions.get") {
            query = sp.web.roleDefinitions();
          }
          break;

        case "lists":
          if (selectedFunction === "lists.getByTitle") {
            const title = params.find((p) => p.key === "title")?.value || "";
            if (!title) throw new Error("需要提供 title 參數");
            query = sp.web.lists.getByTitle(title)();
          } else if (selectedFunction === "lists.getById") {
            const id = params.find((p) => p.key === "id")?.value || "";
            if (!id) throw new Error("需要提供 id 參數");
            query = sp.web.lists.getById(id)();
          } else if (selectedFunction === "list.fields.get") {
            const title =
              params.find((p) => p.key === "listTitle")?.value || "";
            if (!title) throw new Error("需要提供 listTitle 參數");
            query = sp.web.lists.getByTitle(title).fields();
          } else if (selectedFunction === "list.items.get") {
            const title =
              params.find((p) => p.key === "listTitle")?.value || "";
            if (!title) throw new Error("需要提供 listTitle 參數");
            query = sp.web.lists.getByTitle(title).items();
          } else if (selectedFunction === "list.views.get") {
            const title =
              params.find((p) => p.key === "listTitle")?.value || "";
            if (!title) throw new Error("需要提供 listTitle 參數");
            query = sp.web.lists.getByTitle(title).views();
          }
          break;

        case "items": {
          const listTitle =
            params.find((p) => p.key === "listTitle")?.value || "";
          if (!listTitle) throw new Error("需要提供 listTitle 參數");

          if (selectedFunction === "items.getById") {
            const itemId = params.find((p) => p.key === "itemId")?.value || "";
            if (!itemId) throw new Error("需要提供 itemId 參數");
            query = sp.web.lists
              .getByTitle(listTitle)
              .items.getById(parseInt(itemId))();
          } else if (selectedFunction === "items.filter") {
            const filter = params.find((p) => p.key === "filter")?.value || "";
            if (!filter) throw new Error("需要提供 filter 參數");
            query = sp.web.lists.getByTitle(listTitle).items.filter(filter)();
          } else if (selectedFunction === "items.orderBy") {
            const orderBy =
              params.find((p) => p.key === "orderBy")?.value || "";
            const ascending =
              params.find((p) => p.key === "ascending")?.value === "true";
            if (!orderBy) throw new Error("需要提供 orderBy 參數");
            query = sp.web.lists
              .getByTitle(listTitle)
              .items.orderBy(orderBy, ascending)();
          }
          break;
        }
        case "users": {
          if (selectedFunction === "users.get") {
            query = sp.web.siteUsers();
          } else if (selectedFunction === "users.getById") {
            const userId = params.find((p) => p.key === "userId")?.value || "";
            if (!userId) throw new Error("需要提供 userId 參數");
            query = sp.web.siteUsers.getById(parseInt(userId))();
          } else if (selectedFunction === "users.getByEmail") {
            const email = params.find((p) => p.key === "email")?.value || "";
            if (!email) throw new Error("需要提供 email 參數");
            query = sp.web.siteUsers.getByEmail(email)();
          }
          break;
        }
        case "search": {
          if (selectedFunction === "search.query") {
            const searchText =
              params.find((p) => p.key === "querytext")?.value || "";
            if (!searchText) throw new Error("需要提供 querytext 參數");
            query = sp.search({ Querytext: searchText });
          }
          break;
        }
        default:
          throw new Error("不支援的類別");
      }

      return query;
    } catch (error) {
      throw error;
    }
  };

  const executeQuery = async () => {
    if (!selectedFunction) {
      alert("請選擇要測試的功能");
      return;
    }

    setIsLoading(true);
    const start = performance.now();

    try {
      const query = buildPnpQuery();
      if (!query) throw new Error("無法建構查詢");

      const data = await query;

      setResponse({
        status: 200,
        statusText: "OK",
        ok: true,
        time: Math.round(performance.now() - start),
        size: new TextEncoder().encode(JSON.stringify(data)).length,
        data,
        isJson: true,
      });

      // 重置編輯狀態
      setEditedData(JSON.stringify(data, null, 2));
      setHasChanges(false);
      setSaveMessage(null);
    } catch (error: any) {
      console.error("PnP JS Error:", error);
      setResponse({
        error: error.message || "執行查詢時發生錯誤",
        status: 0,
        statusText: "Error",
        isJson: false,
        data: error,
      });
    } finally {
      setIsLoading(false);
    }
  };

  const downloadResponse = () => {
    if (!response?.data) return;
    const blob = new Blob([JSON.stringify(response.data, null, 2)], {
      type: "application/json",
    });

    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `pnpjs-response-${Date.now()}.json`;
    link.click();
    URL.revokeObjectURL(link.href);

    setExportSuccess(true);
    setTimeout(() => setExportSuccess(false), 1500);
  };

  const resetForm = () => {
    setSelectedFunction("");
    setParameters([{ key: "", value: "" }]);
    setQueryOptions("");
    setResponse(null);
    setEditedData("");
    setHasChanges(false);
    setSaveMessage(null);
  };

  // 處理編輯過程中的變更
  const handleEditChange = (changes: string) => {
    setEditedData(changes);
    setHasChanges(true);
    setSaveMessage(null);
  };

  // 處理單筆資料保存
  // const handleSave = (updatedJson: string) => {
  //   console.log("單筆保存:", JSON.parse(updatedJson));
  //   setEditedData(updatedJson);
  //   // 可以在這裡添加個別項目保存邏輯
  // };
  // 根據不同的資料類型執行保存操作
  const performSaveOperation = async (updatedData: any) => {
    if (!sp) throw new Error("PnP JS 尚未初始化");

    // 根據當前選擇的功能決定保存邏輯
    switch (category) {
      case "items":
        await saveListItems(updatedData);
        break;
      case "lists":
        await saveListData(updatedData);
        break;
      case "users":
        await saveUserData(updatedData);
        break;
      case "groups":
        await saveGroupData(updatedData);
        break;
      default:
        // 對於其他類型，只在本地更新
        console.log("本地更新數據:", updatedData);
        break;
    }
  };

  // 主要的保存所有改動功能
  const saveAllChange = async () => {
    console.log("editedData", editedData);
    if (!hasChanges || !editedData) {
      setSaveMessage({
        type: MessageBarType.warning,
        text: "沒有需要保存的變更",
      });
      return;
    }

    setIsLoading(true);
    setSaveMessage(null);

    try {
      // 解析編輯後的數據
      const updatedData = JSON.parse(editedData);
      console.log("updatedData", updatedData);

      // 根據不同的功能類型執行不同的保存邏輯
      await performSaveOperation(updatedData);

      // 更新 response 中的 data
      setResponse((prev) => ({
        ...prev,
        data: updatedData,
      }));

      setHasChanges(false);
      setSaveMessage({
        type: MessageBarType.success,
        text: "所有變更已成功保存！",
      });

      // 3秒後清除訊息
      setTimeout(() => setSaveMessage(null), 3000);
    } catch (error: any) {
      console.error("保存失敗:", error);
      setSaveMessage({
        type: MessageBarType.error,
        text: `保存失敗: ${error.message}`,
      });
    } finally {
      setIsLoading(false);
    }
  };

  // 當類別改變時重置功能選擇
  useEffect(() => {
    setSelectedFunction("");
    setParameters([{ key: "", value: "" }]);
    setResponse(null);
    setEditedData("");
    setHasChanges(false);
    setSaveMessage(null);
  }, [category]);

  // 根據選擇的功能提供參數提示
  const getParameterHints = () => {
    const hints: Record<string, string[]> = {
      // Lists
      "lists.getByTitle": ["title"],
      "lists.getById": ["id"],
      "lists.ensureList": ["title", "description", "template"],
      "list.fields.get": ["listTitle"],
      "list.items.get": ["listTitle"],
      "list.views.get": ["listTitle"],

      // Items
      "items.getById": ["listTitle", "itemId"],
      "items.add": ["listTitle", "itemData (JSON)"],
      "items.update": ["listTitle", "itemId", "itemData (JSON)"],
      "items.delete": ["listTitle", "itemId"],
      "items.filter": ["listTitle", "filter"],
      "items.orderBy": ["listTitle", "orderBy", "ascending (true/false)"],

      // Fields
      "fields.get": ["listTitle (可選)"],
      "fields.getByTitle": ["fieldTitle", "listTitle (可選)"],
      "fields.getById": ["fieldId", "listTitle (可選)"],
      "fields.add": ["fieldData (JSON)", "listTitle (可選)"],
      "fields.update": ["fieldId", "fieldData (JSON)", "listTitle (可選)"],

      // Folders
      "folders.get": ["listTitle"],
      "folders.add": ["listTitle", "folderName"],
      "folders.getByServerRelativeUrl": ["serverRelativeUrl"],
      "folder.files.get": ["serverRelativeUrl"],
      "folder.folders.get": ["serverRelativeUrl"],

      // Files
      "files.get": ["listTitle"],
      "files.getByServerRelativeUrl": ["serverRelativeUrl"],
      "file.getText": ["serverRelativeUrl"],
      "file.getBuffer": ["serverRelativeUrl"],
      "files.add": ["listTitle", "fileName", "fileContent"],

      // Users
      "users.getById": ["userId"],
      "users.getByEmail": ["email"],
      "users.getByLoginName": ["loginName"],
      "users.add": ["loginName"],

      // Groups
      "groups.getById": ["groupId"],
      "groups.getByName": ["groupName"],
      "group.users.get": ["groupId"],
      "groups.add": ["groupData (JSON)"],

      // Search
      "search.query": [
        "querytext",
        "rowlimit (可選)",
        "selectproperties (可選)",
        "sourceid (可選)",
      ],
      "search.suggest": ["querytext"],
      "search.peopleQuery": ["querytext"],

      // Profiles
      "profiles.getPropertiesFor": ["loginName"],
      "profiles.editProfile": [
        "properties (JSON: {accountName, propertyName, propertyValue})",
      ],
    };

    return hints[selectedFunction] || [];
  };
  console.log("editedData", editedData);
  return (
    <Stack
      tokens={{ childrenGap: 20 }}
      styles={{ root: { width: "100%", margin: "auto", padding: 24 } }}
    >
      <Text variant="xLarge">PnP JS v3 測試工具</Text>

      <Pivot>
        <PivotItem headerText="功能測試">
          <Stack tokens={{ childrenGap: 16 }}>
            {/* 類別選擇 */}
            <Stack horizontal tokens={{ childrenGap: 16 }}>
              <Dropdown
                label="選擇類別"
                selectedKey={category}
                options={pnpCategories}
                onChange={(_, option) => setCategory(option?.key as any)}
                styles={{ root: { width: 200 } }}
              />
              <Dropdown
                label="選擇功能"
                selectedKey={selectedFunction}
                options={pnpFunctions[category] || []}
                onChange={(_, option) =>
                  setSelectedFunction(option?.key as string)
                }
                styles={{ root: { flex: 1 } }}
                disabled={!category}
              />
            </Stack>

            {/* 參數設定 */}
            {selectedFunction && (
              <Stack tokens={{ childrenGap: 12 }}>
                <Label>參數設定</Label>
                {getParameterHints().length > 0 && (
                  <MessageBar messageBarType={MessageBarType.info}>
                    建議參數：{getParameterHints().join(", ")}
                  </MessageBar>
                )}

                {parameters.map((param, index) => (
                  <Stack horizontal tokens={{ childrenGap: 8 }} key={index}>
                    <TextField
                      placeholder="參數名稱"
                      value={param.key}
                      onChange={(_, value) =>
                        handleParameterChange(index, "key", value || "")
                      }
                      styles={{ root: { flex: 1 } }}
                    />
                    <TextField
                      placeholder="參數值"
                      value={param.value}
                      onChange={(_, value) =>
                        handleParameterChange(index, "value", value || "")
                      }
                      styles={{ root: { flex: 1 } }}
                    />
                    <IconButton
                      iconProps={{ iconName: "Delete" }}
                      onClick={() => removeParameter(index)}
                      disabled={parameters.length === 1}
                    />
                  </Stack>
                ))}

                <DefaultButton
                  iconProps={{ iconName: "Add" }}
                  text="新增參數"
                  onClick={addParameter}
                />
              </Stack>
            )}

            {/* 查詢選項 */}
            <TextField
              label="額外查詢選項 (可選)"
              placeholder="例如: select, expand, filter 等"
              multiline
              rows={3}
              value={queryOptions}
              onChange={(_, value) => setQueryOptions(value || "")}
            />

            {/* 操作按鈕 */}
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <PrimaryButton
                text="執行查詢"
                onClick={executeQuery}
                disabled={isLoading || !selectedFunction}
              />
              <DefaultButton
                text="重置"
                onClick={resetForm}
                disabled={isLoading}
              />
            </Stack>

            {isLoading && (
              <Spinner label="執行中..." size={SpinnerSize.medium} />
            )}
          </Stack>
        </PivotItem>

        <PivotItem headerText="程式碼範例">
          {selectedFunction && (
            <Stack tokens={{ childrenGap: 12 }}>
              <Label>{"對應的 PnP JS 程式碼："}</Label>
              <pre
                style={{
                  backgroundColor: "#f5f5f5",
                  padding: 12,
                  borderRadius: 4,
                  border: "1px solid #ddd",
                  whiteSpace: "pre-wrap",
                }}
              >
                {`// 初始化 PnP JS
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

const sp = spfi().using(SPFx(this.context));

// 執行查詢
const result = await ${
                  selectedFunction.includes("search")
                    ? "sp.search({ Querytext: 'your search term' })"
                    : `sp.web.${selectedFunction
                        .replace(".", ".")
                        .replace("get", "")}()`
                };

console.log(result);`}
              </pre>
            </Stack>
          )}
        </PivotItem>
      </Pivot>

      {/* 結果顯示 */}
      {response && (
        <Stack tokens={{ childrenGap: 12 }}>
          {/* 保存狀態訊息 */}
          {saveMessage && (
            <MessageBar
              messageBarType={saveMessage.type}
              onDismiss={() => setSaveMessage(null)}
              dismissButtonAriaLabel="關閉"
            >
              {saveMessage.text}
            </MessageBar>
          )}
          <MessageBar
            messageBarType={
              response.ok ? MessageBarType.success : MessageBarType.error
            }
            isMultiline={false}
          >
            {response.ok
              ? `執行成功！耗時：${response.time} ms，資料大小：${response.size} bytes`
              : `執行失敗：${response.error}`}
          </MessageBar>

          {response.ok && (
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <DefaultButton
                text="JSON"
                onClick={() => setResponseMode("json")}
                disabled={!response.isJson}
              />
              <DefaultButton
                text="Raw"
                onClick={() => setResponseMode("raw")}
              />
              <DefaultButton
                text="Edit Mode"
                onClick={() => setResponseMode("editJson")}
              />
              <PrimaryButton
                iconProps={{ iconName: "Download" }}
                text={exportSuccess ? "已匯出" : "下載結果"}
                onClick={downloadResponse}
              />
              {hasChanges ? (
                <PrimaryButton
                  styles={{
                    root: {
                      marginLeft: "auto",
                    },
                  }}
                  iconProps={{ iconName: "Save" }}
                  text={"儲存所有改動 Save All"}
                  onClick={saveAllChange}
                  disabled={!hasChanges || isLoading}
                />
              ) : null}
            </Stack>
          )}

          <Stack
            styles={{
              root: {
                display: responseMode !== "editJson" ? "none" : "flex",
              },
            }}
          >
            {response.data ? (
              <EditableJsonList
                updatableListProperties={properties.updatableListProperties}
                category={category}
                data={editedData || JSON.stringify(response.data, null, 2)}
                //onChange={handleEditChange}
                onSave={handleEditChange} //handleSave}
              />
            ) : null}
          </Stack>

          <pre
            style={{
              display: responseMode === "editJson" ? "none" : "block",
              whiteSpace: "pre-wrap",
              backgroundColor: response.ok ? "#f9f9f9" : "#fff5f5",
              padding: 12,
              borderRadius: 4,
              border: response.ok ? "1px solid #ddd" : "1px solid #ffcccc",
              maxHeight: "500px",
              overflow: "auto",
            }}
          >
            {responseMode === "json" && response.isJson
              ? JSON.stringify(response.data, null, 2)
              : typeof response.data === "string"
              ? response.data
              : JSON.stringify(response.data, null, 2)}
          </pre>
        </Stack>
      )}
    </Stack>
  );
};

export default PnpJsTester;
