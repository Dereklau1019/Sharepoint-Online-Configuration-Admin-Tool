import * as React from "react";
import ReactJson from "react-json-view";
import { useCallback, useMemo, useState } from "react";
import styles from "./WebpartSetting.module.scss";
import {
  Text,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  TextField,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  DetailsList,
  IColumn,
  Stack,
  IconButton,
  SearchBox,
} from "@fluentui/react";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import {
  EWebPartPropertyRow,
  IWebPartPropertyRow,
  IWebpartSettingProps,
} from "./type";

type MessageType = { type: "info" | "error" | "success"; text: string };

const removeEndWithODataContextProps = (obj: any): any => {
  if (Array.isArray(obj)) {
    return obj.map(removeEndWithODataContextProps);
  } else if (obj && typeof obj === "object") {
    const newObj: any = {};
    for (const [key, value] of Object.entries(obj)) {
      if (!key.endsWith("@odata.context")) {
        newObj[key] = removeEndWithODataContextProps(value);
      }
    }
    return newObj;
  }
  return obj;
};

export const WebpartSetting: React.FC<IWebpartSettingProps> = ({ context }) => {
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState<MessageType | null>(null);
  const [searchKey, setSearchKey] = useState("");
  const [pages, setPages] = useState<IWebPartPropertyRow[]>([]);

  const [siteOptions, setSiteOptions] = useState<IDropdownOption[]>([]);
  const [selectedSiteIds, setSelectedSiteIds] = useState<string[]>([]);

  const [pageOptions, setPageOptions] = useState<IDropdownOption[]>([]);
  const [selectedPageId, setSelectedPageId] = useState("ALL");

  const [rows, setRows] = useState<IWebPartPropertyRow[]>([]);
  const [editedRows, setEditedRows] = useState<IWebPartPropertyRow[]>([]);
  const [replaceFrom, setReplaceFrom] = useState(null);
  const [replaceTo, setReplaceTo] = useState(null);

  /** Util */
  const escapeForRegex = (s: string) =>
    s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const showMessage = (type: MessageType["type"], text: string) =>
    setMessage({ type, text });

  const getGraphClient = async (): Promise<MSGraphClientV3> =>
    await context.msGraphClientFactory.getClient("3");

  const loadAllSites = async () => {
    try {
      setLoading(true);
      const client = await getGraphClient();
      // 這裡用 search=* 搜索所有 site (也可以依需要調整)
      const siteRes = await client.api("/sites?search=*").get();
      console.log("siteRes", siteRes);

      setSiteOptions([
        ...siteRes.value.map((s) => ({
          key: s.id!,
          text: s.webUrl!,
        })),
      ]);
      setEditedRows([]);
      setPageOptions([]);
      setSelectedPageId(null);
    } catch (err) {
      console.error(err);
      showMessage(
        "error",
        `Failed to load Site Pages Webparts: ${String(err)}`
      );
    } finally {
      setLoading(false);
    }
  };

  /** Load SitePages for multiple sites */
  const loadAllSitePages = async () => {
    setLoading(true);

    const stringifySafe = (obj: any) => (obj ? JSON.stringify(obj) : "{}");

    try {
      const client = await getGraphClient();

      // 🔹 同步處理多個 site
      const allItemsPerSite = await Promise.all(
        selectedSiteIds.map(async (siteId) => {
          // 取得 site pages
          const res: any = await client
            .api(`/sites/${siteId}/pages/microsoft.graph.sitePage`)
            .expand("canvasLayout")
            .get();

          console.log("res", res);
          // 取得每個 page 的 webparts
          const webpartResults = await Promise.all(
            res.value.map(async (page: any) => {
              const wpRes: any = await client
                .api(
                  `/sites/${siteId}/pages/${page.id}/microsoft.graph.sitePage/webParts`
                )
                .get();
              console.log("wpRes", wpRes);
              // 回傳整理好的 webpart 資料
              return wpRes.value.map((data: any) => ({
                pageFullDetails: page,
                canvasLayout: page.canvasLayout,
                siteId: siteId,
                pageUrl: page.webUrl,
                pageId: page.id,
                pageName: page.name,
                webpartDetails: data,
                webpartId: data?.id,
                webPartType: data?.webPartType,
                webPartName: data?.title,
                properties: stringifySafe(data?.properties),
                serverProcessedContent: stringifySafe(
                  data?.serverProcessedContent
                ),
                innerHtml: data?.innerHtml,
              }));
            })
          );

          // flatten 每個 site 的所有 webpart
          return webpartResults.flat();
        })
      );

      const allItems = allItemsPerSite.flat();

      setPages(allItems);

      // 建立 page 選單
      const uniquePageOptions = Array.from(
        new Map(allItems.map((p) => [p.pageId, p])).values()
      );

      setPageOptions([
        { key: "ALL", text: "All Pages" },
        ...uniquePageOptions.map((p) => ({
          key: p.pageId!,
          text: p.pageUrl!,
        })),
      ]);

      showMessage("success", `Loaded ${allItems.length} Site Pages Webparts.`);
    } catch (err) {
      console.error(err);
      showMessage(
        "error",
        `Failed to load Site Pages Webparts: ${String(err)}`
      );
    } finally {
      setLoading(false);
    }
  };

  /** 搜尋過濾 */
  const filterArrayColumnByKeyword = useCallback(
    (items: IWebPartPropertyRow[], key: string) => {
      if (!key.trim()) return items;
      const lowerKey = key.toLowerCase();
      return items.filter((item) =>
        Object.values(item).some(
          (val) =>
            val !== undefined &&
            val !== null &&
            val.toString().toLowerCase().includes(lowerKey)
        )
      );
    },
    []
  );

  /** Parse selected page(s) WebPart properties */
  const parsePages = () => {
    const selectedPages =
      selectedPageId === "ALL"
        ? pages
        : pages.filter((p) => p.pageId === selectedPageId);

    setRows(selectedPages);
    setEditedRows([]);
    showMessage(
      "success",
      `Parsed ${selectedPages.length} WebPart properties.`
    );
  };

  const replaceInJson = (obj: any, from: string, to: string): any => {
    if (typeof obj === "string") {
      // ✅ 處理字串型別：進行全域替換
      return obj.replace(new RegExp(escapeForRegex(from), "g"), to);
    }

    if (Array.isArray(obj)) {
      // ✅ 處理陣列：遞迴處理每個元素
      return obj.map((item) => replaceInJson(item, from, to));
    }

    if (obj && typeof obj === "object") {
      // ✅ 處理物件：遞迴處理每個 key 對應的 value
      return Object.fromEntries(
        Object.entries(obj).map(([key, value]) => [
          key,
          replaceInJson(value, from, to),
        ])
      );
    }
    return obj; // ✅ 其他型別 (數字、布林、null、undefined) 不變
  };

  /** Replace values */
  const handleReplace = () => {
    if (!replaceFrom || !replaceTo)
      return showMessage("error", "Please provide replace-from-to value.");

    const filtered = filterArrayColumnByKeyword(rows, searchKey);
    const updated = filtered.map((r) => ({
      ...r,
      properties: replaceInJson(r.properties, replaceFrom, replaceTo),
      serverProcessedContent: replaceInJson(
        r.serverProcessedContent,
        replaceFrom,
        replaceTo
      ),
    }));

    setEditedRows(updated);
    setRows(updated);
    showMessage("success", `Replaced ${updated.length} occurrence(s).`);
  };

  /** Save All */
  const saveAllEditedValue = async () => {
    if (!editedRows.length) return;
    let resultMessage = "";
    try {
      setLoading(true);
      const client = await getGraphClient();

      for (const item of editedRows) {
        try {
          const {
            pageFullDetails,
            canvasLayout,
            webpartId,
            webpartDetails,
            properties,
            serverProcessedContent,
            innerHtml,
          } = item;

          // 1️⃣ 更新 WebPart Body
          const body = {
            ...webpartDetails,
            data: {
              //...webpartDetails.data,
              properties: properties ? JSON.parse(properties) : properties,
              serverProcessedContent: serverProcessedContent
                ? JSON.parse(serverProcessedContent)
                : serverProcessedContent,
            },
            innerHtml: innerHtml,
          };

          // let UpdatedCanvasLayout = canvasLayout;
          // for (const [
          //   sectionIndex,
          //   section,
          // ] of UpdatedCanvasLayout.horizontalSections.entries()) {
          //   for (const [columnIndex, column] of section.columns.entries()) {
          //     for (const [webpartIndex, webpart] of column.webparts.entries()) {
          //       if (webpart.id === webpartId) {
          //         UpdatedCanvasLayout.horizontalSections[sectionIndex].columns[
          //           columnIndex
          //         ].webparts[webpartIndex] = body;
          //       }
          //     }
          //   }
          // }

          // UpdatedCanvasLayout = removeEndWithODataContextProps({
          //   ...UpdatedCanvasLayout,
          //   horizontalSections: UpdatedCanvasLayout?.horizontalSections.filter(
          //     (d) => d.layout !== "flexible"
          //   ),
          // });

          await client
            .api(
              `/sites/${item.siteId}/pages/${item.pageId}/microsoft.graph.sitePage/webparts/${item.webpartId}`
            )
            .headers({ "Content-Type": "application/json" })
            .update(body);

          // let updatedPage = {
          //   ...pageFullDetails,
          //   canvasLayout: UpdatedCanvasLayout,
          // };

          // delete updatedPage["contentType"];
          // delete updatedPage["createdBy"];
          // delete updatedPage["lastModifiedBy"];
          // delete updatedPage["parentReference"];
          // delete updatedPage["publishingState"];
          // delete updatedPage["pageLayout"];
          // delete updatedPage["createdDateTime"];
          // delete updatedPage["lastModifiedDateTime"];
          // delete updatedPage["reactions"];
          // delete updatedPage["eTag"];
          // delete updatedPage["id"];
          // delete updatedPage["name"];
          // delete updatedPage["webUrl"];

          // console.log("removeEndWithODataContextProps", updatedPage);

          // await client
          //   .api(
          //     `/sites/${item.siteId}/pages/${item.pageId}/microsoft.graph.sitePage`
          //   )
          //   .headers({
          //     "Content-Type": "application/json",
          //     Accept: "application/json;odata.metadata=none",
          //   })
          //   .update(updatedPage);

          await client
            .api(
              `/sites/${item.siteId}/pages/${item.pageId}/microsoft.graph.sitePage/publish`
            )
            .post({ comment: "Batch updated webparts." });

          resultMessage += `Edited webpart [${item.webPartName}] saved successfully.`;
          showMessage("success", resultMessage);
        } catch (error) {
          console.error(error);
          resultMessage += `Failed to save webpart ${
            item.webPartName
          } changes: ${error.toString()}`;
          showMessage(
            "error",
            `Failed to save webpart ${
              item.webPartName
            } changes: ${error.toString()}`
          );
        }
      }
      setEditedRows([]);
    } catch (err) {
      console.error(err);
      resultMessage += `Failed to save changes : ${String(err)}`;
      showMessage("error", resultMessage);
    } finally {
      setLoading(false);
    }
  };

  const normalizeJson = (value: string): string => {
    try {
      return JSON.stringify(JSON.parse(value));
    } catch {
      return value; // 不是 JSON → 原樣返回
    }
  };

  const handleFieldChange = (
    key: "properties" | "serverProcessedContent" | "innerHtml",
    newValue: string,
    item: IWebPartPropertyRow
  ) => {
    setRows((prev) =>
      prev.map((r) =>
        r.webpartId === item.webpartId ? { ...r, [key]: newValue } : r
      )
    );

    setEditedRows((prev) => {
      const existing = prev.find((r) => r.webpartId === item.webpartId);

      if (!existing) {
        return [...prev, { ...item, [key]: newValue }];
      }

      const oldValue = normalizeJson(existing[key]);
      const newNorm = normalizeJson(newValue);

      const hasChanged = oldValue !== newNorm;

      if (hasChanged) {
        return prev.map((r) =>
          r.webpartId === item.webpartId ? { ...r, [key]: newValue } : r
        );
      }

      return prev;
    });

    // ✅ 避免太頻繁提示，可以 debounce 或只在變更完成後提示
    showMessage("success", `Edited ${item.webPartName} (${item.webpartId}).`);
  };

  /** 列設定 */
  const columns: IColumn[] = [
    {
      key: "innerHtml",
      name: "InnerHtml",
      fieldName: EWebPartPropertyRow.innerHtml,
      minWidth: 350,
      isMultiline: true,
      onRender: (item) => (
        <TextField
          multiline
          autoAdjustHeight
          value={item.innerHtml || ""}
          onChange={(_, text) =>
            handleFieldChange("innerHtml", text || "", item)
          }
          styles={{
            fieldGroup: {
              fontFamily: "monospace",
              whiteSpace: "pre-wrap", // 保留換行與空格
              backgroundColor: "#f9f9f9",
            },
          }}
        />
      ),
    },
    {
      key: "properties",
      name: "Properties",
      fieldName: EWebPartPropertyRow.properties,
      minWidth: 350,
      isMultiline: true,
      onRender: (item) => (
        <ReactJson
          src={JSON.parse(item.properties ?? "{}")}
          name={false}
          collapsed={true}
          displayDataTypes={false}
          theme="rjv-default"
          onEdit={(edit) =>
            handleFieldChange(
              "properties",
              JSON.stringify(edit.updated_src, null, 2),
              item
            )
          }
          onAdd={(add) =>
            handleFieldChange(
              "properties",
              JSON.stringify(add.updated_src, null, 2),
              item
            )
          }
          onDelete={(del) =>
            handleFieldChange(
              "properties",
              JSON.stringify(del.updated_src, null, 2),
              item
            )
          }
        />
      ),
    },
    {
      key: "serverProcessedContent",
      name: "ServerProcessedContent",
      fieldName: EWebPartPropertyRow.serverProcessedContent,
      minWidth: 250,
      isMultiline: true,
      onRender: (item) => (
        <ReactJson
          src={JSON.parse(item.serverProcessedContent ?? "{}")}
          name={false}
          collapsed={true}
          displayDataTypes={false}
          theme="rjv-default"
          onEdit={(edit) =>
            handleFieldChange(
              "serverProcessedContent",
              JSON.stringify(edit.updated_src, null, 2),
              item
            )
          }
          onAdd={(add) =>
            handleFieldChange(
              "serverProcessedContent",
              JSON.stringify(add.updated_src, null, 2),
              item
            )
          }
          onDelete={(del) =>
            handleFieldChange(
              "serverProcessedContent",
              JSON.stringify(del.updated_src, null, 2),
              item
            )
          }
        />
      ),
    },
    { key: "pageUrl", name: "Page Url", fieldName: "pageUrl", minWidth: 100 },
    { key: "pageId", name: "Page ID", fieldName: "pageId", minWidth: 100 },
    {
      key: "pageName",
      name: "Page Name",
      fieldName: "pageName",
      minWidth: 100,
    },
    {
      key: "webpartId",
      name: "WebPart ID",
      fieldName: "webpartId",
      minWidth: 100,
    },
    {
      key: "webPartName",
      name: "WebPart Name",
      fieldName: "webPartName",
      minWidth: 80,
    },
  ].map((c) => ({ ...c, isResizable: true }));

  /** 搜尋處理 */
  const clearSearch = useCallback(() => setSearchKey(""), []);

  const filtered = useMemo(
    () => filterArrayColumnByKeyword(rows, searchKey),
    [rows, searchKey, filterArrayColumnByKeyword]
  );
  console.log("Edited Rows & Filtered Rows", editedRows, filtered);
  return (
    <Stack
      className={styles.webpartSetting}
      tokens={{ childrenGap: 12, padding: 16 }}
    >
      {message && (
        <MessageBar
          styles={{ root: { zIndex: 10000 } }}
          messageBarType={
            message.type === "error"
              ? MessageBarType.error
              : message.type === "success"
              ? MessageBarType.success
              : MessageBarType.info
          }
          onDismiss={() => setMessage(null)}
        >
          {message.text}
        </MessageBar>
      )}
      {loading && (
        <Spinner
          styles={{
            root: {
              backgroundColor: "white",
              opacity: 0.85,
              position: "absolute",
              inset: 0,
              zIndex: 9999,
            },
          }}
          size={SpinnerSize.large}
          label="Working..."
        />
      )}
      <Stack
        horizontal
        tokens={{ childrenGap: 12 }}
        verticalAlign="end"
        styles={{ root: { position: "relative" } }}
      >
        <PrimaryButton
          text="Load Site Pages"
          onClick={loadAllSites}
          disabled={loading}
        />
        <Dropdown
          disabled={loading || siteOptions.length < 1}
          styles={{
            root: { flex: 1 },
            dropdown: { minWidth: 380, maxWidth: "100%" },
          }}
          options={siteOptions}
          selectedKeys={selectedSiteIds}
          onChange={(_, opt) => {
            if (!opt) return;

            setSelectedSiteIds((prev) => {
              if (opt.selected) {
                return [...prev, opt.key as string];
              } else {
                return prev.filter((k) => k !== opt.key);
              }
            });
          }}
          placeholder="Select a or more sites"
          multiSelect
        />
      </Stack>
      <Stack
        horizontal
        tokens={{ childrenGap: 12 }}
        verticalAlign="end"
        styles={{ root: { position: "relative" } }}
      >
        <PrimaryButton
          text="Load Site Pages"
          onClick={loadAllSitePages}
          disabled={loading || selectedSiteIds.length < 1}
        />
        <Dropdown
          disabled={loading || selectedSiteIds.length < 1}
          styles={{ root: { flex: 1 }, dropdown: { minWidth: 380 } }}
          options={pageOptions}
          selectedKey={selectedPageId}
          onChange={(_, opt) => setSelectedPageId(opt?.key as string)}
          placeholder="All or choose page"
        />
        <PrimaryButton
          disabled={
            loading || selectedSiteIds.length < 1 || !pageOptions.length
          }
          text="Parse Page(s)"
          onClick={parsePages}
        />
        <PrimaryButton
          disabled={
            loading || selectedSiteIds.length < 1 || !pageOptions.length
          }
          text="Save All Change"
          onClick={saveAllEditedValue}
        />
      </Stack>
      <Stack
        horizontal
        horizontalAlign="space-between"
        tokens={{ childrenGap: 16 }}
      >
        <Text variant="xLarge" styles={{ root: { width: "auto" } }}>
          Keyword Search:
        </Text>
        <SearchBox
          value={searchKey}
          placeholder="Input Key word ..."
          onChange={(_, v) => setSearchKey(v ?? "")}
          onSearch={(v) => setSearchKey(v ?? "")}
          onClear={clearSearch}
          underlined
          styles={{ root: { flex: 1 } }}
        />
      </Stack>

      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <TextField
          styles={{ root: { flex: 1 } }}
          placeholder="Replace From"
          value={replaceFrom}
          onChange={(_, v) => setReplaceFrom(v || null)}
        />
        <TextField
          styles={{ root: { flex: 1 } }}
          placeholder="Replace To"
          value={replaceTo}
          onChange={(_, v) => setReplaceTo(v || null)}
        />
        <PrimaryButton text="Replace All" onClick={handleReplace} />
      </Stack>
      {filtered.length > 0 && (
        <MessageBar messageBarType={MessageBarType.info}>
          {`Current rows: ${filtered.length}`}
        </MessageBar>
      )}
      <DetailsList
        items={filtered}
        columns={columns}
        setKey="webpartProps"
        selectionMode={0}
        styles={{ root: { maxHeight: 420, overflowY: "auto" } }}
      />
    </Stack>
  );
};

export default WebpartSetting;
