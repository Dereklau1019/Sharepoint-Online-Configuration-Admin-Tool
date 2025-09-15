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
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { EWebPartPropertyRow, IWebPartPropertyRow } from "./type";

export interface IWebpartSettingProps {
  context: WebPartContext;
}

type MessageType = { type: "info" | "error" | "success"; text: string };

export const WebpartSetting: React.FC<IWebpartSettingProps> = ({ context }) => {
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState<MessageType | null>(null);
  const [searchKey, setSearchKey] = useState("");
  const [siteId, setSiteId] = useState("");
  const [pages, setPages] = useState<IWebPartPropertyRow[]>([]);
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

  /** Load SitePages */
  const loadAllSitePages = async () => {
    setLoading(true);
    try {
      const client = await getGraphClient();
      const sitePath = context.pageContext.site.serverRelativeUrl;
      const site = await client
        .api(`/sites/${window.location.hostname}:${sitePath}`)
        .get();
      setSiteId(site.id);

      const res: any = await client
        .api(`/sites/${site.id}/pages/microsoft.graph.sitePage`)
        .expand("canvasLayout")
        .get();

      const promises = res.value.map(async (page: any) => {
        const webpartGraph = `/sites/${site.id}/pages/${page.id}/microsoft.graph.sitePage/webParts`;
        const wpRes: any = await client.api(webpartGraph).get();
        return {
          pageId: page.id,
          pageName: page.name,
          value: wpRes.value,
        };
      });

      const webpartResults = await Promise.all(promises);
      const allItems: IWebPartPropertyRow[] = webpartResults.flatMap((wp) =>
        wp.value.map((data: any) => ({
          pageId: wp.pageId,
          pageName: wp.pageName,
          webpartDetails: data,
          webpartId: data.id,
          webPartType: data.webPartType,
          webpartName: data?.data?.title,
          properties: JSON.stringify(data?.data?.properties),
          serverProcessedContent: JSON.stringify(
            data?.data?.serverProcessedContent
          ),
          innerHtml: data?.innerHtml,
          isEditing: false,
        }))
      );

      setPages(allItems);

      // unique Page 下拉選單
      const uniquePages = Array.from(
        new Map(allItems.map((p) => [p.pageId, p])).values()
      );
      setPageOptions([
        { key: "ALL", text: "All Pages" },
        ...uniquePages.map((p) => ({ key: p.pageId!, text: p.pageName! })),
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
    showMessage(
      "success",
      `Parsed ${selectedPages.length} WebPart properties.`
    );
  };

  const replaceInJson = (obj: any, from: string, to: string): any => {
    // ✅ 處理字串型別：進行全域替換
    if (typeof obj === "string") {
      return obj.replace(new RegExp(escapeForRegex(from), "g"), to);
    }

    // ✅ 處理陣列：遞迴處理每個元素
    if (Array.isArray(obj)) {
      return obj.map((item) => replaceInJson(item, from, to));
    }

    // ✅ 處理物件：遞迴處理每個 key 對應的 value
    if (obj && typeof obj === "object") {
      return Object.fromEntries(
        Object.entries(obj).map(([key, value]) => [
          key,
          replaceInJson(value, from, to),
        ])
      );
    }

    // ✅ 其他型別 (數字、布林、null、undefined) 不變
    return obj;
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

    try {
      setLoading(true);
      const client = await getGraphClient();

      for (const item of editedRows) {
        // 1️⃣ 更新 WebPart
        const body = {
          ...item.webpartDetails,
          data: {
            ...item.webpartDetails.data,
            properties: item.properties
              ? JSON.parse(item.properties)
              : item.properties,
            serverProcessedContent: item.serverProcessedContent
              ? JSON.parse(item.serverProcessedContent)
              : item.serverProcessedContent,
          },
          innerHtml: item.innerHtml,
        };

        await client
          .api(
            `/sites/${siteId}/pages/${item.pageId}/microsoft.graph.sitePage/webparts/${item.webpartId}`
          )
          .headers({ "Content-Type": "application/json" })
          .update(body);

        await client
          .api(
            `/sites/${siteId}/pages/${item.pageId}/microsoft.graph.sitePage/publish`
          )
          .post({ comment: "Batch updated webparts." });
      }

      setEditedRows([]);
      showMessage("success", "All edited values saved successfully.");
    } catch (err) {
      console.error(err);
      showMessage("error", `Failed to save changes: ${String(err)}`);
    } finally {
      setLoading(false);
    }
  };

  /** Input / JSON 編輯變更處理 */
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

      // 如果不存在 → 新增
      if (!existing) {
        return [...prev, { ...item, [key]: newValue }];
      }

      // 如果存在 → 檢查是否有變化
      const hasChanged =
        existing[key] !== newValue ||
        existing.properties !== item.properties ||
        existing.serverProcessedContent !== item.serverProcessedContent;

      if (hasChanged) {
        return prev.map((r) =>
          r.webpartId === item.webpartId ? { ...r, [key]: newValue } : r
        );
      }

      // 沒有變化 → 不更新
      return prev;
    });

    showMessage("success", `Edited ${item.webpartName} (${item.webpartId}).`);
  };

  /** 列設定 */
  const columns: IColumn[] = [
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
      key: "webpartName",
      name: "WebPart Name",
      fieldName: "webpartName",
      minWidth: 80,
    },
    {
      key: "innerHtml",
      name: "InnerHtml",
      fieldName: EWebPartPropertyRow.innerHtml,
      minWidth: 250,
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
      minWidth: 250,
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
  ].map((c) => ({ ...c, isResizable: true }));

  /** 搜尋處理 */
  const clearSearch = useCallback(() => setSearchKey(""), []);

  const filtered = useMemo(
    () => filterArrayColumnByKeyword(rows, searchKey),
    [rows, searchKey, filterArrayColumnByKeyword]
  );
  console.log("Filtered Rows", filtered);
  console.log("Edited Rows", editedRows);
  return (
    <Stack
      className={styles.webpartSetting}
      tokens={{ childrenGap: 12, padding: 16 }}
    >
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
      {message && (
        <MessageBar
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

      <Stack
        horizontal
        tokens={{ childrenGap: 12 }}
        verticalAlign="end"
        styles={{ root: { position: "relative" } }}
      >
        <PrimaryButton
          text="Load Site Pages"
          onClick={loadAllSitePages}
          disabled={loading}
        />
        <Dropdown
          styles={{ root: { flex: 1 }, dropdown: { minWidth: 380 } }}
          options={pageOptions}
          selectedKey={selectedPageId}
          onChange={(_, opt) => setSelectedPageId(opt?.key as string)}
          placeholder="All or choose page"
        />
        <PrimaryButton
          text="Parse Page(s)"
          onClick={parsePages}
          disabled={loading || !pageOptions.length}
        />
        <PrimaryButton
          text="Save All Change"
          onClick={saveAllEditedValue}
          disabled={loading || !pageOptions.length}
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
