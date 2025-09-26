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
  SearchBox,
} from "@fluentui/react";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import {
  EWebPartPropertyRow,
  IWebPartPropertyRow,
  IWebpartSettingProps,
} from "./type";

type MessageType = { type: "info" | "error" | "success"; text: string };

export const WebpartSetting: React.FC<IWebpartSettingProps> = ({ context }) => {
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState<MessageType | null>(null);
  const [searchKey, setSearchKey] = useState("");

  const [sites, setSites] = useState<IDropdownOption[]>([]);
  const [selectedSiteIds, setSelectedSiteIds] = useState<string[]>([]);
  const [pages, setPages] = useState<IWebPartPropertyRow[]>([]);
  const [pageOptions, setPageOptions] = useState<IDropdownOption[]>([]);
  const [selectedPageId, setSelectedPageId] = useState("ALL");
  const [rows, setRows] = useState<IWebPartPropertyRow[]>([]);
  const [editedRows, setEditedRows] = useState<IWebPartPropertyRow[]>([]);
  const [replaceFrom, setReplaceFrom] = useState<string | null>(null);
  const [replaceTo, setReplaceTo] = useState<string | null>(null);

  /** Util */
  const escapeForRegex = (s: string) =>
    s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const showMessage = (type: MessageType["type"], text: string) =>
    setMessage({ type, text });

  const getGraphClient = async (): Promise<MSGraphClientV3> =>
    await context.msGraphClientFactory.getClient("3");

  /** Load Sites */
  const loadAllSites = async () => {
    setLoading(true);
    try {
      const client = await getGraphClient();

      // 這裡可改成實際邏輯，例如呼叫 Graph API 取多個 site
      // 範例：目前先只取當前 site
      const sitePath = context.pageContext.site.serverRelativeUrl;
      const site = await client
        .api(`/sites/${window.location.hostname}:${sitePath}`)
        .get();

      setSites([{ key: site.id, text: site.name || site.webUrl }]);
      showMessage("success", `Loaded site: ${site.name || site.webUrl}`);
    } catch (err) {
      console.error(err);
      showMessage("error", `Failed to load sites: ${String(err)}`);
    } finally {
      setLoading(false);
    }
  };

  /** Load SitePages for multiple sites */
  const loadAllSitePages = async () => {
    if (!selectedSiteIds.length) {
      return showMessage("error", "Please select at least one site.");
    }
    setLoading(true);
    try {
      const client = await getGraphClient();
      let allItems: IWebPartPropertyRow[] = [];

      for (const siteId of selectedSiteIds) {
        const res: any = await client
          .api(`/sites/${siteId}/pages/microsoft.graph.sitePage`)
          .expand("canvasLayout")
          .get();

        const promises = res.value.map(async (page: any) => {
          const webpartGraph = `/sites/${siteId}/pages/${page.id}/microsoft.graph.sitePage/webParts`;
          const wpRes: any = await client.api(webpartGraph).get();
          return {
            siteId,
            pageUrl: page.webUrl,
            pageId: page.id,
            pageName: page.name,
            value: wpRes.value,
          };
        });

        const webpartResults = await Promise.all(promises);

        const siteItems: IWebPartPropertyRow[] = webpartResults.flatMap((wp) =>
          wp.value.map((data: any) => ({
            siteId: wp.siteId,
            pageUrl: wp.pageUrl,
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

        allItems = [...allItems, ...siteItems];
      }

      setPages(allItems);

      setPageOptions([
        { key: "ALL", text: "All Pages" },
        ...allItems.map((p) => ({ key: p.pageId!, text: p.pageUrl! })),
      ]);

      showMessage("success", `Loaded ${allItems.length} Site Pages Webparts.`);
    } catch (err) {
      console.error(err);
      showMessage("error", `Failed to load Site Pages: ${String(err)}`);
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

  /** Replace values */
  const replaceInJson = (obj: any, from: string, to: string): any => {
    if (typeof obj === "string") {
      return obj.replace(new RegExp(escapeForRegex(from), "g"), to);
    }
    if (Array.isArray(obj)) {
      return obj.map((item) => replaceInJson(item, from, to));
    }
    if (obj && typeof obj === "object") {
      return Object.fromEntries(
        Object.entries(obj).map(([key, value]) => [
          key,
          replaceInJson(value, from, to),
        ])
      );
    }
    return obj;
  };

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

  /** Save All (multi-site) */
  const saveAllEditedValue = async () => {
    if (!editedRows.length) return;
    try {
      setLoading(true);
      const client = await getGraphClient();

      for (const item of editedRows) {
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
            `/sites/${item.siteId}/pages/${item.pageId}/microsoft.graph.sitePage/webparts/${item.webpartId}`
          )
          .headers({ "Content-Type": "application/json" })
          .update(body);

        await client
          .api(
            `/sites/${item.siteId}/pages/${item.pageId}/microsoft.graph.sitePage/publish`
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
      if (!existing) {
        return [...prev, { ...item, [key]: newValue }];
      }
      const hasChanged =
        existing[key] !== newValue ||
        existing.properties !== item.properties ||
        existing.serverProcessedContent !== item.serverProcessedContent;
      if (hasChanged) {
        return prev.map((r) =>
          r.webpartId === item.webpartId ? { ...r, [key]: newValue } : r
        );
      }
      return prev;
    });

    showMessage("success", `Edited ${item.webpartName} (${item.webpartId}).`);
  };

  /** 列設定 */
  const columns: IColumn[] = [
    { key: "siteId", name: "Site ID", fieldName: "siteId", minWidth: 100 },
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
  ].map((c) => ({ ...c, isResizable: true }));

  const clearSearch = useCallback(() => setSearchKey(""), []);

  const filtered = useMemo(
    () => filterArrayColumnByKeyword(rows, searchKey),
    [rows, searchKey, filterArrayColumnByKeyword]
  );

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
      <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="end">
        <PrimaryButton
          text="Load All Site"
          onClick={loadAllSites}
          disabled={loading}
        />
        <Dropdown
          styles={{ root: { flex: 1 }, dropdown: { minWidth: 380 } }}
          options={sites}
          selectedKeys={selectedSiteIds}
          onChange={(_, opt) => {
            if (opt?.selected) {
              setSelectedSiteIds((prev) => [...prev, opt.key as string]);
            } else {
              setSelectedSiteIds((prev) => prev.filter((k) => k !== opt.key));
            }
          }}
          placeholder="All or single site"
          multiSelect
        />
        <PrimaryButton text="Load Pages" onClick={loadAllSitePages} />
      </Stack>

      <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="end">
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
          disabled={!pageOptions.length}
        />
        <PrimaryButton
          text="Save All Change"
          onClick={saveAllEditedValue}
          disabled={!pageOptions.length}
        />
      </Stack>

      <Stack
        horizontal
        horizontalAlign="space-between"
        tokens={{ childrenGap: 16 }}
      >
        <Text variant="xLarge">Keyword Search:</Text>
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
          value={replaceFrom ?? ""}
          onChange={(_, v) => setReplaceFrom(v || null)}
        />
        <TextField
          styles={{ root: { flex: 1 } }}
          placeholder="Replace To"
          value={replaceTo ?? ""}
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
