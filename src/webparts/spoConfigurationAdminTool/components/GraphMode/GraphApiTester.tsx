import * as React from "react";
import { useState, useEffect } from "react";
import { MSGraphClientV3 } from "@microsoft/sp-http";
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
} from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";

interface IGraphApiTesterProps {
  context: WebPartContext;
}

const graphMethods: IDropdownOption[] = [
  { key: "GET", text: "GET" },
  { key: "POST", text: "POST" },
  { key: "PATCH", text: "PATCH" },
  { key: "DELETE", text: "DELETE" },
];

const GraphApiTester: React.FC<IGraphApiTesterProps> = ({ context }) => {
  const [url, setUrl] = useState("");
  const [method, setMethod] = useState("GET");
  const [headers, setHeaders] = useState([{ key: "", value: "" }]);
  const [body, setBody] = useState("");
  const [response, setResponse] = useState<any>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [responseMode, setResponseMode] = useState<"json" | "raw">("json");
  const [exportSuccess, setExportSuccess] = useState(false);

  const isJsonString = (str: string) => {
    try {
      JSON.parse(str);
      return true;
    } catch {
      return false;
    }
  };

  const handleHeaderChange = (
    index: number,
    field: "key" | "value",
    value: string
  ) => {
    const updated = [...headers];
    updated[index][field] = value;
    setHeaders(updated);
  };

  const sendRequest = async (): Promise<any> => {
    if (!url.trim()) return alert("請輸入 Graph API URL");

    setIsLoading(true);
    const start = performance.now();

    try {
      const client: MSGraphClientV3 =
        await context.msGraphClientFactory.getClient("3");

      // 建立自訂 headers
      const requestHeaders: Record<string, string> = {};
      headers.forEach((h) => {
        if (h.key.trim()) requestHeaders[h.key.trim()] = h.value.trim();
      });

      let request = client.api(url) as any;

      switch (method) {
        case "GET":
          request = request.get();
          break;
        case "POST":
          if (body && !isJsonString(body)) return alert("Body 必須為有效 JSON");
          request = request.post(JSON.parse(body || "{}"));
          break;
        case "PATCH":
          if (body && !isJsonString(body)) return alert("Body 必須為有效 JSON");
          request = request.patch(JSON.parse(body || "{}"));
          break;
        case "DELETE":
          request = request.delete();
          break;
      }

      // Graph Client 會自動使用 Bearer Token，無需額外 headers
      const data = await request;

      setResponse({
        status: 200,
        statusText: "OK",
        ok: true,
        time: Math.round(performance.now() - start),
        size: new TextEncoder().encode(JSON.stringify(data)).length,
        data,
        isJson: true,
      });
    } catch (e: any) {
      setResponse({
        error: e.message,
        status: 0,
        statusText: "Error",
        isJson: false,
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
    link.download = "response.json";
    link.click();
    URL.revokeObjectURL(link.href);

    setExportSuccess(true);
    setTimeout(() => setExportSuccess(false), 1500);
  };

  useEffect(() => {
    if (response && !response.isJson) setResponseMode("json");
  }, [response]);

  return (
    <Stack
      tokens={{ childrenGap: 20 }}
      styles={{ root: { width: "100%", margin: "auto", padding: 24 } }}
    >
      <Text variant="xLarge">{"Graph API 工具(Fluent UI)"}</Text>
      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <TextField
          styles={{ root: { flex: 1 } }}
          placeholder="Graph API URL"
          value={url}
          onChange={(_, v) => setUrl(v || "")}
          multiline
          draggable={false}
          rows={8}
        />
        <Dropdown
          styles={{ root: { width: "380px" } }}
          selectedKey={method}
          options={graphMethods}
          onChange={(_, o) => setMethod(o?.key as string)}
          disabled={!url}
        />
      </Stack>
      <Stack tokens={{ childrenGap: 8 }}>
        <Text variant="xLarge">{"請求標頭（可選）"}</Text>
        {headers.map((h, i) => (
          <Stack horizontal tokens={{ childrenGap: 8 }} key={i}>
            <TextField
              value={h.key}
              styles={{ root: { flex: 1 } }}
              placeholder="Key"
              onChange={(_, v) => handleHeaderChange(i, "key", v || "")}
            />
            <TextField
              value={h.value}
              styles={{ root: { flex: 1 } }}
              placeholder="Value"
              onChange={(_, v) => handleHeaderChange(i, "value", v || "")}
            />
            <IconButton
              iconProps={{ iconName: "Delete" }}
              styles={{ root: { width: "auto" } }}
              onClick={() => setHeaders(headers.filter((_, idx) => idx !== i))}
              disabled={headers.length === 1}
            />
          </Stack>
        ))}
        <DefaultButton
          iconProps={{ iconName: "Add" }}
          text="新增標頭"
          onClick={() => setHeaders([...headers, { key: "", value: "" }])}
        />
      </Stack>

      {(method === "POST" || method === "PATCH") && (
        <TextField
          label="Request Body (JSON)"
          multiline
          rows={6}
          value={body}
          onChange={(_, v) => setBody(v || "")}
        />
      )}

      <PrimaryButton
        text="送出請求"
        onClick={sendRequest}
        disabled={isLoading || !url.trim()}
      />
      {isLoading && <Spinner label="請求中..." size={SpinnerSize.medium} />}

      {response && (
        <Stack tokens={{ childrenGap: 12 }}>
          <MessageBar
            messageBarType={
              response.ok ? MessageBarType.success : MessageBarType.error
            }
            isMultiline={false}
          >
            狀態：{response.status} {response.statusText}，耗時：{response.time}
            ms
          </MessageBar>

          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <DefaultButton
              text="JSON"
              onClick={() => setResponseMode("json")}
              disabled={!response.isJson}
            />
            <DefaultButton text="Raw" onClick={() => setResponseMode("raw")} />
            <DefaultButton
              iconProps={{ iconName: "Download" }}
              text={exportSuccess ? "已匯出" : "匯出結果"}
              onClick={downloadResponse}
            />
          </Stack>

          <pre
            style={{
              whiteSpace: "pre-wrap",
              backgroundColor: "#f9f9f9",
              padding: 12,
              borderRadius: 4,
            }}
          >
            {responseMode === "json" && response.isJson
              ? JSON.stringify(response.data, null, 2)
              : typeof response.data === "string"
              ? response.data
              : JSON.stringify(response.data)}
          </pre>
        </Stack>
      )}
    </Stack>
  );
};

export default GraphApiTester;
