import * as React from "react";
import styles from "./App.module.scss";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { DefaultButton, PrimaryButton, Stack } from "@fluentui/react";
import GraphApiTester from "./GraphMode/GraphApiTester";
import { useState } from "react";
import PnpJsTester from "./PnpJsMode/PnpJsTester";
import WebpartSetting from "./WebpartMode/WebpartSetting";

export interface IAppProps {
  context: WebPartContext;
}

const App: React.FC<IAppProps> = ({ context }) => {
  const [mode, setMode] = useState<"graph" | "pnpJs" | "webpart">("graph");

  const modes = [
    { key: "graph", label: "Graph Mode" },
    { key: "pnpJs", label: "PnPjs Mode" },
    { key: "webpart", label: "Webpart Mode" },
  ];
  return (
    <Stack className={styles.app} tokens={{ childrenGap: 12 }}>
      <Stack horizontal tokens={{ childrenGap: 12 }}>
        {modes.map((m) =>
          m.key === mode ? (
            <PrimaryButton
              key={m.key}
              styles={{ root: { flex: 1 } }}
              onClick={() => setMode(m.key as "graph" | "pnpJs" | "webpart")}
            >
              {m.label}
            </PrimaryButton>
          ) : (
            <DefaultButton
              key={m.key}
              styles={{ root: { flex: 1 } }}
              onClick={() => setMode(m.key as "graph" | "pnpJs" | "webpart")}
            >
              {m.label}
            </DefaultButton>
          )
        )}
      </Stack>

      {mode === "graph" && (
        <Stack>
          <GraphApiTester context={context} />
        </Stack>
      )}

      {mode === "pnpJs" && (
        <Stack>
          <PnpJsTester context={context} />
        </Stack>
      )}

      {mode === "webpart" && (
        <Stack>
          <WebpartSetting context={context} />
        </Stack>
      )}
    </Stack>
  );
};

export default App;
