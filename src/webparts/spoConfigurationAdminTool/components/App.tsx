import * as React from "react";
import styles from "./App.module.scss";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { PrimaryButton, Stack } from "@fluentui/react";
import GraphApiTester from "./GraphMode/GraphApiTester";
import { useState } from "react";
import PnpJsTester from "./PnpJsMode/PnpJsTester";

export interface IAppProps {
  context: WebPartContext;
}

const App: React.FC<IAppProps> = ({ context }) => {
  const [mode, setMode] = useState<"graph" | "pnpJs">("graph");

  return (
    <Stack className={`${styles.app}`}>
      <Stack horizontal tokens={{ childrenGap: 12 }}>
        <PrimaryButton
          disabled={mode === "graph"}
          styles={{ root: { flex: 1 } }}
          onClick={() => setMode("graph")}
        >
          Graph Mode
        </PrimaryButton>
        <PrimaryButton
          disabled={mode === "pnpJs"}
          styles={{ root: { flex: 1 } }}
          onClick={() => setMode("pnpJs")}
        >
          PnpJs Mode
        </PrimaryButton>
      </Stack>
      <Stack styles={{ root: { display: mode !== "graph" ? "none" : "flex" } }}>
        <GraphApiTester context={context} />
      </Stack>
      <Stack styles={{ root: { display: mode !== "pnpJs" ? "none" : "flex" } }}>
        <PnpJsTester context={context} />
      </Stack>
    </Stack>
  );
};

export default App;
