import * as React from "react";
import { useState, useEffect, Fragment } from "react";
import {
  Text,
  DetailsList,
  IColumn,
  IconButton,
  TextField,
} from "@fluentui/react";
import { EnsureUpdateProperties } from "./type";

interface IEditableJsonListProps {
  category:
    | "web"
    | "lists"
    | "items"
    | "fields"
    | "folders"
    | "files"
    | "users"
    | "groups"
    | "search"
    | "profiles";
  /** 傳入的資料為 JSON 字串 */
  data: string;
  /** 當編輯過程中有變更時回傳 JSON 字串 */
  onChange?: (changes: string) => void;
  /** 當單筆儲存完成時回傳 JSON 字串 */
  onSave?: (updatedJson: string) => void;
}

interface IRow {
  key: string;
  value: any;
  editing: boolean;
  originalValue: any;
}

const EditableJsonList: React.FC<IEditableJsonListProps> = ({
  category,
  data,
  onChange,
  onSave,
}) => {
  const [rows, setRows] = useState<IRow[]>([]);

  /** 初始化：把 JSON string 轉成 rows */
  useEffect(() => {
    try {
      const parsed: Record<string, any> = JSON.parse(data);
      const initRows = Object.entries(parsed).map(([k, v]) => ({
        key: k,
        value: v,
        originalValue: v,
        editing: false,
      }));
      setRows(initRows);
    } catch (err) {
      console.error("JSON 解析錯誤", err);
      setRows([]);
    }
  }, [data]);

  /** 切換編輯模式 */
  const handleEditToggle = (index: number) => {
    const updatedRows = [...rows];
    updatedRows[index].editing = !updatedRows[index].editing;
    setRows(updatedRows);
  };

  /** 修改值（暫存為字串，但會在儲存時轉型） */
  const handleValueChange = (index: number, newValue: string) => {
    const updatedRows = [...rows];
    updatedRows[index].value = newValue;
    setRows(updatedRows);

    const changedRow = updatedRows[index];
    if (changedRow.value !== changedRow.originalValue && onChange) {
      const obj = Object.fromEntries(updatedRows.map((r) => [r.key, r.value]));
      onChange(JSON.stringify(obj, null, 2));
    }
  };

  /** 儲存：轉回正確型別並更新 JSON string */
  const handleSave = (index: number) => {
    const updatedRows = [...rows];
    const row = updatedRows[index];

    // 嘗試還原型別
    let newValue: any = row.value;
    try {
      newValue = JSON.parse(row.value);
    } catch {
      // 如果不能 parse，就維持字串
    }

    row.value = newValue;
    //row.originalValue = newValue;
    row.editing = false;
    setRows(updatedRows);

    const obj = Object.fromEntries(updatedRows.map((r) => [r.key, r.value]));
    const jsonString = JSON.stringify(obj, null, 2);

    onSave?.(jsonString);
  };

  const columns: IColumn[] = [
    {
      key: "colKey",
      name: "Key",
      fieldName: "key",
      minWidth: 80,
      maxWidth: 300,
      isResizable: true,
    },
    {
      key: "colValue",
      name: "Value",
      fieldName: "value",
      minWidth: 80,
      maxWidth: 400,
      isResizable: true,
      onRender: (_item: IRow, index?: number) => {
        if (index === undefined) return null;
        const row = rows[index];
        if (row.editing) {
          return (
            <TextField
              value={row.value !== undefined ? String(row.value) : ""}
              onChange={(_, v) => handleValueChange(index, v || "")}
            />
          );
        }
        return <span>{String(row.value)}</span>;
      },
    },
    {
      key: "colAction",
      name: "Action",
      minWidth: 120,
      maxWidth: 200,
      onRender: (_item: IRow, index?: number) => {
        if (index === undefined) return null;
        const row = rows[index];
        if (!EnsureUpdateProperties.includes(row.key)) {
          return (
            <Text
              styles={{
                root: {
                  color: "GREEN",
                  padding: "8px",
                  borderRadius: "12px",
                },
              }}
              variant="smallPlus"
            >
              {"此屬性無法更新"}
            </Text>
          );
        }
        if (row.editing) {
          const icon = row.value !== row.originalValue ? "Save" : "Cancel";
          return (
            <IconButton
              iconProps={{ iconName: icon }}
              title={icon}
              ariaLabel={icon}
              onClick={() =>
                icon === "Save" ? handleSave(index) : handleEditToggle(index)
              }
            />
          );
        } else {
          return (
            <IconButton
              iconProps={{ iconName: "Edit" }}
              title="Edit"
              ariaLabel="Edit"
              onClick={() => handleEditToggle(index)}
            />
          );
        }
      },
    },
    {
      key: "colAction",
      name: "Status",
      minWidth: 120,
      onRender: (_item: IRow, index?: number) => {
        if (index === undefined) return null;
        const row = rows[index];

        return row.value !== row.originalValue ? (
          <Text
            styles={{
              root: {
                color: "GREEN",
                padding: "8px",
                borderRadius: "12px",
              },
            }}
            variant="smallPlus"
          >
            {"已進行修改 Modified"}
          </Text>
        ) : null;
      },
    },
  ];
  console.log("rows", rows);
  return <DetailsList items={rows} columns={columns} />;
};

export default EditableJsonList;
