import React, { useRef } from "react";
import { Text, makeStyles, tokens, shorthands } from "@fluentui/react-components";
import {
  ReOrderDotsVertical20Regular,
  Add20Regular,
  Dismiss20Regular,
} from "@fluentui/react-icons";
import type { FileNameField } from "../types";

// ============================================================
// Styles
// ============================================================

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    ...shorthands.gap("6px"),
  },
  row: {
    display: "flex",
    ...shorthands.gap("4px"),
    flexWrap: "wrap" as const,
  },
  chip: {
    display: "flex",
    alignItems: "center",
    ...shorthands.gap("4px"),
    ...shorthands.padding("4px", "8px"),
    ...shorthands.borderRadius("16px"),
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground2,
    cursor: "grab",
    userSelect: "none" as const,
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    border: "1px solid transparent",
  },
  chipDragging: {
    opacity: 0.4,
  },
  chipDragOver: {
    border: `1px dashed ${tokens.colorBrandStroke1}`,
  },
  addChip: {
    display: "flex",
    alignItems: "center",
    ...shorthands.gap("4px"),
    ...shorthands.padding("4px", "8px"),
    ...shorthands.borderRadius("16px"),
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground2,
    cursor: "pointer",
    userSelect: "none" as const,
    fontSize: tokens.fontSizeBase200,
    border: `1px dashed ${tokens.colorNeutralStroke1}`,
  },
  iconBtn: {
    background: "none",
    border: "none",
    padding: "0",
    cursor: "pointer",
    display: "flex",
    alignItems: "center",
    color: "inherit",
  },
  preview: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    fontStyle: "italic" as const,
    wordBreak: "break-all" as const,
  },
});

// ============================================================
// Constanten
// ============================================================

const ALL_FIELDS: FileNameField[] = ["date", "subject", "sender", "recipient"];

const FIELD_LABELS: Record<FileNameField, string> = {
  date: "Tijd",
  subject: "Onderwerp",
  sender: "Verzender",
  recipient: "Ontvanger",
};

// ============================================================
// Props
// ============================================================

interface FileNameBuilderProps {
  fields: FileNameField[];
  onChange: (fields: FileNameField[]) => void;
  preview: string;
}

// ============================================================
// Component
// ============================================================

export const FileNameBuilder: React.FC<FileNameBuilderProps> = ({
  fields,
  onChange,
  preview,
}) => {
  const styles = useStyles();
  const dragIndex = useRef<number | null>(null);
  const dragOverIndex = useRef<number | null>(null);

  const inactiveFields = ALL_FIELDS.filter((f) => !fields.includes(f));

  // --- Drag handlers ---
  const handleDragStart = (index: number) => {
    dragIndex.current = index;
  };

  const handleDragOver = (e: React.DragEvent, index: number) => {
    e.preventDefault();
    dragOverIndex.current = index;
  };

  const handleDrop = () => {
    if (dragIndex.current === null || dragOverIndex.current === null) return;
    if (dragIndex.current === dragOverIndex.current) return;

    const updated = [...fields];
    const [moved] = updated.splice(dragIndex.current, 1);
    updated.splice(dragOverIndex.current, 0, moved);

    dragIndex.current = null;
    dragOverIndex.current = null;
    onChange(updated);
  };

  const handleRemove = (field: FileNameField) => {
    onChange(fields.filter((f) => f !== field));
  };

  const handleAdd = (field: FileNameField) => {
    onChange([...fields, field]);
  };

  return (
    <div className={styles.container}>
      {/* Actieve velden (sleepbaar) */}
      <div className={styles.row}>
        {fields.map((field, index) => (
          <div
            key={field}
            className={styles.chip}
            draggable
            onDragStart={() => handleDragStart(index)}
            onDragOver={(e) => handleDragOver(e, index)}
            onDrop={handleDrop}
          >
            <ReOrderDotsVertical20Regular />
            <span>{FIELD_LABELS[field]}</span>
            <button
              className={styles.iconBtn}
              onClick={() => handleRemove(field)}
              title={`${FIELD_LABELS[field]} verwijderen`}
            >
              <Dismiss20Regular />
            </button>
          </div>
        ))}

        {/* Inactieve velden toevoegen */}
        {inactiveFields.map((field) => (
          <div
            key={field}
            className={styles.addChip}
            onClick={() => handleAdd(field)}
            title={`${FIELD_LABELS[field]} toevoegen`}
          >
            <Add20Regular />
            <span>{FIELD_LABELS[field]}</span>
          </div>
        ))}
      </div>

      {/* Preview */}
      <Text className={styles.preview}>{preview || "..."}</Text>
    </div>
  );
};
