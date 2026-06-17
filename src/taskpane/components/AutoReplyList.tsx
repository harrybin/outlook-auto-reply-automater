import { useRef, useState } from "react";
import {
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  DialogTrigger,
  Field,
  Input,
  Switch,
  Textarea,
  tokens,
} from "@fluentui/react-components";
import Editor, { type ContentEditableEvent } from "react-simple-wysiwyg";
import {
  Add24Regular,
  Delete24Regular,
  Edit24Regular,
} from "@fluentui/react-icons";
import type { AutoReplyMessage } from "../types";
import { useStore } from "../useStore";
import { EmojiPicker } from "./EmojiPicker";

export function AutoReplyList() {
  const messages = useStore((s) => s.autoReplyMessages);
  const addMessage = useStore((s) => s.addMessage);
  const updateMessage = useStore((s) => s.updateMessage);
  const deleteMessage = useStore((s) => s.deleteMessage);

  const [editing, setEditing] = useState<AutoReplyMessage | null>(null);
  const [isNew, setIsNew] = useState(false);

  const emptyDraft = (): Omit<
    AutoReplyMessage,
    "id" | "createdAt" | "updatedAt"
  > => ({
    name: "",
    subject: "",
    body: "",
    isHtml: true,
  });

  const [draft, setDraft] = useState(emptyDraft());

  const hasBodyContent = (body: string, isHtml: boolean): boolean => {
    if (!isHtml) {
      return body.trim().length > 0;
    }

    const withoutTags = body
      .replace(/<[^>]*>/g, " ")
      .replace(/&nbsp;/gi, " ")
      .trim();

    return withoutTags.length > 0;
  };

  const openNew = () => {
    setDraft(emptyDraft());
    setIsNew(true);
    setEditing(null);
  };

  const openEdit = (msg: AutoReplyMessage) => {
    setDraft({
      name: msg.name,
      subject: msg.subject,
      body: msg.body,
      isHtml: msg.isHtml,
    });
    setEditing(msg);
    setIsNew(false);
  };

  const handleSave = () => {
    if (isNew) {
      addMessage(draft);
    } else if (editing) {
      updateMessage(editing.id, draft);
    }
    setEditing(null);
    setIsNew(false);
  };

  const isOpen = isNew || editing !== null;

  // Refs for cursor-aware emoji insertion
  const editorContainerRef = useRef<HTMLDivElement | null>(null);
  const blurRangeRef = useRef<Range | null>(null);
  const textareaRef = useRef<HTMLTextAreaElement | null>(null);
  const savedTextareaCursorRef = useRef<{ start: number; end: number } | null>(
    null,
  );

  const handleEditorBlur = () => {
    const sel = window.getSelection();
    if (sel && sel.rangeCount > 0) {
      blurRangeRef.current = sel.getRangeAt(0).cloneRange();
    }
  };

  const handleTextareaBlur = () => {
    const ta = textareaRef.current;
    if (ta) {
      savedTextareaCursorRef.current = {
        start: ta.selectionStart,
        end: ta.selectionEnd,
      };
    }
  };

  const insertEmoji = (emoji: string) => {
    if (draft.isHtml) {
      const container = editorContainerRef.current;
      const contentEditable = container?.querySelector(
        "[contenteditable]",
      ) as HTMLElement | null;
      if (contentEditable && blurRangeRef.current) {
        contentEditable.focus();
        const sel = window.getSelection();
        if (sel) {
          sel.removeAllRanges();
          sel.addRange(blurRangeRef.current);
          const range = sel.getRangeAt(0);
          range.deleteContents();
          const textNode = document.createTextNode(emoji);
          range.insertNode(textNode);
          range.setStartAfter(textNode);
          range.setEndAfter(textNode);
          sel.removeAllRanges();
          sel.addRange(range);
          contentEditable.dispatchEvent(new Event("input", { bubbles: true }));
          return;
        }
      }
      // Fallback: append to body string
      setDraft((p) => ({ ...p, body: p.body + emoji }));
    } else {
      const cursor = savedTextareaCursorRef.current;
      const pos = cursor ?? { start: draft.body.length, end: draft.body.length };
      const newBody =
        draft.body.substring(0, pos.start) +
        emoji +
        draft.body.substring(pos.end);
      setDraft((p) => ({ ...p, body: newBody }));
      // Restore cursor position after React re-renders
      const newCursorPos = pos.start + emoji.length;
      savedTextareaCursorRef.current = { start: newCursorPos, end: newCursorPos };
      requestAnimationFrame(() => {
        const ta = textareaRef.current;
        if (ta) {
          ta.selectionStart = newCursorPos;
          ta.selectionEnd = newCursorPos;
        }
      });
    }
  };

  return (
    <div
      style={{
        display: "flex",
        flexDirection: "column",
        gap: tokens.spacingVerticalM,
      }}
    >
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
        }}
      >
        <span style={{ fontWeight: "600", fontSize: tokens.fontSizeBase400 }}>
          Auto-Reply Messages
        </span>
        <Button
          icon={<Add24Regular />}
          appearance="primary"
          size="small"
          onClick={openNew}
        >
          New
        </Button>
      </div>

      {messages.length === 0 && (
        <p style={{ color: tokens.colorNeutralForeground3 }}>
          No messages yet. Create one to get started.
        </p>
      )}

      {messages.map((msg) => (
        <MessageCard
          key={msg.id}
          message={msg}
          onEdit={() => openEdit(msg)}
          onDelete={() => deleteMessage(msg.id)}
        />
      ))}

      <Dialog
        open={isOpen}
        onOpenChange={(_e, data) => {
          if (!data.open) {
            setEditing(null);
            setIsNew(false);
          }
        }}
      >
        <DialogSurface>
          <DialogBody>
            <DialogTitle>
              {isNew ? "New Auto-Reply Message" : "Edit Auto-Reply Message"}
            </DialogTitle>
            <DialogContent
              style={{
                display: "flex",
                flexDirection: "column",
                gap: tokens.spacingVerticalM,
              }}
            >
              <Field label="Name" required>
                <Input
                  value={draft.name}
                  onChange={(_e, d) =>
                    setDraft((p) => ({ ...p, name: d.value }))
                  }
                  placeholder="e.g. Vacation reply"
                />
              </Field>
              <Field label="Email Subject" required>
                <Input
                  value={draft.subject}
                  onChange={(_e, d) =>
                    setDraft((p) => ({ ...p, subject: d.value }))
                  }
                  placeholder="e.g. Out of office"
                />
              </Field>
              <Field label="Message Body" required>
                <div
                  style={{
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "flex-end",
                    marginBottom: "4px",
                  }}
                >
                  <EmojiPicker onEmojiSelect={insertEmoji} />
                </div>
                {draft.isHtml ? (
                  <Editor
                    value={draft.body}
                    onChange={(event: ContentEditableEvent) =>
                      setDraft((p) => ({ ...p, body: event.target.value }))
                    }
                    placeholder="Enter your auto-reply message..."
                    containerProps={{
                      ref: editorContainerRef,
                      onBlur: handleEditorBlur,
                      style: {
                        minHeight: "240px",
                        border: `1px solid ${tokens.colorNeutralStroke1}`,
                        borderRadius: tokens.borderRadiusMedium,
                        backgroundColor: tokens.colorNeutralBackground1,
                      },
                    }}
                  />
                ) : (
                  <Textarea
                    ref={textareaRef}
                    rows={6}
                    value={draft.body}
                    onChange={(_e, d) =>
                      setDraft((p) => ({ ...p, body: d.value }))
                    }
                    onBlur={handleTextareaBlur}
                    placeholder="Enter your auto-reply message..."
                  />
                )}
              </Field>
              <Field label="HTML format">
                <Switch
                  checked={draft.isHtml}
                  onChange={(_e, d) =>
                    setDraft((p) => ({ ...p, isHtml: d.checked }))
                  }
                  label={draft.isHtml ? "HTML" : "Plain text"}
                />
              </Field>
            </DialogContent>
            <DialogActions>
              <DialogTrigger disableButtonEnhancement>
                <Button appearance="secondary">Cancel</Button>
              </DialogTrigger>
              <Button
                appearance="primary"
                onClick={handleSave}
                disabled={
                  !draft.name.trim() ||
                  !draft.subject.trim() ||
                  !hasBodyContent(draft.body, draft.isHtml)
                }
              >
                Save
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </div>
  );
}

interface MessageCardProps {
  message: AutoReplyMessage;
  onEdit: () => void;
  onDelete: () => void;
}

function MessageCard({ message, onEdit, onDelete }: MessageCardProps) {
  return (
    <div
      style={{
        border: `1px solid ${tokens.colorNeutralStroke2}`,
        borderRadius: tokens.borderRadiusMedium,
        padding: tokens.spacingVerticalS,
        display: "flex",
        justifyContent: "space-between",
        alignItems: "flex-start",
        gap: tokens.spacingHorizontalS,
      }}
    >
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ fontWeight: "600", marginBottom: "2px" }}>
          {message.name}
        </div>
        <div
          style={{
            color: tokens.colorNeutralForeground3,
            fontSize: tokens.fontSizeBase200,
            overflow: "hidden",
            textOverflow: "ellipsis",
            whiteSpace: "nowrap",
          }}
        >
          {message.subject}
        </div>
      </div>
      <div
        style={{
          display: "flex",
          gap: tokens.spacingHorizontalXS,
          flexShrink: 0,
        }}
      >
        <Button
          icon={<Edit24Regular />}
          appearance="subtle"
          size="small"
          onClick={onEdit}
        />
        <Button
          icon={<Delete24Regular />}
          appearance="subtle"
          size="small"
          onClick={onDelete}
        />
      </div>
    </div>
  );
}
