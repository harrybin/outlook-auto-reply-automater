import { useState } from "react";
import {
  Button,
  Popover,
  PopoverSurface,
  PopoverTrigger,
  tokens,
} from "@fluentui/react-components";
import { Emoji24Regular } from "@fluentui/react-icons";

const EMOJI_CATEGORIES: { label: string; emojis: string[] }[] = [
  {
    label: "Smileys",
    emojis: [
      "😀", "😂", "😊", "😍", "🥰", "😎", "🤔", "😅",
      "😭", "🥺", "😤", "🤣", "😇", "😜", "🥳", "😴", "🤯", "😷",
    ],
  },
  {
    label: "Gestures",
    emojis: ["👍", "👎", "👋", "🤝", "🙏", "💪", "👏", "🤞", "✌️", "🤙"],
  },
  {
    label: "Hearts",
    emojis: ["❤️", "🧡", "💛", "💚", "💙", "💜", "🖤", "🤍", "💕", "💯"],
  },
  {
    label: "Objects",
    emojis: [
      "📧", "📅", "📌", "✅", "❌", "⚠️", "🔔", "💡",
      "🎉", "🎊", "🏆", "🚀", "⭐", "🔥", "💬", "📝",
    ],
  },
];

interface EmojiPickerProps {
  onEmojiSelect: (emoji: string) => void;
}

export function EmojiPicker({ onEmojiSelect }: EmojiPickerProps) {
  const [open, setOpen] = useState(false);

  const handleEmojiClick = (emoji: string) => {
    onEmojiSelect(emoji);
    setOpen(false);
  };

  return (
    <Popover
      open={open}
      onOpenChange={(_e, data) => setOpen(data.open)}
      positioning="below-start"
    >
      <PopoverTrigger disableButtonEnhancement>
        <Button
          icon={<Emoji24Regular />}
          appearance="subtle"
          size="small"
          title="Insert emoji"
        />
      </PopoverTrigger>
      <PopoverSurface style={{ padding: tokens.spacingVerticalS }}>
        <div style={{ width: "280px" }}>
          {EMOJI_CATEGORIES.map((category) => (
            <div
              key={category.label}
              style={{ marginBottom: tokens.spacingVerticalS }}
            >
              <div
                style={{
                  fontSize: tokens.fontSizeBase100,
                  color: tokens.colorNeutralForeground3,
                  marginBottom: "4px",
                  fontWeight: "600",
                  textTransform: "uppercase",
                  letterSpacing: "0.05em",
                }}
              >
                {category.label}
              </div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: "2px" }}>
                {category.emojis.map((emoji) => (
                  <button
                    key={emoji}
                    onClick={() => handleEmojiClick(emoji)}
                    aria-label={emoji}
                    title={emoji}
                    style={{
                      background: "none",
                      border: "none",
                      cursor: "pointer",
                      fontSize: "20px",
                      padding: "4px",
                      borderRadius: tokens.borderRadiusSmall,
                      lineHeight: 1,
                    }}
                  >
                    {emoji}
                  </button>
                ))}
              </div>
            </div>
          ))}
        </div>
      </PopoverSurface>
    </Popover>
  );
}
