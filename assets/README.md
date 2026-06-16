# Icons for Outlook Auto-Reply Automater

This directory contains icon files for the Outlook add-in.

## Icon Files

### `icon-16.png` and `icon-32.png`
- **Purpose**: Ribbon icons displayed in the Outlook ribbon UI
- **Size**: 16×16 and 32×32 pixels
- **Usage**: Referenced in `manifest.json` for the ribbon group
- **Design**: Should use Office/Azure branding colors (preferably blue theme)

### `icon-outline.png`
- **Purpose**: Outline icon for Teams app catalog and light backgrounds
- **Size**: 32×32 pixels
- **Usage**: Referenced in `manifest.json` as outline icon
- **Design**: White outline on transparent background

### `icon-color.png`
- **Purpose**: Full-color icon for Teams app catalog and distribution
- **Size**: 32×32 pixels
- **Usage**: Referenced in `manifest.json` as color icon
- **Design**: Full-color icon, preferably blue theme to match Office branding

## Current Status

⚠️ **Placeholder Files**: The current icons are placeholder PNG files. For production use, replace them with actual icon designs.

## How to Regenerate Icons

Run the icon generation script:

```bash
npm run icons
```

Or directly:

```bash
node scripts/generate-icons.js
```

## Design Recommendations

1. **Consistency**: Use the same color scheme across all icon sizes
2. **Clarity**: Ensure icons are clear and recognizable at small sizes (16×16)
3. **Branding**: Consider using Office blue (#0078D4) as the primary color
4. **Transparency**: Use transparent backgrounds for outline and color icons
5. **Format**: Use PNG format with proper compression

## Icon Design Tools

- **Microsoft Office**: Use Office design guidelines for consistency
- **Figma**: Professional design tool for creating scalable icons
- **Photoshop/GIMP**: Traditional image editors for detailed work
- **Online tools**: Icons8, Flaticon, or Iconify for icon inspiration

## Next Steps

1. Create or obtain proper icon designs
2. Replace the placeholder PNG files in this directory
3. Test icons in the add-in UI to ensure visibility and clarity
4. Validate icons meet Microsoft/Office requirements
