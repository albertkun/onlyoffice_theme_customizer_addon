# ONLYOFFICE Theme Customizer Add-on

Presentation plugin for ONLYOFFICE that applies preset brand color palettes as a custom slide theme.

## Included plugin files

- `config.json` – plugin manifest
- `index.html` – plugin UI with preset palettes
- `code.js` – ONLYOFFICE Presentation API integration (`Api.CreateTheme`, `SetTheme`)

## Preset palettes

- Ocean Blue
- Sunset Warm
- Forest Green
- Modern Monochrome

## Install

1. Place this plugin folder somewhere accessible to ONLYOFFICE.
2. Open **ONLYOFFICE Docs** or **ONLYOFFICE Desktop Editors**.
3. Install the plugin:
   - **Desktop Editors**: `Plugins` → `Plugin Manager` → `Install plugin manually` and select this folder.
   - **Docs (self-hosted)**: add this plugin directory to your server plugin path and reload the editors.
4. Open a presentation and run **Theme Customizer** from plugins.

## Usage

1. Open the plugin in a presentation.
2. Click one of the preset palettes.
3. The plugin creates a custom theme and applies it to the current presentation.

## Notes

- The plugin supports ONLYOFFICE Presentation editor (`slide`) only.
- If opened outside ONLYOFFICE, the UI still renders but cannot apply themes.
