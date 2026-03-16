# PCF Position Org Chart

A **Power Apps Component Framework (PCF)** control that renders an interactive org chart for the **Position** table in Microsoft Dataverse. It is designed to be used inside a **Model Driven App** — add it to any Position form, type a position name, hit **Build Chart**, and the control automatically draws the full hierarchy: ancestors above, the selected position highlighted, and all descendants below.

---

## Screenshot

> _Add a screenshot of the rendered org chart here._

---

## Features

- Renders a hierarchical org chart for the **Position** table in Dataverse
- Designed for use in **Model Driven Apps** on a Position form
- Configurable table name, name column, and parent lookup column
- Highlights the selected position and colours its ancestors
- Flexible node boxes that expand to fit long position names
- Scrollable canvas for large org structures
- Integrates directly on a model-driven form via the Custom Controls picker

---

## Properties

| Display Name | Internal Name | Type | Usage | Description |
|---|---|---|---|---|
| **Position** | `defaultPositionName` | Single Line Text | Bound | Binds to a form field. Its value pre-fills the search box when the form loads. |
| **Entity** | `tableLogicalName` | Single Line Text | Input | Logical name of the Dataverse table (e.g. `position`). |
| **Position Name** | `nameColumn` | Single Line Text | Input | Logical name of the column that holds the display name (e.g. `name`). |
| **Position Parent** | `parentLookupColumn` | Single Line Text | Input | Logical name of the parent lookup column, **without** the leading `_` and trailing `_value` (e.g. `parentpositionid`). |

---

## Versions

| Artifact | Version |
|---|---|
| PCF Control | `0.0.8` |
| Dataverse Managed Solution | `2.0.1` |

---

## Getting Started

### Option A — Import the pre-built solution (recommended)

> **Latest release:** Download **`poschart.zip`** from the [**v1.0 Release**](../../releases/tag/v1.0) page — no build tools required.

1. Download **`poschart.zip`** from the [Releases](../../releases) page (see the **v1.0** release).
2. In Power Apps, go to **Solutions → Import solution**.
3. Upload `poschart.zip` and follow the wizard.
4. Open your **Model Driven App**, navigate to a **Position** form, add a **Text** column, then switch its control to **positionchart.positionchart** via **Components → + Component**.
5. Set the following properties:
   - **Entity**: `position`
   - **Position Name**: `name`
   - **Position Parent**: `parentpositionid`

### Option B — Build from source

#### Prerequisites

- [Node.js](https://nodejs.org/) v16+
- [.NET Framework 4.6.2](https://dotnet.microsoft.com/download/dotnet-framework)
- [MSBuild](https://learn.microsoft.com/visualstudio/msbuild/msbuild) (ships with Visual Studio or Build Tools)
- [Power Platform CLI](https://learn.microsoft.com/power-platform/developer/cli/introduction)

#### Build the PCF control

```powershell
cd pcf_positionchart
npm install
npm run build
```

#### Package the managed solution

```powershell
cd ..\poschart
msbuild /t:build /restore /p:configuration=Release
# Output: poschart\bin\Release\poschart.zip
```

---

## Project Structure

```
pcf_positionchart/          # PCF TypeScript project
│   positionchart/
│   ├── index.ts            # Control logic
│   ├── ControlManifest.Input.xml
│   └── css/
│       └── positionchart.css
│   package.json
│   tsconfig.json
│
poschart/                   # Dataverse solution packaging project
│   poschart.cdsproj
│   src/Other/Solution.xml
```

---

## License

This project is licensed under the **MIT License** — see the [LICENSE](LICENSE) file for details.

---

## Contributing

Pull requests are welcome! For major changes, please open an issue first to discuss what you'd like to change.

---

## Author

**Damola Ojoniyi**
