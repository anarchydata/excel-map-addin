# Excel Map

Excel desktop add-in that shows a minimap of the active sheet’s used range (like VS Code’s code minimap). Drag the viewport rectangle to scroll Excel to that area; the rectangle updates when you zoom or scroll outside the map.

## Requirements

- Excel for Windows or Mac (uses `ExcelApiDesktop` window APIs: `visibleRange`, `scrollRow` / `scrollColumn`, `zoom`)
- Node.js 18+

## Dev

```bash
npm install
npm run icons
npm start
```

Dev server: `https://localhost:3002`

```bash
npm stop
```
