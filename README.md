# Excel Grid Viewer

A lightweight React + Vite application for loading, parsing, and viewing Excel files (.xls, .xlsx) from SharePoint or OneDrive links. Designed to be hosted as static files in SharePoint document libraries.

## Features

- **URL-based file loading**: Paste a link to an Excel file and the app fetches and parses it in-browser
- **Multiple sheet support**: Switch between sheets with a dropdown selector
- **Interactive grid**: View data with row and column headers, click any cell to see details
- **Cell details panel**: See cell address, formula, number format, and value
- **Link persistence**: Automatically save and remember the last loaded file URL
- **Graceful error handling**: Helpful messages when files fail to load with tips for SharePoint links
- **Fully responsive**: Works on desktop and mobile browsers

## Technology Stack

- **React 19** - UI framework
- **Vite 7** - Fast build tool and dev server
- **TypeScript** - Type-safe code
- **SheetJS (xlsx)** - Excel file parsing in the browser
- **react-window** - Virtualized grid for performance

## Setup Instructions

### Prerequisites

- Node.js 18+ and npm

### 1. Install Dependencies

```bash
npm install
```

### 2. Development Server

```bash
npm run dev
```

This starts a local dev server at `http://localhost:5173/`. Open it in your browser and the app will hot-reload when you make changes.

### 3. Build for SharePoint

```bash
npm run build
```

This creates a production build in the `dist/` folder with all static files:
- `dist/index.html` - Main HTML file
- `dist/assets/` - JavaScript, CSS, and other assets

**Important**: All paths are relative (using `./`) so the app works when hosted in a SharePoint subfolder.

### 4. Upload to SharePoint

1. Navigate to your SharePoint Online site
2. Go to a document library (e.g., "Shared Documents" or create a new one)
3. Create a folder for the app (e.g., "xls-grid-app")
4. Upload **all contents of `dist/`** into that folder:
   - Upload `index.html`
   - Upload the entire `assets/` folder

### 5. Access the App

Open the app via SharePoint:
```
https://yourtenant.sharepoint.com/sites/YourSite/Shared%20Documents/xls-grid-app/index.html
```

Or use the direct URL to `index.html` from the SharePoint UI (right-click → Copy link).

## Usage

1. **Load a file**: Paste a link to an Excel file in the URL input:
   - Direct links to `.xlsx` or `.xls` files
   - SharePoint/OneDrive document links
   - Make sure the account has view access to the file

2. **Remember the link**: The "Remember this link" checkbox is enabled by default. Disable it if you don't want the URL saved.

3. **View sheets**: If the file has multiple sheets, use the dropdown to switch between them.

4. **Explore the grid**: 
   - Scroll horizontally/vertically to browse data
   - Click any cell to open the details panel
   - The first row is treated as headers

5. **Cell details**: When you click a cell, a panel shows:
   - Cell address (e.g., "A1")
   - Row/column position
   - Cell value
   - Formula (if present)
   - Number format (if present)

## Configuration

### SharePoint Link Tips

**Direct links work best.** To get a direct link from SharePoint:

1. In SharePoint, right-click the Excel file
2. Select "Share"
3. Copy the link
4. Paste into the app

**If the link doesn't work:**
- The file might require authentication—make sure you're logged in
- Try getting a "Copy link" from SharePoint with "Anyone with existing access" permissions
- Some SharePoint links require a `download=1` parameter; the app attempts to handle this

**Non-direct links** (that end in `?d=` without download flag) may not work due to browser CORS policies. The app will show an error message with suggestions.

### How Configuration is Saved

The app uses **localStorage** to remember the last loaded file URL:

- **Key**: `xlsGrid:lastUrl`
- **Value**: The full URL to the Excel file
- **Storage location**: Browser's local storage (per domain)
- **Persistence**: Stays saved until you click "Clear saved link"

On app load, if a saved URL exists:
- The URL is automatically populated in the input field
- You can click "Load" to reload it, or paste a new URL

**Note**: localStorage is domain-specific. If you move the app folder to a different SharePoint site URL, the saved link won't be available.

## Project Structure

```
priority-roadmap/
├── src/
│   ├── components/
│   │   ├── UrlLoader.tsx       # URL input and file loading UI
│   │   ├── SheetSelector.tsx   # Sheet dropdown selector
│   │   ├── GridView.tsx        # Virtualized grid component
│   │   └── CellDetails.tsx     # Cell details modal panel
│   ├── utils/
│   │   └── excel.ts            # Excel parsing and utility functions
│   ├── App.tsx                 # Main app component
│   ├── main.tsx                # React DOM render
│   ├── styles.css              # All styling
│   └── index.css               # Global styles
├── public/                      # Static assets (optional)
├── index.html                   # HTML entry point
├── vite.config.ts              # Vite configuration (base: './')
├── tsconfig.json               # TypeScript configuration
└── package.json                # Dependencies and scripts
```

## Build Scripts

- `npm run dev` - Start dev server with hot reload
- `npm run build` - Build for production to `dist/`
- `npm run lint` - Run ESLint checks
- `npm run preview` - Preview production build locally

## Browser Support

- Chrome/Edge 90+
- Firefox 88+
- Safari 14+
- Modern mobile browsers

## Limitations & Notes

- **File size**: Large files (100MB+) may be slow to parse in the browser
- **Sheet limits**: Files with 1000+ rows/columns may impact scroll performance
- **Auth**: File must be accessible to your SharePoint account
- **CORS**: Some external URLs may be blocked by browser CORS policies
- **Offline**: The app requires internet access to fetch files (no offline support)

## Troubleshooting

| Issue | Solution |
|-------|----------|
| **"Failed to fetch file" error** | Check the URL, ensure you're logged into SharePoint, and verify the file is accessible |
| **"Invalid URL format" error** | Make sure to include `http://` or `https://` at the start of the URL |
| **Grid is slow with large files** | The virtualized grid helps, but very large sheets (10k+ rows) will be slow to parse |
| **Saved link not loading** | Try clearing the URL field and pasting a fresh link; localStorage might be full |
| **App won't load in SharePoint** | Verify all `dist/` files were uploaded correctly, including the `assets/` folder |

## Development Notes

- The app is a **single-page application (SPA)** with no server or API routes
- All Excel parsing happens in the browser using SheetJS
- Relative paths (`./`) ensure it works in SharePoint subfolders
- For local dev testing of SharePoint auth, use an actual SharePoint link

## License

This project is provided as-is for use in SharePoint environments.

- [@vitejs/plugin-react](https://github.com/vitejs/vite-plugin-react/blob/main/packages/plugin-react) uses [Babel](https://babeljs.io/) (or [oxc](https://oxc.rs) when used in [rolldown-vite](https://vite.dev/guide/rolldown)) for Fast Refresh
- [@vitejs/plugin-react-swc](https://github.com/vitejs/vite-plugin-react/blob/main/packages/plugin-react-swc) uses [SWC](https://swc.rs/) for Fast Refresh

## React Compiler

The React Compiler is not enabled on this template because of its impact on dev & build performances. To add it, see [this documentation](https://react.dev/learn/react-compiler/installation).

## Expanding the ESLint configuration

If you are developing a production application, we recommend updating the configuration to enable type-aware lint rules:

```js
export default defineConfig([
  globalIgnores(['dist']),
  {
    files: ['**/*.{ts,tsx}'],
    extends: [
      // Other configs...

      // Remove tseslint.configs.recommended and replace with this
      tseslint.configs.recommendedTypeChecked,
      // Alternatively, use this for stricter rules
      tseslint.configs.strictTypeChecked,
      // Optionally, add this for stylistic rules
      tseslint.configs.stylisticTypeChecked,

      // Other configs...
    ],
    languageOptions: {
      parserOptions: {
        project: ['./tsconfig.node.json', './tsconfig.app.json'],
        tsconfigRootDir: import.meta.dirname,
      },
      // other options...
    },
  },
])
```

You can also install [eslint-plugin-react-x](https://github.com/Rel1cx/eslint-react/tree/main/packages/plugins/eslint-plugin-react-x) and [eslint-plugin-react-dom](https://github.com/Rel1cx/eslint-react/tree/main/packages/plugins/eslint-plugin-react-dom) for React-specific lint rules:

```js
// eslint.config.js
import reactX from 'eslint-plugin-react-x'
import reactDom from 'eslint-plugin-react-dom'

export default defineConfig([
  globalIgnores(['dist']),
  {
    files: ['**/*.{ts,tsx}'],
    extends: [
      // Other configs...
      // Enable lint rules for React
      reactX.configs['recommended-typescript'],
      // Enable lint rules for React DOM
      reactDom.configs.recommended,
    ],
    languageOptions: {
      parserOptions: {
        project: ['./tsconfig.node.json', './tsconfig.app.json'],
        tsconfigRootDir: import.meta.dirname,
      },
      // other options...
    },
  },
])
```
