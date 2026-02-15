# Appa UI

React application with Vite, routing, and API integration.

## Project Structure

```
ui/
├── src/
│   ├── components/
│   │   ├── Layout.jsx        # Main layout with sidebar and outlet
│   │   ├── Layout.css
│   │   ├── Sidebar.jsx       # Navigation sidebar
│   │   ├── Sidebar.css
│   │   ├── Dashboard.jsx     # Dashboard component
│   │   └── Dashboard.css
│   ├── App.jsx               # Main app with routing
│   ├── App.css
│   ├── main.jsx              # Entry point
│   └── index.css             # Global styles
├── index.html
├── package.json
├── vite.config.js            # Vite configuration with API proxy
└── README.md
```

## Installation

```bash
npm install
```

## Development

Start the development server:

```bash
npm run dev
```

The application will run on `http://localhost:5173` with API proxy to `http://localhost:3000`.

## Build

```bash
npm run build
```

## Features

- **Vite** - Fast build tool and dev server
- **React Router** - Client-side routing
- **Axios** - HTTP requests to backend API
- **Sidebar Navigation** - Left-side menu with routing
- **Dashboard** - Sample component demonstrating API integration
- **API Proxy** - Vite development server proxies `/api` requests to backend

## API Integration

The Dashboard component demonstrates API integration using Axios:

```javascript
const response = await axios.get('/api/health')
```

The Vite proxy configuration automatically forwards API requests to `http://localhost:3000`.
