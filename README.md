# Office Bridge

**Live Microsoft Office automation for Claude Code.**

A plugin that connects Claude Code directly to running Office apps (Word, Excel, PowerPoint, Outlook) through Office Add-ins, enabling real-time document reading and editing.

---

## 🚀 Installation

### Option 1: Via SCL Marketplace (Recommended)

Install all SCL plugins at once:

```bash
claude plugins add SanctionedCodeList/SCL_marketplace
```

### Option 2: Standalone Installation

```bash
claude plugins add SanctionedCodeList/office-bridge
```

### Verify Installation

```bash
claude plugins list
# Should show: office-bridge
```

---

## The Problem

Programmatic Office automation typically requires:
- **COM/AppleScript** — Platform-specific, complex setup
- **VBA macros** — Security concerns, no external integration
- **Offline file manipulation** — No live editing, can't see open documents

**Office Bridge** solves this by connecting directly to running Office apps through the Office.js API, enabling real-time bidirectional communication.

---

## Features

| Application | Capabilities |
|-------------|--------------|
| **Word** | Read document structure, edit by paragraph reference, tracked changes |
| **Excel** | Read/write cells and ranges, get sheet structure |
| **PowerPoint** | Create slides, fill placeholders, capture slide images |
| **Outlook** | Read emails, compose replies (limited for Gmail accounts) |

---

## Setup Guide

### Step 1: Install Dependencies

```bash
git clone https://github.com/SanctionedCodeList/office-bridge.git
cd office-bridge
./install.sh
```

### Step 2: Start the Bridge Server

```bash
./server.sh &
```

### Step 3: Start Add-in Dev Servers

Start the server for each Office app you want to use:

```bash
cd addins/word && npm run dev-server &        # Port 3000
cd addins/excel && npm run dev-server &       # Port 3001
cd addins/powerpoint && npm run dev-server &  # Port 3002
cd addins/outlook && npm run dev-server &     # Port 3003
```

### Step 4: Sideload Add-ins

Follow the instructions in `references/setup.md` to sideload add-ins into your Office apps.

---

## How It Works

```
┌─────────────┐     WebSocket     ┌─────────────┐     Office.js    ┌─────────────┐
│ Claude Code │ ◄──────────────► │   Bridge    │ ◄──────────────► │   Office    │
│   (Client)  │                  │   Server    │                  │    Apps     │
└─────────────┘                  └─────────────┘                  └─────────────┘
```

1. **Bridge server** runs locally, listens for WebSocket connections
2. **Office Add-ins** connect to bridge via WebSocket
3. **Claude Code** connects to bridge, discovers available documents
4. **Commands flow**: Claude → Bridge → Add-in → Office.js → Document

---

## Usage Examples

### With Claude Code

Once set up, just describe what you need:

```
> Read the current Word document and summarize it

> Replace all instances of "Acme Corp" with "NewCo Inc" in the open document

> Create a PowerPoint slide with the quarterly results
```

### Programmatic Usage

```typescript
import { connect } from "./src/client.js";

const bridge = await connect();
const docs = await bridge.documents();  // Word documents
const ppts = await bridge.powerpoint(); // PowerPoint presentations

// Get document tree
const tree = await docs[0].getTree();

// Edit by reference
await docs[0].replaceByRef({ p: 3 }, "New text");

// Capture slide image
const slide = await ppts[0].getSlideImage(1);
```

---

## Requirements

| Requirement | Details |
|-------------|---------|
| **Node.js** | 18 or higher |
| **Microsoft Office** | Word, Excel, PowerPoint, or Outlook |
| **Platform** | macOS or Windows |

---

## Documentation

| Document | Description |
|----------|-------------|
| `references/setup.md` | Detailed setup and sideloading instructions |
| `references/word.md` | Word API reference |
| `references/powerpoint-api.md` | PowerPoint API reference |

---

## Troubleshooting

| Problem | Solution |
|---------|----------|
| Add-in not appearing in Office | Ensure dev server is running, try clearing Office cache |
| WebSocket connection failed | Check that bridge server is running on correct port |
| "Permission denied" errors | Office may need to trust the localhost certificate |
| Changes not reflecting | Some operations require document refresh |

---

## Links

- [GitHub](https://github.com/SanctionedCodeList/office-bridge)
- [SCL Marketplace](https://github.com/SanctionedCodeList/SCL_marketplace)
