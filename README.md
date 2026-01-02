# Office Bridge

A Claude Code plugin that connects to Microsoft Office apps (Word, Excel, PowerPoint, Outlook) via Office Add-ins, enabling live document manipulation.

## Features

- **Word**: Read/edit documents, tracked changes, accessibility tree navigation
- **Excel**: Read/write cells and ranges
- **PowerPoint**: Create slides, fill placeholders, capture slide images
- **Outlook**: Read emails, compose replies (limited for Gmail accounts)

## Installation

### As a Claude Code Plugin

```bash
# Clone the repo
git clone https://github.com/yourusername/office-bridge.git

# Install as plugin
claude plugins install /path/to/office-bridge
```

### Manual Setup

1. Install dependencies:
```bash
cd office-bridge
./install.sh
```

2. Start the bridge server:
```bash
./server.sh &
```

3. Start add-in dev servers (as needed):
```bash
cd addins/word && npm run dev-server &        # Port 3000
cd addins/excel && npm run dev-server &       # Port 3001
cd addins/powerpoint && npm run dev-server &  # Port 3002
cd addins/outlook && npm run dev-server &     # Port 3003
```

4. Sideload add-ins into Office apps (see `references/setup.md`)

## Usage

Once connected, Claude Code can interact with your Office documents:

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

## Documentation

- `references/setup.md` - Detailed setup and sideloading instructions
- `references/word.md` - Word API reference
- `references/powerpoint-api.md` - PowerPoint API reference

## Requirements

- Node.js 18+
- Microsoft Office (Word, Excel, PowerPoint, or Outlook)
- macOS or Windows

## License

MIT
