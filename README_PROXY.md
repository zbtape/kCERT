# GPT API Proxy Server - CORS Workaround

## Quick Setup

Since the GPT API doesn't have CORS headers configured, we use a simple local proxy server to bypass browser CORS restrictions.

### Step 1: Start the Proxy Server

Open a terminal in the project directory and run:

```bash
npm run proxy
```

Or directly:

```bash
node proxy-server.js
```

You should see:
```
✅ Proxy server running on http://localhost:3001
   Forwarding requests to: https://digitalmatrix-cat.kpmgcloudops.com/workspace/api/v1/generativeai/chat
   Use endpoint: http://localhost:3001/api/gpt
```

### Step 2: Keep It Running

**Keep the proxy server running** while you use the Excel Add-in. The add-in is already configured to use the proxy when `useProxy = true` in `taskpane.ts`.

### Step 3: Use GPT Features

1. Open the Excel Add-in
2. Go to "Unique Formula Listing" tab
3. Click "GPT Settings" button
4. Enter your Bearer Token and Engagement Code
5. Click "Save & Validate"
6. Now "Ask GPT" buttons should work!

## How It Works

- The proxy runs on `localhost:3001` (same origin as your add-in = no CORS!)
- Your add-in makes requests to `http://localhost:3001/api/gpt`
- The proxy forwards requests to the real API at `digitalmatrix-cat.kpmgcloudops.com`
- The proxy adds CORS headers to responses so your browser accepts them

## Troubleshooting

**"Failed to connect" or "Connection refused"**
- Make sure the proxy server is running (`node proxy-server.js`)
- Check that port 3001 is not in use by another application

**"Authentication failed"**
- Verify your Bearer Token and Engagement Code are correct
- Check the proxy console for any error messages

**Want to disable proxy later?**
- Set `useProxy = false` in `src/taskpane/taskpane.ts` line ~1592
- This requires CORS to be configured on the API server

