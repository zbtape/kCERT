/**
 * Simple Proxy Server for GPT API
 * This runs locally on localhost to bypass CORS restrictions
 * 
 * Usage: node proxy-server.js
 * Then update the API URL in taskpane.ts to: http://localhost:3001/api/gpt
 */

const http = require('http');
const https = require('https');
const url = require('url');

const PROXY_PORT = 3001;
const TARGET_API = 'https://digitalmatrix-cat.kpmgcloudops.com/workspace/api/v1/generativeai/chat';

const server = http.createServer((req, res) => {
    // Enable CORS for all requests - allow any origin (since we don't know the add-in's origin)
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization, Accept');
    res.setHeader('Access-Control-Max-Age', '3600'); // Cache preflight for 1 hour
    
    // Handle preflight requests
    if (req.method === 'OPTIONS') {
        res.writeHead(200);
        res.end();
        return;
    }
    
    // Only handle POST requests to /api/gpt
    if (req.method !== 'POST' || req.url !== '/api/gpt') {
        res.writeHead(404, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'Not found. Use POST /api/gpt' }));
        return;
    }
    
    let body = '';
    
    req.on('data', chunk => {
        body += chunk.toString();
    });
    
    req.on('end', () => {
        try {
            const requestData = JSON.parse(body);
            
            // Extract bearerToken from request body (sent by add-in)
            const bearerToken = requestData.bearerToken;
            
            // Remove bearerToken from the body before sending to API
            const { bearerToken: _, ...apiBody } = requestData;
            
            // Forward the request to the actual API
            const targetUrl = new URL(TARGET_API);
            
            const options = {
                hostname: targetUrl.hostname,
                port: targetUrl.port || 443,
                path: targetUrl.pathname,
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Accept': '*/*',
                    'Authorization': bearerToken ? `Bearer ${bearerToken}` : ''
                },
                rejectUnauthorized: false // Equivalent to verify=False in Python
            };
            
            const apiRequest = https.request(options, (apiRes) => {
                let responseData = '';
                
                apiRes.on('data', chunk => {
                    responseData += chunk.toString();
                });
                
                apiRes.on('end', () => {
                    console.log(`API Response Status: ${apiRes.statusCode}`);
                    console.log(`API Response Headers:`, apiRes.headers);
                    console.log(`API Response Body (first 500 chars):`, responseData.substring(0, 500));
                    
                    res.writeHead(apiRes.statusCode, {
                        'Content-Type': 'application/json',
                        'Access-Control-Allow-Origin': '*'
                    });
                    res.end(responseData);
                });
            });
            
            apiRequest.on('error', (error) => {
                console.error('Proxy error:', error);
                res.writeHead(500, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ error: error.message }));
            });
            
            apiRequest.write(JSON.stringify(apiBody));
            apiRequest.end();
            
        } catch (error) {
            console.error('Parse error:', error);
            res.writeHead(400, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: 'Invalid JSON' }));
        }
    });
});

server.listen(PROXY_PORT, () => {
    console.log(`✅ Proxy server running on http://localhost:${PROXY_PORT}`);
    console.log(`   Forwarding requests to: ${TARGET_API}`);
    console.log(`   Use endpoint: http://localhost:${PROXY_PORT}/api/gpt`);
});

