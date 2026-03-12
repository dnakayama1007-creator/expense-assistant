// ============================================
// 経費精算アシスタント - ローカルサーバー
// Web アプリ + Discord データ同期 + freee API連携
// ============================================
const http = require('http');
const https = require('https');
const fs = require('fs');
const path = require('path');
const url = require('url');
const querystring = require('querystring');
const crypto = require('crypto');

const PORT = 3000;
const ROOT = __dirname;
const DATA_FILE = path.join(ROOT, 'discord_expenses.json');
const FREEE_CONFIG_FILE = path.join(ROOT, 'freee_config.json');
const GOOGLE_SA_KEY_FILE = path.join(ROOT, 'google_service_account.json');
const PROCESSED_IDS_FILE = path.join(ROOT, 'processed_msg_ids.json');

// Helper: load processed msg IDs set
function loadProcessedIds() {
    try {
        if (fs.existsSync(PROCESSED_IDS_FILE)) {
            return new Set(JSON.parse(fs.readFileSync(PROCESSED_IDS_FILE, 'utf-8')));
        }
    } catch (e) { console.error('loadProcessedIds error:', e.message); }
    return new Set();
}
function saveProcessedIds(set) {
    fs.writeFileSync(PROCESSED_IDS_FILE, JSON.stringify([...set], null, 2), 'utf-8');
}

const MIME_TYPES = {
    '.html': 'text/html; charset=utf-8',
    '.css': 'text/css; charset=utf-8',
    '.js': 'application/javascript; charset=utf-8',
    '.json': 'application/json; charset=utf-8',
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.gif': 'image/gif',
    '.webp': 'image/webp',
    '.svg': 'image/svg+xml',
    '.ico': 'image/x-icon',
};

// ============================================
// freee API Helpers
// ============================================
function loadFreeeConfig() {
    try {
        if (fs.existsSync(FREEE_CONFIG_FILE)) {
            return JSON.parse(fs.readFileSync(FREEE_CONFIG_FILE, 'utf-8'));
        }
    } catch (e) { console.error('freee config load error:', e); }
    return {};
}

function saveFreeeConfig(config) {
    fs.writeFileSync(FREEE_CONFIG_FILE, JSON.stringify(config, null, 2), 'utf-8');
}

function httpsRequest(options, postData) {
    return new Promise(function (resolve, reject) {
        var req = https.request(options, function (res) {
            var data = '';
            res.on('data', function (chunk) { data += chunk; });
            res.on('end', function () {
                try {
                    resolve({ status: res.statusCode, data: JSON.parse(data) });
                } catch (e) {
                    resolve({ status: res.statusCode, data: data });
                }
            });
        });
        req.on('error', reject);
        if (postData) req.write(postData);
        req.end();
    });
}

function freeeApiGet(apiPath, accessToken) {
    return httpsRequest({
        hostname: 'api.freee.co.jp',
        path: apiPath,
        method: 'GET',
        headers: {
            'Authorization': 'Bearer ' + accessToken,
            'Accept': 'application/json'
        }
    });
}

function freeeApiPost(apiPath, accessToken, body) {
    var postData = JSON.stringify(body);
    return httpsRequest({
        hostname: 'api.freee.co.jp',
        path: apiPath,
        method: 'POST',
        headers: {
            'Authorization': 'Bearer ' + accessToken,
            'Content-Type': 'application/json',
            'Content-Length': Buffer.byteLength(postData),
            'Accept': 'application/json'
        }
    }, postData);
}

// Upload receipt image to freee
function freeeUploadReceipt(accessToken, companyId, filePath) {
    return new Promise(function (resolve, reject) {
        var boundary = '----FormBoundary' + Date.now();
        var fileName = path.basename(filePath);
        var ext = path.extname(filePath).toLowerCase();
        var mimeType = MIME_TYPES[ext] || 'image/jpeg';

        // Check if filePath is a data URL (base64)
        var fileBuffer;
        if (filePath.startsWith('data:')) {
            var matches = filePath.match(/^data:([^;]+);base64,(.+)$/);
            if (matches) {
                fileBuffer = Buffer.from(matches[2], 'base64');
                mimeType = matches[1];
                fileName = 'receipt' + (mimeType.indexOf('png') >= 0 ? '.png' : '.jpg');
            } else {
                return resolve(null);
            }
        } else {
            // Local file path
            var fullPath = path.join(ROOT, filePath);
            if (!fs.existsSync(fullPath)) {
                console.log('Receipt file not found:', fullPath);
                return resolve(null);
            }
            fileBuffer = fs.readFileSync(fullPath);
        }

        var bodyParts = [];
        // company_id field
        bodyParts.push('--' + boundary + '\r\n');
        bodyParts.push('Content-Disposition: form-data; name="company_id"\r\n\r\n');
        bodyParts.push(companyId.toString() + '\r\n');
        // receipt file field
        bodyParts.push('--' + boundary + '\r\n');
        bodyParts.push('Content-Disposition: form-data; name="receipt"; filename="' + fileName + '"\r\n');
        bodyParts.push('Content-Type: ' + mimeType + '\r\n\r\n');

        var header = Buffer.from(bodyParts.join(''));
        var footer = Buffer.from('\r\n--' + boundary + '--\r\n');
        var bodyBuffer = Buffer.concat([header, fileBuffer, footer]);

        var options = {
            hostname: 'api.freee.co.jp',
            path: '/api/1/receipts',
            method: 'POST',
            headers: {
                'Authorization': 'Bearer ' + accessToken,
                'Content-Type': 'multipart/form-data; boundary=' + boundary,
                'Content-Length': bodyBuffer.length,
                'Accept': 'application/json'
            }
        };

        var req = https.request(options, function (res) {
            var data = '';
            res.on('data', function (chunk) { data += chunk; });
            res.on('end', function () {
                try {
                    var parsed = JSON.parse(data);
                    console.log('Receipt upload result:', res.statusCode, JSON.stringify(parsed).substring(0, 200));
                    if (res.statusCode === 201 && parsed.receipt) {
                        resolve(parsed.receipt.id);
                    } else {
                        resolve(null);
                    }
                } catch (e) {
                    console.error('Receipt upload parse error:', e);
                    resolve(null);
                }
            });
        });
        req.on('error', function (e) {
            console.error('Receipt upload error:', e);
            resolve(null);
        });
        req.write(bodyBuffer);
        req.end();
    });
}

function refreshFreeeToken(config) {
    if (!config.refresh_token || !config.client_id || !config.client_secret) return Promise.resolve(null);
    var postData = querystring.stringify({
        grant_type: 'refresh_token',
        refresh_token: config.refresh_token,
        client_id: config.client_id,
        client_secret: config.client_secret
    });
    return httpsRequest({
        hostname: 'accounts.secure.freee.co.jp',
        path: '/public_api/token',
        method: 'POST',
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
            'Content-Length': Buffer.byteLength(postData)
        }
    }, postData).then(function (result) {
        if (result.status === 200 && result.data.access_token) {
            config.access_token = result.data.access_token;
            config.refresh_token = result.data.refresh_token;
            config.token_expires = Date.now() + (result.data.expires_in * 1000);
            saveFreeeConfig(config);
            return config;
        }
        return null;
    });
}

function getValidToken() {
    var config = loadFreeeConfig();
    if (!config.access_token) return Promise.resolve(null);
    if (config.token_expires && Date.now() > config.token_expires - 300000) {
        return refreshFreeeToken(config);
    }
    return Promise.resolve(config);
}

function parseBody(req) {
    return new Promise(function (resolve, reject) {
        var body = '';
        req.on('data', function (chunk) { body += chunk; });
        req.on('end', function () {
            try { resolve(JSON.parse(body)); }
            catch (e) { resolve({}); }
        });
        req.on('error', reject);
    });
}

// ============================================
// HTTP Server
// ============================================
var server = http.createServer(function (req, res) {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

    if (req.method === 'OPTIONS') {
        res.writeHead(204);
        res.end();
        return;
    }

    var parsedUrl = url.parse(req.url, true);

    // API: Get Discord expenses
    if (parsedUrl.pathname === '/api/discord-expenses') {
        try {
            var rawData = fs.existsSync(DATA_FILE)
                ? JSON.parse(fs.readFileSync(DATA_FILE, 'utf-8'))
                : [];
            // Filter out blacklisted (deleted) IDs
            var blacklist = loadProcessedIds();
            var filtered = rawData.filter(function (e) {
                if (!e.discordMsgId) return true;
                return !blacklist.has(e.discordMsgId);
            });
            res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify(filtered));
        } catch (e) {
            res.writeHead(500);
            res.end(JSON.stringify({ error: e.message }));
        }
        return;
    }

    // API: Get all expenses (shared server-side storage)
    if (parsedUrl.pathname === '/api/expenses' && req.method === 'GET') {
        try {
            var data = fs.existsSync(DATA_FILE)
                ? fs.readFileSync(DATA_FILE, 'utf-8')
                : '[]';
            res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(data);
        } catch (e) {
            res.writeHead(500);
            res.end(JSON.stringify({ error: e.message }));
        }
        return;
    }

    // API: Save all expenses — also auto-register deleted discordMsgIds to blacklist
    if (parsedUrl.pathname === '/api/expenses' && req.method === 'POST') {
        parseBody(req).then(function (body) {
            try {
                var expensesArr = body.expenses || body;
                // Detect deleted entries and register their discordMsgId to blacklist
                var oldData = fs.existsSync(DATA_FILE)
                    ? JSON.parse(fs.readFileSync(DATA_FILE, 'utf-8'))
                    : [];
                var newIds = new Set(expensesArr.map(function (e) { return e.id; }));
                var processedIds = loadProcessedIds();
                var changed = false;
                oldData.forEach(function (e) {
                    if (newIds.has(e.id)) return; // not deleted
                    // Get msgId from discordMsgId field or extract from id
                    var msgId = e.discordMsgId;
                    if (!msgId && e.id && e.id.startsWith('discord_')) {
                        var parts = e.id.split('_');
                        if (parts.length >= 2) msgId = parts[1];
                    }
                    if (msgId) {
                        processedIds.add(msgId);
                        changed = true;
                        console.log('🗑️ Blacklisted msgId:', msgId, '(No.' + e.seqNo + ')');
                    }
                });
                if (changed) saveProcessedIds(processedIds);
                fs.writeFileSync(DATA_FILE, JSON.stringify(expensesArr, null, 2), 'utf-8');
                res.writeHead(200, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ success: true }));
            } catch (e) {
                res.writeHead(500);
                res.end(JSON.stringify({ error: e.message }));
            }
        });
        return;
    }

    // API: Get processed msg IDs blacklist
    if (parsedUrl.pathname === '/api/processed-msg-ids' && req.method === 'GET') {
        var ids = loadProcessedIds();
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify([...ids]));
        return;
    }

    // API: Add IDs to blacklist
    if (parsedUrl.pathname === '/api/processed-msg-ids' && req.method === 'POST') {
        parseBody(req).then(function (body) {
            var ids = loadProcessedIds();
            var toAdd = Array.isArray(body) ? body : (body.ids || []);
            toAdd.forEach(function (id) { ids.add(id); });
            saveProcessedIds(ids);
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ success: true, count: ids.size }));
        });
        return;
    }

    // API: Get receipt image
    if (parsedUrl.pathname.startsWith('/receipts/')) {
        var receiptPath = path.join(ROOT, parsedUrl.pathname);
        if (fs.existsSync(receiptPath)) {
            var ext = path.extname(receiptPath);
            res.writeHead(200, { 'Content-Type': MIME_TYPES[ext] || 'application/octet-stream' });
            fs.createReadStream(receiptPath).pipe(res);
        } else {
            res.writeHead(404);
            res.end('Not Found');
        }
        return;
    }

    // ============================================
    // freee API Endpoints
    // ============================================

    // Save freee config
    if (parsedUrl.pathname === '/api/freee/config' && req.method === 'POST') {
        parseBody(req).then(function (body) {
            var config = loadFreeeConfig();
            if (body.client_id !== undefined) config.client_id = body.client_id;
            if (body.client_secret !== undefined) config.client_secret = body.client_secret;
            if (body.company_id !== undefined) config.company_id = body.company_id;
            saveFreeeConfig(config);
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ success: true }));
        }).catch(function (e) {
            res.writeHead(500);
            res.end(JSON.stringify({ error: e.message }));
        });
        return;
    }

    // Get freee config (masked)
    if (parsedUrl.pathname === '/api/freee/config' && req.method === 'GET') {
        var config = loadFreeeConfig();
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({
            has_client_id: !!config.client_id,
            has_client_secret: !!config.client_secret,
            has_token: !!config.access_token,
            company_id: config.company_id || null,
            companies: config.companies || []
        }));
        return;
    }

    // freee OAuth: redirect to authorization
    if (parsedUrl.pathname === '/api/freee/auth') {
        var config = loadFreeeConfig();
        if (!config.client_id) {
            res.writeHead(400);
            res.end(JSON.stringify({ error: 'client_id not configured' }));
            return;
        }
        var authUrl = 'https://accounts.secure.freee.co.jp/public_api/authorize?' +
            querystring.stringify({
                client_id: config.client_id,
                response_type: 'code',
                redirect_uri: 'http://localhost:' + PORT + '/api/freee/callback'
            });
        res.writeHead(302, { 'Location': authUrl });
        res.end();
        return;
    }

    // freee OAuth: callback
    if (parsedUrl.pathname === '/api/freee/callback') {
        var code = parsedUrl.query.code;
        if (!code) {
            res.writeHead(400);
            res.end('Authorization code missing');
            return;
        }
        var config = loadFreeeConfig();
        var postData = querystring.stringify({
            grant_type: 'authorization_code',
            code: code,
            client_id: config.client_id,
            client_secret: config.client_secret,
            redirect_uri: 'http://localhost:' + PORT + '/api/freee/callback'
        });
        httpsRequest({
            hostname: 'accounts.secure.freee.co.jp',
            path: '/public_api/token',
            method: 'POST',
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded',
                'Content-Length': Buffer.byteLength(postData)
            }
        }, postData).then(function (result) {
            if (result.status === 200 && result.data.access_token) {
                config.access_token = result.data.access_token;
                config.refresh_token = result.data.refresh_token;
                config.token_expires = Date.now() + (result.data.expires_in * 1000);
                return freeeApiGet('/api/1/users/me', config.access_token).then(function (meResult) {
                    if (meResult.status === 200 && meResult.data.user && meResult.data.user.companies) {
                        config.companies = meResult.data.user.companies.map(function (c) {
                            return { id: c.id, name: c.display_name || c.name };
                        });
                        if (!config.company_id && config.companies.length > 0) {
                            config.company_id = config.companies[0].id;
                        }
                    }
                    saveFreeeConfig(config);
                    res.writeHead(302, { 'Location': '/?freee_auth=success' });
                    res.end();
                });
            } else {
                res.writeHead(302, { 'Location': '/?freee_auth=error' });
                res.end();
            }
        }).catch(function (e) {
            console.error('freee callback error:', e);
            res.writeHead(302, { 'Location': '/?freee_auth=error' });
            res.end();
        });
        return;
    }

    // freee: Get status
    if (parsedUrl.pathname === '/api/freee/status') {
        getValidToken().then(function (config) {
            if (!config) {
                res.writeHead(200, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ authenticated: false }));
                return;
            }
            return freeeApiGet('/api/1/users/me', config.access_token).then(function (meResult) {
                console.log('freee /users/me status:', meResult.status);
                if (meResult.status === 200 && meResult.data && meResult.data.user) {
                    var user = meResult.data.user;
                    console.log('freee user:', JSON.stringify({ id: user.id, email: user.email }));
                    // Auto-fetch companies if not set
                    if ((!config.company_id || !config.companies || config.companies.length === 0) && user.companies && user.companies.length > 0) {
                        config.companies = user.companies.map(function (c) {
                            return { id: c.id, name: c.display_name || c.name || ('Company ' + c.id) };
                        });
                        config.company_id = config.companies[0].id;
                        saveFreeeConfig(config);
                        console.log('freee: auto-set company_id to', config.company_id);
                    }
                }
                // If still no company_id, try /api/1/companies
                if (!config.company_id) {
                    return freeeApiGet('/api/1/companies', config.access_token).then(function (compResult) {
                        console.log('freee /companies status:', compResult.status, JSON.stringify(compResult.data).substring(0, 200));
                        if (compResult.status === 200 && compResult.data && compResult.data.companies && compResult.data.companies.length > 0) {
                            config.companies = compResult.data.companies.map(function (c) {
                                return { id: c.id, name: c.display_name || c.name || ('Company ' + c.id) };
                            });
                            config.company_id = config.companies[0].id;
                            saveFreeeConfig(config);
                            console.log('freee: auto-set company_id from /companies to', config.company_id);
                        }
                        res.writeHead(200, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({
                            authenticated: true,
                            company_id: config.company_id,
                            companies: config.companies || []
                        }));
                    });
                }
                res.writeHead(200, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({
                    authenticated: meResult.status === 200,
                    company_id: config.company_id,
                    companies: config.companies || []
                }));
            });
        }).catch(function (e) {
            console.error('freee status error:', e);
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ authenticated: false, error: e.message }));
        });
        return;
    }

    // freee: Get master data
    if (parsedUrl.pathname === '/api/freee/master') {
        getValidToken().then(function (config) {
            if (!config) {
                res.writeHead(401);
                res.end(JSON.stringify({ error: 'Not authenticated' }));
                return;
            }
            var companyId = config.company_id;
            var results = {};
            return freeeApiGet('/api/1/walletables?company_id=' + companyId + '&with_balance=false&limit=3000', config.access_token)
                .then(function (r) {
                    results.wallets = (r.status === 200 && r.data.walletables) ? r.data.walletables : [];
                    return freeeApiGet('/api/1/account_items?company_id=' + companyId + '&limit=3000', config.access_token);
                })
                .then(function (r) {
                    results.account_items = (r.status === 200 && r.data.account_items) ? r.data.account_items : [];
                    return freeeApiGet('/api/1/items?company_id=' + companyId + '&limit=3000', config.access_token);
                })
                .then(function (r) {
                    results.items = (r.status === 200 && r.data.items) ? r.data.items : [];
                    return freeeApiGet('/api/1/tags?company_id=' + companyId + '&limit=3000', config.access_token);
                })
                .then(function (r) {
                    results.tags = (r.status === 200 && r.data.tags) ? r.data.tags : [];
                    return freeeApiGet('/api/1/taxes/codes?company_id=' + companyId + '&limit=3000', config.access_token);
                })
                .then(function (r) {
                    results.tax_codes = (r.status === 200 && r.data.taxes) ? r.data.taxes : [];
                    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
                    res.end(JSON.stringify(results));
                });
        }).catch(function (e) {
            res.writeHead(500);
            res.end(JSON.stringify({ error: e.message }));
        });
        return;
    }

    // ============================================
    // Google Sheets Read Endpoint
    // ============================================
    if (parsedUrl.pathname === '/api/sheets/read' && req.method === 'GET') {
        var sheetName = parsedUrl.query.sheet || '';
        var spreadsheetId = '19ftXRJtYQgfVaowzF5KZT6UVl5Ogp0EaHocP8PoYGLM';
        // Use Google Sheets CSV export
        var csvUrl = '/spreadsheets/d/' + spreadsheetId + '/gviz/tq?tqx=out:csv&sheet=' + encodeURIComponent(sheetName);

        var csvOptions = {
            hostname: 'docs.google.com',
            path: csvUrl,
            method: 'GET',
            headers: { 'User-Agent': 'Mozilla/5.0' }
        };

        httpsRequest(csvOptions).then(function (result) {
            if (result.status === 200) {
                // Parse CSV to JSON
                var lines = [];
                var currentLine = '';
                var inQuotes = false;
                var csvText = typeof result.data === 'string' ? result.data : JSON.stringify(result.data);

                for (var ci = 0; ci < csvText.length; ci++) {
                    var ch = csvText[ci];
                    if (ch === '"') {
                        inQuotes = !inQuotes;
                        currentLine += ch;
                    } else if (ch === '\n' && !inQuotes) {
                        lines.push(currentLine);
                        currentLine = '';
                    } else {
                        currentLine += ch;
                    }
                }
                if (currentLine) lines.push(currentLine);

                var rows = lines.map(function (line) {
                    var cols = [];
                    var val = '';
                    var inQ = false;
                    for (var j = 0; j < line.length; j++) {
                        var c = line[j];
                        if (c === '"') {
                            if (inQ && j + 1 < line.length && line[j + 1] === '"') {
                                val += '"';
                                j++;
                            } else {
                                inQ = !inQ;
                            }
                        } else if (c === ',' && !inQ) {
                            cols.push(val);
                            val = '';
                        } else {
                            val += c;
                        }
                    }
                    cols.push(val);
                    return cols;
                });

                res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ rows: rows, count: rows.length }));
            } else {
                console.log('Sheets CSV error:', result.status, typeof result.data === 'string' ? result.data.substring(0, 200) : '');
                res.writeHead(result.status);
                res.end(JSON.stringify({ error: 'Failed to fetch sheet data', status: result.status }));
            }
        }).catch(function (e) {
            console.error('Sheets read error:', e);
            res.writeHead(500);
            res.end(JSON.stringify({ error: e.message }));
        });
        return;
    }

    // ============================================
    // Google Cloud Vision OCR Endpoint
    // ============================================
    if (parsedUrl.pathname === '/api/ocr' && req.method === 'POST') {
        parseBody(req).then(function (body) {
            var imagePath = body.imagePath;
            if (!imagePath) {
                res.writeHead(400, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ error: 'imagePath is required' }));
                return;
            }

            // Resolve image path
            var fullPath = path.join(ROOT, imagePath.replace(/^\//, ''));
            if (!fs.existsSync(fullPath)) {
                res.writeHead(404, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ error: 'Image file not found: ' + imagePath }));
                return;
            }

            // Read service account key
            if (!fs.existsSync(GOOGLE_SA_KEY_FILE)) {
                res.writeHead(500, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ error: 'Google service account key file not found' }));
                return;
            }
            var saKey = JSON.parse(fs.readFileSync(GOOGLE_SA_KEY_FILE, 'utf-8'));

            // Create JWT
            var now = Math.floor(Date.now() / 1000);
            var jwtHeader = { alg: 'RS256', typ: 'JWT' };
            var jwtClaim = {
                iss: saKey.client_email,
                scope: 'https://www.googleapis.com/auth/cloud-vision',
                aud: 'https://oauth2.googleapis.com/token',
                iat: now,
                exp: now + 3600
            };
            function base64url(obj) {
                return Buffer.from(JSON.stringify(obj)).toString('base64')
                    .replace(/=/g, '').replace(/\+/g, '-').replace(/\//g, '_');
            }
            var signInput = base64url(jwtHeader) + '.' + base64url(jwtClaim);
            var sign = crypto.createSign('RSA-SHA256');
            sign.update(signInput);
            var signature = sign.sign(saKey.private_key, 'base64')
                .replace(/=/g, '').replace(/\+/g, '-').replace(/\//g, '_');
            var jwt = signInput + '.' + signature;

            // Get access token
            var tokenData = querystring.stringify({
                grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
                assertion: jwt
            });
            var tokenOptions = {
                hostname: 'oauth2.googleapis.com',
                path: '/token',
                method: 'POST',
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                    'Content-Length': Buffer.byteLength(tokenData)
                }
            };

            httpsRequest(tokenOptions, tokenData).then(function (tokenResult) {
                if (tokenResult.status !== 200 || !tokenResult.data.access_token) {
                    res.writeHead(500, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify({ error: 'Failed to get access token' }));
                    return;
                }
                var accessToken = tokenResult.data.access_token;

                // Read image file and encode to base64
                var imageBuffer = fs.readFileSync(fullPath);
                var imageBase64 = imageBuffer.toString('base64');

                // Call Vision API
                var visionBody = JSON.stringify({
                    requests: [{
                        image: { content: imageBase64 },
                        features: [{ type: 'TEXT_DETECTION' }]
                    }]
                });
                var visionOptions = {
                    hostname: 'vision.googleapis.com',
                    path: '/v1/images:annotate',
                    method: 'POST',
                    headers: {
                        'Authorization': 'Bearer ' + accessToken,
                        'Content-Type': 'application/json',
                        'Content-Length': Buffer.byteLength(visionBody)
                    }
                };

                return httpsRequest(visionOptions, visionBody);
            }).then(function (visionResult) {
                if (!visionResult) return;
                if (visionResult.status !== 200) {
                    console.error('Vision API error:', visionResult.status, visionResult.data);
                    res.writeHead(500, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify({ error: 'Vision API error', details: visionResult.data }));
                    return;
                }

                var fullText = '';
                if (visionResult.data.responses && visionResult.data.responses[0] && visionResult.data.responses[0].fullTextAnnotation) {
                    fullText = visionResult.data.responses[0].fullTextAnnotation.text;
                }

                // Normalize text: convert backslash-yen (\¥) to yen sign
                var normalizedText = fullText.replace(/\\[¥￥]/g, '¥');

                // Detect Amazon receipt (no order number on receipt)
                var isAmazon = /amazon|アマゾン/i.test(fullText) || /amazon|アマゾン/i.test(body.payment || '') || /amazon|アマゾン/i.test(body.supplier || '');

                // Order number patterns (skip for Amazon - they don't print order numbers)
                var orderNumber = '';
                if (!isAmazon) {
                    var orderPatterns = [
                        /([a-z]+-\d{6,})/i,                       // Store-prefixed: joshin-12345678
                        /注文番号\n[^\n]*\n([A-Za-z0-9\-]+)/,     // Multi-line with date between
                        /注文番号\n([A-Za-z0-9\-]+)/,             // Multi-line: 注文番号\nvalue
                        /注文番号[:\s：]+([A-Za-z0-9\-]+)/,       // Same line: 注文番号: value
                        /注文ID[:\s：]+([A-Za-z0-9\-]+)/,
                        /Order\s*(?:Number|ID|#)[:\s]*([A-Za-z0-9\-]+)/i,
                        /(?:受注|伝票)[番号]*[:\s：]+([A-Za-z0-9\-]+)/,
                        /(\d{3}-\d{7}-\d{7})/                     // Amazon 注文ID (just in case)
                    ];
                    for (var i = 0; i < orderPatterns.length; i++) {
                        var m = fullText.match(orderPatterns[i]);
                        if (m) { orderNumber = m[1]; break; }
                    }
                }

                // Total amount patterns (Amazon-aware)
                var totalAmount = '';
                var totalPatterns = [
                    /支払い金額\n([0-9,]+)円/,                 // 支払い金額\n17,580円
                    /小計[（\(税込\)）]*\n([0-9,]+)円/,        // 小計(税込)\n6,480円
                    /注文金額\n(?:商品合計[^\n]*\n(?:送料[^\n]*\n)?)?.*?([0-9,]+)円/s,
                    /商品合計[^0-9]*([0-9,]+)円/,
                    /([0-9,]+)円\s*[x×]\s*\d+/,               // 62,800円 x1
                    /合計[金額（）\(\)税込]*[:\s：]*[¥￥]?\s*([0-9,]+)/,
                    /お支払い[金額合計]*[:\s：]*[¥￥]?\s*([0-9,]+)/,
                    /ご請求[金額額]*[:\s：]*[¥￥]?\s*([0-9,]+)/,
                    /単価[:\s：]*([0-9,]+)円/,                 // 単価: 17,580円
                    /(?:Total|TOTAL)[:\s]*[¥￥$]?\s*([0-9,]+)/i,
                    /([0-9,]+)円\n0円\n/,                      // Amount followed by 送料0円
                    // Amazon specific: ¥13,700 or ¥ 13,700
                    /[¥￥]\s*([0-9,]{4,})/,
                    // Plain number patterns for Amazon (e.g. 13,700 on its own line after item)
                    /お届け先[\s\S]{1,200}?([0-9]{2,3},[0-9]{3})(?:\s|\n|$)/,
                ];
                for (var j = 0; j < totalPatterns.length; j++) {
                    var m2 = normalizedText.match(totalPatterns[j]);
                    if (m2) { totalAmount = m2[1].replace(/,/g, ''); break; }
                }

                console.log('OCR result - Order:', orderNumber, 'Total:', totalAmount);
                res.writeHead(200, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({
                    success: true,
                    orderNumber: orderNumber,
                    totalAmount: totalAmount,
                    fullText: fullText
                }));
            }).catch(function (err) {
                console.error('OCR error:', err);
                res.writeHead(500, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ error: err.message }));
            });
        });
        return;
    }

    // ============================================
    // Google Sheets Write Endpoint
    // ============================================
    if (parsedUrl.pathname === '/api/sheets/write' && req.method === 'POST') {
        parseBody(req).then(function (body) {
            // Read service account key
            if (!fs.existsSync(GOOGLE_SA_KEY_FILE)) {
                res.writeHead(500, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ error: 'Google service account key file not found. Place google_service_account.json in the project root.' }));
                return;
            }
            var saKey = JSON.parse(fs.readFileSync(GOOGLE_SA_KEY_FILE, 'utf-8'));

            // Create JWT for Google Sheets API
            var now = Math.floor(Date.now() / 1000);
            var jwtHeader = { alg: 'RS256', typ: 'JWT' };
            var jwtClaim = {
                iss: saKey.client_email,
                scope: 'https://www.googleapis.com/auth/spreadsheets',
                aud: 'https://oauth2.googleapis.com/token',
                iat: now,
                exp: now + 3600
            };

            function base64url(obj) {
                return Buffer.from(JSON.stringify(obj)).toString('base64')
                    .replace(/=/g, '').replace(/\+/g, '-').replace(/\//g, '_');
            }
            var signInput = base64url(jwtHeader) + '.' + base64url(jwtClaim);
            var sign = crypto.createSign('RSA-SHA256');
            sign.update(signInput);
            var signature = sign.sign(saKey.private_key, 'base64')
                .replace(/=/g, '').replace(/\+/g, '-').replace(/\//g, '_');
            var jwt = signInput + '.' + signature;

            // Exchange JWT for access token
            var tokenData = querystring.stringify({
                grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
                assertion: jwt
            });
            var tokenOptions = {
                hostname: 'oauth2.googleapis.com',
                path: '/token',
                method: 'POST',
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                    'Content-Length': Buffer.byteLength(tokenData)
                }
            };

            httpsRequest(tokenOptions, tokenData).then(function (tokenResult) {
                if (tokenResult.status !== 200 || !tokenResult.data.access_token) {
                    console.error('Google token error:', tokenResult.status, tokenResult.data);
                    res.writeHead(500, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify({ error: 'Failed to get Google access token', details: tokenResult.data }));
                    return;
                }
                var accessToken = tokenResult.data.access_token;
                var spreadsheetId = body.spreadsheetId || '19ftXRJtYQgfVaowzF5KZT6UVl5Ogp0EaHocP8PoYGLM';
                var rowsAF = body.rowsAF || [];
                var rowsHK = body.rowsHK || [];

                // First, get the actual sheet name from spreadsheet metadata
                var metaOptions = {
                    hostname: 'sheets.googleapis.com',
                    path: '/v4/spreadsheets/' + spreadsheetId + '?fields=sheets.properties.title',
                    method: 'GET',
                    headers: {
                        'Authorization': 'Bearer ' + accessToken
                    }
                };

                return httpsRequest(metaOptions).then(function (metaResult) {
                    var sheetName = 'Sheet1';
                    if (metaResult.status === 200 && metaResult.data.sheets && metaResult.data.sheets.length > 0) {
                        sheetName = metaResult.data.sheets[0].properties.title;
                        console.log('Using sheet name:', sheetName);
                    }

                    // Step 1: Clear A-F and H-K (preserve G column)
                    var clearRangeAF = encodeURIComponent(sheetName + '!A2:F');
                    var clearRangeHK = encodeURIComponent(sheetName + '!H2:L');
                    var clearHeaders = {
                        'Authorization': 'Bearer ' + accessToken,
                        'Content-Type': 'application/json',
                        'Content-Length': 2
                    };
                    var clearOptionsAF = {
                        hostname: 'sheets.googleapis.com',
                        path: '/v4/spreadsheets/' + spreadsheetId + '/values/' + clearRangeAF + ':clear',
                        method: 'POST',
                        headers: clearHeaders
                    };
                    var clearOptionsHK = {
                        hostname: 'sheets.googleapis.com',
                        path: '/v4/spreadsheets/' + spreadsheetId + '/values/' + clearRangeHK + ':clear',
                        method: 'POST',
                        headers: clearHeaders
                    };

                    return httpsRequest(clearOptionsAF, '{}').then(function () {
                        return httpsRequest(clearOptionsHK, '{}');
                    }).then(function (clearResult) {
                        console.log('Sheets clear result:', clearResult.status);

                        // Step 2: Write A-F data starting from A2
                        var writeDataAF = JSON.stringify({ values: rowsAF });
                        var writeRangeAF = encodeURIComponent(sheetName + '!A2');
                        var writeOptionsAF = {
                            hostname: 'sheets.googleapis.com',
                            path: '/v4/spreadsheets/' + spreadsheetId + '/values/' + writeRangeAF + '?valueInputOption=USER_ENTERED',
                            method: 'PUT',
                            headers: {
                                'Authorization': 'Bearer ' + accessToken,
                                'Content-Type': 'application/json',
                                'Content-Length': Buffer.byteLength(writeDataAF)
                            }
                        };

                        return httpsRequest(writeOptionsAF, writeDataAF);
                    }).then(function (writeResultAF) {
                        if (writeResultAF.status !== 200) {
                            throw new Error('Failed to write A-F: ' + JSON.stringify(writeResultAF.data));
                        }

                        // Step 3: Write H-L data starting from H2
                        var writeDataHK = JSON.stringify({ values: rowsHK });
                        var writeRangeHK = encodeURIComponent(sheetName + '!H2');
                        var writeOptionsHK = {
                            hostname: 'sheets.googleapis.com',
                            path: '/v4/spreadsheets/' + spreadsheetId + '/values/' + writeRangeHK + '?valueInputOption=USER_ENTERED',
                            method: 'PUT',
                            headers: {
                                'Authorization': 'Bearer ' + accessToken,
                                'Content-Type': 'application/json',
                                'Content-Length': Buffer.byteLength(writeDataHK)
                            }
                        };

                        return httpsRequest(writeOptionsHK, writeDataHK);
                    }).then(function (writeResultHK) {
                        if (writeResultHK.status === 200) {
                            console.log('Sheets write success (A-F, H-K)');
                            res.writeHead(200, { 'Content-Type': 'application/json' });
                            res.end(JSON.stringify({ success: true }));
                        } else {
                            console.error('Sheets write error H-K:', writeResultHK.status, writeResultHK.data);
                            res.writeHead(writeResultHK.status, { 'Content-Type': 'application/json' });
                            res.end(JSON.stringify({ error: 'Failed to write H-K', details: writeResultHK.data }));
                        }
                    });
                });
            }).catch(function (err) {
                console.error('Sheets write error:', err);
                res.writeHead(500, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ error: err.message }));
            });
        });
        return;
    }

    // freee: Create manual journals (振替伝票)
    if (parsedUrl.pathname === '/api/freee/deals' && req.method === 'POST') {
        Promise.all([getValidToken(), parseBody(req)]).then(function (arr) {
            var config = arr[0];
            var body = arr[1];
            if (!config) {
                res.writeHead(401);
                res.end(JSON.stringify({ error: 'Not authenticated' }));
                return;
            }
            var companyId = config.company_id;
            var results = [];
            var chain = Promise.resolve();

            console.log('freee journal params:', JSON.stringify({
                debit_account_item_id: body.account_item_id,
                debit_tax_code: body.tax_code,
                credit_account_item_id: body.credit_account_item_id,
                credit_tax_code: body.credit_tax_code,
                item_id: body.item_id,
                tag_ids: body.tag_ids
            }));

            body.expenses.forEach(function (exp, i) {
                chain = chain.then(function () {
                    var receiptPromise = Promise.resolve(null);
                    if (exp.receipt) {
                        console.log('Uploading receipt for exp:', exp.id);
                        receiptPromise = freeeUploadReceipt(config.access_token, companyId, exp.receipt);
                    }
                    return receiptPromise.then(function (receiptId) {
                        var debitEntry = {
                            entry_side: 'debit',
                            account_item_id: parseInt(body.account_item_id),
                            tax_code: parseInt(body.tax_code),
                            amount: parseInt(exp.amount),
                            description: exp.description
                        };
                        if (body.item_id) debitEntry.item_id = parseInt(body.item_id);
                        if (body.tag_ids && body.tag_ids.length > 0) {
                            debitEntry.tag_ids = body.tag_ids.map(function (id) { return parseInt(id); });
                        }
                        var creditEntry = {
                            entry_side: 'credit',
                            account_item_id: parseInt(body.credit_account_item_id),
                            tax_code: parseInt(body.credit_tax_code),
                            amount: parseInt(exp.amount),
                            description: exp.description
                        };
                        var journalBody = {
                            company_id: parseInt(companyId),
                            issue_date: exp.date,
                            adjustment: false,
                            details: [debitEntry, creditEntry]
                        };
                        if (receiptId) journalBody.receipt_ids = [receiptId];
                        console.log('=== freee manual_journal body ===');
                        console.log(JSON.stringify(journalBody, null, 2));

                        return freeeApiPost('/api/1/manual_journals', config.access_token, journalBody).then(function (r) {
                            console.log('freee journal result:', r.status, JSON.stringify(r.data).substring(0, 300));
                            results.push({
                                expense_id: exp.id,
                                seq_no: exp.seqNo,
                                success: r.status === 201 || r.status === 200,
                                status: r.status,
                                deal_id: (r.data && r.data.manual_journal) ? r.data.manual_journal.id : null,
                                error: r.status >= 400 ? r.data : null
                            });
                            if (i < body.expenses.length - 1) {
                                return new Promise(function (resolve) { setTimeout(resolve, 500); });
                            }
                        });
                    });
                });
            });

            return chain.then(function () {
                res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ results: results }));
            });
        }).catch(function (e) {
            console.error('freee deals error:', e);
            res.writeHead(500);
            res.end(JSON.stringify({ error: e.message }));
        });
        return;
    }

    // Static files
    var filePath = parsedUrl.pathname === '/' ? '/index.html' : parsedUrl.pathname;
    filePath = path.join(ROOT, filePath);

    if (!filePath.startsWith(ROOT)) {
        res.writeHead(403);
        res.end('Forbidden');
        return;
    }

    if (fs.existsSync(filePath) && fs.statSync(filePath).isFile()) {
        var ext = path.extname(filePath);
        res.writeHead(200, {
            'Content-Type': MIME_TYPES[ext] || 'application/octet-stream',
            'Cache-Control': 'no-cache, no-store, must-revalidate',
            'Pragma': 'no-cache',
            'Expires': '0'
        });
        fs.createReadStream(filePath).pipe(res);
    } else {
        res.writeHead(404);
        res.end('Not Found');
    }
});

server.listen(PORT, '127.0.0.1', function () {
    console.log('');
    console.log('\u2554\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2557');
    console.log('\u2551   \ud83c\udf10 \u7d4c\u8cbb\u7cbe\u7b97\u30a2\u30b7\u30b9\u30bf\u30f3\u30c8 Web\u30b5\u30fc\u30d0\u30fc    \u2551');
    console.log('\u2560\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2563');
    console.log('\u2551   URL: http://localhost:' + PORT + '             \u2551');
    console.log('\u2551   \u30b9\u30c6\u30fc\u30bf\u30b9: \u8d77\u52d5\u4e2d \u2705                  \u2551');
    console.log('\u2551   freee API\u9023\u643a: \u6709\u52b9 \ud83d\udd17                 \u2551');
    console.log('\u255a\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u255d');
    console.log('');
});
