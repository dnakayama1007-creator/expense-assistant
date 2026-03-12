// ============================================
// Cloudflare Tunnel Starter + URL Writer
// Starts cloudflared, captures URL, writes to Google Sheets & Discord
// ============================================
require('dotenv').config();
const { spawn } = require('child_process');
const https = require('https');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const querystring = require('querystring');

const ROOT = __dirname;
const GOOGLE_SA_KEY_FILE = path.join(ROOT, 'google_service_account.json');
const SPREADSHEET_ID = '19ftXRJtYQgfVaowzF5KZT6UVl5Ogp0EaHocP8PoYGLM';
const SHEET_NAME = 'URL';
const CELL = 'A1';
const DISCORD_BOT_TOKEN = process.env.DISCORD_BOT_TOKEN;
const DISCORD_CHANNEL_ID = '1477594706249777224';

function httpsRequest(options, postData) {
    return new Promise(function (resolve, reject) {
        var req = https.request(options, function (res) {
            var body = '';
            res.on('data', function (chunk) { body += chunk; });
            res.on('end', function () {
                try { resolve({ status: res.statusCode, data: JSON.parse(body) }); }
                catch (e) { resolve({ status: res.statusCode, data: body }); }
            });
        });
        req.on('error', reject);
        if (postData) req.write(postData);
        req.end();
    });
}

function writeUrlToSheet(tunnelUrl) {
    if (!fs.existsSync(GOOGLE_SA_KEY_FILE)) {
        console.error('❌ Google service account key file not found');
        return;
    }
    var saKey = JSON.parse(fs.readFileSync(GOOGLE_SA_KEY_FILE, 'utf-8'));

    // Create JWT
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
            console.error('❌ Failed to get access token:', tokenResult.data);
            return;
        }
        var accessToken = tokenResult.data.access_token;

        // Write URL to sheet
        var writeData = JSON.stringify({ values: [[tunnelUrl]] });
        var range = encodeURIComponent(SHEET_NAME + '!' + CELL);
        var writeOptions = {
            hostname: 'sheets.googleapis.com',
            path: '/v4/spreadsheets/' + SPREADSHEET_ID + '/values/' + range + '?valueInputOption=USER_ENTERED',
            method: 'PUT',
            headers: {
                'Authorization': 'Bearer ' + accessToken,
                'Content-Type': 'application/json',
                'Content-Length': Buffer.byteLength(writeData)
            }
        };

        return httpsRequest(writeOptions, writeData);
    }).then(function (result) {
        if (result && result.status === 200) {
            console.log('✅ URLをGoogleスプレッドシートに書き込みました: ' + tunnelUrl);
        } else {
            console.error('❌ Sheets write error:', result ? result.data : 'no result');
        }
    }).catch(function (err) {
        console.error('❌ Error writing URL:', err.message);
    });
}

// Start cloudflared
console.log('🚀 Cloudflare Tunnel を起動中...');
var cloudflared = spawn(path.join(ROOT, 'cloudflared.exe'), ['tunnel', '--url', 'http://127.0.0.1:3000'], {
    stdio: ['ignore', 'pipe', 'pipe']
});

var urlFound = false;

function processOutput(data) {
    var text = data.toString();
    process.stderr.write(text);

    if (!urlFound) {
        var match = text.match(/https:\/\/[a-z0-9\-]+\.trycloudflare\.com/);
        if (match) {
            urlFound = true;
            var url = match[0];
            console.log('\n');
            console.log('╔══════════════════════════════════════════════════════╗');
            console.log('║   🌐 外部アクセスURL                                ║');
            console.log('║   ' + url + '  ║');
            console.log('╚══════════════════════════════════════════════════════╝');
            console.log('\n📝 GoogleスプレッドシートにURLを書き込み中...');
            writeUrlToSheet(url);
            console.log('📨 DiscordにURLを投稿中...');
            postUrlToDiscord(url);
        }
    }
}

cloudflared.stdout.on('data', processOutput);
cloudflared.stderr.on('data', processOutput);

cloudflared.on('close', function (code) {
    console.log('Cloudflared exited with code:', code);
});

process.on('SIGINT', function () {
    cloudflared.kill();
    process.exit();
});

// ============================================
// Post URL to Discord
// ============================================
function postUrlToDiscord(tunnelUrl) {
    if (!DISCORD_BOT_TOKEN) {
        console.error('❌ DISCORD_BOT_TOKEN not found in .env');
        return;
    }
    var messageBody = JSON.stringify({
        content: '🌐 **経費アシスタント 外部アクセスURL更新**\n' + tunnelUrl
    });
    var options = {
        hostname: 'discord.com',
        path: '/api/v10/channels/' + DISCORD_CHANNEL_ID + '/messages',
        method: 'POST',
        headers: {
            'Authorization': 'Bot ' + DISCORD_BOT_TOKEN,
            'Content-Type': 'application/json',
            'Content-Length': Buffer.byteLength(messageBody)
        }
    };
    httpsRequest(options, messageBody).then(function (result) {
        if (result.status === 200 || result.status === 201) {
            console.log('✅ DiscordにURLを投稿しました');
        } else {
            console.error('❌ Discord post error:', result.status, result.data);
        }
    }).catch(function (err) {
        console.error('❌ Discord post error:', err.message);
    });
}
