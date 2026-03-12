// ============================================
// 経費精算アシスタント - Discord Bot
// ============================================
require('dotenv').config();
const { Client, GatewayIntentBits, Partials, EmbedBuilder } = require('discord.js');
const fs = require('fs');
const path = require('path');
const https = require('https');
const http = require('http');

// ============================================
// Config
// ============================================
const DATA_FILE = path.join(__dirname, 'discord_expenses.json');
const IMAGES_DIR = path.join(__dirname, 'receipts');

// Ensure directories exist
if (!fs.existsSync(IMAGES_DIR)) {
    fs.mkdirSync(IMAGES_DIR, { recursive: true });
}

// ============================================
// Data Management
// ============================================
function loadExpenses() {
    try {
        if (fs.existsSync(DATA_FILE)) {
            return JSON.parse(fs.readFileSync(DATA_FILE, 'utf-8'));
        }
    } catch (e) {
        console.error('Error loading expenses:', e);
    }
    return [];
}

function saveExpenses(expenses) {
    fs.writeFileSync(DATA_FILE, JSON.stringify(expenses, null, 2), 'utf-8');
}

// ============================================
// Text Parser (same logic as web app)
// ============================================
function parseMemoLine(line) {
    let date = null;
    let amount = null;
    let category = null;
    let store = null;
    let description = line;

    const now = new Date();
    const currentYear = now.getFullYear();

    // Normalize full-width numbers
    let normalized = line;
    normalized = normalized.replace(/[０-９]/g, (s) => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
    normalized = normalized.replace(/，/g, ',');
    normalized = normalized.replace(/￥/g, '¥');

    // --- Extract Date ---
    // YYYY/M/D or YYYY-M-D
    let dateMatch = normalized.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    if (dateMatch) {
        date = `${dateMatch[1]}-${dateMatch[2].padStart(2, '0')}-${dateMatch[3].padStart(2, '0')}`;
        description = description.replace(dateMatch[0], '').trim();
    }

    // M/D
    if (!date) {
        dateMatch = normalized.match(/(?:^|\s)(\d{1,2})\/(\d{1,2})(?:\s|$|[^\d])/);
        if (dateMatch) {
            const m = parseInt(dateMatch[1]);
            const d = parseInt(dateMatch[2]);
            if (m >= 1 && m <= 12 && d >= 1 && d <= 31) {
                date = `${currentYear}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
                description = description.replace(dateMatch[0], ' ').trim();
            }
        }
    }

    // M月D日
    if (!date) {
        dateMatch = normalized.match(/(\d{1,2})月(\d{1,2})日/);
        if (dateMatch) {
            const m = parseInt(dateMatch[1]);
            const d = parseInt(dateMatch[2]);
            if (m >= 1 && m <= 12 && d >= 1 && d <= 31) {
                date = `${currentYear}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
                description = description.replace(dateMatch[0], '').trim();
            }
        }
    }

    // Default to today
    if (!date) {
        const today = new Date();
        date = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
    }

    // --- Extract Amount ---
    let amountMatch = normalized.match(/[¥\\]\s*([0-9]{1,3}(?:,?[0-9]{3})*)/);
    if (amountMatch) {
        amount = parseInt(amountMatch[1].replace(/,/g, ''));
        description = description.replace(amountMatch[0], '').trim();
    }

    if (!amount) {
        amountMatch = normalized.match(/([0-9]{1,3}(?:,?[0-9]{3})*)\s*円/);
        if (amountMatch) {
            amount = parseInt(amountMatch[1].replace(/,/g, ''));
            description = description.replace(amountMatch[0], '').trim();
        }
    }

    if (!amount) {
        // Find ALL standalone numbers (not adjacent to letters) and pick the LAST one
        // This avoids treating product model numbers like "9060xt" as amounts
        const allMatches = [...normalized.matchAll(/(?:^|[\s,、])([0-9]{3,7})(?=[\s,、。]|$)/g)];
        // Filter out numbers that are part of alphanumeric strings (e.g., "9060xt", "16G")
        const validMatches = allMatches.filter(m => {
            const pos = m.index + m[0].length;
            const before = m.index > 0 ? normalized[m.index - 1] : '';
            const after = pos < normalized.length ? normalized[pos] : '';
            // Skip if adjacent to a letter (part of a model number)
            return !/[a-zA-Z]/.test(before) && !/[a-zA-Z]/.test(after);
        });
        // Pick the last valid match (most likely to be the price)
        if (validMatches.length > 0) {
            const lastMatch = validMatches[validMatches.length - 1];
            const val = parseInt(lastMatch[1]);
            if (val >= 100 && val < 10000000) {
                amount = val;
                description = description.replace(lastMatch[1], '').trim();
            }
        }
    }

    if (!amount) return null;

    // --- Detect Category ---
    const categoryKeywords = {
        '交通費': ['タクシー', 'タクシ', '電車', '新幹線', 'バス', '地下鉄', '駅', '交通', '定期', 'Suica', 'PASMO', 'IC', '乗車', '往復', '片道', '→', '➡', 'JR', '私鉄', '飛行機', '航空'],
        '宿泊費': ['ホテル', '旅館', '宿泊', '宿', '民泊', 'ビジネスホテル', '泊'],
        '会議費': ['スタバ', 'スターバックス', 'カフェ', 'コーヒー', '喫茶', '打ち合わせ', '打合せ', '会議', 'ミーティング', 'MTG', 'ドトール', 'タリーズ'],
        '接待交際費': ['レストラン', '居酒屋', '飲み会', '会食', '接待', '食事', 'ランチ', 'ディナー', '懇親', '忘年会', '新年会', '歓迎会', '送別会', '焼肉', '寿司'],
        '通信費': ['通信', 'Wi-Fi', 'WiFi', '電話', '携帯', 'SIM', 'ポケット', 'インターネット'],
        '消耗品費': ['文具', 'コピー', '用紙', 'ペン', 'ノート', 'USB', 'ケーブル', '電池', '消耗品', '事務用品'],
        '書籍・研修費': ['本', '書籍', '雑誌', '研修', 'セミナー', '講座', '勉強会', '技術書']
    };

    const categoryIcons = {
        '交通費': '🚃', '宿泊費': '🏨', '会議費': '☕',
        '接待交際費': '🍽️', '通信費': '📱', '消耗品費': '🖊️',
        '書籍・研修費': '📚', 'その他': '📦'
    };

    for (const [cat, keywords] of Object.entries(categoryKeywords)) {
        for (const keyword of keywords) {
            if (line.includes(keyword)) {
                category = cat;
                break;
            }
        }
        if (category) break;
    }
    if (!category) category = 'その他';

    // --- Extract Store (購入店舗) ---
    // Look for store patterns at the beginning of the original line
    // Pattern: "StoreName(detail)" or "StoreName（detail）" at start
    const storeMatch = line.match(/^([A-Za-z\u3040-\u9FFF]+[\(（][^\)）]+[\)）])/)
        || line.match(/^([A-Za-z]+(?:\.[A-Za-z]+)*)\s/)
        || line.match(/^(楽天|Amazon|アマゾン|Yahoo|ヤフー|メルカリ|ヨドバシ|ビック|ジョーシン|ケーズ|エディオン|ドンキ|コストコ|IKEA|ニトリ|ダイソー|セリア|モノタロウ|AliExpress)\s?/i);
    if (storeMatch) {
        store = storeMatch[1] || storeMatch[0];
        store = store.trim();
        // Remove store from description
        description = description.replace(storeMatch[0], '').trim();
    }

    // Clean description
    description = description.replace(/\s+/g, ' ').trim();
    description = description.replace(/^[\s,、。・\-\/]+|[\s,、。・\-\/]+$/g, '').trim();
    if (!description) {
        description = line.replace(/[¥\\]\s*[0-9,.]+|[0-9,.]+\s*円/g, '').replace(/\s+/g, ' ').trim() || 'Discord経費';
    }

    return {
        date,
        amount,
        category,
        categoryIcon: categoryIcons[category] || '📦',
        store: store || '',
        description,
        originalLine: line
    };
}

// ============================================
// Image Download
// ============================================
function downloadImage(url, filename) {
    return new Promise((resolve, reject) => {
        const filePath = path.join(IMAGES_DIR, filename);
        const file = fs.createWriteStream(filePath);
        const protocol = url.startsWith('https') ? https : http;

        protocol.get(url, (response) => {
            response.pipe(file);
            file.on('finish', () => {
                file.close();
                resolve(filePath);
            });
        }).on('error', (err) => {
            fs.unlink(filePath, () => { });
            reject(err);
        });
    });
}

// ============================================
// Discord Bot
// ============================================
const client = new Client({
    intents: [
        GatewayIntentBits.Guilds,
        GatewayIntentBits.GuildMessages,
        GatewayIntentBits.MessageContent,
        GatewayIntentBits.DirectMessages,
    ],
    partials: [Partials.Channel, Partials.Message],
});

client.once('ready', async () => {
    console.log('');
    console.log('╔══════════════════════════════════════════╗');
    console.log('║   💰 経費精算アシスタント Bot 起動完了   ║');
    console.log('╠══════════════════════════════════════════╣');
    console.log(`║   Bot: ${client.user.tag.padEnd(32)}║`);
    console.log(`║   サーバー数: ${String(client.guilds.cache.size).padEnd(26)}║`);
    console.log('║   ステータス: オンライン ✅              ║');
    console.log('╚══════════════════════════════════════════╝');
    console.log('');
    console.log('📝 使い方:');
    console.log('  テキスト送信 → 「2/24 タクシー 1500円」');
    console.log('  画像送信     → 領収書の写真を送信');
    console.log('  !一覧        → 最近の経費を表示');
    console.log('  !合計        → 今月の合計を表示');
    console.log('  !ヘルプ      → コマンド一覧');
    console.log('');

    // Fetch missed messages while bot was offline
    await fetchMissedMessages();
});

// ============================================
// Fetch past messages that were missed while offline
// ============================================
async function fetchMissedMessages() {
    try {
        const CHANNEL_ID = '1477594706249777224';
        const channel = await client.channels.fetch(CHANNEL_ID);
        if (!channel) return;

        const existingExpenses = loadExpenses();

        // Build a set of ALL known Discord message IDs from existing data
        const processedMsgIds = new Set();

        // 1. Load blacklist (deleted expenses' msgIds)
        const PROCESSED_IDS_FILE = path.join(__dirname, 'processed_msg_ids.json');
        try {
            if (fs.existsSync(PROCESSED_IDS_FILE)) {
                const blacklist = JSON.parse(fs.readFileSync(PROCESSED_IDS_FILE, 'utf-8'));
                blacklist.forEach(id => processedMsgIds.add(id));
                console.log(`📋 ブラックリスト読み込み: ${blacklist.length}件`);
            }
        } catch (e) { console.error('blacklist load error:', e.message); }

        // 2. Add currently existing expense msgIds
        existingExpenses.forEach(e => {
            if (e.discordMsgId) processedMsgIds.add(e.discordMsgId);
            // Also extract from id format: discord_MSGID_random
            if (e.id) {
                const parts = e.id.split('_');
                if (parts.length >= 2 && parts[0] === 'discord') {
                    processedMsgIds.add(parts[1]);
                }
            }
        });

        // Fetch last 100 messages (Discord API max per request)
        const messages = await channel.messages.fetch({ limit: 100 });
        const sorted = [...messages.values()].sort((a, b) => a.createdTimestamp - b.createdTimestamp);

        // ボット自身の最後のメッセージのタイムスタンプを取得 (ユーザー提案の重複防止策)
        let lastBotReplyTime = 0;
        for (const msg of sorted) {
            if (msg.author.id === client.user.id) {
                lastBotReplyTime = Math.max(lastBotReplyTime, msg.createdTimestamp);
            }
        }
        
        if (lastBotReplyTime > 0) {
            console.log(`🤖 ボット最終応答時刻: ${new Date(lastBotReplyTime).toLocaleString()} (これ以前の投稿は無視されます)`);
        }

        let newCount = 0;
        for (const msg of sorted) {
            if (msg.author.bot) continue;
            if (msg.content.startsWith('!')) continue;

            // ユーザー提案ロジック: ボットの最終応答より前のメッセージはスキップ
            if (lastBotReplyTime > 0 && msg.createdTimestamp <= lastBotReplyTime) {
                continue;
            }

            // Skip if already processed
            if (processedMsgIds.has(msg.id)) continue;

            // Mark as processed immediately to prevent duplicates within this run
            processedMsgIds.add(msg.id);

            // Try to process as expense (silent - no reply)
            const added = await processExpenseMessage(msg, true);
            if (added) newCount++;
        }

        if (newCount > 0) {
            console.log(`✅ オフライン中に受信した ${newCount} 件のメッセージを取り込みました`);
        } else {
            console.log('📋 未取り込みのメッセージはありませんでした');
        }
    } catch (err) {
        console.error('過去メッセージ取得エラー:', err.message);
    }
}

// ============================================
// 処理済みmsgIdをブラックリストに登録
// ============================================
function registerProcessedMsgId(msgId) {
    const PIDS_FILE = require('path').join(__dirname, 'processed_msg_ids.json');
    try {
        let ids = [];
        if (fs.existsSync(PIDS_FILE)) ids = JSON.parse(fs.readFileSync(PIDS_FILE, 'utf-8'));
        if (!ids.includes(msgId)) {
            ids.push(msgId);
            fs.writeFileSync(PIDS_FILE, JSON.stringify(ids, null, 2), 'utf-8');
        }
    } catch (e) { console.error('registerProcessedMsgId error:', e.message); }
}

// ============================================
// 購入者マッピング (Discord表示名 → 購入者名)
// ============================================
function getBuyer(author) {
    const name = author.global_name || author.globalName || author.username || '';
    const username = author.username || '';
    const map = {
        'かなっぺ': 'かなっぺ',
        'kanatsupe0505': 'かなっぺ',
        'みーフォン': 'パパ＆し',
        'mihuon': 'パパ＆し',
        '新泥': 'みーくん',
        'xinni0078': 'みーくん',
        'ぱっせん': 'ぱっせん',
        'patsusen1277': 'ぱっせん'
    };
    return map[name] || map[username] || 'その他';
}

// ============================================
// Process a single expense message (shared logic)
// silent=true: no Discord reply (for catch-up processing)
// ============================================
async function processExpenseMessage(message, silent = false) {
    let receiptPath = null;
    const buyer = getBuyer(message.author);

    // Handle image attachments
    if (message.attachments.size > 0) {
        const imageAttachment = message.attachments.find(a =>
            a.contentType && a.contentType.startsWith('image/')
        );
        if (imageAttachment) {
            try {
                const ext = imageAttachment.contentType.split('/')[1] || 'png';
                const filename = `receipt_${message.id}.${ext}`;
                const filePath = path.join(IMAGES_DIR, filename);
                if (!fs.existsSync(filePath)) {
                    receiptPath = await downloadImage(imageAttachment.url, filename);
                } else {
                    receiptPath = filePath;
                }
            } catch (err) {
                console.error('画像ダウンロードエラー:', err);
            }
        }
    }

    const content = message.content.trim();
    if (!content || content.startsWith('!')) {
        // Image only
        if (receiptPath) {
            const msgDate = new Date(message.createdTimestamp);
            const dateStr = `${msgDate.getFullYear()}-${String(msgDate.getMonth() + 1).padStart(2, '0')}-${String(msgDate.getDate()).padStart(2, '0')}`;
            const expenses = loadExpenses();
            expenses.push({
                id: `discord_${message.id}_img`,
                discordMsgId: message.id,
                date: dateStr,
                orderDate: dateStr,
                amount: 0, unitPrice: 0, quantity: 1,
                category: 'その他', payment: '', description: '領収書（金額未入力）',
                receipt: '/receipts/' + path.basename(receiptPath),
                status: '未清算', source: 'discord', buyer: buyer,
                createdAt: new Date(message.createdTimestamp).toISOString()
            });
            saveExpenses(expenses);
            return true;
        }
        return false;
    }

    const lines = content.split('\n').map(l => l.trim()).filter(l => l.length > 0);
    const msgDate = new Date(message.createdTimestamp);
    const currentYear = msgDate.getFullYear();

    let store = '', description = '', unitPrice = 0, quantity = 1, date = null;

    if (lines.length >= 3) {
        store = lines[0];
        description = lines[1];
        const amountRaw = lines[2];
        const dateRaw = lines.length >= 4 ? lines[3] : null;

        let priceNormalized = amountRaw
            .replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
            .replace(/＠/g, '@');

        if (priceNormalized.includes('@')) {
            const parts = priceNormalized.split('@');
            unitPrice = parseInt(parts[0].replace(/[^0-9]/g, ''), 10) || 0;
            quantity = parseInt(parts[1].replace(/[^0-9]/g, ''), 10) || 1;
        } else {
            unitPrice = parseInt(priceNormalized.replace(/[^0-9]/g, ''), 10) || 0;
        }

        if (dateRaw) {
            let dn = dateRaw.replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
            let dm = dn.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
            if (dm) date = `${dm[1]}-${dm[2].padStart(2, '0')}-${dm[3].padStart(2, '0')}`;
            if (!date) { dm = dn.match(/(\d{1,2})[\/](\d{1,2})/); if (dm) date = `${currentYear}-${dm[1].padStart(2, '0')}-${dm[2].padStart(2, '0')}`; }
            if (!date) { dm = dn.match(/(\d{1,2})月(\d{1,2})日/); if (dm) date = `${currentYear}-${dm[1].padStart(2, '0')}-${dm[2].padStart(2, '0')}`; }
        }
    } else if (lines.length === 1) {
        let priceNormalized = lines[0]
            .replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
            .replace(/[^0-9]/g, '');
        unitPrice = parseInt(priceNormalized, 10) || 0;
        quantity = 1;
    } else {
        return false;
    }

    if (!date) {
        date = `${msgDate.getFullYear()}-${String(msgDate.getMonth() + 1).padStart(2, '0')}-${String(msgDate.getDate()).padStart(2, '0')}`;
    }

    if (unitPrice <= 0) return false;

    const amount = unitPrice * quantity;
    const expenses = loadExpenses();
    expenses.push({
        id: `discord_${message.id}_${Math.random().toString(36).substr(2, 6)}`,
        discordMsgId: message.id,
        date, orderDate: date, unitPrice, quantity, amount,
        category: 'その他', payment: store, description,
        receipt: receiptPath ? '/receipts/' + path.basename(receiptPath) : null,
        status: '未清算', source: 'discord', buyer: buyer,
        createdAt: new Date(message.createdTimestamp).toISOString()
    });
    saveExpenses(expenses);
    console.log(`📥 過去メッセージ取り込み: ${date} ${description || store || unitPrice + '円'}`);
    return true;
}



client.on('messageCreate', async (message) => {
    // Ignore bot messages
    if (message.author.bot) return;

    const content = message.content.trim();

    // ============================================
    // Commands
    // ============================================
    if (content === '!ヘルプ' || content === '!help') {
        const helpEmbed = new EmbedBuilder()
            .setColor(0x6C5CE7)
            .setTitle('💰 経費精算アシスタント - ヘルプ')
            .setDescription('経費メモや領収書を送るだけで自動登録！')
            .addFields(
                { name: '📝 経費を登録', value: '以下の形式でメッセージを送信:\n```\n購入店舗\n内容\n単価@数量\n日付（省略可）\n```\n例:\n```\nYahoo(ジョーシン)\nhdd 4tb\n16800@2\n3/1\n```\n※数量1の場合は `16800` だけでOK', inline: false },
                { name: '📸 領収書を登録', value: '画像を添付して上記テキストと一緒に送信', inline: false },
                { name: '📋 !一覧', value: '最近登録した経費を表示', inline: true },
                { name: '💴 !合計', value: '今月の合計金額を表示', inline: true },
                { name: '❓ !ヘルプ', value: 'このヘルプを表示', inline: true },
            )
            .setFooter({ text: '経費アシスタント | Webアプリと自動同期されます' })
            .setTimestamp();

        await message.reply({ embeds: [helpEmbed] });
        return;
    }

    if (content === '!一覧' || content === '!list') {
        const expenses = loadExpenses();
        const recent = expenses.slice(-10).reverse();

        if (recent.length === 0) {
            await message.reply('📋 まだ経費が登録されていません。');
            return;
        }

        const listEmbed = new EmbedBuilder()
            .setColor(0x00CEC9)
            .setTitle('📋 最近の経費（最新10件）')
            .setDescription(
                recent.map((e, i) => {
                    const icon = e.categoryIcon || '📦';
                    return `**${i + 1}.** ${icon} \`${e.date}\` ${e.description} — **¥${e.amount.toLocaleString()}**`;
                }).join('\n')
            )
            .setFooter({ text: `全${expenses.length}件` })
            .setTimestamp();

        await message.reply({ embeds: [listEmbed] });
        return;
    }

    if (content === '!合計' || content === '!total') {
        const expenses = loadExpenses();
        const now = new Date();
        const thisMonth = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
        const monthExpenses = expenses.filter(e => e.date && e.date.startsWith(thisMonth));
        const total = monthExpenses.reduce((sum, e) => sum + (e.amount || 0), 0);

        const catTotals = {};
        monthExpenses.forEach(e => {
            const cat = e.category || 'その他';
            catTotals[cat] = (catTotals[cat] || 0) + e.amount;
        });

        const breakdown = Object.entries(catTotals)
            .sort((a, b) => b[1] - a[1])
            .map(([cat, amt]) => `${cat}: ¥${amt.toLocaleString()}`)
            .join('\n') || 'なし';

        const totalEmbed = new EmbedBuilder()
            .setColor(0x6C5CE7)
            .setTitle(`💴 ${now.getMonth() + 1}月の経費合計`)
            .addFields(
                { name: '合計金額', value: `**¥${total.toLocaleString()}**`, inline: true },
                { name: '件数', value: `**${monthExpenses.length}件**`, inline: true },
                { name: 'カテゴリ別内訳', value: breakdown, inline: false },
            )
            .setTimestamp();

        await message.reply({ embeds: [totalEmbed] });
        return;
    }

    // ============================================
    // Expense Registration (text + optional image)
    // ============================================
    let hasExpenseData = false;
    let receiptPath = null;

    // Handle image attachments
    if (message.attachments.size > 0) {
        const imageAttachment = message.attachments.find(a =>
            a.contentType && a.contentType.startsWith('image/')
        );

        if (imageAttachment) {
            try {
                const ext = imageAttachment.contentType.split('/')[1] || 'png';
                const filename = `receipt_${Date.now()}.${ext}`;
                receiptPath = await downloadImage(imageAttachment.url, filename);
                console.log(`📸 領収書を保存: ${receiptPath}`);
            } catch (err) {
                console.error('画像ダウンロードエラー:', err);
            }
        }
    }

    // Parse text content - Multi-line format:
    // Line 1: 購入店舗 (Store)
    // Line 2: 内容 (Description)
    // Line 3: 金額 (Amount)
    // Line 4: 日付 (Date) - optional, defaults to today
    if (content && !content.startsWith('!')) {
        const lines = content.split('\n').map(l => l.trim()).filter(l => l.length > 0);

        if (lines.length >= 3) {
            const store = lines[0];
            const description = lines[1];
            const amountRaw = lines[2];
            const dateRaw = lines.length >= 4 ? lines[3] : null;

            // Parse line 3: unitPrice@quantity format (e.g. "16800@2")
            // If no @, quantity defaults to 1
            let priceNormalized = amountRaw
                .replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
                .replace(/＠/g, '@');

            let unitPrice = 0;
            let quantity = 1;

            if (priceNormalized.includes('@')) {
                const parts = priceNormalized.split('@');
                unitPrice = parseInt(parts[0].replace(/[^0-9]/g, ''), 10) || 0;
                quantity = parseInt(parts[1].replace(/[^0-9]/g, ''), 10) || 1;
            } else {
                unitPrice = parseInt(priceNormalized.replace(/[^0-9]/g, ''), 10) || 0;
            }

            const amount = unitPrice * quantity;

            // Parse date
            let date = null;
            const now = new Date();
            const currentYear = now.getFullYear();

            if (dateRaw) {
                let dn = dateRaw.replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
                let dm = dn.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
                if (dm) {
                    date = `${dm[1]}-${dm[2].padStart(2, '0')}-${dm[3].padStart(2, '0')}`;
                }
                if (!date) {
                    dm = dn.match(/(\d{1,2})[\/](\d{1,2})/);
                    if (dm) {
                        date = `${currentYear}-${dm[1].padStart(2, '0')}-${dm[2].padStart(2, '0')}`;
                    }
                }
                if (!date) {
                    dm = dn.match(/(\d{1,2})月(\d{1,2})日/);
                    if (dm) {
                        date = `${currentYear}-${dm[1].padStart(2, '0')}-${dm[2].padStart(2, '0')}`;
                    }
                }
            }

            if (!date) {
                date = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
            }

            // Auto-detect category
            const categoryKeywords = {
                '交通費': ['タクシー', '電車', '新幹線', 'バス', '地下鉄', '交通', '→', 'JR', '飛行機'],
                '宿泊費': ['ホテル', '旅館', '宿泊', '民泊'],
                '会議費': ['スタバ', 'カフェ', 'コーヒー', '打ち合わせ', '会議', 'MTG'],
                '接待交際費': ['レストラン', '居酒屋', '飲み会', '会食', '接待', 'ランチ', 'ディナー'],
                '通信費': ['通信', 'WiFi', '電話', '携帯', 'SIM'],
                '消耗品費': ['文具', 'コピー', 'USB', 'ケーブル', '電池', '消耗品'],
                '書籍・研修費': ['本', '書籍', '研修', 'セミナー', '技術書']
            };
            let category = 'その他';
            const categoryIcons = {
                '交通費': '🚃', '宿泊費': '🏨', '会議費': '☕',
                '接待交際費': '🍽️', '通信費': '📱', '消耗品費': '🖊️',
                '書籍・研修費': '📚', 'その他': '📦'
            };

            for (const [cat, keywords] of Object.entries(categoryKeywords)) {
                if (keywords.some(kw => description.includes(kw) || store.includes(kw))) {
                    category = cat;
                    break;
                }
            }

            if (unitPrice > 0) {
                const expenses = loadExpenses();
                const expense = {
                    id: `discord_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`,
                    date,
                    orderDate: date,
                    unitPrice,
                    quantity,
                    amount,
                    category,
                    payment: store,
                    description,
                    receipt: receiptPath ? '/receipts/' + path.basename(receiptPath) : null,
                    status: '未清算',
                    source: 'discord',
                    buyer: getBuyer(message.author),
                    discordMsgId: message.id,
                    createdAt: new Date().toISOString()
                };
                expenses.push(expense);
                registerProcessedMsgId(message.id);
                saveExpenses(expenses);

                const embed = new EmbedBuilder()
                    .setColor(0x00B894)
                    .setTitle('✅ 経費を登録しました')
                    .addFields(
                        { name: '🏪 購入店舗', value: store, inline: true },
                        { name: '📝 内容', value: description, inline: true },
                        { name: '💰 単価', value: `¥${unitPrice.toLocaleString()}`, inline: true },
                        { name: '📦 数量', value: `${quantity}`, inline: true },
                        { name: '💴 金額', value: `¥${amount.toLocaleString()}`, inline: true },
                        { name: '📅 日付', value: date, inline: true },
                        { name: '📁 カテゴリ', value: `${categoryIcons[category] || '📦'} ${category}`, inline: true },
                    )
                    .setFooter({ text: 'Webアプリのダッシュボードに反映されます' })
                    .setTimestamp();

                await message.reply({ embeds: [embed] });
                hasExpenseData = true;
                receiptPath = null;
            } else {
                await message.reply('⚠️ 単価を読み取れませんでした。3行目に `単価@数量` の形式で入力してください。\n\n📝 **入力形式:**\n```\n購入店舗\n内容\n単価@数量\n日付（省略可）\n```\n例:\n```\nYahoo(ジョーシン)\nhdd 4tb\n16800@2\n3/1\n```');
            }
        } else if (lines.length === 1) {
            // Single line: unit price only
            let priceNormalized = lines[0]
                .replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
                .replace(/[^0-9]/g, '');
            let unitPrice = parseInt(priceNormalized, 10) || 0;

            if (unitPrice > 0) {
                const now = new Date();
                const date = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
                const expenses = loadExpenses();
                const expense = {
                    id: `discord_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`,
                    date,
                    orderDate: date,
                    unitPrice,
                    quantity: 1,
                    amount: unitPrice,
                    category: 'その他',
                    payment: '',
                    description: '',
                    receipt: receiptPath ? '/receipts/' + path.basename(receiptPath) : null,
                    status: '未清算',
                    source: 'discord',
                    buyer: getBuyer(message.author),
                    discordMsgId: message.id,
                    createdAt: new Date().toISOString()
                };
                expenses.push(expense);
                registerProcessedMsgId(message.id);
                saveExpenses(expenses);

                const embed = new EmbedBuilder()
                    .setColor(0x00B894)
                    .setTitle('✅ 経費を登録しました')
                    .addFields(
                        { name: '💰 単価', value: `¥${unitPrice.toLocaleString()}`, inline: true },
                        { name: '📦 数量', value: '1', inline: true },
                        { name: '💴 金額', value: `¥${unitPrice.toLocaleString()}`, inline: true },
                        { name: '📅 日付', value: date, inline: true },
                    )
                    .setFooter({ text: 'Webアプリのダッシュボードに反映されます' })
                    .setTimestamp();

                await message.reply({ embeds: [embed] });
                hasExpenseData = true;
                receiptPath = null;
            }
        } else if (lines.length === 2) {
            await message.reply('📝 **入力形式（各行に1つずつ）:**\n```\n購入店舗\n内容\n単価@数量\n日付（省略可）\n```\n例:\n```\nYahoo(ジョーシン)\nhdd 4tb\n16800@2\n3/1\n```\n※数量1の場合は `16800` だけでOK\n※単価のみの場合は1行で送信OK');
        }
    }

    // Image only (no parseable text)
    if (receiptPath && !hasExpenseData) {
        const expenses = loadExpenses();
        const expense = {
            id: `discord_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`,
            date: `${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, '0')}-${String(new Date().getDate()).padStart(2, '0')}`,
            orderDate: `${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, '0')}-${String(new Date().getDate()).padStart(2, '0')}`,
            amount: 0,
            category: 'その他',
            payment: '',
            description: '領収書（金額未入力）',
            receipt: '/receipts/' + path.basename(receiptPath),
            status: '未清算',
            source: 'discord',
            buyer: getBuyer(message.author),
            discordMsgId: message.id,
            createdAt: new Date().toISOString()
        };
        expenses.push(expense);
        registerProcessedMsgId(message.id);
        saveExpenses(expenses);

        const embed = new EmbedBuilder()
            .setColor(0xFDCB6E)
            .setTitle('📸 領収書を保存しました')
            .setDescription('画像を保存しました。金額はWebアプリで入力してください。\n\n💡 ヒント: 画像と一緒にテキスト（例: `タクシー 1500円`）を送ると自動登録されます。')
            .setTimestamp();

        await message.reply({ embeds: [embed] });
    }
});

// ============================================
// Start Bot
// ============================================
const token = process.env.DISCORD_BOT_TOKEN;
if (!token) {
    console.error('❌ DISCORD_BOT_TOKEN が .env ファイルに設定されていません');
    process.exit(1);
}

client.login(token).catch(err => {
    console.error('❌ Botのログインに失敗しました:', err.message);
    console.error('トークンが正しいか確認してください。');
    process.exit(1);
});
