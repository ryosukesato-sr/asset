/**
 * IT資産管理システム - テストデータ生成
 */

function generateTestData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'テストデータ生成',
    '約200件の資産テストデータと30名のユーザーデータを生成します。既存データは削除されます。続行しますか？',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(SHEETS.ASSETS)) {
    setup();
  }

  clearExistingData_(ss);
  const testUsers = generateUserData_(ss);
  const testAssets = generateAssetData_(testUsers);
  writeAssetData_(ss, testAssets);
  generateHistoryData_(ss, testAssets);

  ui.alert(`テストデータの生成が完了しました。\nユーザー数: ${testUsers.length}名\n資産数: ${testAssets.length}件`);
}

function clearExistingData_(ss) {
  [SHEETS.ASSETS, SHEETS.HISTORY, SHEETS.USERS].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet && sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }
  });
}

// ===== ユーザーデータ生成 =====

function generateUserData_(ss) {
  const departments = DEFAULT_DEPARTMENTS.map(d => d[1]);

  const people = [
    { last: '田中', first: '太郎', lastR: 'tanaka', firstR: 'taro' },
    { last: '鈴木', first: '花子', lastR: 'suzuki', firstR: 'hanako' },
    { last: '佐藤', first: '一郎', lastR: 'sato', firstR: 'ichiro' },
    { last: '高橋', first: '美咲', lastR: 'takahashi', firstR: 'misaki' },
    { last: '渡辺', first: '健二', lastR: 'watanabe', firstR: 'kenji' },
    { last: '伊藤', first: '由美', lastR: 'ito', firstR: 'yumi' },
    { last: '山本', first: '大輔', lastR: 'yamamoto', firstR: 'daisuke' },
    { last: '中村', first: '和子', lastR: 'nakamura', firstR: 'kazuko' },
    { last: '小林', first: '直樹', lastR: 'kobayashi', firstR: 'naoki' },
    { last: '加藤', first: '恵子', lastR: 'kato', firstR: 'keiko' },
    { last: '吉田', first: '翔太', lastR: 'yoshida', firstR: 'shota' },
    { last: '山田', first: '裕子', lastR: 'yamada', firstR: 'yuko' },
    { last: '松本', first: '隆', lastR: 'matsumoto', firstR: 'takashi' },
    { last: '井上', first: '明美', lastR: 'inoue', firstR: 'akemi' },
    { last: '木村', first: '拓也', lastR: 'kimura', firstR: 'takuya' },
    { last: '林', first: '真理子', lastR: 'hayashi', firstR: 'mariko' },
    { last: '斎藤', first: '正', lastR: 'saito', firstR: 'tadashi' },
    { last: '清水', first: '香織', lastR: 'shimizu', firstR: 'kaori' },
    { last: '山口', first: '勇', lastR: 'yamaguchi', firstR: 'isamu' },
    { last: '阿部', first: '愛', lastR: 'abe', firstR: 'ai' },
    { last: '森田', first: '誠', lastR: 'morita', firstR: 'makoto' },
    { last: '池田', first: 'さくら', lastR: 'ikeda', firstR: 'sakura' },
    { last: '橋本', first: '浩', lastR: 'hashimoto', firstR: 'hiroshi' },
    { last: '石川', first: '麻衣', lastR: 'ishikawa', firstR: 'mai' },
    { last: '前田', first: '亮', lastR: 'maeda', firstR: 'ryo' },
    { last: '藤田', first: '美穂', lastR: 'fujita', firstR: 'miho' },
    { last: '岡田', first: '修', lastR: 'okada', firstR: 'osamu' },
    { last: '後藤', first: '真由美', lastR: 'goto', firstR: 'mayumi' },
    { last: '長谷川', first: '学', lastR: 'hasegawa', firstR: 'manabu' },
    { last: '近藤', first: '理恵', lastR: 'kondo', firstR: 'rie' }
  ];

  const titles = ['部長', '課長', '係長', '主任', '担当', '担当', '担当', '担当'];
  const domain = 'example.com';

  const userRows = people.map((p, i) => {
    const id = `USR-${String(i + 1).padStart(4, '0')}`;
    const fullName = `${p.last} ${p.first}`;
    const email = `${p.lastR}.${p.firstR}@${domain}`;
    const dept = departments[i % departments.length];
    const title = pickRandom_(titles);
    const phone = `03-${String(Math.floor(Math.random() * 9000) + 1000)}-${String(Math.floor(Math.random() * 9000) + 1000)}`;
    const status = i < 28 ? '在籍' : '退職';

    return [id, fullName, email, dept, title, phone, status];
  });

  const sheet = ss.getSheetByName(SHEETS.USERS);
  if (userRows.length > 0) {
    sheet.getRange(2, 1, userRows.length, USER_HEADERS.length).setValues(userRows);
  }

  // オブジェクト配列として返す
  return userRows.map(r => ({
    id: r[0], name: r[1], email: r[2], department: r[3], title: r[4], phone: r[5], status: r[6]
  }));
}

// ===== 資産データ生成 =====

function generateAssetData_(testUsers) {
  const departments = DEFAULT_DEPARTMENTS.map(d => d[1]);
  const activeUsers = testUsers.filter(u => u.status === '在籍');

  const manufacturers = {
    'ノートPC': ['Dell', 'HP', 'Lenovo', 'Apple', 'NEC', 'Fujitsu', 'ASUS', 'Microsoft'],
    'デスクトップPC': ['Dell', 'HP', 'Lenovo', 'NEC', 'Fujitsu', 'Mouse Computer'],
    'モニター': ['Dell', 'HP', 'LG', 'EIZO', 'BenQ', 'Philips', 'ASUS', 'iiyama'],
    'サーバー': ['Dell', 'HP Enterprise', 'Lenovo', 'Fujitsu', 'NEC'],
    'ネットワーク機器': ['Cisco', 'Yamaha', 'Buffalo', 'NETGEAR', 'Juniper', 'Aruba'],
    'プリンター': ['Canon', 'Epson', 'Brother', 'Ricoh', 'HP', 'SHARP'],
    'モバイル': ['Apple', 'Samsung', 'Google', 'Sony', 'SHARP'],
    'タブレット': ['Apple', 'Samsung', 'Lenovo', 'Microsoft', 'NEC'],
    '周辺機器': ['Logicool', 'Elecom', 'Anker', 'Jabra', 'Poly'],
    'その他': ['各社', 'その他']
  };

  const models = {
    'Dell': ['Latitude 5540', 'Latitude 7440', 'XPS 15', 'OptiPlex 7010', 'PowerEdge R750', 'U2723QE', 'P2422H'],
    'HP': ['EliteBook 860', 'ProBook 450', 'ProDesk 400', 'ProLiant DL380', 'E24 G5', 'LaserJet Pro M404'],
    'Lenovo': ['ThinkPad X1 Carbon', 'ThinkPad T14s', 'ThinkCentre M70q', 'ThinkVision T24i', 'Tab P11'],
    'Apple': ['MacBook Pro 14"', 'MacBook Air M2', 'iMac 24"', 'iPhone 15 Pro', 'iPhone 14', 'iPad Pro 12.9"', 'iPad Air'],
    'NEC': ['VersaPro VK', 'Mate MK', 'Express5800', 'MultiSync EA', 'LAVIE Tab'],
    'Fujitsu': ['LIFEBOOK U9', 'ESPRIMO D7', 'PRIMERGY RX2540', 'VL-244SSV'],
    'ASUS': ['ExpertBook B9', 'ProArt PA278QV', 'ZenBook 14'],
    'Microsoft': ['Surface Pro 9', 'Surface Laptop 5', 'Surface Go 3'],
    'LG': ['27UL850', '32UN880', '34WN80C'],
    'EIZO': ['FlexScan EV2490', 'FlexScan EV2780', 'ColorEdge CS2740'],
    'BenQ': ['PD2700U', 'GW2780', 'EW3270U'],
    'Cisco': ['Catalyst 9200', 'Catalyst 2960', 'ISR 4331', 'Meraki MR46'],
    'Yamaha': ['RTX1220', 'RTX830', 'SWX2310'],
    'Buffalo': ['BS-GS2016', 'WXR-5950AX12'],
    'Canon': ['imageRUNNER C3530', 'Satera LBP622C'],
    'Epson': ['LP-S3290', 'PX-M6711FT'],
    'Brother': ['MFC-L3770CDW', 'HL-L3230CDW'],
    'Ricoh': ['IM C3500', 'SP C261'],
    'Samsung': ['Galaxy S24', 'Galaxy Tab S9'],
    'Google': ['Pixel 8 Pro', 'Pixel 7a'],
    'Sony': ['Xperia 5 V', 'Xperia 1 V'],
    'SHARP': ['AQUOS sense8', 'BP-70C26'],
    'Logicool': ['MX Keys', 'MX Master 3S', 'C920 HD Pro', 'Rally Bar'],
    'Elecom': ['TK-FDM110', 'M-XGM20DL'],
    'Anker': ['PowerConf C300', 'Soundcore A30i'],
    'Jabra': ['Evolve2 85', 'Speak2 75'],
    'Poly': ['Studio P15', 'Voyager Focus 2'],
    'Mouse Computer': ['mouse B5-I7', 'mouse DT7'],
    'HP Enterprise': ['ProLiant DL360', 'ProLiant DL380'],
    'NETGEAR': ['GS316EP', 'WAX620'],
    'Juniper': ['EX2300', 'SRX300'],
    'Aruba': ['AP-535', 'CX 6200'],
    'Philips': ['272E2FA', '243V7Q'],
    'iiyama': ['ProLite XUB2493HS', 'ProLite XB2783HSU'],
    '各社': ['汎用品'],
    'その他': ['その他']
  };

  const osOptions = {
    'ノートPC': ['Windows 11 Pro', 'Windows 10 Pro', 'macOS Sonoma', 'macOS Ventura', 'Ubuntu 22.04'],
    'デスクトップPC': ['Windows 11 Pro', 'Windows 10 Pro', 'Ubuntu 22.04'],
    'サーバー': ['Windows Server 2022', 'Windows Server 2019', 'Ubuntu 22.04 LTS', 'Red Hat Enterprise Linux 9', 'CentOS 7'],
    'モバイル': ['iOS 17', 'iOS 16', 'Android 14', 'Android 13'],
    'タブレット': ['iPadOS 17', 'iPadOS 16', 'Android 13', 'Windows 11'],
    'モニター': [''], 'ネットワーク機器': [''], 'プリンター': [''], '周辺機器': [''], 'その他': ['']
  };

  const locations = [
    '本社 1F', '本社 2F', '本社 3F', '本社 4F', '本社 5F',
    '大阪支社', '名古屋支社', '福岡支社', '札幌支社',
    'サーバールーム', 'テレワーク', '会議室A', '会議室B', '受付'
  ];

  const priceRange = {
    'ノートPC': [80000, 350000], 'デスクトップPC': [60000, 250000],
    'モニター': [20000, 120000], 'サーバー': [300000, 2000000],
    'ネットワーク機器': [10000, 500000], 'プリンター': [30000, 800000],
    'モバイル': [30000, 200000], 'タブレット': [30000, 180000],
    '周辺機器': [3000, 80000], 'その他': [5000, 100000]
  };

  const categoryDistribution = {
    'ノートPC': 55, 'デスクトップPC': 20, 'モニター': 40,
    'サーバー': 8, 'ネットワーク機器': 15, 'プリンター': 12,
    'モバイル': 20, 'タブレット': 10, '周辺機器': 15, 'その他': 5
  };

  const assets = [];
  const now = new Date();

  // ステータスに応じてユーザーを割り当てるかどうか
  const statusNeedsUser = ['利用中', '受取待ち', '返却待ち', '回収連絡済み'];

  Object.keys(categoryDistribution).forEach(category => {
    const count = categoryDistribution[category];
    const mfrs = manufacturers[category] || ['その他'];
    const [minPrice, maxPrice] = priceRange[category] || [5000, 100000];
    const osOpts = osOptions[category] || [''];

    for (let i = 0; i < count; i++) {
      const mfr = pickRandom_(mfrs);
      const modelList = models[mfr] || ['標準モデル'];
      const model = pickRandom_(modelList);
      const status = weightedStatus_();
      const os = pickRandom_(osOpts);

      // ユーザー割り当て
      let userId = '', userName = '', userEmail = '', dept = '', location = '';
      if (statusNeedsUser.includes(status)) {
        const user = pickRandom_(activeUsers);
        userId = user.id;
        userName = user.name;
        userEmail = user.email;
        dept = user.department;
        location = pickRandom_(locations);
      } else {
        dept = pickRandom_(departments);
        if (status === '保管' || status === '在庫') {
          location = pickRandom_(['倉庫', 'サーバールーム', '本社 1F 保管庫']);
        }
      }

      const purchaseDate = randomDate_(new Date(2020, 0, 1), new Date(2025, 11, 31));
      const price = Math.round((Math.random() * (maxPrice - minPrice) + minPrice) / 1000) * 1000;
      const serial = generateSerial_(mfr);
      const ip = needsIP_(category) ? generateIP_() : '';
      const mac = needsIP_(category) ? generateMAC_() : '';
      const createdDate = new Date(purchaseDate.getTime() + Math.random() * 7 * 86400000);
      const updatedDate = new Date(createdDate.getTime() + Math.random() * (now.getTime() - createdDate.getTime()));
      const id = `AST-${String(assets.length + 1).padStart(5, '0')}`;

      const leaseEnd = Math.random() < 0.2 ? randomDate_(new Date(), new Date(2026, 11, 31)) : null;
      const warrantyEnd = randomDate_(new Date(purchaseDate), new Date(purchaseDate.getTime() + 3 * 365 * 86400000));
      const returnDue = status === '返却待ち' ? randomDate_(new Date(), new Date(2025, 11, 31)) : null;

      assets.push([
        id,
        `${mfr} ${model}`,
        category,
        mfr,
        model,
        serial,
        purchaseDate,
        price,
        leaseEnd || '',
        warrantyEnd,
        returnDue || '',
        userId,
        userName,
        userEmail,
        dept,
        location,
        status,
        ip,
        mac,
        os,
        '',
        '',
        '',
        createdDate,
        updatedDate
      ]);
    }
  });

  return assets;
}

function writeAssetData_(ss, assets) {
  const sheet = ss.getSheetByName(SHEETS.ASSETS);
  if (assets.length > 0) {
    sheet.getRange(2, 1, assets.length, ASSET_HEADERS.length).setValues(assets);
  }
}

function generateHistoryData_(ss, assets) {
  const sheet = ss.getSheetByName(SHEETS.HISTORY);
  const historyRows = [];
  const changeTypes = ['新規登録', '更新', 'ステータス変更', '使用者変更', '設置場所変更'];
  const emails = ['admin@example.com', 'tanaka.taro@example.com', 'suzuki.hanako@example.com', 'sato.ichiro@example.com'];

  assets.forEach(asset => {
    historyRows.push([
      Utilities.getUuid(),
      asset[0],
      asset[18], // 登録日
      '新規登録',
      `資産「${asset[1]}」を登録`,
      pickRandom_(emails)
    ]);
  });

  for (let i = 0; i < 50; i++) {
    const asset = pickRandom_(assets);
    const changeType = pickRandom_(changeTypes.slice(1));
    const changeDate = randomDate_(new Date(asset[18]), new Date());
    let description = '';

    switch (changeType) {
      case '更新':
        description = '資産情報を更新';
        break;
      case 'ステータス変更':
        const fromStatus = pickRandom_(STATUS_OPTIONS);
        const toStatus = pickRandom_(STATUS_OPTIONS.filter(s => s !== fromStatus));
        description = `ステータス: ${fromStatus} → ${toStatus}`;
        break;
      case '使用者変更':
        description = '使用者を変更';
        break;
      case '設置場所変更':
        description = '設置場所を変更';
        break;
    }

    historyRows.push([
      Utilities.getUuid(),
      asset[0],
      changeDate,
      changeType,
      description,
      pickRandom_(emails)
    ]);
  }

  historyRows.sort((a, b) => new Date(a[2]) - new Date(b[2]));

  if (historyRows.length > 0) {
    sheet.getRange(2, 1, historyRows.length, HISTORY_HEADERS.length).setValues(historyRows);
  }
}

// ===== ヘルパー =====

function pickRandom_(arr) {
  return arr[Math.floor(Math.random() * arr.length)];
}

function weightedStatus_() {
  const r = Math.random();
  if (r < 0.50) return '利用中';
  if (r < 0.62) return '在庫';
  if (r < 0.69) return '受取待ち';
  if (r < 0.76) return '返却待ち';
  if (r < 0.82) return '回収連絡済み';
  if (r < 0.87) return '回収済み';
  if (r < 0.92) return 'リース終了';
  if (r < 0.96) return '紛失';
  return '保管';
}

function randomDate_(start, end) {
  return new Date(start.getTime() + Math.random() * (end.getTime() - start.getTime()));
}

function generateSerial_(manufacturer) {
  const prefix = manufacturer.substring(0, 2).toUpperCase();
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let serial = prefix;
  for (let i = 0; i < 10; i++) {
    serial += chars[Math.floor(Math.random() * chars.length)];
  }
  return serial;
}

function generateIP_() {
  const subnet = pickRandom_(['192.168.1', '192.168.2', '192.168.10', '10.0.1', '10.0.2']);
  return `${subnet}.${Math.floor(Math.random() * 254) + 1}`;
}

function generateMAC_() {
  const hex = '0123456789ABCDEF';
  let mac = '';
  for (let i = 0; i < 6; i++) {
    if (i > 0) mac += ':';
    mac += hex[Math.floor(Math.random() * 16)] + hex[Math.floor(Math.random() * 16)];
  }
  return mac;
}

function needsIP_(category) {
  return ['ノートPC', 'デスクトップPC', 'サーバー', 'ネットワーク機器', 'プリンター'].includes(category);
}
