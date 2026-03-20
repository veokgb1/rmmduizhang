// ==========================================
// Constants.gs（结构化资产审计输出）
// 说明：此文件集中管理全局常量/密钥/资源定位/运行参数。
// ⚠️ 强烈建议：将真实密钥迁移到 Script Properties，并尽快轮换已暴露的 Key。
// ==========================================

// ==========================================
// 【流水明细表列序号】LEDGER_COLS（P5 修复）
// 规则：序号均为 1-indexed，与 getRange(row, col) 保持一致。
// - getRange 直接用：getRange(row, LEDGER_COLS.AMOUNT)
// - getValues 数组用：data[i][LEDGER_COLS.AMOUNT - 1]
// - 整行操作用：getRange(r, 1, 1, LEDGER_COLS.TOTAL_COLS)
// ==========================================
const LEDGER_COLS = {
  DATE:        1,  // 日期
  MONTH:       2,  // 月份（yyyy-MM）
  TYPE:        3,  // 收支类型（收入 / 支出）
  CATEGORY:    4,  // 分类
  AMOUNT:      5,  // 金额
  SUMMARY:     6,  // 摘要
  SOURCE:      7,  // 来源（快捷文本录入 / 批量文本录入 等）
  VOUCHER:     8,  // 凭证（Rich Text 超链接）
  STATUS:      9,  // 关联状态（智能关联 / 人工关联 / 未关联 / 🔖 待续关联）
  EXTRA:       10, // 解绑时清空（备用字段，用途待确认）
  VOUCHER_V2:  11, // 凭证迁移目标（migrateVouchers 写入 "v:id1|id2" 格式）

  TOTAL_COLS:  9,  // 整行格式化范围（DATE ~ STATUS，即第 1~9 列）
};

const DUKA_CONFIG = {
  // ==========================================================
  // 【第一层：核心资产 AUTH】
  // 最敏感、最重要的授权信息统一放在顶部，便于集中管控与轮换。
  // ==========================================================
  AUTH: {
    // Gemini / Google API Keys
    GEMINI_API_KEY: PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY'),
    SHEETS_API_KEY: 'YOUR_API_KEY_HERE',

    // Token 计算盐（用于生成/验证 token、simpleHash 等）
    TOKEN_SECRET: PropertiesService.getScriptProperties().getProperty('TOKEN_SECRET'),

    // Gitee 同步（当前为占位符，待填真实值）
    GITEE: {
      TOKEN: '在此填入令牌',
      OWNER: '在此填入用户名',
      REPO: 'myledger',
      BRANCH: 'master',
    },
  },

  // ==========================================================
  // 【第二层：资源定位 RESOURCES】
  // 与外部资源绑定的定位信息：Spreadsheet ID、Folder IDs、URL 链接等。
  // ==========================================================
  RESOURCES: {
    // Google Apps Script 项目信息（来自 .clasp.json）
    APPS_SCRIPT: {
      SCRIPT_ID: '1wncm_ToJpTgC1Ezc3jaVXNMEvg7REMkjH3GW3fhLO4TS4BF94ec2ZYyQ',
      ROOT_DIR: './',
    },

    // 表格资源
    SPREADSHEET: {
      ID: '1moQy0qsxBTSQ3VvLRVD9onoZuvycsbe5tMqGyhAiVdc',
    },

    // Drive 文件夹 / 文件定位
    DRIVE: {
      // 默认待处理文件夹（原代码以 URL 形式存在）
      DEFAULT_PENDING_URL: 'https://script.google.com/macros/s/AKfycbz-m8GuPED56ezJ4MvlNC_kqxcuZ4oFXFTmnR5yIQE01wbOKF_rh_34mICwMVYAblwX/exec',

      // 各类 Folder ID
      UPLOAD_FOLDER_ID: '11F2YTXriOElRxGZi1ScYlRirTwnkCdl2',
      FOLDER_ARCHIVED_ID: '13wIu-LW37XRcQuEQYxAezU907xeQu-wj',
      FOLDER_TAGGED_PENDING_ID: '1rgm5dWxUxE3aEU66uHjQyjoHIkDpAG5l',

      // 去重雷达 / 扫描候选文件夹列表
      DEDUPE_FOLDERS: [
        { name: '02_收支凭证', id: '1kV5jq1GPAnNj8eJO8FC9JCYLDKUsdOs0' },
        { name: '2025-12', id: '11F2YTXriOElRxGZi1ScYlRirTwnkCdl2' },
        { name: '00_财务凭证存档', id: '13wIu-LW37XRcQuEQYxAezU907xeQu-wj' },
        { name: '03_手动待关联区', id: '1GTtgAb64EveVijxQ2ydMggDwOqzosUq6' },
        { name: '04_带标记未完结凭证', id: '1rgm5dWxUxE3aEU66uHjQyjoHIkDpAG5l' },
      ],
    },

    // 外部 API 地址（集中归档，方便排查/替换）
    URLS: {
      GEMINI: {
        // 【统一主力版本】全项目统一使用 gemini-2.5-flash
        // 使用方式：DUKA_CONFIG.RESOURCES.URLS.GEMINI.PRIMARY + '?key=' + API_KEY
        PRIMARY:
          'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent',
        // 【已废弃】原 05_网页API代理 使用的 2.0 版本，待该文件完成统一后删除
        LEGACY_20_FLASH:
          'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent',
      },
      GOOGLE_DRIVE: {
        THUMBNAIL: 'https://drive.google.com/thumbnail',
        FILE_VIEW_PREFIX: 'https://drive.google.com/file/d/',
        DRIVE_V3_FILES: 'https://www.googleapis.com/drive/v3/files',
      },
      GITEE: {
        CONTENTS_API_PREFIX: 'https://gitee.com/api/v5/repos/',
      },
    },
  },

  // ==========================================================
  // 【第三层：运行参数 SETTINGS】
  // 运行阈值、重试/分页参数、匹配阈值等非敏感配置项。
  // ==========================================================
  SETTINGS: {
    // 统一日期时区/格式
    TIMEZONE: {
      // 代码中同时出现 "GMT+8" 与 "Asia/Shanghai"，这里统一显式记录
      GMT_OFFSET: 'GMT+8',
      IANA: 'Asia/Shanghai',
      DATE: 'yyyy-MM-dd',
      DATETIME: 'yyyy-MM-dd HH:mm:ss',
      STAMP: 'yyyyMMdd_HHmmss',
    },

    // 图片/内容大小限制
    LIMITS: {
      // 原代码出现两处：CONFIG.MAX_BYTES 与 imageProxy 内 MAX_BYTES
      MAX_BYTES: 20 * 1024 * 1024,
      THUMB_SIZE: 'w800',
      THUMB_SIZE_UPLOAD_RETURN: 'w400',
    },

    // 匹配引擎阈值（analyzeSingleImage / calculateMatchScore）
    MATCHING: {
      MIN_SCORE: 60,
      AMOUNT: {
        EXACT_EPS: 0.01,
        SMALL_AMOUNT_MAX: 100,
        SMALL_AMOUNT_ABS_DIFF_MAX: 5,
        RELATIVE_DIFF_MAX: 0.2,
      },
      DATE: {
        SAME_DAY_SCORE: 20,
        WITHIN_3_DAYS_SCORE: 15,
        WITHIN_15_DAYS_SCORE: 10,
        WITHIN_3_DAYS: 3,
        WITHIN_15_DAYS: 15,
      },
      LEDGER_ROW_START: 3,
    },

    // Drive API 扫描/分页
    DRIVE_SCAN: {
      PAGE_SIZE: 1000,
      FIELDS: 'nextPageToken,files(id,md5Checksum,mimeType)',
      MIME_IMAGE_PREFIX: 'image/',
    },

    // Web API（05_网页API代理）相关
    WEB_API: {
      TOKEN: {
        // generateToken 返回：base64(username):ts:md5(payload).slice(0,16)
        HASH_SLICE_LEN: 16,
        DIGEST_ALGO: 'MD5',
        SEPARATOR: '|',
      },
      VOUCHER_MIGRATE: {
        // migrateVouchers 写入 "v:" + ids.join('|')
        PREFIX: 'v:',
        JOIN_CHAR: '|',
        SOURCE_COL: 8,
        TARGET_COL: 11,
      },
    },
  },

  // ==========================================================
  // 【第四层：元数据 METADATA】
  // 本次抓取时间、来源文件等非运行必需信息，仅用于审计追踪。
  // ==========================================================
  METADATA: {
    AUDIT_AT: '2026-03-18',
    SOURCE_FILES: [
      '01_配置与菜单.js',
      '02_核心引擎.js',
      '03_前端UI弹窗.js',
      '04_辅助工具箱.js',
      '05_网页API代理升级版.js',
      '.clasp.json',
    ],
    NOTES: [
      '项目当前未发现 .gs 源文件，实际脚本以 .js 形式存在（Apps Script 依然可用）。',
      '建议把 AUTH 下的 Key/Secret 迁移到 Script Properties，并在代码中统一从 DUKA_CONFIG 读取。',
    ],
  },
};

