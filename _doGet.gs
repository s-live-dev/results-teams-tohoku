// テスト用URL
// https://script.google.com/macros/s/AKfycbypYylvBCZr7_v9CAyKXRE0GDsiZ3gKWEpYoyE7f88/dev?s=1VjvZD1mQ9sR80VWh_izIAw5GeS9o7y3CBdySpM2MyiM&qr=team%2FTohoku
// Opus4 クリーンアップ実行済み

/**
 * WebページとしてアクセスされたときにHTMLを返す
 * @param {object} e - イベントオブジェクト
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  // URLパラメータからスプレッドシートIDとQRパラメータを取得してテンプレートに渡す
  template.s = e.parameter.s || '';
  template.qr = e.parameter.qr || ''; // QRコード用パラメータを追加

  return template.evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * 他のHTMLファイル（CSS、JSなど）をインクルードするためのヘルパー関数
 * @param {string} filename - インクルードするファイル名
 * @returns {string} ファイルの内容
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * HTML側から呼び出され、団体戦の全データをJSONオブジェクトとして返す
 * @param {string} s - スプレッドシートID
 * @param {string} qrParam - QRコード生成用パラメータ
 * @returns {object} 大会情報、競技別団体、総合団体のデータを含むオブジェクト
 */
function getTeamData(s, qrParam) {
  if (!s) {
    const errorMessage = "Error: Spreadsheet ID(s) is missing.";
    console.error(errorMessage);
    return { error: errorMessage };
  }

  try {
    const ss = SpreadsheetApp.openById(s);
    // シート名を固定で指定
    const trapSheet = ss.getSheetByName("トラップ");
    const skeetSheet = ss.getSheetByName("スキート");

    if (!trapSheet) return { error: `Sheet "トラップ" not found.` };
    if (!skeetSheet) return { error: `Sheet "スキート" not found.` };

    // 大会情報を取得（QRコードパラメータも渡す）
    const eventInfo = getEventInfoFromSheet(s, qrParam);

    // 各シートから選手データをパース
    const trapPlayers = parsePlayersData(trapSheet, 'trap');
    const skeetPlayers = parsePlayersData(skeetSheet, 'skeet');
    const allPlayers = [...trapPlayers, ...skeetPlayers];

    // 団体戦の結果を計算
    const teamResults = calculateTeamResults(allPlayers);

    // 最終的なデータオブジェクトを構築して返す
    return {
      eventInfo: eventInfo,
      teamEvent: {
        trap: teamResults.eventTrap,
        skeet: teamResults.eventSkeet
      },
      teamOverall: teamResults.overall
    };

  } catch (e) {
    console.error("getTeamData Error: " + e.message + " Stack: " + e.stack);
    return { error: e.toString() };
  }
}

/**
 * オリジナルの getEventInfoData をベースに団体戦用に調整
 * @param {string} s - スプレッドシートID
 * @param {string} qrParam - QRコード生成用パラメータ
 * @returns {object} 大会情報オブジェクト
 */
function getEventInfoFromSheet(s, qrParam) {
  // オリジナルと同じく「大会情報」シートから取得
  var sheet = SpreadsheetApp.openById(s).getSheetByName('大会情報');
  var eData = sheet.getDataRange().getValues().slice(1, 3); // 最大2件のデータを取得

  // eData から列　主催協会:[0] が空の行を削除
  eData = eData.filter(function (row) {
    return row[0] !== ''; // インデックス0の列が空ではない行だけを残す
  });

  if (eData.length === 0) {
    // データがない場合のデフォルト値
    return {
      name: "団体戦結果",
      flagUrl: "",
      place: "",
      date: "",
      days: "",
      weather: "",
      lastUpdate: "最終更新: " + new Date().toLocaleTimeString('ja-JP'),
      qrCodeUrl: "",
      status: {
        trap: "---",
        skeet: "---"
      }
    };
  }

  // 最初の行のデータを使用（オリジナルと同じ構造）
  var row = eData[0];

  // OpenWeatherMap APIから気象情報を取得（オリジナルと同じ）
  var weatherData;
  try {
    var location = row[7].split(',');
    var latitude = parseFloat(location[0].trim());
    var longitude = parseFloat(location[1].trim());
    var apiKey = PropertiesService.getScriptProperties().getProperty('AK_openWeather');
    var url = `https://api.openweathermap.org/data/2.5/weather?units=metric&lat=${latitude}&lon=${longitude}&appid=${apiKey}`;
    var response = UrlFetchApp.fetch(url);
    var json = response.getContentText();
    weatherData = JSON.parse(json);
  } catch (error) {
    weatherData = {
      weather: [{ description: 'N/A ' }],
      main: { temp: 'N/A ', humidity: 'N/A ', pressure: 'N/A ' },
      wind: { speed: 'N/A ' }
    };
    console.log('S-LIVE: caught an error,set default values:', error);
  }

  // QRコード生成（修正版）
  var qrCodeUrl = "";
  if (qrParam) {
    // パラメータが指定されている場合
    var targetUrl = "https://s-live.org/" + qrParam;
    qrCodeUrl = "https://api.qrserver.com/v1/create-qr-code/?data=" +
      encodeURIComponent(targetUrl) +
      '&format=png&margin=10&size=150x150';
  } else {
    // デフォルトのQRコード（s-live.orgのトップページ）
    qrCodeUrl = "https://api.qrserver.com/v1/create-qr-code/?data=" +
      encodeURIComponent("https://s-live.org/") +
      '&format=png&margin=10&size=150x150';
  }

  // 状況アイコンを生成する関数
  function getStatusIcon(status) {
    switch (status) {
      case '競技前': return '<i class="fa-regular fa-circle-pause"></i>';
      case '競技中': return '<i class="fa-regular fa-circle-play"></i>';
      case '競技終了': return '<i class="fa-regular fa-circle-check"></i>';
      case '1日目終了': return '<i class="fa-regular fa-circle-pause"></i>';
      default: return '';
    }
  }

  // オリジナルの形式でデータを構築
  return {
    name: row[1], // 大会名
    flagUrl: 'https://s-live.org/wp-content/plugins/s-live/resource/flag/' + encodeURIComponent(row[0]) + '.png',
    place: row[6], // 場所
    date: '<i class="fa-regular fa-calendar-days"></i> ' + Utilities.formatDate(new Date(row[5]), "Asia/Tokyo", "yy/MM/dd"),
    days: row[4] + 'Day(s)',
    weather: '<i class="fa-solid fa-sun"></i> ' + weatherData.weather[0].description + ' ' +
      '<i class="fa-solid fa-temperature-three-quarters"></i> ' + weatherData.main.temp + 'c ' +
      '<i class="fa-solid fa-droplet"></i> ' + weatherData.main.humidity + '% ' +
      '<i class="fa-solid fa-wind"></i> ' + weatherData.wind.speed + 'm/s ' +
      '<i class="fa-solid fa-gauge-simple"></i> ' + weatherData.main.pressure + 'hPa',
    lastUpdate: '<i class="fa-regular fa-clock"></i> ' + Utilities.formatDate(new Date(row[2]), 'Asia/Tokyo', "yy/MM/dd HH:mm"),
    qrCodeUrl: qrCodeUrl,
    status: {
      trap: row[3] || "---", // 状況
      skeet: row[3] || "---"  // 団体戦では同じ状況を想定
    },
    statusIcon: getStatusIcon(row[3]) // 状況アイコンを追加
  };
}

/**
 * シートから選手データをパースしてオブジェクトの配列を返す
 * @param {Sheet} sheet - スプレッドシートのシートオブジェクト
 * @param {string} discipline - 'trap' または 'skeet'
 * @returns {Array<object>} 選手オブジェクトの配列
 */
function parsePlayersData(sheet, discipline) {
  const data = sheet.getRange("A2:W" + sheet.getLastRow()).getValues();
  const players = [];

  data.forEach(row => {
    if (!row[5] || !row[6]) return;
    if (row[22] === "RPO") return;

    // 順位の取得（修正：デフォルトを900に）
    const pos = Number(row[2]) || 900;

    // 900以上（初期値・DNS）は除外
    if (pos >= 900) return;

    // タイムスタンプの処理を改善
    let updateTimeString = null;
    if (row[19] instanceof Date) {
      // Date型の場合、日本時間でフォーマット
      updateTimeString = Utilities.formatDate(row[19], "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss");
    } else if (row[19]) {
      // 既に文字列の場合はそのまま使用
      updateTimeString = row[19].toString();
    }

    players.push({
      discipline: discipline,
      team: row[6],
      name: row[5],
      pos: pos,
      r1: Number(row[7]) || 0,
      r2: Number(row[8]) || 0,
      r3: Number(row[9]) || 0,
      r4: Number(row[10]) || 0,
      total: Number(row[17]) || 0,
      updateTime: updateTimeString // 文字列として保存
    });
  });
  return players;
}

/**
 * 選手データから団体戦の結果を計算する
 * @param {Array<object>} players - 全選手オブジェクトの配列
 * @returns {object} 計算された団体戦の結果
 */
function calculateTeamResults(players) {
  const teams = players.reduce((acc, player) => {
    if (!acc[player.team]) {
      acc[player.team] = { trap: [], skeet: [] };
    }
    acc[player.team][player.discipline].push(player);
    return acc;
  }, {});

  const eventTrap = [], eventSkeet = [], overall = [];

  for (const teamName in teams) {
    const teamPlayers = teams[teamName];

    // 個人順位（pos）で昇順ソート（小さい数字が上位）
    teamPlayers.trap.sort((a, b) => a.pos - b.pos);
    teamPlayers.skeet.sort((a, b) => a.pos - b.pos);

    // トラップ種目別団体（上位3名）
    const eventTrapPlayers = teamPlayers.trap.slice(0, 3);
    // 有効な選手が存在する場合のみ追加
    if (eventTrapPlayers.length > 0) {
      eventTrap.push({
        name: teamName,
        total: eventTrapPlayers.reduce((sum, p) => sum + p.total, 0),
        players: eventTrapPlayers.map((p, i) => {
          let rank = 1;
          for (let j = 0; j < i; j++) {
            if (eventTrapPlayers[j].total !== p.total) {
              rank = j + 2;
            }
          }
          return { ...p, rank: rank, pos: p.pos };
        })
      });
    }

    // スキート種目別団体（上位3名）
    const eventSkeetPlayers = teamPlayers.skeet.slice(0, 3);
    // 有効な選手が存在する場合のみ追加
    if (eventSkeetPlayers.length > 0) {
      eventSkeet.push({
        name: teamName,
        total: eventSkeetPlayers.reduce((sum, p) => sum + p.total, 0),
        players: eventSkeetPlayers.map((p, i) => {
          let rank = 1;
          for (let j = 0; j < i; j++) {
            if (eventSkeetPlayers[j].total !== p.total) {
              rank = j + 2;
            }
          }
          return { ...p, rank: rank, pos: p.pos };
        })
      });
    }

    // 総合団体（トラップ上位5名、スキート上位3名）
    const overallTrapPlayers = teamPlayers.trap.slice(0, 5);
    const overallSkeetPlayers = teamPlayers.skeet.slice(0, 3);

    // 総合団体のマッピング部分を修正
    // どちらかの種目に有効な選手がいる場合のみ追加
    if (overallTrapPlayers.length > 0 || overallSkeetPlayers.length > 0) {
      const trapTotal = overallTrapPlayers.reduce((sum, p) => sum + p.total, 0);
      const skeetTotal = overallSkeetPlayers.reduce((sum, p) => sum + p.total, 0);
      overall.push({
        name: teamName,
        trapTotal: trapTotal,
        skeetTotal: skeetTotal,
        overallTotal: trapTotal + skeetTotal,
        trapPlayers: overallTrapPlayers.map((p, i) => {
          let rank = 1;
          for (let j = 0; j < i; j++) {
            if (overallTrapPlayers[j].total !== p.total) {
              rank = j + 2;
            }
          }
          return {
            name: p.name,
            total: p.total,
            rank: rank,
            pos: p.pos,
            r1: p.r1,
            r2: p.r2,
            r3: p.r3,
            r4: p.r4,
            updateTime: p.updateTime
          };
        }),
        skeetPlayers: overallSkeetPlayers.map((p, i) => {
          let rank = 1;
          for (let j = 0; j < i; j++) {
            if (overallSkeetPlayers[j].total !== p.total) {
              rank = j + 2;
            }
          }
          return {
            name: p.name,
            total: p.total,
            rank: rank,
            pos: p.pos,
            r1: p.r1,
            r2: p.r2,
            r3: p.r3,
            r4: p.r4,
            updateTime: p.updateTime
          };
        })
      });
    }
  }

  // トラップ種目団体の同点同順位制
  eventTrap.sort((a, b) => b.total - a.total);
  let currentRank = 1;
  for (let i = 0; i < eventTrap.length; i++) {
    if (i > 0 && eventTrap[i].total !== eventTrap[i - 1].total) {
      currentRank = i + 1;
    }
    eventTrap[i].rank = currentRank;
  }

  // スキート種目団体の同点同順位制
  eventSkeet.sort((a, b) => b.total - a.total);
  currentRank = 1;
  for (let i = 0; i < eventSkeet.length; i++) {
    if (i > 0 && eventSkeet[i].total !== eventSkeet[i - 1].total) {
      currentRank = i + 1;
    }
    eventSkeet[i].rank = currentRank;
  }

  // 総合団体の同点同順位制
  overall.sort((a, b) => b.overallTotal - a.overallTotal);
  currentRank = 1;
  for (let i = 0; i < overall.length; i++) {
    if (i > 0 && overall[i].overallTotal !== overall[i - 1].overallTotal) {
      currentRank = i + 1;
    }
    overall[i].rank = currentRank;
  }

  return { eventTrap, eventSkeet, overall };
}

/**
 * 外部サーバーの音声ファイルをプロキシ経由で取得
 * @param {string} soundName - 音声名（'rankUp' または 'playerUpdate'）
 */
function getExternalSound(soundName) {
  const soundUrls = {
    rankUp: 'https://s-live.org/sounds/punch-it_team_rankup.mp3',
    playerUpdate: 'https://s-live.org/sounds/athlete_updated.mp3'
  };

  try {
    const url = soundUrls[soundName];
    if (!url) {
      return { success: false, error: 'Invalid sound name' };
    }

    // 外部URLから音声データを取得
    const response = UrlFetchApp.fetch(url);
    const blob = response.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());

    return {
      success: true,
      mimeType: blob.getContentType(),
      data: base64
    };
  } catch (e) {
    console.error('外部音声取得エラー:', e);
    return {
      success: false,
      error: e.toString()
    };
  }
}