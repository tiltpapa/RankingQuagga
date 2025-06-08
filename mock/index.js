import { fakerJA, fakerEN } from '@faker-js/faker';

// rankによって結果をソートするヘルパー関数
function sortByRank(results) {
  return results.sort((a, b) => {
    if (a.point !== b.point) {
      return b.point - a.point; // pointが多い方が上位
    }
    // pointが同じ場合はtimeが少ない方が上位
    return parseFloat(a.time) - parseFloat(b.time);
  });
}

// ランダムな日本人名を生成する関数
function generateRandomName() {
  const name = fakerJA.person.fullName();
  
  // 10%の確率でグループ名を追加
  if (Math.random() < 0.1) {
    const groupName = fakerJA.lorem.word() + fakerJA.lorem.word();
    return `${name}(${groupName})`;
  }

  // 10%の確率で英語名を生成
  if (Math.random() < 0.1) {
    const englishName = fakerEN.person.fullName();
    return englishName;
  }

  return name;
}

// 成績情報を生成する関数
function generateMember(id) {
  return {
    id: id,
    admin: false,
    user: {
      id: id,
      name: generateRandomName()
    }
  };
}

// 特定のピリオドの成績を生成する関数
function generatePeriodResults() {
  const results = [];
  
  for (let i = 0; i < 15; i++) {
    const point = Math.floor(Math.random() * 11); // 0-10のランダムな整数
    const time = ((Math.random() * 10) * point).toFixed(3);
    
    results.push({
      rank: 0, // 後で正しい順位を設定
      point: point,
      time: time,
      member: generateMember(i + 1)
    });
  }
  
  // 結果をソートして順位を設定
  const sortedResults = sortByRank(results);
  sortedResults.forEach((result, index) => {
    result.rank = index + 1;
  });
  
  return sortedResults;
}

// 暫定総合成績を生成する関数
function generateTotalResults(top = 35) {
  const results = [];
  
  for (let i = 0; i < top; i++) {
    const point = Math.floor(Math.random() * 50) + 1; // 1-50のランダムな整数
    const time = ((Math.random() * 10) * point + 5).toFixed(3);
    const money = (Math.random() * 1000000).toFixed(1); // ランダムな賞金額
    
    results.push({
      rank: 0, // 後で正しい順位を設定
      point: point,
      time: time,
      member: generateMember(i + 1)
    });
  }
  
  // 結果をソートして順位を設定
  const sortedResults = sortByRank(results);
  sortedResults.forEach((result, index) => {
    result.rank = index + 1;
  });
  
  return sortedResults;
}

// リクエストハンドラー
async function handleRequest(request) {
  const url = new URL(request.url);
  const path = url.pathname;
  
  // CORS ヘッダーを設定
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Content-Type': 'application/json'
  };
  
  // 特定のピリオドの成績
  if (path.match(/\/programs\/[^\/]+\/result\/period$/)) {
    return new Response(JSON.stringify({ results: generatePeriodResults() }), { headers });
  }
  
  // 暫定総合成績
  if (path.match(/\/programs\/[^\/]+\/result\/total$/)) {
    const params = new URLSearchParams(url.search);
    const top = params.get('top') ? parseInt(params.get('top')) : 35;
    
    return new Response(JSON.stringify({ results: generateTotalResults(top) }), { headers });
  }
  
  // 該当するエンドポイントがない場合は404を返す
  return new Response(JSON.stringify({ error: 'Not found' }), { 
    status: 404,
    headers
  });
}

// Cloudflare Workersのエントリーポイント
addEventListener('fetch', event => {
  event.respondWith(handleRequest(event.request));
}); 