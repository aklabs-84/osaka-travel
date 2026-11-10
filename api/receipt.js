/* 영수증 읽기 함수 (Vercel)
 * 앱이 보낸 사진 → Gemini → 합계·가게·품목·종류.
 * 열쇠는 Vercel 환경변수에만 있다: GEMINI_KEY_FREE(먼저 씀), GEMINI_KEY_PAID(무료가 막히면 씀),
 * GEMINI_MODEL(기본 모델), GEMINI_MODEL_STRONG(선택: 합계를 못 읽었을 때 한 번 더 읽는 모델).
 * 여행 코드가 맞는 사람만, 여행마다 하루 횟수 제한(슈퍼베이스 rc_take). 열쇠는 응답에 절대 싣지 않는다. */
const SB = 'https://eyjiqepgcjhtytserrlp.supabase.co';
const SB_KEY = 'sb_publishable_Vu7hE7D7Z3uVbEeRIoW4OQ__mHe36St'; // 공개용 키(프론트에도 있음)
const CATS = ['food', 'move', 'shop', 'do', 'etc'];
const PROMPT = '이것은 일본 여행 중 받은 영수증 사진이다. 아래 항목을 읽어 JSON으로만 답해라.\n' +
  '- total: 세금 포함 최종 결제 합계(엔, 정수). 합계(合計)·税込·お会計 중 최종 지불액. 못 찾으면 0\n' +
  '- store: 가게 이름(한국어로 읽기 쉽게, 15자 이내. 모르면 빈 문자열)\n' +
  '- items: 산 품목을 한국어로 짧게 요약(예: "라멘, 교자, 맥주"). 20자 이내, 많으면 대표 2~3개 + "외"\n' +
  '- cat: food(식당·카페·편의점 음식), move(교통), shop(쇼핑), do(관광·입장), etc(그 외) 중 하나\n' +
  '영수증이 아니거나 읽을 수 없으면 total을 0으로 한다.';

async function take(trip, kind) {
  try {
    const r = await fetch(SB + '/rest/v1/rpc/rc_take', {
      method: 'POST', headers: { apikey: SB_KEY, 'Content-Type': 'application/json' },
      body: JSON.stringify({ p_trip: trip, p_kind: kind })
    });
    if (!r.ok) return 'net';
    return await r.json();
  } catch (e) { return 'net'; }
}

// 반환: {s:'ok',r} | {s:'retry'}(한도·서버 문제 → 다음 열쇠) | {s:'bad'}(사진 문제 등 → 멈춤)
async function gemini(key, model, img) {
  const payload = {
    contents: [{ parts: [{ text: PROMPT }, { inline_data: { mime_type: 'image/jpeg', data: img } }] }],
    generationConfig: {
      temperature: 0, responseMimeType: 'application/json',
      responseSchema: { type: 'OBJECT', properties: {
        total: { type: 'INTEGER' }, store: { type: 'STRING' }, items: { type: 'STRING' },
        cat: { type: 'STRING', enum: CATS } }, required: ['total', 'store', 'items', 'cat'] }
    }
  };
  let res;
  try {
    res = await fetch('https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent', {
      method: 'POST', headers: { 'Content-Type': 'application/json', 'x-goog-api-key': key },
      body: JSON.stringify(payload), signal: AbortSignal.timeout(25000)
    });
  } catch (e) { return { s: 'retry' }; }
  if (res.status === 429 || res.status >= 500 || res.status === 403) return { s: 'retry' };
  if (res.status !== 200) return { s: 'bad' };
  try {
    const t = (await res.json()).candidates[0].content.parts[0].text;
    const r = JSON.parse(t);
    return { s: 'ok', r: {
      total: Math.max(0, Math.round(+r.total || 0)), store: String(r.store || '').slice(0, 20),
      items: String(r.items || '').slice(0, 30), cat: CATS.includes(r.cat) ? r.cat : 'etc' } };
  } catch (e) { return { s: 'retry' }; }
}

module.exports = async (req, res) => {
  res.setHeader('Cache-Control', 'no-store');
  if (req.method !== 'POST') return res.status(405).json({ ok: 0, error: 'method' });
  let b = req.body;
  if (typeof b === 'string') { try { b = JSON.parse(b); } catch (e) { b = null; } }
  const img = b && b.img, trip = b && String(b.trip || '');
  if (!img || typeof img !== 'string' || img.length > 4000000) return res.status(400).json({ ok: 0, error: 'img' });
  const free = process.env.GEMINI_KEY_FREE, paid = process.env.GEMINI_KEY_PAID;
  if (!free && !paid) return res.status(200).json({ ok: 0, error: 'nokey' });
  const model = process.env.GEMINI_MODEL || 'gemini-3.1-flash-lite';
  const strong = process.env.GEMINI_MODEL_STRONG || '';

  const t = await take(trip, 'free');                    // 여행 코드 확인 + 하루 횟수
  if (t !== 'ok') return res.status(200).json({ ok: 0, error: t === 'limit' ? 'limit' : t === 'trip' ? 'trip' : 'net' });

  let out = null, usedPaid = false;
  for (const k of [free, paid]) {
    if (!k) continue;
    if (k === paid && k !== free) {                      // 유료로 넘어갈 때만 따로 세고 상한을 건다
      const p = await take(trip, 'paid');
      if (p !== 'ok') break;
      usedPaid = true;
    }
    let a = await gemini(k, model, img);
    if (a.s === 'ok' && !a.r.total && strong && strong !== model) {   // 합계를 못 읽었으면 센 모델로 한 번 더
      const a2 = await gemini(k, strong, img);
      if (a2.s === 'ok') a = a2;
    }
    if (a.s === 'ok') { out = a.r; break; }
    if (a.s === 'bad') return res.status(200).json({ ok: 0, error: 'img' });
  }
  if (!out) return res.status(200).json({ ok: 0, error: usedPaid ? 'ai' : paid ? 'limit' : 'ai' });
  return res.status(200).json(Object.assign({ ok: 1 }, out));
};
