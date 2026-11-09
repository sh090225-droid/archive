export async function onRequest({ request, env, params }) {
  if (request.method !== 'POST') {
    return Response.json({ error: 'POST 요청만 허용돼.' }, { status: 405 });
  }
  if (!env.GAS_WEB_APP_URL || !env.GAS_API_SECRET) {
    return Response.json({ error: 'Cloudflare 환경 변수 GAS_WEB_APP_URL 또는 GAS_API_SECRET이 설정되지 않았어.' }, { status: 500 });
  }

  const action = Array.isArray(params.path) ? params.path.join('/') : String(params.path || '');
  const allowed = new Set([
    'getInitialData', 'getCharacter', 'getPublicCharacter',
    'saveCharacter', 'uploadCharacterImage', 'deleteCharacter',
    'saveFolder', 'deleteFolder'
  ]);
  if (!allowed.has(action)) return Response.json({ error: '허용되지 않은 API야.' }, { status: 404 });

  let incoming;
  try { incoming = await request.json(); }
  catch { return Response.json({ error: '요청 본문이 올바른 JSON이 아니야.' }, { status: 400 }); }

  try {
    const upstream = await fetch(env.GAS_WEB_APP_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ secret: env.GAS_API_SECRET, action, args: Array.isArray(incoming.args) ? incoming.args : [] })
    });
    const text = await upstream.text();
    let data;
    try { data = JSON.parse(text); }
    catch { throw new Error('Apps Script가 JSON이 아닌 응답을 반환했어. 웹앱 배포 주소와 배포 권한을 확인해줘.'); }
    if (!upstream.ok || data.error) return Response.json({ error: data.error || 'Apps Script 요청 실패' }, { status: 502 });
    return Response.json(data);
  } catch (error) {
    return Response.json({ error: error.message || '백엔드 연결에 실패했어.' }, { status: 502 });
  }
}
