export const config = { runtime: 'edge' };

const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbzBzq0RFa1kfNKmLx7IoeeWK3BFR8cA4hcYESpy56kNs6AGKuYIxJ1FGET40IrNUEs5/exec';

export default async function handler(req) {
  const { searchParams } = new URL(req.url);
  const action = searchParams.get('action') || 'read';
  const month = searchParams.get('month') || '';
  
  // Build Apps Script URL
  let url = APPS_SCRIPT_URL + '?action=' + action;
  if (month) url += '&month=' + month;
  
  // Forward all other params (for write actions)
  for (const [key, val] of searchParams.entries()) {
    if (key !== 'action' && key !== 'month') {
      url += '&' + key + '=' + encodeURIComponent(val);
    }
  }
  
  try {
    const res = await fetch(url, { redirect: 'follow' });
    const data = await res.text();
    
    // For read actions: cache for 5 minutes, stale-while-revalidate for 10 min
    // For write actions: no cache
    const isRead = action === 'read';
    
    return new Response(data, {
      status: 200,
      headers: {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': '*',
        'Cache-Control': isRead 
          ? 's-maxage=300, stale-while-revalidate=600'  // 5min cache, 10min stale OK
          : 'no-store',
      }
    });
  } catch (e) {
    return new Response(JSON.stringify({ error: e.message }), {
      status: 500,
      headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' }
    });
  }
}
