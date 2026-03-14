export const config = { runtime: 'edge' };

const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbzBzq0RFa1kfNKmLx7IoeeWK3BFR8cA4hcYESpy56kNs6AGKuYIxJ1FGET40IrNUEs5/exec';

// Allowed actions whitelist
const ALLOWED_ACTIONS = ['read', 'write', 'update_cheque'];

// Allowed param keys whitelist
const ALLOWED_PARAMS = [
  'action', 'month', 'bancoCuenta', 'saldoMP', 'efectivoCaja', 'inversiones',
  'deudaProveedores', 'deudaServicios', 'deudaMP', 'inversionPublicidad',
  'gastoMarketing', 'ventasCorporativas', 'gastosOperativos', 'numero', 'nuevoEstado'
];

// Google OAuth token verification
async function verifyGoogleToken(token) {
  if (!token) return null;
  try {
    const res = await fetch('https://www.googleapis.com/oauth2/v3/tokeninfo?id_token=' + token);
    if (!res.ok) return null;
    const data = await res.json();
    // Token valid and not expired
    if (data.email && data.email_verified === 'true') return data.email;
    return null;
  } catch (e) {
    return null;
  }
}

// Allowed emails (set your email here)
const ALLOWED_EMAILS = [
  'rosecaloficial@gmail.com',
  'damianap9@gmail.com',
  'damianp@poletdigital.com',
  'llaverosvisibles@gmail.com',
  'martinp10105@gmail.com'
];

export default async function handler(req) {
  // CORS preflight
  if (req.method === 'OPTIONS') {
    return new Response(null, {
      status: 204,
      headers: {
        'Access-Control-Allow-Origin': 'https://rosecal.vercel.app',
        'Access-Control-Allow-Methods': 'GET, OPTIONS',
        'Access-Control-Allow-Headers': 'Authorization, Content-Type',
        'Access-Control-Max-Age': '86400',
      }
    });
  }

  // Only allow GET
  if (req.method !== 'GET') {
    return new Response(JSON.stringify({ error: 'Method not allowed' }), {
      status: 405,
      headers: { 'Content-Type': 'application/json' }
    });
  }

  // Auth check
  const authHeader = req.headers.get('Authorization');
  const token = authHeader ? authHeader.replace('Bearer ', '') : null;
  const email = await verifyGoogleToken(token);
  
  if (!email || !ALLOWED_EMAILS.includes(email.toLowerCase())) {
    return new Response(JSON.stringify({ error: 'unauthorized' }), {
      status: 401,
      headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': 'https://rosecal.vercel.app' }
    });
  }

  const { searchParams } = new URL(req.url);
  const action = searchParams.get('action') || 'read';

  // Validate action
  if (!ALLOWED_ACTIONS.includes(action)) {
    return new Response(JSON.stringify({ error: 'Invalid action' }), {
      status: 400,
      headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': 'https://rosecal.vercel.app' }
    });
  }

  // Build URL with only allowed params
  let url = APPS_SCRIPT_URL + '?action=' + encodeURIComponent(action);
  for (const [key, val] of searchParams.entries()) {
    if (key !== 'action' && ALLOWED_PARAMS.includes(key)) {
      // Sanitize: only allow alphanumeric, dots, commas, dashes, spaces
      const sanitized = String(val).replace(/[^\w.,\-\s]/g, '').substring(0, 100);
      url += '&' + encodeURIComponent(key) + '=' + encodeURIComponent(sanitized);
    }
  }

  try {
    const res = await fetch(url, { redirect: 'follow' });
    const data = await res.text();

    const isRead = action === 'read';

    return new Response(data, {
      status: 200,
      headers: {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': 'https://rosecal.vercel.app',
        'Cache-Control': isRead
          ? 's-maxage=300, stale-while-revalidate=600'
          : 'no-store',
        'X-Content-Type-Options': 'nosniff',
        'X-Frame-Options': 'DENY',
      }
    });
  } catch (e) {
    return new Response(JSON.stringify({ error: 'Server error' }), {
      status: 500,
      headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': 'https://rosecal.vercel.app' }
    });
  }
}
