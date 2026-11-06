export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  let body = req.body;
  if (typeof body === 'object') {
    body = JSON.stringify(body);
  }

  try {
    const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbz9og-at_yQlXuWxNte0zPMjmdhj96G7Od6_4MQ6wl3IIeBeQZ5P6fH28K-UlA3ZyrOow/exec';

    const response = await fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      redirect: 'follow',
      headers: {
        'Content-Type': 'text/plain;charset=utf-8'
      },
      body: String(body)
    });

    const text = await response.text();
    return res.status(200).send(text);
  } catch (error) {
    return res.status(500).json({ error: error.message });
  }
}
