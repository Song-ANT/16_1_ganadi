export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate, max-age=0');

  try {
    const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbz9og-at_yQlXuWxNte0zPMjmdhj96G7Od6_4MQ6wl3IIeBeQZ5P6fH28K-UlA3ZyrOow/exec';
    const response = await fetch(APPS_SCRIPT_URL, {
      headers: {
        'Cache-Control': 'no-cache'
      }
    });
    const data = await response.json();
    return res.status(200).json(data);
  } catch (error) {
    return res.status(500).json({ error: error.message });
  }
}
