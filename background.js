chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message?.type !== 'fetchGoogleSheet') return;
  (async () => {
    try {
      const id = String(message.id || '').trim();
      if (!/^[a-zA-Z0-9-_]+$/.test(id)) throw new Error('Invalid Google Sheets file ID.');
      const url = `https://docs.google.com/spreadsheets/d/${id}/export?format=xlsx`;
      const response = await fetch(url, { credentials: 'omit', redirect: 'follow' });
      const contentType = response.headers.get('content-type') || '';
      if (!response.ok) throw new Error(`Google Sheets returned HTTP ${response.status}.`);
      if (contentType.includes('text/html')) throw new Error('Google returned a web page instead of the spreadsheet. Make sure the sheet is shared as “Anyone with the link”.');
      const buffer = await response.arrayBuffer();
      if (buffer.byteLength < 100) throw new Error('Google returned an empty or invalid spreadsheet.');
      const bytes = new Uint8Array(buffer);
      let binary = '';
      const chunk = 0x8000;
      for (let i = 0; i < bytes.length; i += chunk) binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
      sendResponse({ ok: true, base64: btoa(binary) });
    } catch (error) {
      sendResponse({ ok: false, error: error?.message || String(error) });
    }
  })();
  return true;
});
