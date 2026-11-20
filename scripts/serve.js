// Optional: a tiny static file server for local use.  node scripts/serve.js [port]
// (You can also just open index.html in a browser: nothing here needs a server.)
const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');

const root = path.resolve(__dirname, '..');
const port = Number(process.argv[2]) || 8080;
const types = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8',
  '.json': 'application/json', '.svg': 'image/svg+xml', '.xlsx': 'application/octet-stream',
};

http.createServer((req, res) => {
  const urlPath = decodeURIComponent(new URL(req.url, 'http://localhost').pathname);
  const file = path.resolve(root, '.' + (urlPath.endsWith('/') ? urlPath + 'index.html' : urlPath));
  if (!file.startsWith(root + path.sep) || !fs.existsSync(file) || !fs.statSync(file).isFile()) {
    res.writeHead(404, { 'Content-Type': 'text/plain' });
    return res.end('Not found');
  }
  res.writeHead(200, { 'Content-Type': types[path.extname(file)] || 'application/octet-stream' });
  fs.createReadStream(file).pipe(res);
}).listen(port, () => console.log('Excel Table: http://localhost:' + port));
