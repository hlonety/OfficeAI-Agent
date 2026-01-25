const https = require('https');
const fs = require('fs');
const path = require('path');

const PORT = 3000;
const MIME_TYPES = {
    '.html': 'text/html',
    '.js': 'text/javascript',
    '.css': 'text/css',
    '.json': 'application/json',
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.gif': 'image/gif',
    '.svg': 'image/svg+xml',
    '.xml': 'application/xml',
    '.ico': 'image/x-icon',
};

// 尝试加载证书
let options = {};
try {
    options = {
        key: fs.readFileSync(path.join(__dirname, 'config/ssl/key.pem')),
        cert: fs.readFileSync(path.join(__dirname, 'config/ssl/cert.pem'))
    };
} catch (e) {
    console.error("无法加载SSL证书，请确保已运行 node setup-certs.js");
    process.exit(1);
}

const server = https.createServer(options, (req, res) => {
    console.log(`${req.method} ${req.url}`);

    // 处理 CORS
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

    if (req.method === 'OPTIONS') {
        res.writeHead(204);
        res.end();
        return;
    }

    // 路径处理
    let filePath = '.' + req.url;
    if (filePath === './') {
        filePath = './src/taskpane/taskpane.html';
    }

    // 移除 query parameters
    const queryIndex = filePath.indexOf('?');
    if (queryIndex !== -1) {
        filePath = filePath.substring(0, queryIndex);
    }

    const extname = path.extname(filePath);
    let contentType = MIME_TYPES[extname] || 'application/octet-stream';

    fs.readFile(filePath, (error, content) => {
        if (error) {
            if (error.code == 'ENOENT') {
                fs.readFile('./404.html', (error, content) => {
                    res.writeHead(200, { 'Content-Type': 'text/html' });
                    res.end(content, 'utf-8');
                });
            } else {
                res.writeHead(500);
                res.end('Sorry, check with the site admin for error: ' + error.code + ' ..\n');
            }
        } else {
            res.writeHead(200, { 'Content-Type': contentType });
            res.end(content, 'utf-8');
        }
    });

});

server.listen(PORT, () => {
    console.log(`Server running at https://localhost:${PORT}/`);
    console.log(`Test URL: https://localhost:${PORT}/src/taskpane/taskpane.html`);
});
