const devCerts = require("office-addin-dev-certs");
const fs = require("fs");
const path = require("path");

async function createCert() {
    const sslDir = path.resolve(__dirname, "config/ssl");
    if (!fs.existsSync(sslDir)) {
        fs.mkdirSync(sslDir, { recursive: true });
    }

    try {
        console.log("Generating certificates...");
        const options = { days: 365 };
        const result = await devCerts.getHttpsServerOptions(options);

        fs.writeFileSync(path.join(sslDir, "cert.pem"), result.cert);
        fs.writeFileSync(path.join(sslDir, "key.pem"), result.key);
        console.log("Certificates created successfully in " + sslDir);
    } catch (err) {
        console.error("Error generating certificates:", err);
    }
}

createCert();
