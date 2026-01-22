const fs = require('fs');
const selfsigned = require('selfsigned');

const attrs = [{ name: 'commonName', value: 'localhost' }];
const pems = selfsigned.generate(attrs, { days: 365 });

fs.writeFileSync('cert.pem', pems.cert);
fs.writeFileSync('cert.crt', pems.cert);
fs.writeFileSync('key.pem', pems.private);

console.log('Certificates generated successfully: cert.pem, key.pem');
