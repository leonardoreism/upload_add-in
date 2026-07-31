const fs = require("fs");
const path = require("path");

const distPath = path.resolve(__dirname, "..", "dist");
const indexPath = path.join(distPath, "index.html");

if (!fs.existsSync(distPath)) {
  throw new Error("A pasta dist não foi encontrada. Execute o build primeiro.");
}

const html = `<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Projeto Integrador</title>
</head>
<body>
  <h1>Projeto Integrador</h1>
  <p>Office Add-in publicado com sucesso.</p>
  <p>Abra o suplemento pelo Microsoft Excel.</p>
</body>
</html>`;

fs.writeFileSync(indexPath, html, "utf8");

console.log("Arquivo dist/index.html criado com sucesso.");