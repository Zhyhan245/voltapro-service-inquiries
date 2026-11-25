const fs = require('node:fs');
const path = require('node:path');

function walkJsFiles(dir) {
  const files = [];
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      files.push(...walkJsFiles(full));
    } else if (entry.isFile() && entry.name.endsWith('.js')) {
      files.push(full);
    }
  }
  return files;
}

function analyzeFile(filePath, rootDir) {
  const content = fs.readFileSync(filePath, 'utf8');
  const lines = content.split(/\r?\n/);
  const totalLines = lines.length;
  const nonEmptyLines = lines.filter((l) => l.trim().length > 0).length;
  const commentLines = lines.filter((l) => {
    const t = l.trim();
    return t.startsWith('//') || t.startsWith('/*') || t.startsWith('*');
  }).length;
  const fnCount = (content.match(/\bfunction\b|=>/g) || []).length;
  const branchCount = (content.match(/\b(if|else if|switch|case|for|while|catch|\?)\b|&&|\|\|/g) || []).length;

  return {
    file: path.relative(rootDir, filePath).replaceAll('\\', '/'),
    totalLines,
    nonEmptyLines,
    commentLines,
    functions: fnCount,
    approxBranches: branchCount
  };
}

function main() {
  const rootDir = path.join(__dirname, '..');
  const targets = [
    ...walkJsFiles(path.join(rootDir, 'src')),
    ...walkJsFiles(path.join(rootDir, 'public', 'js'))
  ];

  const files = targets.map((f) => analyzeFile(f, rootDir));
  const summary = {
    filesCount: files.length,
    totalLines: files.reduce((s, f) => s + f.totalLines, 0),
    nonEmptyLines: files.reduce((s, f) => s + f.nonEmptyLines, 0),
    commentLines: files.reduce((s, f) => s + f.commentLines, 0),
    totalFunctions: files.reduce((s, f) => s + f.functions, 0),
    largestFile: [...files].sort((a, b) => b.totalLines - a.totalLines)[0]
  };

  const outPath = path.join(rootDir, 'evidence', 'reports', 'code-metrics.json');
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  fs.writeFileSync(outPath, JSON.stringify({ generatedAt: new Date().toISOString(), summary, files }, null, 2));
  console.log('Метрики коду:', summary);
}

main();
