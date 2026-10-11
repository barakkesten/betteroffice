import { expect, test } from 'bun:test';
import { existsSync, readFileSync } from 'node:fs';
import { dirname, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import ts from 'typescript';

const root = fileURLToPath(new URL('../../../../', import.meta.url));
const barrel = resolve(root, 'packages/pptx/src/index.ts');

function source(path: string): ts.SourceFile {
  const name = relative(root, path).replace(/\\/g, '/');
  if (name.startsWith('../') || /(^|\/)(?:fixtures|__fixtures__|results)(\/|$)/.test(name) ||
    name.startsWith('tmp/perf/') || name.endsWith('.docx')) throw new Error(`Forbidden source: ${name}`);
  return ts.createSourceFile(path, readFileSync(path, 'utf8'), ts.ScriptTarget.Latest, true);
}

function modulePath(from: string, specifier: string): string {
  const base = resolve(dirname(from), specifier);
  const stem = base.replace(/\.(?:js|mjs)$/, '');
  for (const path of [base, `${stem}.ts`, `${stem}.tsx`, `${base}/index.ts`, `${base}/index.tsx`]) {
    if (/\.(?:ts|tsx|js|mjs)$/.test(path) && existsSync(path)) return path;
  }
  throw new Error(`Unresolved source: ${relative(root, from)} -> ${specifier}`);
}

test('p2_peerless_import_graph_has_no_wasm_modules', () => {
  const exports = new Map<string, string>();
  for (const statement of source(barrel).statements) {
    if (!ts.isExportDeclaration(statement) || statement.isTypeOnly || !statement.moduleSpecifier ||
      !ts.isStringLiteral(statement.moduleSpecifier) || !statement.exportClause ||
      !ts.isNamedExports(statement.exportClause)) continue;
    for (const entry of statement.exportClause.elements) {
      if (!entry.isTypeOnly) exports.set(entry.name.text, modulePath(barrel, statement.moduleSpecifier.text));
    }
  }
  const seen = new Set<string>();
  const visit = (path: string): void => {
    if (seen.has(path)) return;
    seen.add(path);
    const name = relative(root, path).replace(/\\/g, '/');
    expect(name).not.toMatch(/packages\/pptx\/src\/(?:render\/image\.ts|wasm\/(?:loader|generated\/))/);
    const follow = (specifier: string, names?: readonly string[]) => {
      if (specifier.startsWith('.')) visit(modulePath(path, specifier));
      else if (specifier === '@betteroffice/pptx') {
        if (!names) throw new Error(`Expected named pptx import: ${name}`);
        for (const imported of names) {
          const target = exports.get(imported);
          if (!target) throw new Error(`Missing source export: ${imported}`);
          visit(target);
        }
      }
    };
    const walk = (node: ts.Node): void => {
      if (ts.isImportDeclaration(node) && ts.isStringLiteral(node.moduleSpecifier)) {
        if (node.importClause?.isTypeOnly) return;
        const bindings = node.importClause?.namedBindings;
        const names = bindings && ts.isNamedImports(bindings)
          ? bindings.elements.filter((entry) => !entry.isTypeOnly).map((entry) => (entry.propertyName ?? entry.name).text)
          : undefined;
        if (names?.length === 0 && !node.importClause?.name) return;
        follow(node.moduleSpecifier.text, names);
      } else if (ts.isExportDeclaration(node) && !node.isTypeOnly && node.moduleSpecifier &&
        ts.isStringLiteral(node.moduleSpecifier)) {
        const names = node.exportClause && ts.isNamedExports(node.exportClause)
          ? node.exportClause.elements.filter((entry) => !entry.isTypeOnly).map((entry) => (entry.propertyName ?? entry.name).text)
          : undefined;
        if (names?.length !== 0) follow(node.moduleSpecifier.text, names);
      } else if (ts.isCallExpression(node) && node.expression.kind === ts.SyntaxKind.ImportKeyword &&
        node.arguments[0] && ts.isStringLiteral(node.arguments[0])) follow(node.arguments[0].text);
      ts.forEachChild(node, walk);
    };
    walk(source(path));
  };
  for (const name of ['PeerlessPptxSurface', 'PeerlessTextInput', 'commandFramePainter',
    'localEchoPainter', 'peerlessViewStore', 'usePeerlessPresentation']) {
    visit(resolve(root, `packages/pptx-react/src/worker/${name}.${name.startsWith('Peerless') ? 'tsx' : 'ts'}`));
  }
  expect(seen.has(resolve(root, 'packages/pptx/src/render/browserFrameImages.ts'))).toBe(true);
});
