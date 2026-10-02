// XPath (XLSForm subset) to JavaScript translator for the muso port.
// Every XPath value goes through the runtime X of runtime.js (XPath 1.0 semantics:
// strings, numbers, booleans, dates as days since 1970-01-01).

export class Unsupported extends Error {}

// A straight quote closes only on a straight quote (keeps "l’enfant" whole); a curly opener
// closes on any quote of its family, straight included (the sources mix ’yes').
const QUOTE_CLOSE = { "'": ["'"], '"': ['"'], '‘': ['’', '‘', "'"], '’': ['’', '‘', "'"], '“': ['”', '“', '"'], '”': ['”', '“', '"'] };
const NAME_START = /[A-Za-z_]/;
const NAME_CHAR = /[A-Za-z0-9_.:\-]/;

export function tokenize(src) {
  const tokens = [];
  let i = 0;
  const prevIsOperand = () => {
    const p = tokens[tokens.length - 1];
    return p && ['num', 'str', 'var', 'path', 'name', 'rparen', 'dot'].includes(p.k);
  };
  while (i < src.length) {
    const c = src[i];
    if (/\s/.test(c)) { i++; continue; }
    if (c === '$' && src[i + 1] === '{') {
      const j = src.indexOf('}', i);
      if (j < 0) throw new Unsupported(`unclosed \${ in: ${src}`);
      tokens.push({ k: 'var', v: src.slice(i + 2, j).trim() });
      i = j + 1;
      continue;
    }
    if (QUOTE_CLOSE[c]) {
      let j = i + 1;
      while (j < src.length && !QUOTE_CLOSE[c].includes(src[j])) j++;
      if (j >= src.length) throw new Unsupported(`unclosed string in: ${src}`);
      tokens.push({ k: 'str', v: src.slice(i + 1, j) });
      i = j + 1;
      continue;
    }
    if (/[0-9]/.test(c) || (c === '.' && /[0-9]/.test(src[i + 1] ?? ''))) {
      const m = /^\d*\.?\d+/.exec(src.slice(i));
      tokens.push({ k: 'num', v: Number(m[0]) });
      i += m[0].length;
      continue;
    }
    if (c === '.' || c === '/') {
      // A path: '.', '..', '../x', '/data/x', './x'.
      const m = /^[./A-Za-z0-9_:\-]+/.exec(src.slice(i));
      const path = m[0];
      tokens.push(path === '.' ? { k: 'dot' } : { k: 'path', v: path });
      i += path.length;
      continue;
    }
    const two = src.slice(i, i + 2);
    if (['!=', '<=', '>='].includes(two)) { tokens.push({ k: 'op', v: two }); i += 2; continue; }
    if ('=<>+|,'.includes(c)) { tokens.push({ k: c === ',' ? 'comma' : 'op', v: c }); i++; continue; }
    if (c === '-') { tokens.push({ k: 'op', v: '-' }); i++; continue; }
    if (c === '*') { tokens.push({ k: 'op', v: '*' }); i++; continue; }
    if (c === '(') { tokens.push({ k: 'lparen' }); i++; continue; }
    if (c === ')') { tokens.push({ k: 'rparen' }); i++; continue; }
    if (NAME_START.test(c)) {
      let j = i + 1;
      while (j < src.length && NAME_CHAR.test(src[j])) j++;
      let name = src.slice(i, j);
      // A name glued to '-' then a digit or space is a subtraction, not part of the name.
      while (name.endsWith('-') || name.endsWith('.')) { name = name.slice(0, -1); j--; }
      if (prevIsOperand() && ['and', 'or', 'div', 'mod'].includes(name)) {
        tokens.push({ k: 'op', v: name });
      } else if (src[j] === '/') {
        const m = /^[./A-Za-z0-9_:\-]+/.exec(src.slice(i));
        tokens.push({ k: 'path', v: m[0] });
        j = i + m[0].length;
      } else {
        tokens.push({ k: 'name', v: name });
      }
      i = j;
      continue;
    }
    throw new Unsupported(`unexpected character '${c}' in: ${src}`);
  }
  return tokens;
}

/** Recursive descent parser: or < and < = != < relational < + - < * div mod < unary < primary. */
export function parse(src) {
  const tokens = tokenize(src);
  let p = 0;
  const peek = () => tokens[p];
  const isOp = (...ops) => peek()?.k === 'op' && ops.includes(peek().v);
  const expect = (k) => {
    if (peek()?.k !== k) throw new Unsupported(`expected ${k} at token ${p} in: ${src}`);
    return tokens[p++];
  };
  const binary = (next, ops) => () => {
    let left = next();
    while (isOp(...ops)) {
      const op = tokens[p++].v;
      left = { t: 'bin', op, left, right: next() };
    }
    return left;
  };
  const primary = () => {
    const tok = peek();
    if (!tok) throw new Unsupported(`unexpected end in: ${src}`);
    p++;
    switch (tok.k) {
      case 'num': return { t: 'num', v: tok.v };
      case 'str': return { t: 'str', v: tok.v };
      case 'var': return { t: 'var', v: tok.v };
      case 'dot': return { t: 'dot' };
      case 'path': return { t: 'path', v: tok.v };
      case 'lparen': { const e = orExpr(); expect('rparen'); return e; }
      case 'name': {
        if (peek()?.k === 'lparen') {
          p++;
          const args = [];
          if (peek()?.k !== 'rparen') {
            args.push(orExpr());
            while (peek()?.k === 'comma') { p++; args.push(orExpr()); }
          }
          expect('rparen');
          if (tok.v === 'instance' && peek()?.k === 'path') {
            return { t: 'instance', name: args[0]?.v, path: tokens[p++].v };
          }
          return { t: 'call', name: tok.v, args };
        }
        return { t: 'name', v: tok.v };
      }
      default: throw new Unsupported(`unexpected token ${tok.k} in: ${src}`);
    }
  };
  const unary = () => {
    if (isOp('-')) { p++; return { t: 'neg', e: unary() }; }
    return primary();
  };
  const mul = binary(unary, ['*', 'div', 'mod']);
  const add = binary(mul, ['+', '-']);
  const rel = binary(add, ['<', '<=', '>', '>=']);
  const eq = binary(rel, ['=', '!=']);
  const andExpr = binary(eq, ['and']);
  const orExpr = binary(andExpr, ['or']);
  const ast = orExpr();
  if (p !== tokens.length) throw new Unsupported(`trailing tokens after position ${p} in: ${src}`);
  return ast;
}

const FUNCTIONS = new Set([
  'concat', 'coalesce', 'floor', 'ceiling', 'round', 'number', 'string', 'string-length', 'substr',
  'substring', 'substring-before', 'substring-after', 'contains', 'starts-with', 'normalize-space',
  'regex', 'selected', 'count-selected', 'selected-at', 'decimal-date-time', 'decimal-date',
  'date-time', 'date', 'format-date', 'format-date-time', 'today', 'now', 'difference-in-months',
  'int', 'abs', 'min', 'max', 'boolean', 'once', 'uuid',
]);
const COMPARE = { '=': 'eq', '!=': 'ne', '<': 'lt', '<=': 'le', '>': 'gt', '>=': 'ge' };
const ARITH = { '+': '+', '-': '-', '*': '*', div: '/', mod: '%' };

/**
 * JavaScript of an XPath expression. ctx: { self (key of '.'), resolve(name) -> key,
 * choiceList(name) -> list id of a select field, choiceFilter (bare names are option
 * properties) }.
 */
export function toJs(ast, ctx) {
  const js = (n) => toJs(n, ctx);
  switch (ast.t) {
    case 'num': return String(ast.v);
    case 'str': return JSON.stringify(ast.v);
    case 'var': return ctx.varJs ? ctx.varJs(ast.v) : `X.v(app, ${JSON.stringify(ctx.resolve(ast.v))})`;
    case 'dot':
      if (!ctx.self) throw new Unsupported("'.' outside a field");
      return `X.v(app, ${JSON.stringify(ctx.self)})`;
    case 'name':
      if (ctx.choiceFilter) return `X.opt(app, ${JSON.stringify(ast.v)})`;
      throw new Unsupported(`bare name '${ast.v}'`);
    case 'path': {
      const m = /^(?:(?:\.\.\/)+|\/data\/)inputs\/(.+)$/.exec(ast.v);
      if (m) return `X.inp(app, ${JSON.stringify(m[1])})`;
      const rel = /^\.\.\/([A-Za-z_][\w-]*)$/.exec(ast.v);
      if (rel) return `X.v(app, ${JSON.stringify(ctx.resolve(rel[1]))})`;
      throw new Unsupported(`path '${ast.v}'`);
    }
    case 'instance': throw new Unsupported(`instance('${ast.name}')${ast.path}`);
    case 'neg': return `(-X.num(${js(ast.e)}))`;
    case 'bin': {
      if (ast.op === 'and') return `(X.bool(${js(ast.left)}) && X.bool(${js(ast.right)}))`;
      if (ast.op === 'or') return `(X.bool(${js(ast.left)}) || X.bool(${js(ast.right)}))`;
      if (COMPARE[ast.op]) return `X.${COMPARE[ast.op]}(${js(ast.left)}, ${js(ast.right)})`;
      if (ARITH[ast.op]) return `(X.num(${js(ast.left)}) ${ARITH[ast.op]} X.num(${js(ast.right)}))`;
      throw new Unsupported(`operator ${ast.op}`);
    }
    case 'call': {
      const name = ast.name;
      const lower = name.toLowerCase();
      if (lower === 'true' || lower === 'false') return lower;
      if (name === 'if') {
        if (ast.args.length !== 3) throw new Unsupported('if() needs 3 arguments');
        return `(X.bool(${js(ast.args[0])}) ? ${js(ast.args[1])} : ${js(ast.args[2])})`;
      }
      if (name === 'not') return `(!X.bool(${js(ast.args[0])}))`;
      if (name === 'jr:choice-name') {
        const ref = ast.args[1];
        const field = ref?.t === 'str' ? /^\$\{([^}]+)\}$/.exec(ref.v.trim())?.[1] : ref?.t === 'var' ? ref.v : null;
        if (!field) throw new Unsupported('jr:choice-name without a ${field} reference');
        return `X.choiceName(${JSON.stringify(ctx.choiceList(field))}, ${js(ast.args[0])})`;
      }
      if (!FUNCTIONS.has(name)) throw new Unsupported(`function ${name}()`);
      return `X.f[${JSON.stringify(name)}](app${ast.args.map(a => `, ${js(a)}`).join('')})`;
    }
    default: throw new Unsupported(`node ${ast.t}`);
  }
}
