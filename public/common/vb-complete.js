/* Completion for the Visual Basic window. Names come from the caret's
   procedure, the other standard modules on this host, and the procedures
   the virtual machine actually installs. A half-written buffer is scanned;
   it is not compiled and nothing is called. */
(function () {
  'use strict';
  const L = window.L;
  const SCALARS = ['Boolean', 'Byte', 'Integer', 'Long', 'Single', 'Double', 'Currency', 'Date', 'String', 'Variant', 'Object'];
  const FILE_NAMES = new Set('freefile eof lof loc seek input filelen getattr setattr fileattr filedatetime filecopy kill reset curdir chdir mkdir rmdir dir input$ dir$'.split(' '));
  const ORDER = { variable: 0, procedure: 1, keyword: 1, member: 2, module: 3, type: 4, constant: 5 };
  /* VB6 statement keywords (the language's own list). Me is a class instance, and this window edits a standard module. Not the editor mode's VB.NET hint words. */
  const KEYWORDS = 'DefBool DefByte DefInt DefLng DefCur DefSng DefDbl DefDate DefStr DefObj DefVar Implements Option Explicit Base Compare Text Binary Private Public Global Friend Static Dim ReDim Preserve Const As New Set Let Call Sub Function Property Get End Exit If Then Else ElseIf For Each To Step Next Do Loop While Until Wend Select Case Is With WithEvents GoTo GoSub Return On Error Resume Stop And Or Xor Not Eqv Imp Mod Like True False Nothing Null Empty ByVal ByRef Optional ParamArray Type TypeOf LSet RSet Enum RaiseEvent Event Open Close Print Write Input Line Append Output Random Access Read Shared Lock Unlock Integer Long Single Double Currency Decimal Byte Boolean String Date Object Variant Debug Load Unload'.split(' ').filter(function (name, index, all) {
    return name && name !== 'Me' && all.indexOf(name) === index;
  });

  function id(name) {
    return String(name || '').toLowerCase();
  }

  function paramText(params, variadic) {
    if (!params || !params.length) return variadic ? '...' : '';
    const text = params.map(function (param) {
      if (!param) return '';
      if (typeof param === 'string') {
        const optional = /^Optional\b/i.test(param);
        const name = param.replace(/^(?:(?:Optional|ByVal|ByRef|ParamArray)\s+)*/i, '').split(/[\s(]/)[0];
        return optional ? '[' + name + ']' : name;
      }
      return param.optional ? '[' + param.name + ']' : param.name;
    }).filter(Boolean).join(', ');
    return variadic ? text + ', ...' : text;
  }

  function kindOf(symbol) {
    if (symbol.kind === 'constant') return 'constant';
    if (symbol.kind === 'variable' || symbol.kind === 'parameter' || symbol.kind === 'field') return 'variable';
    if (symbol.accessor || String(symbol.kind).indexOf('property') === 0) return 'member';
    return 'procedure';
  }

  function fromSymbol(symbol, host) {
    return {
      name: symbol.name,
      kind: kindOf(symbol),
      type: symbol.type || '',
      params: paramText(symbol.parameters || symbol.params),
      host: !!host,
      accessor: symbol.accessor || '',
    };
  }

  function prefer(current, next) {
    if (!current) return next;
    if (current.accessor === 'get') return current;
    if (next.accessor === 'get') return next;
    return current;
  }

  function moduleSymbols(index, includePrivate, host) {
    const map = new Map();
    const symbols = index && index.symbols || [];
    for (let i = 0; i < symbols.length; i++) {
      const symbol = symbols[i];
      if (symbol.ownerId || symbol.kind === 'event' || symbol.kind === 'type' || symbol.kind === 'enum') continue;
      if (!includePrivate && symbol.scope === 'private') continue;
      const item = fromSymbol(symbol, host);
      const key = id(item.name);
      map.set(key, prefer(map.get(key), item));
    }
    return [...map.values()];
  }

  function procedureSymbols(index, procedure) {
    if (!procedure) return [];
    const map = new Map();
    const symbols = index.symbols || [];
    for (let i = 0; i < symbols.length; i++) {
      const symbol = symbols[i];
      if (symbol.ownerId !== procedure.id || symbol.kind === 'event') continue;
      const item = fromSymbol(symbol, false);
      map.set(id(item.name), item);
    }
    return [...map.values()];
  }

  function frameSymbols(frame, index) {
    const map = new Map();
    const locals = frame && frame.locals || [];
    for (let i = 0; i < locals.length; i++) {
      const local = locals[i];
      if (!local || !local.name) continue;
      map.set(id(local.name), { name: local.name, kind: 'variable', type: local.type || '', params: '', host: false, accessor: '' });
    }
    const procedure = (index.procedures || []).find(function (item) { return id(item.name) === id(frame && frame.procedure); });
    if (procedure) {
      const symbols = index.symbols || [];
      for (let i = 0; i < symbols.length; i++) {
        const symbol = symbols[i];
        if (symbol.ownerId !== procedure.id || symbol.kind !== 'parameter' || map.has(id(symbol.name))) continue;
        map.set(id(symbol.name), fromSymbol(symbol, false));
      }
    }
    return [...map.values()];
  }

  function effectiveType(item, types) {
    const declared = item && item.type || '';
    if (item && item.host && types) {
      const mapped = types[item.name];
      if (mapped && (!declared || /^object$/i.test(declared) || /^variant$/i.test(declared))) return mapped;
    }
    return declared;
  }

  function domTypeName(desc) {
    const ret = desc && desc.ret;
    if (!ret || ret === 'scalar' || ret === 'empty' || ret === 'parent' || ret === 'live') return '';
    return ret;
  }

  function domMembers(typeName, member) {
    if (!typeName || typeof member !== 'function') return [];
    const names = ['Application', 'Parent', 'Creator'];
    const seen = new Set(names.map(id));
    const spec = arguments[2];
    if (spec && spec.members) {
      for (const name in spec.members) {
        if (!seen.has(id(name))) { names.push(name); seen.add(id(name)); }
      }
    }
    const items = [];
    for (let i = 0; i < names.length; i++) {
      const desc = member(typeName, names[i]);
      if (!desc || desc.kind === 'gap') continue;
      const call = desc.kind === 'call' || desc.kind === 'both';
      items.push({
        name: names[i],
        kind: call ? 'procedure' : 'member',
        type: domTypeName(desc),
        params: '',
        host: false,
        accessor: '',
      });
    }
    return items;
  }

  function recordOf(typeName, indexes) {
    const key = id(typeName);
    if (!key) return null;
    for (let i = 0; i < indexes.length; i++) {
      const records = indexes[i].records || [];
      for (let j = 0; j < records.length; j++) if (id(records[j].name) === key) return records[j];
    }
    return null;
  }

  function membersOf(item, ctx) {
    if (!item) return [];
    if (item.members) {
      return item.members.map(function (member) {
        return {
          name: member.name,
          kind: member.kind || 'member',
          type: member.type || '',
          params: paramText(member.params),
          host: false,
          accessor: '',
        };
      });
    }
    if (item.kind === 'module' && item.moduleName) {
      const index = ctx.byName.get(id(item.moduleName));
      if (!index) return [];
      return moduleSymbols(index, item.includePrivate, item.hostModule);
    }
    const type = effectiveType(item, ctx.types);
    const record = recordOf(type, ctx.indexes);
    if (record) {
      if (record.kind === 'enum' && item.kind !== 'enum') return [];
      return (record.members || []).map(function (member) {
        return fromSymbol(member, false);
      });
    }
    if (type && ctx.domTypes && ctx.domTypes[type]) return domMembers(type, ctx.member, ctx.domTypes[type]);
    return [];
  }

  function splitChain(text) {
    const parts = [];
    const src = String(text || '').trim();
    let i = 0;
    while (i < src.length) {
      while (src[i] === ' ' || src[i] === '\t') i++;
      if (src[i] === '.') { i++; continue; }
      const match = /^[A-Za-z_\u0080-\uffff][\w\u0080-\uffff]*[$%&!#@]?/.exec(src.slice(i));
      if (!match) break;
      i += match[0].length;
      let call = false;
      while (src[i] === ' ' || src[i] === '\t') i++;
      if (src[i] === '(') {
        call = true;
        let depth = 1;
        i++;
        while (i < src.length && depth) {
          if (src[i] === '(') depth++;
          else if (src[i] === ')') depth--;
          i++;
        }
      }
      parts.push({ name: match[0], call: call });
    }
    return parts;
  }

  function findNamed(items, name) {
    const key = id(name);
    for (let i = 0; i < items.length; i++) if (id(items[i].name) === key) return items[i];
    return null;
  }

  function resolve(text, ctx) {
    const parts = splitChain(text);
    if (!parts.length) return null;
    let item = findNamed(ctx.roots, parts[0].name);
    for (let i = 1; item && i < parts.length; i++) item = findNamed(membersOf(item, ctx), parts[i].name);
    return item;
  }

  function withText(index, offset) {
    const blocks = (index.withBlocks || []).filter(function (block) { return offset >= block.start && offset <= block.end; });
    let block = blocks[blocks.length - 1];
    if (!block) return '';
    let expr = String(block.expression || '').trim();
    let guard = 0;
    while (/^\./.test(expr) && block.parent && guard < 8) {
      expr = String(block.parent.expression || '').trim() + expr;
      block = block.parent;
      guard++;
    }
    return expr.replace(/^\.+/, '').trim();
  }

  function allowedIntrinsic(item, files) {
    const key = id(item.name);
    if (key === 'shell') return false;
    if (!files && FILE_NAMES.has(key)) return false;
    return true;
  }

  function publish(item, types) {
    const type = effectiveType(item, types);
    const out = { name: item.name, kind: item.kind };
    if (type && type !== 'Void') out.type = type;
    if (item.params) out.params = item.params;
    return out;
  }

  function finish(items, span, matches, prefix, types) {
    const seen = new Set();
    const out = [];
    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      if (!item || !item.name || !matches(item.name, prefix)) continue;
      const key = id(item.name);
      if (seen.has(key)) continue;
      seen.add(key);
      out.push(publish(item, types));
    }
    out.sort(function (a, b) {
      return (ORDER[a.kind] - ORDER[b.kind]) || a.name.toLowerCase().localeCompare(b.name.toLowerCase());
    });
    return { start: span.start, end: span.end, items: out };
  }

  L.vbComplete = function (request) {
    const engine = request && (request.engine || window.VB6Engine);
    const code = String(request && request.code || '');
    const offset = Math.max(0, Math.min(code.length, request && request.offset || 0));
    const empty = { start: offset, end: offset, items: [] };
    if (!engine || typeof engine.scanDeclarations !== 'function' || typeof engine.statementBefore !== 'function') return empty;
    const statement = engine.statementBefore(code, offset);
    if (statement.state !== 'code' || /^\s*(?:#|Rem\b)/i.test(statement.masked) || /^\s*\d+\s*$/.test(statement.masked)) return empty;
    const span = engine.completionSpan(code, offset);
    const matches = engine.completionMatches;
    const before = statement.masked.slice(0, Math.max(0, span.start - statement.start));
    const typeMatch = before.match(/\b(As(?:\s+New)?|New|Implements)\s+((?:[A-Za-z_][\w]*\s*\.\s*)*)$/i);
    if (typeMatch) {
      if (/New|Implements/i.test(typeMatch[1]) || typeMatch[2]) return { start: span.start, end: span.end, items: [] };
      return finish(SCALARS.map(function (name) { return { name: name, kind: 'type', type: name, params: '', host: false }; }), span, matches, span.prefix, null);
    }
    const framed = !!(request && request.frame);
    const moduleName = request && request.moduleName || 'Module1';
    const edited = { name: moduleName, code: framed ? String(request.source || '') : code, host: false };
    const extras = (request && request.modules || []).filter(function (mod) { return mod && mod.name && id(mod.name) !== id(moduleName); });
    const indexes = [];
    const byName = new Map();
    function take(mod) {
      const index = engine.scanDeclarations({ id: mod.name, name: mod.name, kind: 'module', code: mod.code || '' });
      indexes.push(index);
      byName.set(id(mod.name), index);
      return index;
    }
    const editedIndex = take(edited);
    for (let i = 0; i < extras.length; i++) take(extras[i]);
    const procedure = framed ? null : (editedIndex.procedures || []).find(function (item) { return offset >= item.offset && offset <= item.endOffset; });
    const roots = [];
    if (framed) roots.push(...frameSymbols(request.frame, byName.get(id(request.frame.module || moduleName)) || editedIndex));
    else roots.push(...procedureSymbols(editedIndex, procedure));
    roots.push(...moduleSymbols(editedIndex, true, false));
    const modules = [edited].concat(extras);
    for (let i = 0; i < modules.length; i++) {
      roots.push({
        name: modules[i].name,
        kind: 'module',
        type: modules[i].name,
        params: '',
        host: false,
        moduleName: modules[i].name,
        includePrivate: id(modules[i].name) === id(moduleName),
        hostModule: !!modules[i].host,
        accessor: '',
      });
    }
    for (let i = 0; i < extras.length; i++) roots.push(...moduleSymbols(byName.get(id(extras[i].name)), false, !!extras[i].host));
    const intrinsics = request && request.intrinsics || {};
    const files = !!(request && request.files);
    const lists = [intrinsics.procedures || [], intrinsics.constants || [], intrinsics.objects || []];
    for (let i = 0; i < lists.length; i++) {
      for (let j = 0; j < lists[i].length; j++) {
        const item = lists[i][j];
        if (!allowedIntrinsic(item, files)) continue;
        roots.push({
          name: item.name,
          kind: item.kind || 'procedure',
          type: item.type || '',
          params: paramText(item.params, item.variadic),
          host: false,
          members: item.members || null,
          accessor: '',
        });
      }
    }
    const memberAccess = /\.\s*$/.test(code.slice(0, span.start));
    if (!memberAccess) {
      for (let i = 0; i < KEYWORDS.length; i++) {
        roots.push({ name: KEYWORDS[i], kind: 'keyword', type: '', params: '', host: false, accessor: '' });
      }
    }
    const ctx = { roots: roots, indexes: indexes, byName: byName, types: request && request.types || {}, member: request && request.member, domTypes: request && request.domTypes || null };
    let items = roots;
    if (memberAccess) {
      let left = '';
      if (typeof engine.expressionBefore === 'function') left = engine.expressionBefore(code, span.start).text || '';
      left = String(left).trim().replace(/\.$/, '').trim();
      if (!left || left === '.') left = framed ? '' : withText(editedIndex, offset);
      const receiver = left ? resolve(left, ctx) : null;
      items = receiver ? membersOf(receiver, ctx) : [];
    }
    return finish(items, span, matches, span.prefix, ctx.types);
  };
})();
