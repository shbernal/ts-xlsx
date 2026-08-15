/**
 * The tools. One implementation, two surfaces.
 *
 * Four rules govern every response here, and each exists because breaking it
 * produces a *confidently wrong* answer rather than a visible failure:
 *
 * 1. **Always name the profile.** Something true in Transitional and false in
 *    Strict, returned unlabelled, is a wrong answer.
 * 2. **Always resolve, never hand back a raw ref.** A `type_ref` string or an
 *    unexpanded `attributeGroup` ref makes the agent do the join, badly.
 * 3. **Absent is not prohibited.** "Not in this profile" and "no such symbol"
 *    lead to opposite next actions, so they are never the same response.
 * 4. **No raw XSD.** If a tool wants to quote, it quotes structured fields.
 */
import {DEFAULT_DB_PATH, openGraph} from './db.mjs';
import {candidateVocabularies, formatQName, loadVocabularyIndex, parseQName} from './qname.mjs';

export const PROFILE_KEYS = ['transitional', 'strict'];
const DEFAULT_PROFILE = 'transitional';

/** Content models nest; groups reference groups. Bounded, and truncation is always reported. */
const MAX_DEPTH = 12;

export function createGraph({path = DEFAULT_DB_PATH, db = null} = {}) {
  const handle = db ?? openGraph(path);
  const index = loadVocabularyIndex(handle);
  const profiles = new Map(
    handle.all('SELECT id, key, label FROM profiles ORDER BY id').map((p) => [p.key, p]),
  );

  const profileId = (key) => {
    const profile = profiles.get(key);
    if (profile === undefined) {
      throw new Error(
        `unknown profile "${key}" — expected one of: ${[...profiles.keys()].join(', ')}`,
      );
    }
    return profile.id;
  };

  // ---------------------------------------------------------------- lookup --

  /**
   * Symbols matching a qname.
   *
   * Deliberately returns *every* match rather than the best one. A bare name is
   * routinely ambiguous, and local element declarations mean even a qualified
   * name can name several symbols — `w:p` is declared inside more than one
   * type. Collapsing that to one row is how a tool ends up describing a
   * different element than the one asked about.
   *
   * @param {string} qname
   * @param {{kind?: string | string[] | null, includeAnonymous?: boolean}} [options]
   */
  function lookup(qname, {kind = null, includeAnonymous = false} = {}) {
    const parsed = parseQName(qname);
    if (parsed.localName === '') return {parsed, vocabularies: null, rows: []};
    const vocabularies = candidateVocabularies(parsed, index);
    if (vocabularies !== null && vocabularies.length === 0) {
      return {parsed, vocabularies, rows: []};
    }

    const clauses = ['s.local_name = ? COLLATE NOCASE'];
    const params = [parsed.localName];
    if (vocabularies !== null) {
      clauses.push(`v.key IN (${vocabularies.map(() => '?').join(', ')})`);
      params.push(...vocabularies);
    }
    if (kind !== null) {
      const kinds = Array.isArray(kind) ? kind : [kind];
      clauses.push(`s.kind IN (${kinds.map(() => '?').join(', ')})`);
      params.push(...kinds);
    }
    if (!includeAnonymous) clauses.push('s.is_anonymous = 0');

    const rows = handle.all(
      `SELECT s.id, s.local_name, s.kind, s.parent_symbol_id, s.type_ref, s.type_symbol_id,
              s.is_anonymous, v.key AS vocabulary
         FROM symbols s JOIN vocabularies v ON v.id = s.vocabulary_id
        WHERE ${clauses.join(' AND ')}
        ORDER BY s.parent_symbol_id, v.key, s.kind, s.id`,
      ...params,
    );
    return {parsed, vocabularies, rows};
  }

  const symbolById = (id) =>
    handle.get(
      `SELECT s.id, s.local_name, s.kind, s.parent_symbol_id, s.type_ref, s.type_symbol_id,
              s.is_anonymous, v.key AS vocabulary
         FROM symbols s JOIN vocabularies v ON v.id = s.vocabulary_id WHERE s.id = ?`,
      id,
    );

  const symbolProfiles = (id) =>
    handle
      .all(
        `SELECT p.key FROM symbol_profiles sp JOIN profiles p ON p.id = sp.profile_id
          WHERE sp.symbol_id = ? ORDER BY p.id`,
        id,
      )
      .map((r) => r.key);

  const namespaceOf = (vocabularyKey, profile) =>
    index.byKey.get(vocabularyKey)?.namespaces.find((n) => n.profile === profile) ?? null;

  const display = (row) => formatQName(row.vocabulary, row.local_name, index);

  /**
   * A stored reference, rewritten for a reader.
   *
   * The database spells references `vocabularyKey:localName` (`wml:CT_P`)
   * because the vocabulary is the identity. Nobody writes documents that way —
   * they write `w:CT_P` — so every ref leaving this module goes through here.
   * Rule 2 is about resolution, but a resolved answer in an unfamiliar spelling
   * is still work pushed back onto the caller.
   */
  const displayRef = (ref) => {
    if (ref === null || ref === undefined) return ref;
    const colon = ref.indexOf(':');
    if (colon === -1) return ref;
    const vocabulary = ref.slice(0, colon);
    if (vocabulary === 'xsd' || vocabulary === 'xml') return ref;
    if (!index.byKey.has(vocabulary)) return ref;
    return formatQName(vocabulary, ref.slice(colon + 1), index);
  };

  /** A type reference, resolved. Built-ins stay named but carry no symbol. */
  function describeType(typeRef, typeSymbolId) {
    if (typeRef === null) return null;
    if (typeSymbolId === null) return {qname: displayRef(typeRef), builtin: true};
    const target = symbolById(typeSymbolId);
    return {qname: display(target), kind: target.kind, builtin: false, id: target.id};
  }

  function describeSymbol(row, profile) {
    const owner = row.parent_symbol_id === 0 ? null : symbolById(row.parent_symbol_id);
    const namespace = namespaceOf(row.vocabulary, profile);
    return {
      id: row.id,
      qname: display(row),
      name: row.local_name,
      kind: row.kind,
      vocabulary: row.vocabulary,
      namespace: namespace === null ? null : {uri: namespace.uri, prefix: namespace.prefix},
      profiles: symbolProfiles(row.id),
      scope: owner === null ? 'global' : {declared_in: display(owner)},
      type: describeType(row.type_ref, row.type_symbol_id),
    };
  }

  /**
   * The envelope every tool returns for a name it could not resolve.
   *
   * Rule 3 lives here: three different "no" answers, because the next action
   * differs. Nothing found at all is a typo; found but not in this profile is a
   * profile switch; a named vocabulary that does not exist is a bad prefix.
   */
  function notFound(qname, parsed, vocabularies, profile) {
    if (vocabularies !== null && vocabularies.length === 0) {
      return {
        found: false,
        reason: 'unknown_vocabulary',
        message:
          `no vocabulary is named "${parsed.prefix ?? parsed.uri}". ` +
          `Known prefixes: ${knownPrefixes().join(', ')}.`,
      };
    }
    // Was it found in the *other* profile? That is a different answer.
    const elsewhere = lookup(qname).rows.filter((row) =>
      symbolProfiles(row.id).some((key) => key !== profile),
    );
    if (elsewhere.length > 0) {
      const others = [...new Set(elsewhere.flatMap((row) => symbolProfiles(row.id)))].filter(
        (key) => key !== profile,
      );
      return {
        found: false,
        reason: 'not_in_profile',
        message: `"${qname}" exists, but not in the ${profile} profile. It is in: ${others.join(', ')}.`,
        profiles: others,
      };
    }
    return {
      found: false,
      reason: 'unknown_symbol',
      message: `no symbol named "${parsed.localName}" — this is a name lookup, not a search. Try search("${parsed.localName}").`,
    };
  }

  // Aliases are included because this list is advice for a caller who just
  // mistyped a prefix, and "the prefixes that work" is the useful set — a
  // spelling we accept but omit here reads as unsupported.
  const knownPrefixes = () =>
    [
      ...new Set(
        index.all.flatMap((v) => [
          ...v.namespaces.map((n) => n.prefix),
          ...v.aliasPrefixes.map((a) => a.prefix),
        ]),
      ),
    ]
      .filter(Boolean)
      .sort();

  /** Only symbols actually present in the requested profile. */
  const inProfile = (rows, profile) =>
    rows.filter((row) => symbolProfiles(row.id).includes(profile));

  /**
   * What an ambiguous answer says, which depends on why it is ambiguous.
   *
   * The two causes need opposite advice. A name matching several *vocabularies*
   * is the caller's to fix — qualifying it picks one. A name declared in
   * several places *within* a vocabulary cannot be fixed that way: `w:tblPr` is
   * genuinely two content models and which applies depends on the parent, so
   * the answer has to carry both and say where each holds.
   */
  const ambiguityMessage = (qname, resolved, noun) =>
    resolved.ambiguity === 'vocabulary'
      ? `"${qname}" names ${resolved.variants.length} unrelated symbols that share a local name: ` +
        `${resolved.variants.map((v) => display(v.symbol)).join(', ')}. ` +
        'Qualify the name to pick one.'
      : `${qname} has ${resolved.variants.length} different ${noun} depending on where it appears.`;

  // ----------------------------------------------------------------- tools --

  /** Canonical record for a name. */
  function element(qname, {profile = DEFAULT_PROFILE} = {}) {
    profileId(profile);
    const {parsed, vocabularies, rows} = lookup(qname);
    const matches = inProfile(rows, profile);
    if (matches.length === 0)
      return {query: qname, profile, ...notFound(qname, parsed, vocabularies, profile)};
    return {
      query: qname,
      profile,
      found: true,
      count: matches.length,
      // More than one is normal, not an error: the same element name is
      // declared inside several types, and they can carry different types.
      symbols: matches.map((row) => describeSymbol(row, profile)),
    };
  }

  /** A complexType or simpleType: what it derives from and what shape it has. */
  function type(qname, {profile = DEFAULT_PROFILE} = {}) {
    const pid = profileId(profile);
    const {parsed, vocabularies, rows} = lookup(qname, {kind: ['complexType', 'simpleType']});
    const matches = inProfile(rows, profile);
    if (matches.length === 0)
      return {query: qname, profile, ...notFound(qname, parsed, vocabularies, profile)};

    return {
      query: qname,
      profile,
      found: true,
      count: matches.length,
      types: matches.map((row) => {
        const base = handle.get(
          `SELECT relation, content_model, base_type_ref, base_symbol_id
             FROM inheritance_edges WHERE symbol_id = ? AND profile_id = ?`,
          row.id,
          pid,
        );
        const summary =
          row.kind === 'simpleType'
            ? {value_space: describeValueSpace(row.id, pid, 0)}
            : {
                content: contentSummary(row.id, pid),
                attribute_count: handle.get(
                  'SELECT COUNT(*) AS n FROM attr_edges WHERE symbol_id = ? AND profile_id = ?',
                  row.id,
                  pid,
                ).n,
              };
        return {
          ...describeSymbol(row, profile),
          derivation:
            base === undefined
              ? null
              : {
                  relation: base.relation,
                  content_model: base.content_model,
                  base:
                    base.base_symbol_id === null
                      ? {qname: displayRef(base.base_type_ref), builtin: true}
                      : {qname: display(symbolById(base.base_symbol_id)), builtin: false},
                },
          ...summary,
        };
      }),
    };
  }

  /** A one-line shape for `type`, so it does not have to embed the whole tree. */
  function contentSummary(symbolId, pid) {
    const compositors = handle.all(
      'SELECT kind FROM compositors WHERE parent_symbol_id = ? AND profile_id = ? ORDER BY order_index',
      symbolId,
      pid,
    );
    const children = handle.get(
      'SELECT COUNT(*) AS n FROM child_edges WHERE parent_symbol_id = ? AND profile_id = ?',
      symbolId,
      pid,
    ).n;
    const groups = handle.get(
      `SELECT COUNT(*) AS n FROM group_edges
        WHERE parent_symbol_id = ? AND profile_id = ? AND ref_kind = 'group'`,
      symbolId,
      pid,
    ).n;
    if (compositors.length === 0 && groups === 0) return 'empty';
    return {top: compositors.map((c) => c.kind), direct_children: children, group_refs: groups};
  }

  // ------------------------------------------------------------- children ---

  /**
   * Legal children, in schema order.
   *
   * The differentiator, so the output is both shapes at once: the **compositor
   * tree**, which is what actually answers "can I put this here, and where",
   * and a **flat ordered list** for the common case. A flat list alone loses
   * the choice/sequence distinction, which is most of the question.
   *
   * Two resolutions happen before anything is returned, and skipping either
   * produces a plausible wrong answer:
   *
   * - **Inheritance.** For an `extension`, the base's particle comes *first*,
   *   then the derived one. Every contributed node says which type it came
   *   from, so the agent can see why. A `restriction` replaces the base's
   *   particle rather than extending it, and is not prepended.
   * - **Group refs.** A type whose whole content is a `group ref` has zero
   *   `child_edges` — `CT_SolidColorFillProperties` is the standard example —
   *   so a query that reads only `child_edges` answers "accepts nothing" for a
   *   large part of DrawingML. Groups are expanded in place, carrying the
   *   *ref site's* cardinality, not the definition's.
   */
  function children(qname, {profile = DEFAULT_PROFILE, expandGroups = true} = {}) {
    const pid = profileId(profile);
    const resolved = resolveToType(qname, profile);
    if (resolved.found === false) return {query: qname, profile, ...resolved};

    const describe = ({symbol, via, declaredIn}) => {
      const state = {seenGroups: new Set(), truncated: false};
      const tree = particlesOf(symbol.id, pid, state, 0, expandGroups);
      return {
        type: display(symbol),
        ...(via === null ? {} : {resolved_from: via}),
        ...(declaredIn.length === 0 ? {} : {applies_when_declared_in: declaredIn}),
        truncated: state.truncated,
        tree,
        order: flatten(tree),
      };
    };

    // The common case stays flat; only a genuinely ambiguous name pays for the
    // extra nesting, and then it pays it visibly.
    if (resolved.variants.length === 1) {
      return {query: qname, profile, found: true, ...describe(resolved.variants[0])};
    }
    return {
      query: qname,
      profile,
      found: true,
      ambiguous: true,
      message: ambiguityMessage(qname, resolved, 'content models'),
      variants: resolved.variants.map(describe),
    };
  }

  /**
   * An element resolves to its type; a type is already one.
   *
   * Returns **every** distinct type the name resolves to, because in this
   * corpus one element name genuinely has several. `w:tblPr` is `CT_TblPr`
   * inside `CT_Tbl` and `CT_TblPrBase` in three other places — a real
   * difference, and the second most common thing anyone asks about tables.
   *
   * Refusing to answer until the caller disambiguates would be correct and
   * useless; picking one silently would be wrong. So each variant is resolved
   * and labelled with where it applies, and the caller decides.
   */
  function resolveToType(qname, profile) {
    const {parsed, vocabularies, rows} = lookup(qname);
    const matches = inProfile(rows, profile);
    if (matches.length === 0) return notFound(qname, parsed, vocabularies, profile);

    // A global type name is unique per vocabulary but not across them:
    // `CT_Shape` is declared in six, and they are unrelated types that happen
    // to share a name. Returning the first is the silent pick rule 3 forbids,
    // and it is the wrong one five times out of six.
    const asType = matches.filter((r) => r.kind === 'complexType' || r.kind === 'group');
    if (asType.length > 0) {
      return {
        found: true,
        ambiguity: 'vocabulary',
        variants: asType.map((symbol) => ({symbol, via: null, declaredIn: []})),
      };
    }

    const declarations = matches.filter((r) => r.type_symbol_id !== null);
    if (declarations.length === 0) {
      const builtin = matches.find((r) => r.type_ref !== null);
      return {
        found: false,
        reason: 'no_content_model',
        message: builtin
          ? `${qname} has type ${displayRef(builtin.type_ref)}, a built-in with no element content.`
          : `${qname} declares no type, so it has no children in the schema.`,
      };
    }

    // Group the declarations by the type they point at, so the answer is one
    // entry per genuinely different content model rather than one per site.
    const byType = new Map();
    for (const row of declarations) {
      if (!byType.has(row.type_symbol_id)) byType.set(row.type_symbol_id, []);
      byType.get(row.type_symbol_id).push(row);
    }

    return {
      found: true,
      ambiguity: 'declaration_site',
      variants: [...byType.entries()].map(([typeSymbolId, sites]) => ({
        symbol: symbolById(typeSymbolId),
        via: `${display(sites[0])} -> ${displayRef(sites[0].type_ref)}`,
        declaredIn: sites.map((r) =>
          r.parent_symbol_id === 0 ? 'global' : display(symbolById(r.parent_symbol_id)),
        ),
      })),
    };
  }

  /**
   * The particle tree of one type, with inheritance and groups resolved.
   *
   * `contributedBy` threads the originating type down through the recursion so
   * every node in the answer can say where it came from — the difference
   * between "these are the legal children" and "these are the legal children
   * *and here is why*".
   */
  /** @param {string | null} contributedBy the type that contributed these particles, if inherited */
  function particlesOf(symbolId, pid, state, depth, expandGroups, contributedBy = null) {
    if (depth > MAX_DEPTH) {
      state.truncated = true;
      return [];
    }
    const nodes = [];

    const base = handle.get(
      `SELECT relation, content_model, base_symbol_id FROM inheritance_edges
        WHERE symbol_id = ? AND profile_id = ?`,
      symbolId,
      pid,
    );
    // Extension prepends the base particle; restriction replaces it. Getting
    // this backwards reorders the legal children of every derived type.
    if (base !== undefined && base.relation === 'extension' && base.base_symbol_id !== null) {
      const from = display(symbolById(base.base_symbol_id));
      nodes.push(...particlesOf(base.base_symbol_id, pid, state, depth + 1, expandGroups, from));
    }

    for (const node of topLevelParticles(
      symbolId,
      pid,
      state,
      depth,
      expandGroups,
      contributedBy,
    )) {
      nodes.push(node);
    }
    return nodes;
  }

  /** Compositors and group refs hanging directly off a definition, in document order. */
  function topLevelParticles(symbolId, pid, state, depth, expandGroups, contributedBy) {
    const compositors = handle
      .all(
        `SELECT id, kind, min_occurs, max_occurs, order_index FROM compositors
          WHERE parent_symbol_id = ? AND profile_id = ? ORDER BY order_index`,
        symbolId,
        pid,
      )
      .map((c) => ({...c, _sort: c.order_index, _kind: 'compositor'}));

    const groups = handle
      .all(
        `SELECT id, group_symbol_id, min_occurs, max_occurs, order_index FROM group_edges
          WHERE parent_symbol_id = ? AND profile_id = ? AND ref_kind = 'group' AND compositor_id IS NULL
          ORDER BY order_index`,
        symbolId,
        pid,
      )
      .map((g) => ({...g, _sort: g.order_index, _kind: 'group'}));

    return [...compositors, ...groups]
      .sort((a, b) => a._sort - b._sort)
      .map((entry) =>
        entry._kind === 'compositor'
          ? compositorNode(entry, pid, state, depth, expandGroups, contributedBy)
          : groupNode(entry, pid, state, depth, expandGroups, contributedBy),
      );
  }

  function compositorNode(compositor, pid, state, depth, expandGroups, contributedBy) {
    return {
      kind: compositor.kind,
      min: compositor.min_occurs,
      max: compositor.max_occurs,
      ...(contributedBy === null ? {} : {from: contributedBy}),
      children: compositorChildren(compositor.id, pid, state, depth, expandGroups, contributedBy),
    };
  }

  /** Elements, wildcards, nested compositors and group refs inside one compositor. */
  function compositorChildren(compositorId, pid, state, depth, expandGroups, contributedBy) {
    if (depth > MAX_DEPTH) {
      state.truncated = true;
      return [];
    }

    const elements = handle
      .all(
        `SELECT child_symbol_id, min_occurs, max_occurs, order_index,
                is_wildcard, wildcard_namespace, wildcard_process_contents
           FROM child_edges WHERE compositor_id = ? AND profile_id = ? ORDER BY order_index`,
        compositorId,
        pid,
      )
      .map((e) => ({...e, _sort: e.order_index, _kind: 'child'}));

    const nested = handle
      .all(
        `SELECT id, kind, min_occurs, max_occurs, order_index FROM compositors
          WHERE parent_compositor_id = ? AND profile_id = ? ORDER BY order_index`,
        compositorId,
        pid,
      )
      .map((c) => ({...c, _sort: c.order_index, _kind: 'compositor'}));

    const groups = handle
      .all(
        `SELECT id, group_symbol_id, min_occurs, max_occurs, order_index FROM group_edges
          WHERE compositor_id = ? AND profile_id = ? AND ref_kind = 'group' ORDER BY order_index`,
        compositorId,
        pid,
      )
      .map((g) => ({...g, _sort: g.order_index, _kind: 'group'}));

    return [...elements, ...nested, ...groups]
      .sort((a, b) => a._sort - b._sort)
      .map((entry) => {
        if (entry._kind === 'compositor') {
          return compositorNode(entry, pid, state, depth + 1, expandGroups, contributedBy);
        }
        if (entry._kind === 'group') {
          return groupNode(entry, pid, state, depth + 1, expandGroups, contributedBy);
        }
        if (entry.is_wildcard === 1) {
          return {
            kind: 'any',
            namespace: entry.wildcard_namespace,
            process_contents: entry.wildcard_process_contents,
            min: entry.min_occurs,
            max: entry.max_occurs,
            ...(contributedBy === null ? {} : {from: contributedBy}),
          };
        }
        const child = symbolById(entry.child_symbol_id);
        return {
          kind: 'element',
          qname: display(child),
          type: displayRef(child.type_ref),
          min: entry.min_occurs,
          max: entry.max_occurs,
          ...(contributedBy === null ? {} : {from: contributedBy}),
        };
      });
  }

  /**
   * A group reference, expanded in place.
   *
   * The min/max reported are the **ref site's**: `<xsd:group ref="EG_X"
   * maxOccurs="unbounded"/>` says the group repeats here, which is not what the
   * group's own definition says.
   *
   * A group already open on this path is marked `recursive` rather than
   * expanded again. Measured on the built graph, ECMA-376 has 19
   * group-to-group references and **no cycles among them**, so this guard does
   * not currently fire — the recursion in OOXML runs through *types*
   * (CT_Tbl -> CT_Row -> CT_Cell -> block content), which this function does
   * not follow. It stays anyway: the failure it prevents is a hang rather than
   * a wrong answer, and one added group ref in a future edition is all it takes.
   */
  function groupNode(edge, pid, state, depth, expandGroups, contributedBy) {
    const group = symbolById(edge.group_symbol_id);
    /** @type {{kind: string, qname: string, min: number, max: number, from?: string, children?: any[]}} */
    const node = {
      kind: 'group',
      qname: display(group),
      min: edge.min_occurs,
      max: edge.max_occurs,
      ...(contributedBy === null ? {} : {from: contributedBy}),
    };
    if (!expandGroups) return node;
    if (state.seenGroups.has(group.id)) return {...node, recursive: true};
    if (depth > MAX_DEPTH) {
      state.truncated = true;
      return {...node, truncated: true};
    }

    state.seenGroups.add(group.id);
    node.children = topLevelParticles(group.id, pid, state, depth + 1, expandGroups, contributedBy);
    state.seenGroups.delete(group.id);
    return node;
  }

  /** Every element the tree permits, in order, flattened for the common case. */
  function flatten(nodes, path = [], out = []) {
    for (const node of nodes) {
      if (node.kind === 'element') {
        out.push({
          qname: node.qname,
          type: node.type,
          min: node.min,
          max: node.max,
          in: path.join(' > ') || 'top',
          ...(node.from === undefined ? {} : {from: node.from}),
        });
      } else if (node.kind === 'any') {
        out.push({
          wildcard: node.namespace,
          process_contents: node.process_contents,
          in: path.join(' > ') || 'top',
        });
      } else if (node.children !== undefined) {
        const label = node.kind === 'group' ? `group ${node.qname}` : node.kind;
        flatten(node.children, [...path, label], out);
      }
    }
    return out;
  }

  // ----------------------------------------------------------- attributes ---

  /**
   * Attributes of a type, including inherited ones and attributeGroup
   * expansions.
   *
   * Everything is resolved: an `attributeGroup` ref is expanded to the
   * attributes it carries, and each attribute names where it came from. An
   * agent asked to add an attribute needs the whole legal set, not the set this
   * one type happens to declare directly.
   */
  function attributes(qname, {profile = DEFAULT_PROFILE} = {}) {
    const pid = profileId(profile);
    const resolved = resolveToType(qname, profile);
    if (resolved.found === false) return {query: qname, profile, ...resolved};

    const collectFor = (rootSymbolId) => {
      const collected = [];
      const seenGroups = new Set();

      const collect = (symbolId, from, depth) => {
        if (depth > MAX_DEPTH) return;

        const base = handle.get(
          `SELECT relation, base_symbol_id FROM inheritance_edges
          WHERE symbol_id = ? AND profile_id = ?`,
          symbolId,
          pid,
        );
        // Both extension and restriction inherit attributes — unlike particles,
        // where restriction replaces. A restriction may narrow an attribute's
        // use, and the derived declaration wins, which is why the base is
        // collected first and later duplicates override.
        if (base !== undefined && base.base_symbol_id !== null) {
          collect(base.base_symbol_id, display(symbolById(base.base_symbol_id)), depth + 1);
        }

        for (const row of handle.all(
          `SELECT local_name, attr_use, is_qualified, default_value, fixed_value,
                type_ref, type_symbol_id, attr_symbol_id, order_index
           FROM attr_edges WHERE symbol_id = ? AND profile_id = ? ORDER BY order_index`,
          symbolId,
          pid,
        )) {
          collected.push({
            name: row.local_name,
            qualified: row.is_qualified === 1,
            use: row.attr_use,
            type: describeType(row.type_ref, row.type_symbol_id),
            ...(row.default_value === null ? {} : {default: row.default_value}),
            ...(row.fixed_value === null ? {} : {fixed: row.fixed_value}),
            ...(from === null ? {} : {from}),
          });
        }

        for (const row of handle.all(
          `SELECT group_symbol_id FROM group_edges
          WHERE parent_symbol_id = ? AND profile_id = ? AND ref_kind = 'attributeGroup'
          ORDER BY order_index`,
          symbolId,
          pid,
        )) {
          if (seenGroups.has(row.group_symbol_id)) continue;
          seenGroups.add(row.group_symbol_id);
          collect(row.group_symbol_id, display(symbolById(row.group_symbol_id)), depth + 1);
        }
      };

      collect(rootSymbolId, null, 0);

      // A derived type redeclaring an inherited attribute wins; it was
      // collected later, so the last write is the right one.
      const byName = new Map();
      for (const attribute of collected) byName.set(attribute.name, attribute);
      return [...byName.values()].sort((a, b) => a.name.localeCompare(b.name, 'en'));
    };

    const describe = ({symbol, via, declaredIn}) => {
      const list = collectFor(symbol.id);
      return {
        type: display(symbol),
        ...(via === null ? {} : {resolved_from: via}),
        ...(declaredIn.length === 0 ? {} : {applies_when_declared_in: declaredIn}),
        count: list.length,
        attributes: list,
      };
    };

    if (resolved.variants.length === 1) {
      return {query: qname, profile, found: true, ...describe(resolved.variants[0])};
    }
    return {
      query: qname,
      profile,
      found: true,
      ambiguous: true,
      message: ambiguityMessage(qname, resolved, 'types'),
      variants: resolved.variants.map(describe),
    };
  }

  // --------------------------------------------------------------- values ---

  /**
   * Every distinct simple type a name resolves to.
   *
   * `resolveToType`'s counterpart for the value-space tools, and the sharper
   * case of the same hazard: `ST_Direction` is `ltr|rtl` in wml, `horz|vert` in
   * pml and `norm|rev` in dml-diagram. Collapsing that to one is not a partial
   * answer, it is a wrong one that reads as authoritative.
   *
   * Deduplicated on the resolved *target*, so several declarations pointing at
   * one type stay a single answer rather than becoming false ambiguity.
   */
  function resolveToSimpleType(qname, profile) {
    const {parsed, vocabularies, rows} = lookup(qname, {
      kind: ['simpleType', 'attribute', 'element'],
    });
    const matches = inProfile(rows, profile);
    if (matches.length === 0) return notFound(qname, parsed, vocabularies, profile);

    const targets = [];
    const seen = new Set();
    for (const row of matches) {
      const symbolId = row.kind === 'simpleType' ? row.id : row.type_symbol_id;
      const key = symbolId === null ? `ref:${row.type_ref}` : `id:${symbolId}`;
      if (seen.has(key)) continue;
      seen.add(key);
      targets.push({row, symbolId});
    }
    return {found: true, targets};
  }

  const targetName = ({row, symbolId}) =>
    symbolId === null ? displayRef(row.type_ref) : display(symbolById(symbolId));

  const simpleTypeAmbiguityMessage = (qname, targets) =>
    `"${qname}" resolves to ${targets.length} unrelated types that share a local name: ` +
    `${targets.map(targetName).join(', ')}. Qualify the name to pick one.`;

  /** Enumeration values only. Says so usefully when the type is not enumerated. */
  function enumValues(qname, {profile = DEFAULT_PROFILE} = {}) {
    const pid = profileId(profile);
    const resolved = resolveToSimpleType(qname, profile);
    if (resolved.found === false) return {query: qname, profile, ...resolved};

    const describe = (target) => {
      if (target.symbolId === null) {
        return {
          type: targetName(target),
          enumerated: false,
          reason: 'not_a_simple_type',
          message: `${display(target.row)} has no simple type to enumerate.`,
          values: [],
        };
      }
      const name = targetName(target);
      const values = handle.all(
        'SELECT value FROM enums WHERE symbol_id = ? AND profile_id = ? ORDER BY order_index',
        target.symbolId,
        pid,
      );
      // Not an error — most simple types are bounded by facets or unions rather
      // than enumerated, and `values` is the tool that answers those.
      if (values.length === 0) {
        return {
          type: name,
          enumerated: false,
          message: `${name} is not an enumerated type; use values() for its facets and union members.`,
          values: [],
        };
      }
      return {
        type: name,
        enumerated: true,
        count: values.length,
        values: values.map((v) => v.value),
      };
    };

    if (resolved.targets.length === 1) {
      return {query: qname, profile, found: true, ...describe(resolved.targets[0])};
    }
    return {
      query: qname,
      profile,
      found: true,
      ambiguous: true,
      message: simpleTypeAmbiguityMessage(qname, resolved.targets),
      variants: resolved.targets.map(describe),
    };
  }

  /**
   * The whole legal value space of a simple type: base, facets, enumeration,
   * union members and list item type, resolved recursively.
   *
   * This is what `enum` cannot answer. `ST_TwipsMeasure` is a union,
   * `ST_Percentage` is a pattern, and the numeric family is bounded by
   * inclusive facets; an agent *writing* OOXML needs the bounds and the
   * pattern, not "it restricts xsd:string".
   */
  function values(qname, {profile = DEFAULT_PROFILE} = {}) {
    const pid = profileId(profile);
    const resolved = resolveToSimpleType(qname, profile);
    if (resolved.found === false) return {query: qname, profile, ...resolved};

    const describe = (target) => {
      if (target.symbolId === null) {
        if (target.row.type_ref === null) {
          return {
            type: null,
            message: `${display(target.row)} declares no type, so it has no value space in the schema.`,
          };
        }
        const name = targetName(target);
        return {
          type: name,
          builtin: true,
          message: `${display(target.row)} is typed ${name}, an XSD built-in; its value space is the XSD one.`,
        };
      }
      return {type: targetName(target), ...describeValueSpace(target.symbolId, pid, 0)};
    };

    if (resolved.targets.length === 1) {
      return {query: qname, profile, found: true, ...describe(resolved.targets[0])};
    }
    return {
      query: qname,
      profile,
      found: true,
      ambiguous: true,
      message: simpleTypeAmbiguityMessage(qname, resolved.targets),
      variants: resolved.targets.map(describe),
    };
  }

  function describeValueSpace(symbolId, pid, depth) {
    if (depth > MAX_DEPTH) return {truncated: true};

    const base = handle.get(
      `SELECT base_type_ref, base_symbol_id FROM inheritance_edges
        WHERE symbol_id = ? AND profile_id = ? AND content_model = 'simpleType'`,
      symbolId,
      pid,
    );
    const enumerated = handle.all(
      'SELECT value FROM enums WHERE symbol_id = ? AND profile_id = ? ORDER BY order_index',
      symbolId,
      pid,
    );
    const facets = handle.all(
      'SELECT facet, value FROM simple_type_facets WHERE symbol_id = ? AND profile_id = ? ORDER BY order_index',
      symbolId,
      pid,
    );
    const members = handle.all(
      `SELECT member_kind, member_type_ref, member_symbol_id, order_index FROM union_members
        WHERE symbol_id = ? AND profile_id = ? ORDER BY order_index`,
      symbolId,
      pid,
    );

    const space = {};
    if (base !== undefined) space.base = displayRef(base.base_type_ref);
    if (enumerated.length > 0) space.enumeration = enumerated.map((e) => e.value);
    if (facets.length > 0) {
      space.facets = Object.fromEntries(facets.map((f) => [f.facet, f.value]));
    }
    if (members.length > 0) {
      space[members[0].member_kind === 'list' ? 'list_of' : 'one_of'] = members.map((member) => {
        if (member.member_symbol_id === null)
          return {type: displayRef(member.member_type_ref), builtin: true};
        const symbol = symbolById(member.member_symbol_id);
        return {
          // Anonymous members have no usable name — they are the inline
          // alternatives of a union — so they are described inline instead.
          ...(symbol.is_anonymous === 1 ? {inline: true} : {type: display(symbol)}),
          ...describeValueSpace(member.member_symbol_id, pid, depth + 1),
        };
      });
    }
    if (Object.keys(space).length === 0) space.unconstrained = true;
    return space;
  }

  // ------------------------------------------------------------ namespace ---

  /**
   * URI to vocabulary and back, in both directions, for both profiles.
   *
   * An alias prefix matches here too, and says so: `prefix` keeps reporting what
   * the schemas bind (NULL when they bind nothing), and `aliases` carries the
   * accepted spellings with the citation for each. Folding an alias into
   * `prefix` would make this tool claim the standard mandates a prefix it does
   * not, which is the claim the NULL exists to avoid.
   */
  function namespace(query, {profile = null} = {}) {
    const text = String(query ?? '').trim();
    const matches = [];
    for (const vocabulary of index.all) {
      const aliases = vocabulary.aliasPrefixes;
      const aliasHit = aliases.some((a) => a.prefix === text);
      for (const ns of vocabulary.namespaces) {
        if (profile !== null && ns.profile !== profile) continue;
        if (ns.uri === text || ns.prefix === text || vocabulary.key === text || aliasHit) {
          matches.push({
            vocabulary: vocabulary.key,
            uri: ns.uri,
            prefix: ns.prefix,
            profile: ns.profile,
            ...(aliases.length === 0 ? {} : {aliases}),
          });
        }
      }
    }
    if (matches.length === 0) {
      return {
        query: text,
        found: false,
        reason: 'unknown_namespace',
        message: `no namespace, prefix or vocabulary matches "${text}".`,
        known_prefixes: knownPrefixes(),
      };
    }
    return {query: text, found: true, count: matches.length, namespaces: matches};
  }

  // --------------------------------------------------------------- search ---

  /**
   * Substring match on symbol names. **Not semantic, and it says so.**
   *
   * There are no embeddings here by design. That is fine for the real use — an
   * agent asking about OOXML usually already knows the qname and wants the
   * record — but the tool must not imply otherwise, and an empty result must
   * say "no name contains X" rather than returning nothing ambiguously.
   */
  /**
   * @param {string} text
   * @param {{profile?: string, kind?: string | string[] | null, limit?: number}} [options]
   */
  function search(text, {profile = DEFAULT_PROFILE, kind = null, limit = 40} = {}) {
    const pid = profileId(profile);
    const needle = String(text ?? '').trim();
    if (needle === '') {
      return {
        query: needle,
        profile,
        found: false,
        reason: 'empty_query',
        message: 'search needs a substring.',
      };
    }

    const params = [pid, `%${needle}%`];
    let kindClause = '';
    if (kind !== null) {
      const kinds = Array.isArray(kind) ? kind : [kind];
      kindClause = ` AND s.kind IN (${kinds.map(() => '?').join(', ')})`;
      params.push(...kinds);
    }

    const rows = handle.all(
      `SELECT s.id, s.local_name, s.kind, s.parent_symbol_id, s.type_ref, v.key AS vocabulary
         FROM symbols s
         JOIN vocabularies v ON v.id = s.vocabulary_id
         JOIN symbol_profiles sp ON sp.symbol_id = s.id AND sp.profile_id = ?
        WHERE s.local_name LIKE ? COLLATE NOCASE AND s.is_anonymous = 0${kindClause}
        ORDER BY LENGTH(s.local_name), v.key, s.kind, s.local_name
        LIMIT ?`,
      ...params,
      limit + 1,
    );

    const truncated = rows.length > limit;
    const page = rows.slice(0, limit);
    return {
      query: needle,
      profile,
      match: 'substring, case-insensitive — this is a name search, not a semantic one',
      found: page.length > 0,
      ...(page.length === 0
        ? {message: `no symbol name contains "${needle}".`}
        : {
            count: page.length,
            truncated,
            results: page.map((row) => ({
              qname: display(row),
              kind: row.kind,
              vocabulary: row.vocabulary,
              ...(row.parent_symbol_id === 0
                ? {}
                : {declared_in: display(symbolById(row.parent_symbol_id))}),
              ...(row.type_ref === null ? {} : {type: displayRef(row.type_ref)}),
            })),
          }),
    };
  }

  // -------------------------------------------------------- diff_profiles ---

  /**
   * What Transitional adds to Strict for one symbol.
   *
   * Falls out of the profile model for free, and it is a genuinely one-sided
   * comparison: Strict contains no symbol Transitional lacks, verified in the
   * ingest, so every difference is something Transitional adds back for legacy
   * compatibility.
   */
  function diffProfiles(qname) {
    const {parsed, vocabularies, rows} = lookup(qname);
    if (rows.length === 0) {
      return {query: qname, ...notFound(qname, parsed, vocabularies, DEFAULT_PROFILE)};
    }

    return {
      query: qname,
      found: true,
      count: rows.length,
      symbols: rows.map((row) => {
        const present = symbolProfiles(row.id);
        const perProfile = {};
        for (const key of present) {
          const pid = profileId(key);
          perProfile[key] = {
            namespace: namespaceOf(row.vocabulary, key)?.uri ?? null,
            children: handle.get(
              'SELECT COUNT(*) AS n FROM child_edges WHERE parent_symbol_id = ? AND profile_id = ?',
              row.id,
              pid,
            ).n,
            attributes: handle.get(
              'SELECT COUNT(*) AS n FROM attr_edges WHERE symbol_id = ? AND profile_id = ?',
              row.id,
              pid,
            ).n,
            enumeration: handle
              .all(
                'SELECT value FROM enums WHERE symbol_id = ? AND profile_id = ? ORDER BY order_index',
                row.id,
                pid,
              )
              .map((e) => e.value),
            union_members: handle
              .all(
                `SELECT member_type_ref FROM union_members
                  WHERE symbol_id = ? AND profile_id = ? ORDER BY order_index`,
                row.id,
                pid,
              )
              .map((m) => displayRef(m.member_type_ref)),
          };
        }

        const missing = PROFILE_KEYS.filter((key) => !present.includes(key));
        return {
          qname: display(row),
          kind: row.kind,
          in_profiles: present,
          ...(missing.length > 0 ? {absent_from: missing} : {}),
          differences: describeDifferences(perProfile, present),
          detail: perProfile,
        };
      }),
    };
  }

  function describeDifferences(perProfile, present) {
    if (present.length < 2) return [`present only in ${present.join(', ')}`];
    const [a, b] = present;
    const notes = [];
    for (const field of ['children', 'attributes']) {
      if (perProfile[a][field] !== perProfile[b][field]) {
        notes.push(`${field}: ${a}=${perProfile[a][field]}, ${b}=${perProfile[b][field]}`);
      }
    }
    for (const field of ['enumeration', 'union_members']) {
      const onlyA = perProfile[a][field].filter((v) => !perProfile[b][field].includes(v));
      const onlyB = perProfile[b][field].filter((v) => !perProfile[a][field].includes(v));
      if (onlyA.length > 0) notes.push(`${field} only in ${a}: ${onlyA.join(', ')}`);
      if (onlyB.length > 0) notes.push(`${field} only in ${b}: ${onlyB.join(', ')}`);
    }
    if (perProfile[a].namespace !== perProfile[b].namespace) {
      notes.push('namespace URI differs (expected — the profiles alias the same vocabulary)');
    }
    return notes.length === 0 ? ['identical apart from the namespace URI'] : notes;
  }

  return {
    element,
    type,
    children,
    attributes,
    enum: enumValues,
    values,
    namespace,
    search,
    diff_profiles: diffProfiles,
    // Shared with core/explain.mjs, which is a composition of these rather than
    // a second query layer.
    _internal: {handle, index, profileId, lookup, inProfile, display, symbolById, resolveToType},
    close: () => handle.close(),
  };
}
