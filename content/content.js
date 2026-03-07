// =============================================================================
// ADO Template Finder — Content Script (v3)
//
// v3: Uses ADO REST API to fetch FULL file content (no more Monaco scraping
//     issues with virtualized rendering). Also fetches external variable
//     template files to resolve ${{ variables.X }} expressions.
// =============================================================================

(function () {
  "use strict";

  if (window.__adoTemplateFinder) return;
  window.__adoTemplateFinder = true;

  const PANEL_ID = "ado-tf-panel";
  const TOGGLE_ID = "ado-tf-toggle";
  const DEBOUNCE_MS = 1200;
  const POLL_INTERVAL_MS = 2500;

  // Cache fetched file contents to avoid repeated API calls
  const fileCache = new Map();
  let lastAnalyzedUrl = null;
  let isRunning = false;

  // =========================================================================
  // URL helpers
  // =========================================================================

  function parseAdoUrl() {
    const url = new URL(window.location.href);
    let org, project, repo, filePath, branch;

    if (url.hostname === "dev.azure.com") {
      const parts = url.pathname.split("/").filter(Boolean);
      org = parts[0];
      project = parts[1];
      const gitIdx = parts.indexOf("_git");
      repo = gitIdx >= 0 ? decodeURIComponent(parts[gitIdx + 1]) : null;
    } else if (url.hostname.endsWith(".visualstudio.com")) {
      org = url.hostname.split(".")[0];
      const parts = url.pathname.split("/").filter(Boolean);
      project = parts[0];
      const gitIdx = parts.indexOf("_git");
      repo = gitIdx >= 0 ? decodeURIComponent(parts[gitIdx + 1]) : null;
    }

    // Extract file path from query param
    filePath = url.searchParams.get("path") || null;

    // Extract branch/version from query
    const ver = url.searchParams.get("version");
    if (ver) {
      if (ver.startsWith("GB")) branch = ver.substring(2);
    }

    return { org, project, repo, filePath, branch, hostname: url.hostname };
  }

  function buildApiBase(ctx) {
    if (ctx.hostname === "dev.azure.com") {
      return `https://dev.azure.com/${ctx.org}/${encodeURIComponent(ctx.project)}/_apis`;
    } else {
      return `https://${ctx.hostname}/${encodeURIComponent(ctx.project)}/_apis`;
    }
  }

  function buildAdoFileUrl(ctx, repoName, filePath, ref, line) {
    let normPath = filePath.replace(/\\/g, "/").replace(/^\.\//, "").replace(/^\//, "");
    normPath = "/" + normPath;

    // The repo name might be "Project/RepoName" — for the URL we need just the repo name
    const repoForUrl = repoName.includes("/") ? repoName.split("/").pop() : repoName;

    let base;
    if (ctx.hostname === "dev.azure.com") {
      base = `https://dev.azure.com/${ctx.org}/${ctx.project}/_git/${encodeURIComponent(repoForUrl)}`;
    } else {
      base = `https://${ctx.hostname}/${ctx.project}/_git/${encodeURIComponent(repoForUrl)}`;
    }

    let version = "";
    if (ref) {
      if (ref.startsWith("refs/tags/")) {
        version = `&version=GT${ref.replace("refs/tags/", "")}`;
      } else if (ref.startsWith("refs/heads/")) {
        version = `&version=GB${ref.replace("refs/heads/", "")}`;
      } else {
        version = `&version=GB${ref}`;
      }
    }

    let lineParam = "";
    if (line) {
      lineParam = `&line=${line}&lineEnd=${line}&lineStartColumn=1&lineEndColumn=999&lineStyle=plain`;
    }

    return `${base}?path=${encodeURIComponent(normPath)}${version}${lineParam}`;
  }

  // =========================================================================
  // ADO REST API — fetch file contents
  // =========================================================================

  /**
   * Fetch a file's raw text content from the ADO Git API.
   * Uses the session cookie for auth (same-origin request).
   *
   * @param {object} ctx - parsed ADO URL context
   * @param {string} repoName - repo name (or Project/RepoName)
   * @param {string} path - file path within the repo
   * @param {string} [ref] - optional ref (refs/tags/X, refs/heads/X, or branch name)
   * @returns {Promise<string|null>} raw file content or null on error
   */
  async function fetchFileContent(ctx, repoName, path, ref) {
    const cacheKey = `${repoName}::${path}::${ref || "default"}`;
    if (fileCache.has(cacheKey)) return fileCache.get(cacheKey);

    const repoForApi = repoName.includes("/") ? repoName.split("/").pop() : repoName;
    let normPath = path.replace(/\\/g, "/").replace(/^\.\//, "").replace(/^\//, "");
    normPath = "/" + normPath;

    let versionParams = "";
    if (ref) {
      if (ref.startsWith("refs/tags/")) {
        versionParams = `&versionDescriptor.version=${encodeURIComponent(ref.replace("refs/tags/", ""))}&versionDescriptor.versionType=tag`;
      } else if (ref.startsWith("refs/heads/")) {
        versionParams = `&versionDescriptor.version=${encodeURIComponent(ref.replace("refs/heads/", ""))}&versionDescriptor.versionType=branch`;
      } else {
        versionParams = `&versionDescriptor.version=${encodeURIComponent(ref)}&versionDescriptor.versionType=branch`;
      }
    }

    // Build API URL — try current project first, then cross-project if repo name has Project/Repo format
    const apiBase = buildApiBase(ctx);
    const url = `${apiBase}/git/repositories/${encodeURIComponent(repoForApi)}/items?path=${encodeURIComponent(normPath)}&includeContent=true&api-version=7.0${versionParams}&$format=text`;

    console.log(`[ADO Template Finder] Fetching: ${repoForApi}:${normPath} (ref: ${ref || 'default'})`);

    try {
      const resp = await fetch(url, {
        credentials: "same-origin",
        headers: { Accept: "text/plain" },
      });
      if (!resp.ok) {
        console.debug(`[ADO Template Finder] API ${resp.status} for ${repoName}:${normPath} (expected for cross-project repos)`);
        // If 404, try without the leading directory in path (some repos mount differently)
        // Also try alternate extensions (.yml vs .yaml)
        if (resp.status === 404) {
          const altExt = normPath.endsWith('.yaml')
            ? normPath.replace(/\.yaml$/, '.yml')
            : normPath.endsWith('.yml')
            ? normPath.replace(/\.yml$/, '.yaml')
            : null;
          if (altExt) {
            console.log(`[ADO Template Finder] Retrying with alternate extension: ${altExt}`);
            const altUrl = url.replace(encodeURIComponent(normPath), encodeURIComponent(altExt));
            const altResp = await fetch(altUrl, {
              credentials: "same-origin",
              headers: { Accept: "text/plain" },
            });
            if (altResp.ok) {
              const text = await altResp.text();
              console.log(`[ADO Template Finder] ✓ Fetched (alt ext): ${repoForApi}:${altExt} (${text.length} chars)`);
              fileCache.set(cacheKey, text);
              return text;
            }
          }
        }
        fileCache.set(cacheKey, null);
        return null;
      }
      const text = await resp.text();
      console.log(`[ADO Template Finder] ✓ Fetched: ${repoForApi}:${normPath} (${text.length} chars)`);
      fileCache.set(cacheKey, text);
      return text;
    } catch (err) {
      console.debug(`[ADO Template Finder] Fetch error for ${repoName}:${normPath}`, err);
      fileCache.set(cacheKey, null);
      return null;
    }
  }

  // =========================================================================
  // YAML parsing
  // =========================================================================

  function parseResourceRepos(yaml) {
    const repos = {};
    const repoBlockRegex =
      /-\s*repository:\s*(\S+)[\s\S]*?name:\s*(\S+)(?:[\s\S]*?ref:\s*(\S+))?/g;
    let m;
    while ((m = repoBlockRegex.exec(yaml)) !== null) {
      repos[m[1]] = { name: m[2], ref: m[3] || null };
    }
    return repos;
  }

  function parseVariables(yaml) {
    const vars = {};
    // Track line numbers: varName → lineNumber (1-based)
    const varLines = {};

    // Format 1: - name: X \n    value: Y (standard ADO variable template)
    // Use line-by-line to track line numbers
    const lines = yaml.split('\n');
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      const nameMatch = line.match(/^-\s*name:\s*['"]?([^'"\n]+?)['"]?\s*$/);
      if (nameMatch && i + 1 < lines.length) {
        const valueLine = lines[i + 1].trim();
        const valueMatch = valueLine.match(/^value:\s*['"]?([^'"\n]+?)['"]?\s*$/);
        if (valueMatch) {
          const key = nameMatch[1].trim();
          vars[key] = valueMatch[1].trim();
          varLines[key] = i + 1; // 1-based, on the name line
        }
      }
    }

    // Format 2: Simple key: value under a variables: block
    let inVariablesBlock = false;
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      const trimmed = line.trim();

      if (/^variables\s*:/.test(trimmed)) {
        inVariablesBlock = true;
        continue;
      }

      if (inVariablesBlock && /^\S/.test(line) && !/^\s*#/.test(line) && !line.startsWith(' ')) {
        if (/^\w[\w.]*\s*:/.test(trimmed) && !trimmed.startsWith('-')) {
          inVariablesBlock = false;
        }
      }

      if (!inVariablesBlock) continue;

      const simpleKV = trimmed.match(/^([A-Za-z_]\w*)\s*:\s*['"]?(.+?)['"]?\s*$/);
      if (simpleKV && !trimmed.startsWith('-')) {
        const key = simpleKV[1];
        const val = simpleKV[2];
        if (!vars[key]) {
          vars[key] = val;
          varLines[key] = i + 1; // 1-based
        }
      }
    }

    return { vars, varLines };
  }

  /**
   * Find ALL template includes in the YAML and resolve variable expressions
   * in their paths using localVars.
   *
   * Uses line-by-line processing to avoid the regex newline-consumption bug
   * where adjacent `- template:` lines would cause one to be skipped.
   *
   * Returns array of { path, repoAlias, raw, resolved }
   */
  function findVariableTemplateIncludes(yaml, localVars) {
    const results = [];
    const seen = new Set();
    const lines = yaml.split('\n');

    for (const line of lines) {
      const trimmed = line.trim();

      // Match: - template: <value>  (with or without leading dash)
      const m = trimmed.match(/^-?\s*template:\s*['"]?(.+?)['"]?\s*$/);
      if (!m) continue;

      let raw = m[1].trim();

      // Resolve ${{ variables.X }} in the path using local variables (case-insensitive)
      let resolvedRaw = raw.replace(
        /\$\{\{\s*variables\.(\w+)\s*\}\}/g,
        (match, varName) => {
          const found = lookupVar(localVars, varName);
          return found ? found.value : match;
        }
      );

      // Skip if still has unresolved variable expressions — can't fetch
      if (/\$\{\{/.test(resolvedRaw)) continue;

      // Must end in .yml or .yaml after resolution
      const pathPart = resolvedRaw.split("@")[0];
      if (!/\.ya?ml$/i.test(pathPart)) continue;

      // Split on @RepoAlias
      const atIdx = resolvedRaw.lastIndexOf("@");
      let path, repoAlias;
      if (atIdx > 0) {
        path = resolvedRaw.substring(0, atIdx);
        repoAlias = resolvedRaw.substring(atIdx + 1);
      } else {
        path = resolvedRaw;
        repoAlias = null;
      }

      // Deduplicate
      const key = `${path}@${repoAlias || ""}`;
      if (seen.has(key)) continue;
      seen.add(key);

      results.push({ path, repoAlias, raw, resolved: resolvedRaw });
    }

    console.log(`[ADO Template Finder] Found ${results.length} template includes to fetch:`,
      results.map(r => r.resolved));
    return results;
  }

  /**
   * Case-insensitive variable lookup. ADO variables are case-insensitive.
   * Returns { key, value } where key is the actual key in the map, or null.
   */
  function lookupVar(vars, name) {
    // Try exact match first (fast path)
    if (vars[name] !== undefined) return { key: name, value: vars[name] };
    // Case-insensitive fallback
    const lower = name.toLowerCase();
    for (const k of Object.keys(vars)) {
      if (k.toLowerCase() === lower) return { key: k, value: vars[k] };
    }
    return null;
  }

  function resolveVariableExpr(expr, vars) {
    const resolved = expr.replace(
      /\$\{\{\s*variables\.(\w+)\s*\}\}|\$\{\{\s*variables\['([^']+)'\]\s*\}\}/g,
      (match, name1, name2) => {
        const name = name1 || name2;
        const found = lookupVar(vars, name);
        return found ? found.value : match;
      }
    );
    return /\$\{\{.*?\}\}/.test(resolved) ? null : resolved;
  }

  function parseTemplateRefs(yaml) {
    const refs = [];
    const seen = new Set();
    const lines = yaml.split('\n');

    for (let i = 0; i < lines.length; i++) {
      const trimmed = lines[i].trim();
      const m = trimmed.match(/^-?\s*template:\s*(.+)/);
      if (!m) continue;

      let raw = m[1].trim().replace(/^['"]|['"]$/g, "").replace(/\s*#.*$/, "");
      if (seen.has(raw)) continue;
      seen.add(raw);

      const hasVars = /\$\{\{.*?\}\}/.test(raw);
      const isPureVar = /^\$\{\{.*\}\}$/.test(raw);

      const cleanedForAt = raw.replace(/\$\{\{.*?\}\}/g, "___EXPR___");
      const atIdx = cleanedForAt.lastIndexOf("@");
      let path, repoAlias;
      if (atIdx > 0) {
        path = raw.substring(0, atIdx);
        repoAlias = raw.substring(atIdx + 1);
      } else {
        path = raw;
        repoAlias = null;
      }

      refs.push({ raw, path, repoAlias, resolvable: !hasVars, isPureVar, type: "direct", lineNumber: i + 1 });
    }
    return refs;
  }

  function findIndirectTemplateVars(vars) {
    const results = [];
    for (const [name, value] of Object.entries(vars)) {
      const cleanVal = value.replace(/^['"]|['"]$/g, "");
      if (!/\.ya?ml$/i.test(cleanVal.split("@")[0])) continue;
      const atIdx = cleanVal.lastIndexOf("@");
      results.push({
        varName: name,
        raw: cleanVal,
        path: atIdx > 0 ? cleanVal.substring(0, atIdx) : cleanVal,
        repoAlias: atIdx > 0 ? cleanVal.substring(atIdx + 1) : null,
        type: "indirect",
      });
    }
    return results;
  }

  // =========================================================================
  // Main analysis (async — fetches external files)
  // =========================================================================

  async function analyzeYaml(yaml, ctx) {
    if (!ctx.org || !ctx.project) return [];

    const repoMap = parseResourceRepos(yaml);
    const { vars: localVars, varLines: localVarLines } = parseVariables(yaml);

    console.log(`[ADO Template Finder] Local variables found: ${Object.keys(localVars).length}`,
      Object.keys(localVars));
    console.log(`[ADO Template Finder] Resource repos:`, repoMap);

    // --- Fetch ALL referenced template files to extract variables from them ---
    // Pass localVars so ${{variables.buildType}} etc. are resolved in paths
    const varTemplateIncludes = findVariableTemplateIncludes(yaml, localVars);
    const externalVars = {};
    // Track which file each variable was defined in: varName → { path, repoName, ref, line }
    const varSources = {};

    // Fetch in parallel
    const fetchPromises = varTemplateIncludes.map(async (vti) => {
      const repo = vti.repoAlias ? repoMap[vti.repoAlias] : null;
      const repoName = repo ? repo.name : ctx.repo;
      const ref = repo ? repo.ref : ctx.branch;

      if (!repoName) {
        console.debug(`[ADO Template Finder] No repo found for alias "${vti.repoAlias}" (may be injected by extends template)`);
        return;
      }

      console.log(`[ADO Template Finder] Fetching variable template: ${repoName}:${vti.path} (ref: ${ref})`);
      const content = await fetchFileContent(ctx, repoName, vti.path, ref);
      if (content) {
        const { vars, varLines } = parseVariables(content);
        console.log(`[ADO Template Finder] Variables found in ${vti.path}: ${Object.keys(vars).length}`,
          Object.keys(vars).slice(0, 20));
        // Record source for each variable
        for (const varName of Object.keys(vars)) {
          varSources[varName] = {
            path: vti.path,
            repoName: repoName,
            ref: ref,
            repoAlias: vti.repoAlias,
            line: varLines[varName] || null,
          };
        }
        Object.assign(externalVars, vars);
      } else {
        console.debug(`[ADO Template Finder] Could not fetch ${repoName}:${vti.path} (may be in a different project)`);
      }
    });

    await Promise.all(fetchPromises);

    console.log(`[ADO Template Finder] Total external variables: ${Object.keys(externalVars).length}`);

    // Merge: local vars take precedence over external
    const allVars = { ...externalVars, ...localVars };
    console.log(`[ADO Template Finder] All variables available: ${Object.keys(allVars).length}`,
      Object.keys(allVars));

    const templateRefs = parseTemplateRefs(yaml);

    // Collect variable names actually referenced in template: lines (lowercase for case-insensitive matching)
    const usedVarNames = new Set();
    for (const tpl of templateRefs) {
      const m = tpl.raw.match(/\$\{\{\s*variables\.(\w+)\s*\}\}/g);
      if (m) {
        for (const expr of m) {
          const nameMatch = expr.match(/variables\.(\w+)/);
          if (nameMatch) usedVarNames.add(nameMatch[1].toLowerCase());
        }
      }
    }

    const indirectVars = findIndirectTemplateVars(allVars);
    const results = [];

    // --- Process direct template references ---
    for (const tpl of templateRefs) {
      const entry = {
        raw: tpl.raw,
        resolvable: tpl.resolvable,
        url: null,
        label: tpl.raw,
        repoName: null,
        repoAlias: tpl.repoAlias,
        path: tpl.path,
        type: tpl.type,
        resolvedVia: null,
        lineNumber: tpl.lineNumber || null,
      };

      if (!tpl.resolvable && tpl.isPureVar) {
        const varNameMatch = tpl.raw.match(/\$\{\{\s*variables\.(\w+)\s*\}\}/);
        if (varNameMatch) {
          const found = lookupVar(allVars, varNameMatch[1]);
          if (found) {
            const resolvedValue = found.value;
            const actualKey = found.key;
            const cleanVal = resolvedValue.replace(/^['"]|['"]$/g, "");
            const atIdx = cleanVal.lastIndexOf("@");
            entry.resolvable = true;
            entry.path = atIdx > 0 ? cleanVal.substring(0, atIdx) : cleanVal;
            entry.repoAlias = atIdx > 0 ? cleanVal.substring(atIdx + 1) : null;
            entry.resolvedVia = actualKey;

            // Add source info — where was this variable defined?
            const src = varSources[actualKey];
            if (src) {
              entry.definedInUrl = buildAdoFileUrl(ctx, src.repoName, src.path, src.ref, src.line);
              entry.definedInPath = src.path;
              entry.definedInRepo = src.repoName;
              entry.definedInLine = src.line;
            }
          }
        }
      } else if (!tpl.resolvable) {
        const resolved = resolveVariableExpr(tpl.raw, allVars);
        if (resolved) {
          const cleanVal = resolved.replace(/^['"]|['"]$/g, "");
          const atIdx = cleanVal.lastIndexOf("@");
          entry.resolvable = true;
          entry.path = atIdx > 0 ? cleanVal.substring(0, atIdx) : cleanVal;
          entry.repoAlias = atIdx > 0 ? cleanVal.substring(atIdx + 1) : null;
          entry.resolvedVia = "(variable substitution)";
        }
      }

      if (!entry.resolvable) {
        // Even though the final destination can't be resolved, try to find
        // where the variable is defined so we can still link to the source
        const varNameMatch = tpl.raw.match(/\$\{\{\s*variables\.(\w+)\s*\}\}/);
        if (varNameMatch) {
          const found = lookupVar(allVars, varNameMatch[1]);
          if (found) {
            entry.label = `${found.value}`;
            entry.resolvedVia = found.key;
            const src = varSources[found.key];
            if (src) {
              entry.definedInUrl = buildAdoFileUrl(ctx, src.repoName, src.path, src.ref, src.line);
              entry.definedInPath = src.path;
              entry.definedInRepo = src.repoName;
              entry.definedInLine = src.line;
            }
          } else {
            entry.label = tpl.raw + "  (unresolved variable)";
          }
        } else {
          entry.label = tpl.raw + "  (unresolved variable)";
        }
        results.push(entry);
        continue;
      }

      const effectivePath = entry.path || tpl.path;
      const effectiveAlias = entry.repoAlias ?? tpl.repoAlias;

      if (effectiveAlias) {
        const repo = repoMap[effectiveAlias];
        if (repo) {
          entry.repoName = repo.name;
          entry.url = buildAdoFileUrl(ctx, repo.name, effectivePath, repo.ref);
          entry.label = effectivePath;
        } else {
          entry.label = `${effectivePath} @${effectiveAlias}`;
          entry.url = null; // Can't link — repo not in resources
          // But still show where the variable is defined
          if (entry.resolvedVia && !entry.definedInUrl) {
            const src = varSources[entry.resolvedVia];
            if (src) {
              entry.definedInUrl = buildAdoFileUrl(ctx, src.repoName, src.path, src.ref, src.line);
              entry.definedInPath = src.path;
              entry.definedInRepo = src.repoName;
              entry.definedInLine = src.line;
            }
          }
        }
      } else {
        entry.repoName = ctx.repo;
        entry.url = buildAdoFileUrl(ctx, ctx.repo, effectivePath);
        entry.label = effectivePath;
      }

      results.push(entry);
    }

    // --- Indirect template variables (from external files too) ---
    // Only include variables that are either:
    //   (a) actually used in a template: line in the main YAML, or
    //   (b) have a resolvable repo alias, and are shown as "available"
    //   (c) we're viewing a standalone variable file (no resources block) — show all
    const hasResourceRepos = Object.keys(repoMap).length > 0;
    const directPaths = new Set(
      results.filter((r) => r.url).map((r) => (r.path || "") + "@" + (r.repoAlias || ""))
    );

    for (const iv of indirectVars) {
      const key = iv.path + "@" + (iv.repoAlias || "");
      if (directPaths.has(key)) continue;

      // Skip if repo alias can't be resolved (e.g. VstsTemplates not in resources)
      // unless the variable is actually used in a template: line
      // In standalone files (no resources block), still show them but as unresolved
      if (iv.repoAlias && hasResourceRepos) {
        const repo = repoMap[iv.repoAlias];
        if (!repo && !usedVarNames.has(iv.varName.toLowerCase())) {
          // Not resolvable and not used — skip (reduces noise)
          continue;
        }
      }

      const entry = {
        raw: iv.raw,
        resolvable: true,
        url: null,
        label: iv.path,
        repoName: null,
        repoAlias: iv.repoAlias,
        path: iv.path,
        type: usedVarNames.has(iv.varName.toLowerCase()) ? "indirect-used" : "indirect",
        varName: iv.varName,
      };

      if (iv.repoAlias) {
        const repo = repoMap[iv.repoAlias];
        if (repo) {
          entry.repoName = repo.name;
          entry.url = buildAdoFileUrl(ctx, repo.name, iv.path, repo.ref);
        } else {
          // Repo alias can't be resolved — mark as unresolved
          // (even in standalone files — an explicit @Alias means a specific repo)
          entry.resolvable = false;
          entry.label = `${iv.path} @${iv.repoAlias}`;
        }
      } else {
        entry.repoName = ctx.repo;
        entry.url = buildAdoFileUrl(ctx, ctx.repo, iv.path);
      }

      results.push(entry);
    }

    return results;
  }

  // =========================================================================
  // UI — Floating sidebar panel
  // =========================================================================

  function removePanel() {
    document.getElementById(PANEL_ID)?.remove();
    document.getElementById(TOGGLE_ID)?.remove();
  }

  function createPanel(results, loading) {
    removePanel();
    if (results.length === 0 && !loading) return;

    const toggle = document.createElement("button");
    toggle.id = TOGGLE_ID;
    toggle.title = "ADO Template Finder";
    if (loading) {
      toggle.innerHTML = '<span class="ado-tf-spinner"></span> <strong>TF</strong>';
    } else {
      toggle.innerHTML = `<strong>TF</strong>: ${results.length}`;
    }
    document.body.appendChild(toggle);

    const panel = document.createElement("div");
    panel.id = PANEL_ID;
    panel.classList.add("ado-tf-collapsed");

    const header = document.createElement("div");
    header.className = "ado-tf-header";
    const statusText = loading ? "Loading..." : `(${results.length})`;
    header.innerHTML = `
      <span class="ado-tf-title">Template References ${statusText}</span>
      <button class="ado-tf-close" title="Close">&times;</button>
    `;
    panel.appendChild(header);
    header.querySelector(".ado-tf-close").addEventListener("click", () => {
      panel.classList.add("ado-tf-collapsed");
    });

    const body = document.createElement("div");
    body.className = "ado-tf-body";

    if (loading) {
      const loadingEl = document.createElement("div");
      loadingEl.className = "ado-tf-loading";
      loadingEl.innerHTML = '<span class="ado-tf-spinner"></span> Fetching files &amp; resolving variables...';
      body.appendChild(loadingEl);
    } else {
      const directResolved = results.filter(
        (r) => r.type === "direct" && r.url && !r.resolvedVia
      );
      const varResolved = results.filter(
        (r) => r.type === "direct" && r.url && r.resolvedVia
      );
      const indirect = results.filter((r) => (r.type === "indirect" || r.type === "indirect-used") && r.url);
      const unresolved = results.filter((r) => !r.url);

      // Count only actionable (linked) items for the toggle
      const actionableCount = directResolved.length + varResolved.length + indirect.length;

      // Update toggle text to show actionable count
      const toggleEl = document.getElementById(TOGGLE_ID);
      if (toggleEl) {
        toggleEl.innerHTML = `<strong>TF</strong>: ${actionableCount}` + (unresolved.length > 0 ? ` <span class="ado-tf-toggle-extra">(+${unresolved.length})</span>` : '');
      }

      renderSection(body, "Templates", directResolved, false);
      renderSection(body, "Resolved via Variables", varResolved, false);
      renderSection(body, "Available Template Variables", indirect, true);
      renderSection(body, "Unresolved", unresolved, true);
    }

    panel.appendChild(body);
    document.body.appendChild(panel);

    toggle.addEventListener("click", () => {
      panel.classList.toggle("ado-tf-collapsed");
    });
  }

  function renderSection(body, title, items, startCollapsed) {
    if (items.length === 0) return;

    const section = document.createElement("div");
    section.className = "ado-tf-section";

    const sectionHeader = document.createElement("div");
    sectionHeader.className = "ado-tf-section-header";
    sectionHeader.setAttribute("role", "button");
    sectionHeader.innerHTML = `<span class="ado-tf-section-toggle">${startCollapsed ? '\u25b6' : '\u25bc'}</span> ${title} (${items.length})`;
    section.appendChild(sectionHeader);

    const sectionBody = document.createElement("div");
    sectionBody.className = "ado-tf-section-body";
    if (startCollapsed) sectionBody.classList.add("ado-tf-section-collapsed");

    // Toggle collapse on click
    sectionHeader.addEventListener("click", () => {
      const isCollapsed = sectionBody.classList.toggle("ado-tf-section-collapsed");
      sectionHeader.querySelector(".ado-tf-section-toggle").textContent = isCollapsed ? '▶' : '▼';
    });

    const groups = {};
    for (const r of items) {
      const key = r.repoAlias || r.repoName || "(this repo)";
      if (!groups[key]) groups[key] = [];
      groups[key].push(r);
    }

    for (const [, groupItems] of Object.entries(groups)) {
      const groupHeader = document.createElement("div");
      groupHeader.className = "ado-tf-group-header";
      groupHeader.textContent = groupItems[0].repoName || groupItems[0].repoAlias || "(this repo)";
      sectionBody.appendChild(groupHeader);

      for (const item of groupItems) {
        const row = document.createElement("div");
        row.className = "ado-tf-row";

        if (item.url) {
          const a = document.createElement("a");
          a.href = item.url;
          a.target = "_blank";
          a.rel = "noopener";
          a.className = "ado-tf-link";
          a.textContent = (item.path || item.label || item.raw).replace(/\\/g, "/");
          a.title = item.url;
          row.appendChild(a);

          // Show line number badge if available
          if (item.lineNumber) {
            const lineBadge = document.createElement("span");
            lineBadge.className = "ado-tf-line-badge";
            lineBadge.textContent = `L${item.lineNumber}`;
            lineBadge.title = `Line ${item.lineNumber} in this file`;
            row.appendChild(lineBadge);
          }

          if (item.varName) {
            const badge = document.createElement("span");
            badge.className = "ado-tf-badge";
            badge.textContent = `$${item.varName}`;
            badge.title = `Variable: ${item.varName}`;
            row.appendChild(badge);
          } else if (item.resolvedVia) {
            const badge = document.createElement("span");
            badge.className = "ado-tf-badge ado-tf-badge-resolved";
            badge.textContent = `via $${item.resolvedVia}`;
            badge.title = `Resolved from: ${item.resolvedVia}`;
            row.appendChild(badge);
          }

          // Show definition source link (where the variable is defined)
          if (item.definedInUrl) {
            const defRow = document.createElement("div");
            defRow.className = "ado-tf-def-row";
            const defLabel = document.createElement("span");
            defLabel.className = "ado-tf-def-label";
            defLabel.textContent = "defined in ";
            defRow.appendChild(defLabel);
            const defLink = document.createElement("a");
            defLink.href = item.definedInUrl;
            defLink.target = "_blank";
            defLink.rel = "noopener";
            defLink.className = "ado-tf-def-link";
            defLink.textContent = item.definedInPath + (item.definedInLine ? `:${item.definedInLine}` : '');
            defLink.title = `Defined in ${item.definedInRepo}: ${item.definedInPath}${item.definedInLine ? ' (line ' + item.definedInLine + ')' : ''}`;
            defRow.appendChild(defLink);
            row.appendChild(defRow);
          }
        } else {
          const span = document.createElement("span");
          span.className = "ado-tf-unresolved";
          span.textContent = item.label || item.raw;
          row.appendChild(span);

          if (item.resolvedVia) {
            const badge = document.createElement("span");
            badge.className = "ado-tf-badge ado-tf-badge-resolved";
            badge.textContent = `via $${item.resolvedVia}`;
            row.appendChild(badge);
          }

          if (item.lineNumber) {
            const lineBadge = document.createElement("span");
            lineBadge.className = "ado-tf-line-badge";
            lineBadge.textContent = `L${item.lineNumber}`;
            lineBadge.title = `Line ${item.lineNumber} in this file`;
            row.appendChild(lineBadge);
          }

          // Show "defined in" link even for unresolved entries
          if (item.definedInUrl) {
            const defRow = document.createElement("div");
            defRow.className = "ado-tf-def-row";
            const defLabel = document.createElement("span");
            defLabel.className = "ado-tf-def-label";
            defLabel.textContent = "defined in ";
            defRow.appendChild(defLabel);
            const defLink = document.createElement("a");
            defLink.href = item.definedInUrl;
            defLink.target = "_blank";
            defLink.rel = "noopener";
            defLink.className = "ado-tf-def-link";
            defLink.textContent = item.definedInPath + (item.definedInLine ? `:${item.definedInLine}` : '');
            defLink.title = `Defined in ${item.definedInRepo}: ${item.definedInPath}`;
            defRow.appendChild(defLink);
            row.appendChild(defRow);
          }
        }

        sectionBody.appendChild(row);
      }
    }

    section.appendChild(sectionBody);
    body.appendChild(section);
  }

  // =========================================================================
  // Lifecycle
  // =========================================================================

  async function run() {
    if (isRunning) return;

    const url = window.location.href;
    const isFileView = /\/_git\//.test(url);
    const isPipelineView = /\/_build\b/.test(url);

    if (!isFileView && !isPipelineView) {
      removePanel();
      return;
    }

    const ctx = parseAdoUrl();
    if (!ctx.repo || !ctx.filePath) {
      // No specific file selected — nothing to do
      return;
    }

    // Check if it looks like a YAML file
    if (!/\.ya?ml$/i.test(ctx.filePath)) {
      removePanel();
      return;
    }

    // Avoid re-analyzing the same URL
    const analysisKey = `${ctx.repo}::${ctx.filePath}::${ctx.branch}`;
    if (analysisKey === lastAnalyzedUrl) return;

    isRunning = true;

    try {
      // Show loading state
      createPanel([], true);

      // Fetch the main file via API
      const yaml = await fetchFileContent(ctx, ctx.repo, ctx.filePath, ctx.branch);
      if (!yaml) {
        // API failed — fall back to DOM scraping
        const domYaml = extractYamlTextFromDom();
        if (domYaml && (/template\s*:/i.test(domYaml) || /\w+:\s*\S+\.ya?ml/i.test(domYaml))) {
          const results = await analyzeYaml(domYaml, ctx);
          createPanel(results, false);
          lastAnalyzedUrl = analysisKey;
        } else {
          removePanel();
        }
        return;
      }

      if (!/template\s*:/i.test(yaml) && !/-\s*name:.*\n\s*value:.*\.ya?ml/i.test(yaml) && !/\w+:\s*\S+\.ya?ml/i.test(yaml)) {
        removePanel();
        lastAnalyzedUrl = analysisKey;
        return;
      }

      const results = await analyzeYaml(yaml, ctx);
      createPanel(results, false);
      lastAnalyzedUrl = analysisKey;
    } catch (err) {
      console.error("[ADO Template Finder] Error:", err);
      removePanel();
    } finally {
      isRunning = false;
    }
  }

  /**
   * Fallback: extract YAML from DOM (for pipeline run YAML tab, etc.)
   */
  function extractYamlTextFromDom() {
    const candidates = [];

    // Monaco
    const monacoContainers = document.querySelectorAll(".view-lines");
    for (const container of monacoContainers) {
      const lines = container.querySelectorAll(".view-line");
      if (lines.length > 0) {
        const text = Array.from(lines).map((el) => el.textContent).join("\n");
        if (text.length > 30) candidates.push(text);
      }
    }

    // Pre/code blocks
    for (const sel of ["pre.content", ".file-content pre", ".raw-file-content", "code.hljs"]) {
      for (const el of document.querySelectorAll(sel)) {
        if (el.textContent.length > 30) candidates.push(el.textContent);
      }
    }

    // File content containers
    for (const sel of [".file-editor-view", ".file-content", ".repos-file-content"]) {
      const el = document.querySelector(sel);
      if (el && el.textContent.length > 30) candidates.push(el.textContent);
    }

    if (candidates.length === 0) return null;
    candidates.sort((a, b) => b.length - a.length);
    return candidates[0];
  }

  // Debounced runner
  let timer = null;
  function scheduleRun() {
    clearTimeout(timer);
    timer = setTimeout(run, DEBOUNCE_MS);
  }

  scheduleRun();

  // Watch for SPA navigation
  let lastUrl = window.location.href;
  function handleUrlChange() {
    if (window.location.href !== lastUrl) {
      lastUrl = window.location.href;
      lastAnalyzedUrl = null;
      fileCache.clear();
      scheduleRun();
    }
  }

  const observer = new MutationObserver(handleUrlChange);
  observer.observe(document.body, { childList: true, subtree: true });

  // Also poll in case URL doesn't change but content does (e.g., branch switch)
  setInterval(handleUrlChange, POLL_INTERVAL_MS);
})();
