/* global Asc, Api */
(function () {
  function reply(requestId, result, error) {
    var payload = {
      type: "agent-doc",
      requestId: requestId,
      result: result == null ? null : result,
      error: error || null
    }
    var raw = JSON.stringify(payload)
    // Walk every ancestor. COOP/cross-origin can make window.top throw, and
    // parent-only delivery stops at frameEditor (which does not forward us).
    var seen = []
    var win = window
    while (win) {
      try {
        if (seen.indexOf(win) >= 0) break
        seen.push(win)
        win.postMessage(raw, "*")
      } catch (err) {}
      var next = null
      try {
        next = win.parent
      } catch (err) {
        break
      }
      if (!next || next === win) break
      win = next
    }
  }

  function normalizeNewlines(text) {
    return String(text == null ? "" : text)
      .replace(/\\r\\n/g, "\n")
      .replace(/\\n/g, "\n")
      .replace(/\\r/g, "\n")
      .replace(/\\t/g, "\t")
      .replace(/\r\n/g, "\n")
      .replace(/\r/g, "\n")
  }

  var INLINE_TAG_FLAGS = {
    b: "bold",
    strong: "bold",
    i: "italic",
    em: "italic",
    u: "underline",
    s: "strike",
    strike: "strike",
    del: "strike"
  }

  function decodeHtmlEntities(text) {
    return String(text)
      .replace(/&nbsp;/gi, " ")
      .replace(/&amp;/gi, "&")
      .replace(/&lt;/gi, "<")
      .replace(/&gt;/gi, ">")
      .replace(/&quot;/gi, "\"")
      .replace(/&#39;/g, "'")
      .replace(/&apos;/gi, "'")
  }

  function parseInlineMarkup(text) {
    var source = String(text == null ? "" : text)
    var tagRe = /<\/?(b|strong|i|em|u|s|strike|del|br|p)\b[^>]*\/?>/gi
    var flags = { bold: 0, italic: 0, underline: 0, strike: 0 }
    var paragraphs = [[]]
    var current = paragraphs[0]
    var lastIndex = 0
    var hasMarkup = false
    var match

    function formatOf() {
      return {
        bold: flags.bold > 0,
        italic: flags.italic > 0,
        underline: flags.underline > 0,
        strike: flags.strike > 0
      }
    }

    function appendRun(chunk) {
      if (!chunk) return
      var decoded = decodeHtmlEntities(chunk)
      if (!decoded) return
      var fmt = formatOf()
      var last = current[current.length - 1]
      if (
        last &&
        last.text != null &&
        last.break !== "line" &&
        last.bold === fmt.bold &&
        last.italic === fmt.italic &&
        last.underline === fmt.underline &&
        last.strike === fmt.strike
      ) {
        last.text += decoded
        return
      }
      current.push({
        text: decoded,
        bold: fmt.bold,
        italic: fmt.italic,
        underline: fmt.underline,
        strike: fmt.strike
      })
    }

    function appendText(raw) {
      if (!raw) return
      var parts = raw.split("\n")
      var i
      for (i = 0; i < parts.length; i++) {
        if (i > 0) {
          paragraphs.push([])
          current = paragraphs[paragraphs.length - 1]
        }
        appendRun(parts[i])
      }
    }

    while ((match = tagRe.exec(source))) {
      hasMarkup = true
      appendText(source.slice(lastIndex, match.index))
      lastIndex = tagRe.lastIndex
      var token = match[0]
      var name = match[1].toLowerCase()
      var closing = token.charAt(1) === "/"
      if (name === "br") {
        current.push({ break: "line" })
        continue
      }
      if (name === "p") {
        if (current.length) {
          paragraphs.push([])
          current = paragraphs[paragraphs.length - 1]
        }
        continue
      }
      var key = INLINE_TAG_FLAGS[name]
      if (!key) continue
      if (closing) {
        if (flags[key] > 0) flags[key]--
      } else {
        flags[key]++
      }
    }
    appendText(source.slice(lastIndex))
    while (paragraphs.length > 1 && paragraphs[paragraphs.length - 1].length === 0) paragraphs.pop()
    return { hasMarkup: hasMarkup, paragraphs: paragraphs }
  }

  function paragraphsFromPayload(payload, text) {
    if (payload && Array.isArray(payload.paragraphs) && payload.paragraphs.length) {
      return { hasMarkup: true, paragraphs: payload.paragraphs }
    }
    return parseInlineMarkup(text)
  }

  function isTableElement(el) {
    if (!el) return false
    if (typeof el.GetRow === "function") return true
    var klass = el.GetClassType ? String(el.GetClassType()) : ""
    return klass.toLowerCase().indexOf("table") >= 0
  }

  function tablePreviewText(table) {
    if (!table || typeof table.GetRow !== "function") return ""
    var rows = []
    var r
    var max = typeof table.GetRowsCount === "function" ? Math.min(table.GetRowsCount(), 6) : 6
    for (r = 0; r < max; r++) {
      var row = table.GetRow(r)
      if (!row) break
      var cells = []
      var c
      var cols = typeof row.GetCellsCount === "function" ? Math.min(row.GetCellsCount(), 8) : 8
      for (c = 0; c < cols; c++) {
        var cell = row.GetCell(c)
        if (!cell) break
        var content = cell.GetContent && cell.GetContent()
        var para = content && content.GetElement ? content.GetElement(0) : null
        cells.push(para && para.GetText ? String(para.GetText()) : "")
      }
      rows.push(cells.join(" | "))
    }
    return rows.join("\n")
  }

  function elementText(el) {
    if (!el) return ""
    if (isTableElement(el)) return tablePreviewText(el)
    if (typeof el.GetText === "function") return String(el.GetText() || "")
    return ""
  }

  function clipText(text, limit) {
    var value = String(text == null ? "" : text)
    var max = limit || 500
    if (value.length <= max) return value
    return value.slice(0, max) + "…"
  }

  /**
   * Stable Agent "lines": one per top-level document element, same order as
   * doc_get_outline.elements. line is 1-based; element_index is 0-based (pass
   * to doc_goto / doc_select). Tables are one line (preview text only — edit
   * cells with table_index + row + col tools).
   */
  function listLogicalLines(doc) {
    var lines = []
    var n = doc.GetElementsCount ? doc.GetElementsCount() : 0
    var tableOrdinal = -1
    var i
    for (i = 0; i < n; i++) {
      var el = doc.GetElement(i)
      var isTable = isTableElement(el)
      if (isTable) tableOrdinal += 1
      var text = ""
      if (el) {
        text = isTable
          ? tablePreviewText(el)
          : typeof el.GetText === "function"
            ? String(el.GetText() || "")
            : ""
      }
      var row = {
        line: i + 1,
        element_index: i,
        type: isTable ? "table" : "paragraph",
        text: text
      }
      if (isTable) row.table_index = tableOrdinal
      lines.push(row)
    }
    return lines
  }

  function rewriteParagraphText(para, newText) {
    if (!para) return false
    var text = String(newText == null ? "" : newText)
    if (typeof para.RemoveAllElements === "function") {
      try {
        para.RemoveAllElements()
        if (typeof para.AddText === "function") {
          para.AddText(text)
          return true
        }
      } catch (err) {}
    }
    if (typeof para.GetElementsCount === "function" && typeof para.RemoveElement === "function") {
      try {
        var guard = 0
        while (para.GetElementsCount() > 0 && guard < 500) {
          para.RemoveElement(0)
          guard += 1
        }
        if (typeof para.AddText === "function") {
          para.AddText(text)
          return true
        }
      } catch (err2) {}
    }
    return false
  }

  function countLiteralMatches(haystack, needle) {
    if (!needle) return 0
    var count = 0
    var pos = 0
    while (true) {
      var idx = haystack.indexOf(needle, pos)
      if (idx < 0) break
      count += 1
      pos = idx + needle.length
    }
    return count
  }

  function elementStyleName(el) {
    if (!el || typeof el.GetStyle !== "function") return ""
    var style = el.GetStyle()
    return style && typeof style.GetName === "function" ? String(style.GetName()) : ""
  }

  function normalizePlacement(raw) {
    var value = String(raw == null ? "" : raw)
      .trim()
      .toLowerCase()
      .replace(/-/g, "_")
      .replace(/\s+/g, "_")
    if (!value || value === "cursor") return "cursor"
    if (value === "before" || value === "before_element") return "before"
    if (value === "after" || value === "after_element") return "after"
    if (
      value === "before_table" ||
      value === "before_next_table" ||
      value === "before_following_table" ||
      value === "antes_da_tabela" ||
      value === "antes_tabela" ||
      value.indexOf("before_table") === 0
    ) {
      return "before_table"
    }
    if (
      value === "after_table" ||
      value === "after_next_table" ||
      value === "depois_da_tabela" ||
      value === "apos_a_tabela" ||
      value === "após_a_tabela" ||
      value.indexOf("after_table") === 0
    ) {
      return "after_table"
    }
    return value
  }

  /**
   * Resolve a document element index for navigation/insert.
   * Never defaults to 0 when a heading/table target was requested but missing.
   */
  function resolveDocTarget(doc, opts) {
    var n = doc.GetElementsCount ? doc.GetElementsCount() : 0
    var placement = normalizePlacement(opts && opts.placement)
    var heading =
      opts && opts.heading != null && String(opts.heading).trim()
        ? String(opts.heading).trim()
        : ""
    var tableIndex =
      opts && opts.table_index != null && opts.table_index !== ""
        ? Number(opts.table_index)
        : null
    var index =
      opts && opts.index != null && opts.index !== "" ? Number(opts.index) : null
    var targetKind =
      opts && opts.target != null ? String(opts.target).trim().toLowerCase() : ""

    if (tableIndex != null && !isNaN(tableIndex)) {
      var tableOrdinal = -1
      var ti
      for (ti = 0; ti < n; ti++) {
        if (!isTableElement(doc.GetElement(ti))) continue
        tableOrdinal += 1
        // table_index is 0-based among tables only (see get_outline.tables).
        if (tableOrdinal === tableIndex) {
          index = ti
          break
        }
      }
      if (index == null || isNaN(index)) {
        return { ok: false, error: "Table not found (table_index=" + tableIndex + ")." }
      }
    }

    if (heading && (index == null || isNaN(index))) {
      var needle = heading.toLowerCase()
      var hi
      var best = null
      var bestScore = -1
      for (hi = 0; hi < n; hi++) {
        var el = doc.GetElement(hi)
        var text = elementText(el).toLowerCase().trim()
        if (text.indexOf(needle) < 0) continue
        // Prefer exact / heading-styled matches over early TOC substring hits.
        var score = 0
        if (text === needle) score += 100
        else if (text.indexOf(needle) === 0) score += 40
        var styleName = elementStyleName(el).toLowerCase()
        if (styleName.indexOf("heading") >= 0 || styleName.indexOf("título") >= 0 || styleName.indexOf("titulo") >= 0) {
          score += 50
        }
        // For table placement, prefer a heading that actually has a following table.
        if (placement === "before_table" || placement === "after_table") {
          var look
          // Prefer a match that has a *following* table (skip self when match is a table).
          for (look = hi + 1; look < n && look < hi + 40; look++) {
            if (isTableElement(doc.GetElement(look))) {
              score += 80
              break
            }
          }
        }
        // Slight preference for later body matches over front-matter/TOC.
        score += Math.min(hi, 30) / 30
        if (score > bestScore) {
          bestScore = score
          best = hi
        }
      }
      if (best == null) {
        return { ok: false, error: "Heading/text not found: " + heading }
      }
      index = best
    }

    if (targetKind === "table" && (index == null || isNaN(index))) {
      return { ok: false, error: "Pass table_index or heading to target a table." }
    }

    if (
      (placement === "before_table" || placement === "after_table") &&
      index != null &&
      !isNaN(index)
    ) {
      // table_index already names the target table. Heading matches that live
      // inside a table (RAC-style forms) must advance to the *following* table
      // (e.g. monthly grid after the 2.2.1 section table), not the same one.
      var resolvedViaTableIndex = tableIndex != null && !isNaN(tableIndex)
      // Heading/paragraph path: always search from the next element so a
      // section table that contains "2.2.1" does not become the insert target.
      var start = resolvedViaTableIndex ? index : index + 1
      var tj
      var foundTable = null
      for (tj = start; tj < n; tj++) {
        if (isTableElement(doc.GetElement(tj))) {
          foundTable = tj
          break
        }
      }
      if (foundTable == null) {
        return {
          ok: false,
          error: "No table found after the target (placement=" + placement + ")."
        }
      }
      index = foundTable
    }

    if (index == null || isNaN(index)) {
      if (heading || tableIndex != null || targetKind === "table" || placement.indexOf("table") >= 0) {
        return { ok: false, error: "Target not found in document." }
      }
      return { ok: false, error: "Pass heading, index, or table_index (no silent start-of-document)." }
    }

    if (index < 0 || index >= n) {
      return { ok: false, error: "Element index out of range: " + index }
    }

    var insertAt = index
    if (placement === "after" || placement === "after_table") insertAt = index + 1
    if (placement === "before" || placement === "before_table") insertAt = index

    return {
      ok: true,
      index: index,
      insert_at: insertAt,
      placement: placement,
      heading: heading || null,
      type: isTableElement(doc.GetElement(index)) ? "table" : "paragraph"
    }
  }

  function selectAtElementStart(el) {
    if (!el) return false
    if (typeof el.GetRange === "function") {
      try {
        var range = el.GetRange(0, 0)
        if (range && typeof range.Select === "function") {
          range.Select()
          return true
        }
      } catch (err) {}
    }
    if (typeof el.Select === "function") {
      el.Select()
      return true
    }
    return false
  }

  function parseRowCol(opts) {
    var row =
      opts && opts.row != null && opts.row !== ""
        ? Number(opts.row)
        : opts && opts.row_index != null && opts.row_index !== ""
          ? Number(opts.row_index)
          : null
    var col =
      opts && opts.col != null && opts.col !== ""
        ? Number(opts.col)
        : opts && opts.column != null && opts.column !== ""
          ? Number(opts.column)
          : opts && opts.col_index != null && opts.col_index !== ""
            ? Number(opts.col_index)
            : null
    return {
      row: row != null && !isNaN(row) ? row : null,
      col: col != null && !isNaN(col) ? col : null
    }
  }

  function getCellPlainText(cell) {
    if (!cell) return ""
    if (typeof cell.GetText === "function") {
      try {
        return String(
          cell.GetText({
            TableCellSeparator: "",
            TableRowSeparator: "",
            ParaSeparator: "\n",
            Numbering: false
          }) || ""
        )
      } catch (err) {}
    }
    var content = cell.GetContent && cell.GetContent()
    var para = content && content.GetElement && content.GetElement(0)
    if (para && typeof para.GetText === "function") return String(para.GetText() || "")
    return ""
  }

  function tableDims(table) {
    var rows = typeof table.GetRowsCount === "function" ? table.GetRowsCount() : 0
    var cols = 0
    if (typeof table.GetColumnsCount === "function") {
      try {
        cols = table.GetColumnsCount() || 0
      } catch (err) {
        cols = 0
      }
    }
    if (!cols && rows > 0 && typeof table.GetRow === "function") {
      var first = table.GetRow(0)
      cols = first && typeof first.GetCellsCount === "function" ? first.GetCellsCount() : 0
    }
    return { rows: rows, cols: cols }
  }

  function tableSummary(table, tableIndex, elementIndex) {
    var dims = tableDims(table)
    var preview = tablePreviewText(table)
    return {
      table_index: tableIndex,
      index: elementIndex,
      rows: dims.rows,
      cols: dims.cols,
      text: preview.length > 240 ? preview.slice(0, 240) : preview
    }
  }

  function resolveTableByIndex(doc, tableIndex) {
    var n = doc.GetElementsCount ? doc.GetElementsCount() : 0
    var ordinal = -1
    var i
    for (i = 0; i < n; i++) {
      var el = doc.GetElement(i)
      if (!isTableElement(el)) continue
      ordinal += 1
      if (ordinal === tableIndex) {
        return { ok: true, table: el, index: i, table_index: tableIndex }
      }
    }
    return { ok: false, error: "Table not found (table_index=" + tableIndex + ")." }
  }

  function resolveFlatTable(doc, flatIndex) {
    if (typeof doc.GetAllTables !== "function") {
      return { ok: false, error: "GetAllTables unavailable (need newer OnlyOffice)." }
    }
    var list = doc.GetAllTables()
    if (!list) return { ok: false, error: "No tables in document." }
    var len = typeof list.length === "number" ? list.length : 0
    if (typeof list.GetCount === "function") len = list.GetCount()
    if (flatIndex < 0 || flatIndex >= len) {
      return {
        ok: false,
        error: "flat_table_index out of range: " + flatIndex + " (count=" + len + ")."
      }
    }
    var table = typeof list.Get === "function" ? list.Get(flatIndex) : list[flatIndex]
    if (!table) return { ok: false, error: "Table missing at flat_table_index=" + flatIndex + "." }
    return { ok: true, table: table, flat_table_index: flatIndex }
  }

  function getTableCell(table, rowIndex, colIndex) {
    if (!table) return null
    if (typeof table.GetCell === "function") {
      try {
        return table.GetCell(rowIndex, colIndex)
      } catch (err) {
        return null
      }
    }
    var row = typeof table.GetRow === "function" ? table.GetRow(rowIndex) : null
    if (!row || typeof row.GetCell !== "function") return null
    try {
      return row.GetCell(colIndex)
    } catch (err2) {
      return null
    }
  }

  /**
   * Land the caret in table_index (or flat_table_index) at row/col (0-based).
   * Uses ApiTable.GetCell + ApiTableCell.Select when available.
   */
  function selectTableCell(doc, opts) {
    var rc = parseRowCol(opts)
    var tableIndex =
      opts && opts.table_index != null && opts.table_index !== ""
        ? Number(opts.table_index)
        : null
    var flatIndex =
      opts && opts.flat_table_index != null && opts.flat_table_index !== ""
        ? Number(opts.flat_table_index)
        : null
    if (rc.row == null || rc.col == null) {
      return { ok: false, error: "row and col are required for cell targeting (0-based)." }
    }
    var resolved
    if (flatIndex != null && !isNaN(flatIndex)) {
      resolved = resolveFlatTable(doc, flatIndex)
    } else if (tableIndex != null && !isNaN(tableIndex)) {
      resolved = resolveTableByIndex(doc, tableIndex)
    } else {
      return {
        ok: false,
        error: "Pass table_index (top-level) or flat_table_index (GetAllTables) with row/col."
      }
    }
    if (!resolved.ok) return resolved
    var table = resolved.table
    var dims = tableDims(table)
    if (rc.row < 0 || rc.row >= dims.rows) {
      return {
        ok: false,
        error: "Row out of range: " + rc.row + " (rows=" + dims.rows + ")."
      }
    }
    var rowObj = table.GetRow(rc.row)
    var rowCols =
      rowObj && typeof rowObj.GetCellsCount === "function" ? rowObj.GetCellsCount() : dims.cols
    if (rc.col < 0 || rc.col >= rowCols) {
      return {
        ok: false,
        error:
          "Col out of range: " + rc.col + " (cols=" + rowCols + " in row " + rc.row + ")."
      }
    }
    var cell = getTableCell(table, rc.row, rc.col)
    if (!cell) {
      return {
        ok: false,
        error: "Cell not found at (" + rc.row + "," + rc.col + ")."
      }
    }
    var selected = false
    if (typeof cell.Select === "function") {
      try {
        cell.Select()
        selected = true
      } catch (err) {
        selected = false
      }
    }
    if (!selected) {
      var content = cell.GetContent && cell.GetContent()
      var para = content && content.GetElement ? content.GetElement(0) : null
      selected = selectAtElementStart(para || content)
    }
    if (!selected) {
      return {
        ok: false,
        error: "Could not select cell (" + rc.row + "," + rc.col + ")."
      }
    }
    var verified = null
    if (typeof doc.GetActiveTable === "function") {
      try {
        var active = doc.GetActiveTable()
        verified = {
          in_table: !!active,
          row: typeof cell.GetRowIndex === "function" ? cell.GetRowIndex() : rc.row,
          col: typeof cell.GetIndex === "function" ? cell.GetIndex() : rc.col
        }
        if (verified.in_table && verified.row != null && verified.row !== rc.row) {
          return {
            ok: false,
            error:
              "Cell select verify failed: expected row " +
              rc.row +
              " got " +
              verified.row +
              "."
          }
        }
        if (verified.in_table && verified.col != null && verified.col !== rc.col) {
          return {
            ok: false,
            error:
              "Cell select verify failed: expected col " +
              rc.col +
              " got " +
              verified.col +
              "."
          }
        }
      } catch (err2) {
        verified = { in_table: null }
      }
    }
    return {
      ok: true,
      table_index: resolved.table_index != null ? resolved.table_index : null,
      flat_table_index: resolved.flat_table_index != null ? resolved.flat_table_index : null,
      index: resolved.index != null ? resolved.index : null,
      row: rc.row,
      col: rc.col,
      cell_text: getCellPlainText(cell),
      rows: dims.rows,
      cols: dims.cols,
      verified: verified,
      type: "cell"
    }
  }

  function writeTableCell(doc, opts, text, mode) {
    var landed = selectTableCell(doc, opts)
    if (!landed.ok) return landed
    var table
    if (landed.flat_table_index != null) {
      var flat = resolveFlatTable(doc, landed.flat_table_index)
      if (!flat.ok) return flat
      table = flat.table
    } else {
      var top = resolveTableByIndex(doc, landed.table_index)
      if (!top.ok) return top
      table = top.table
    }
    var cell = getTableCell(table, landed.row, landed.col)
    if (!cell) {
      return { ok: false, error: "Cell disappeared after select." }
    }
    var writeMode = mode === "replace" || mode === "set" ? "replace" : "append"
    var chunk = text == null ? "" : String(text)
    if (writeMode === "replace" && typeof cell.SetText === "function") {
      cell.SetText(chunk)
    } else if (typeof cell.AddText === "function") {
      cell.AddText(chunk)
    } else {
      var content = cell.GetContent && cell.GetContent()
      var para = content && content.GetElement ? content.GetElement(0) : null
      if (para && para.AddText) para.AddText(chunk)
      else return { ok: false, error: "Cell AddText/SetText unavailable." }
    }
    return {
      ok: true,
      table_index: landed.table_index,
      flat_table_index: landed.flat_table_index,
      index: landed.index,
      row: landed.row,
      col: landed.col,
      cell_mode: writeMode,
      cell_text: getCellPlainText(cell),
      inserted: chunk.length,
      type: "cell",
      verified: landed.verified
    }
  }

  function insertParasAt(doc, paras, insertAt, placement) {
    if (!paras || !paras.length) return false
    var pos = insertAt == null || insertAt === "" ? null : Number(insertAt)
    // When a target index was resolved, ALWAYS use AddElement — never fall back to
    // InsertContent at the live cursor (that re-introduced markers at document start).
    if (pos != null && !isNaN(pos) && typeof doc.AddElement === "function") {
      var i
      // AddElement(nPos) inserts before the current element at nPos; add last→first.
      for (i = paras.length - 1; i >= 0; i--) {
        doc.AddElement(pos, paras[i])
      }
      return true
    }
    var usePos =
      placement === "before" ||
      placement === "after" ||
      placement === "before_table" ||
      placement === "after_table"
    if (usePos) return false
    if (typeof doc.InsertContent === "function") {
      doc.InsertContent(paras)
      return true
    }
    return false
  }

  function replyCommand(requestId, info, error) {
    if (error) {
      reply(requestId, null, error)
      return
    }
    if (info === false || info == null) {
      reply(requestId, null, "Target not found in document.")
      return
    }
    if (typeof info === "object" && info.ok === false) {
      reply(requestId, null, info.error || "Target not found in document.")
      return
    }
    reply(requestId, info)
  }

  function setScope(values) {
    if (!window.Asc) window.Asc = {}
    // Never replace Asc.scope: callCommand does JSON.stringify(window.Asc.scope).
    if (!window.Asc.scope || typeof window.Asc.scope !== "object") window.Asc.scope = {}
    var scope = window.Asc.scope
    var key
    for (key in scope) {
      if (Object.prototype.hasOwnProperty.call(scope, key) && key !== "prototype") delete scope[key]
    }
    if (!values || typeof values !== "object") return scope
    for (key in values) {
      if (!Object.prototype.hasOwnProperty.call(values, key)) continue
      if (key === "prototype") continue
      var value = values[key]
      if (value === undefined) continue
      if (typeof value === "function") continue
      scope[key] = value
    }
    return scope
  }

  /**
   * OnlyOffice callCommand stringifies the function and evals it in the document
   * frame — outer plugin closures are NOT available there. Bake helpers into the
   * command source so get_outline / resolveDocTarget / inserts keep working.
   */
  function buildDocCommand(fn) {
    var helpers = [
      isTableElement,
      tablePreviewText,
      elementText,
      clipText,
      listLogicalLines,
      rewriteParagraphText,
      countLiteralMatches,
      elementStyleName,
      normalizePlacement,
      resolveDocTarget,
      selectAtElementStart,
      parseRowCol,
      getCellPlainText,
      tableDims,
      tableSummary,
      resolveTableByIndex,
      resolveFlatTable,
      getTableCell,
      selectTableCell,
      writeTableCell,
      insertParasAt
    ]
      .map(function (f) {
        return f.toString()
      })
      .join("\n")
    return new Function(
      helpers + "\nreturn (" + fn.toString() + ").apply(this, arguments);"
    )
  }

  function callDoc(fn, done, opts) {
    var timeoutMs = opts && opts.timeoutMs ? opts.timeoutMs : 20000
    var retriesLeft = opts && opts.retries != null ? opts.retries : 0
    function attempt(left) {
      var finished = false
      var timer = window.setTimeout(function () {
        if (finished) return
        finished = true
        if (left > 0) {
          attempt(left - 1)
          return
        }
        done(
          undefined,
          "OnlyOffice callCommand timeout. Editor may still be loading after open; retry once or wait for ready."
        )
      }, timeoutMs)
      function finish(value, error) {
        if (finished) return
        finished = true
        window.clearTimeout(timer)
        done(value, error)
      }
      try {
        if (!window.Asc || !window.Asc.plugin) {
          finish(undefined, "Asc.plugin missing.")
          return
        }
        if (typeof window.Asc.plugin.callCommand !== "function") {
          finish(undefined, "callCommand unavailable.")
          return
        }
        if (!window.Asc.plugin.info) window.Asc.plugin.info = {}
        // Sub-window frames refuse callCommand/executeMethod ("does not allow in window frame").
        try {
          delete window.Asc.plugin.windowID
        } catch (err) {
          window.Asc.plugin.windowID = undefined
        }
        var command = buildDocCommand(fn)
        window.Asc.plugin.callCommand(command, false, true, function (value) {
          finish(value)
        })
      } catch (err) {
        finish(undefined, err && err.message ? err.message : String(err))
      }
    }
    attempt(retriesLeft)
  }

  function runMethod(name, params, done) {
    var finished = false
    var timer = window.setTimeout(function () {
      if (finished) return
      finished = true
      done(undefined, "OnlyOffice executeMethod timeout: " + name)
    }, 10000)
    function finish(value, error) {
      if (finished) return
      finished = true
      window.clearTimeout(timer)
      done(value, error)
    }
    try {
      if (!window.Asc || !window.Asc.plugin || typeof window.Asc.plugin.executeMethod !== "function") {
        finish(undefined, "executeMethod unavailable.")
        return
      }
      if (!window.Asc.plugin.info) window.Asc.plugin.info = {}
      try {
        delete window.Asc.plugin.windowID
      } catch (err) {
        window.Asc.plugin.windowID = undefined
      }
      window.Asc.plugin.executeMethod(name, params || [], function (value) {
        finish(value)
      })
    } catch (err) {
      finish(undefined, err && err.message ? err.message : String(err))
    }
  }

  function handle(data) {
    if (!data || typeof data !== "object") return
    var requestId = data.requestId
    var type = data.type
    var payload = data.payload && typeof data.payload === "object" ? data.payload : {}
    if (type === "ready" || type === "ping") {
      reply(requestId, { ok: true })
      return
    }

    try {
      if (type === "get_text") {
        setScope({})
        callDoc(function () {
          function tableText(table) {
            if (!table || typeof table.GetRow !== "function") return ""
            var rows = []
            var r
            var max = typeof table.GetRowsCount === "function" ? table.GetRowsCount() : 40
            for (r = 0; r < max; r++) {
              var row = table.GetRow(r)
              if (!row) break
              var cells = []
              var c
              var cols = typeof row.GetCellsCount === "function" ? row.GetCellsCount() : 12
              for (c = 0; c < cols; c++) {
                var cell = row.GetCell(c)
                if (!cell) break
                var content = cell.GetContent && cell.GetContent()
                var para = content && content.GetElement ? content.GetElement(0) : null
                cells.push(para && para.GetText ? para.GetText() : "")
              }
              rows.push(cells.join(" | "))
            }
            return rows.join("\n")
          }
          if (typeof Api.GetDocument === "function") {
            var doc = Api.GetDocument()
            var n = doc.GetElementsCount()
            var parts = []
            var i
            for (i = 0; i < n; i++) {
              var el = doc.GetElement(i)
              if (!el) continue
              var klass = el.GetClassType ? String(el.GetClassType()) : ""
              if (klass.toLowerCase().indexOf("table") >= 0 || el.GetRow) {
                var grid = tableText(el)
                if (grid) parts.push(grid)
                continue
              }
              if (typeof el.GetText === "function") parts.push(el.GetText())
            }
            return parts.join("\n")
          }
          if (typeof Api.GetPresentation === "function") {
            var pres = Api.GetPresentation()
            var slides = []
            var slideCount = pres.GetSlidesCount ? pres.GetSlidesCount() : 0
            var si
            for (si = 0; si < slideCount; si++) {
              var slide = pres.GetSlideByIndex ? pres.GetSlideByIndex(si) : null
              if (!slide || !slide.GetAllShapes) continue
              var shapes = slide.GetAllShapes()
              var texts = []
              var sj
              for (sj = 0; sj < (shapes ? shapes.length : 0); sj++) {
                var content = shapes[sj].GetDocContent && shapes[sj].GetDocContent()
                if (content && content.GetText) texts.push(content.GetText())
              }
              if (texts.length) slides.push("Slide " + (si + 1) + ":\n" + texts.join("\n"))
            }
            return slides.join("\n\n")
          }
          if (typeof Api.GetActiveSheet === "function") {
            var range = Api.GetActiveSheet().GetUsedRange()
            return range && typeof range.GetValue === "function" ? range.GetValue() : ""
          }
          return ""
        }, function (text, error) {
          if (error) reply(requestId, null, error)
          else reply(requestId, { text: text == null ? "" : text })
        })
        return
      }

      if (type === "get_outline") {
        callDoc(
          function () {
            if (typeof Api.GetDocument !== "function") {
              return { outline: [], elements: [], tables: [] }
            }
            var doc = Api.GetDocument()
            var items = []
            var elements = []
            var tables = []
            var flat_tables = []
            var tableOrdinal = 0
            var n = doc.GetElementsCount()
            var i
            for (i = 0; i < n; i++) {
              var el = doc.GetElement(i)
              if (!el) continue
              var preview = elementText(el)
              var styleName = elementStyleName(el)
              var isTable = isTableElement(el)
              var row = {
                index: i,
                type: isTable ? "table" : "paragraph",
                text: preview.length > 240 ? preview.slice(0, 240) : preview,
                style: styleName || null
              }
              if (isTable) {
                row.table_index = tableOrdinal
                var summary = tableSummary(el, tableOrdinal, i)
                row.rows = summary.rows
                row.cols = summary.cols
                tables.push(summary)
                tableOrdinal += 1
              }
              elements.push(row)
              var match = styleName.match(/heading\s*(\d)|t[ií]tulo\s*(\d)/i)
              var headingText =
                typeof el.GetText === "function" ? String(el.GetText() || "") : preview
              if (match) {
                items.push({
                  level: parseInt(match[1] || match[2], 10),
                  text: headingText,
                  index: i
                })
                continue
              }
              // Forms like RAC: numbered titles are often Normal paragraphs OR the
              // first lines inside a wrapping table (not Heading styles).
              var trimmed = String(headingText || preview || "").trim()
              if (isTable) {
                var firstLine = trimmed.split(/\r?\n/)[0] || ""
                trimmed = firstLine.trim()
              }
              if (
                trimmed.length > 0 &&
                trimmed.length < 160 &&
                /^\d+(\.\d+){0,5}\.?\s+\S/.test(trimmed)
              ) {
                var dots = (trimmed.match(/\./g) || []).length
                items.push({
                  level: Math.min(dots + 1, 9),
                  text: trimmed,
                  index: i
                })
              }
            }
            // Nested tables (e.g. yearly grid inside a RAC wrapper) via GetAllTables.
            if (typeof doc.GetAllTables === "function") {
              try {
                var all = doc.GetAllTables()
                var allLen = all && typeof all.length === "number" ? all.length : 0
                if (all && typeof all.GetCount === "function") allLen = all.GetCount()
                var fi
                for (fi = 0; fi < allLen; fi++) {
                  var nest = typeof all.Get === "function" ? all.Get(fi) : all[fi]
                  if (!nest) continue
                  var nestDims = tableDims(nest)
                  var nestText = tablePreviewText(nest)
                  flat_tables.push({
                    flat_table_index: fi,
                    rows: nestDims.rows,
                    cols: nestDims.cols,
                    text: nestText.length > 240 ? nestText.slice(0, 240) : nestText
                  })
                }
              } catch (errFlat) {}
            }
            return {
              outline: items,
              elements: elements,
              tables: tables,
              flat_tables: flat_tables
            }
          },
          function (info, error) {
            if (error) reply(requestId, null, error)
            else if (info == null) {
              reply(
                requestId,
                null,
                "get_outline returned empty (document command failed)."
              )
            } else {
              var body = info && typeof info === "object" ? info : { outline: info || [] }
              if (Array.isArray(body)) body = { outline: body, elements: [], tables: [] }
              reply(requestId, {
                outline: body.outline || [],
                elements: body.elements || [],
                tables: body.tables || [],
                flat_tables: body.flat_tables || []
              })
            }
          },
          { timeoutMs: 20000, retries: 1 }
        )
        return
      }

      if (type === "get_selection") {
        runMethod("GetSelectedText", [{ Numbering: false }], function (text, error) {
          if (error) reply(requestId, null, error)
          else reply(requestId, { text: text || "" })
        })
        return
      }

      if (type === "insert_text" || type === "type" || type === "replace_selection") {
        var insert = normalizeNewlines(payload.text || "")
        var markup = paragraphsFromPayload(payload, insert)
        var insertPlacement = normalizePlacement(payload.placement || payload.at)
        var hasTarget =
          !!(
            (payload.heading != null && String(payload.heading).trim()) ||
            (payload.text_target != null && String(payload.text_target).trim()) ||
            (payload.table_index != null && payload.table_index !== "") ||
            (payload.index != null &&
              payload.index !== "" &&
              insertPlacement !== "cursor") ||
            insertPlacement === "before" ||
            insertPlacement === "after" ||
            insertPlacement === "before_table" ||
            insertPlacement === "after_table"
          )
        var needsFormat = !!(
          markup.hasMarkup ||
          payload.font ||
          payload.size ||
          payload.bold != null ||
          payload.italic != null ||
          payload.underline != null ||
          payload.strike != null ||
          payload.color ||
          payload.highlight ||
          payload.align ||
          payload.style ||
          payload.list ||
          payload.break ||
          hasTarget
        )
        // Plain text at the live cursor: PasteText. Targeted placement uses callDoc + AddElement.
        if (!needsFormat) {
          runMethod("PasteText", [insert], function (_value, error) {
            if (error) reply(requestId, null, error)
            else reply(requestId, type === "replace_selection" ? { replaced: true } : { inserted: insert.length })
          })
          return
        }
        setScope({
          text: insert,
          paragraphs: markup.paragraphs,
          kind: type,
          bold: payload.bold,
          italic: payload.italic,
          underline: payload.underline,
          strike: payload.strike,
          font: payload.font,
          size: payload.size,
          color: payload.color,
          highlight: payload.highlight,
          align: payload.align,
          style: payload.style,
          list: payload.list,
          break: payload.break,
          placement: insertPlacement,
          heading:
            payload.heading != null && String(payload.heading).trim()
              ? String(payload.heading).trim()
              : payload.text_target != null && String(payload.text_target).trim()
                ? String(payload.text_target).trim()
                : null,
          index: payload.index,
          table_index: payload.table_index,
          flat_table_index: payload.flat_table_index,
          row: payload.row != null ? payload.row : payload.row_index,
          col:
            payload.col != null
              ? payload.col
              : payload.column != null
                ? payload.column
                : payload.col_index,
          cell_mode: payload.cell_mode || payload.mode || "append",
          target: payload.target
        })
        callDoc(function () {
          var s = Asc.scope
          function applyRunFormat(run, piece) {
            if (!run) return
            if ((piece && piece.bold === true) || s.bold === true) run.SetBold(true)
            if ((piece && piece.italic === true) || s.italic === true) run.SetItalic(true)
            if (((piece && piece.underline === true) || s.underline === true) && run.SetUnderline) {
              run.SetUnderline("single")
            }
            if (((piece && piece.strike === true) || s.strike === true) && run.SetStrikeout) {
              run.SetStrikeout(true)
            }
            if (s.font && run.SetFontFamily) run.SetFontFamily(s.font)
            if (s.size && run.SetFontSize) run.SetFontSize(Number(s.size) * 2)
            if (s.color && run.SetColor) run.SetColor(s.color.r, s.color.g, s.color.b)
            if (s.highlight && run.SetHighlight) run.SetHighlight(s.highlight)
          }
          function plainFromParagraphs(blocks) {
            var lines = []
            var pi
            for (pi = 0; pi < (blocks ? blocks.length : 0); pi++) {
              var bits = []
              var ri
              for (ri = 0; ri < blocks[pi].length; ri++) {
                if (blocks[pi][ri] && blocks[pi][ri].text) bits.push(blocks[pi][ri].text)
              }
              lines.push(bits.join(""))
            }
            return lines.join("\n")
          }
          if (typeof Api.GetPresentation === "function" && typeof Api.GetDocument !== "function") {
            var pres = Api.GetPresentation()
            var slide = pres.GetCurrentSlide ? pres.GetCurrentSlide() : null
            if (!slide) return { ok: false, error: "No slide." }
            var shape = typeof Api.CreateShape === "function" ? Api.CreateShape("rect", 100 * 36000, 40 * 36000) : null
            if (!shape) return { ok: false, error: "CreateShape unavailable." }
            var box = shape.GetDocContent && shape.GetDocContent()
            var slidePara = box && box.GetElement ? box.GetElement(0) : null
            if (slidePara && slidePara.AddText) slidePara.AddText(plainFromParagraphs(s.paragraphs) || s.text || "")
            if (slide.AddObject) slide.AddObject(shape)
            return { ok: true, inserted: (s.text || "").length }
          }
          if (typeof Api.GetDocument !== "function") return { ok: false, error: "No document." }
          var doc = Api.GetDocument()
          var cellTarget = parseRowCol(s)
          // In-cell write: ApiTableCell.AddText/SetText — never AddElement (breaks table).
          if (cellTarget.row != null && cellTarget.col != null) {
            var plain = plainFromParagraphs(s.paragraphs)
            if (!plain && s.text) plain = String(s.text)
            return writeTableCell(
              doc,
              {
                table_index: s.table_index,
                flat_table_index: s.flat_table_index,
                row: cellTarget.row,
                col: cellTarget.col
              },
              plain,
              s.cell_mode
            )
          }
          var placement = normalizePlacement(s.placement)
          var resolved = null
          var wantsTarget =
            placement === "before" ||
            placement === "after" ||
            placement === "before_table" ||
            placement === "after_table" ||
            (s.heading && String(s.heading).trim()) ||
            (s.table_index != null && s.table_index !== "") ||
            (s.index != null && s.index !== "" && placement !== "cursor")
          if (wantsTarget) {
            var effectivePlacement = placement
            if (
              effectivePlacement === "cursor" &&
              ((s.heading && String(s.heading).trim()) ||
                (s.table_index != null && s.table_index !== "") ||
                (s.index != null && s.index !== ""))
            ) {
              // Heading/table without explicit placement → insert before that element.
              effectivePlacement = "before"
            }
            resolved = resolveDocTarget(doc, {
              heading: s.heading,
              index: s.index,
              table_index: s.table_index,
              target: s.target,
              placement: effectivePlacement
            })
            if (!resolved.ok) return resolved
          }
          var blocks = s.paragraphs
          if (!blocks || !blocks.length) {
            var lines = String(s.text || "").split("\n")
            if (!lines.length) lines = [""]
            blocks = []
            var li
            for (li = 0; li < lines.length; li++) blocks.push([{ text: lines[li] }])
          }
          var paras = []
          var i
          for (i = 0; i < blocks.length; i++) {
            var para = Api.CreateParagraph()
            var pieces = blocks[i] || []
            var j
            for (j = 0; j < pieces.length; j++) {
              var piece = pieces[j] || {}
              if (piece.break === "line") {
                if (para.AddLineBreak) para.AddLineBreak()
                continue
              }
              var run = typeof Api.CreateRun === "function" ? Api.CreateRun() : null
              var chunk = piece.text == null ? "" : String(piece.text)
              if (run) {
                run.AddText(chunk)
                applyRunFormat(run, piece)
                if (para.AddElement) para.AddElement(run)
              } else if (para.AddText) {
                para.AddText(chunk)
              }
            }
            if (s.align && para.SetJc) para.SetJc(s.align)
            if (s.style && i === 0) {
              var style = doc.GetStyle(s.style)
              if (!style && s.style === "Heading 1") style = doc.GetStyle("Título 1")
              if (!style && s.style === "Heading 2") style = doc.GetStyle("Título 2")
              if (style && para.SetStyle) para.SetStyle(style)
            }
            if (i === 0 && s.break === "page" && para.AddPageBreak) para.AddPageBreak()
            if (s.break === "line" && para.AddLineBreak) para.AddLineBreak()
            if (s.list && s.list !== "none" && para.SetNumbering && Api.CreateNumbering) {
              var numbering = Api.CreateNumbering(s.list === "number" ? "numbered" : "bullet")
              if (numbering && numbering.GetLevel) para.SetNumbering(numbering.GetLevel(0))
            }
            paras.push(para)
          }
          if (resolved) {
            if (!insertParasAt(doc, paras, resolved.insert_at, resolved.placement)) {
              return { ok: false, error: "AddElement/InsertContent unavailable." }
            }
            return {
              ok: true,
              inserted: (s.text || "").length,
              index: resolved.index,
              insert_at: resolved.insert_at,
              placement: resolved.placement,
              type: resolved.type
            }
          }
          if (doc.InsertContent) doc.InsertContent(paras)
          return { ok: true, inserted: (s.text || "").length, placement: "cursor" }
        }, function (info, error) {
          if (type === "replace_selection") {
            if (error) reply(requestId, null, error)
            else if (info && info.ok === false) reply(requestId, null, info.error || "Replace failed.")
            else reply(requestId, { replaced: true })
            return
          }
          replyCommand(requestId, info, error)
        })
        return
      }

      if (type === "format") {
        setScope(payload)
        callDoc(function () {
          var s = Asc.scope
          var doc = Api.GetDocument()
          var target = typeof doc.GetRangeBySelect === "function" ? doc.GetRangeBySelect() : null
          if (!target && typeof doc.GetCurrentParagraph === "function") target = doc.GetCurrentParagraph()
          if (!target) return false
          if (s.bold === true && target.SetBold) target.SetBold(true)
          if (s.bold === false && target.SetBold) target.SetBold(false)
          if (s.italic === true && target.SetItalic) target.SetItalic(true)
          if (s.italic === false && target.SetItalic) target.SetItalic(false)
          if (s.underline === true && target.SetUnderline) target.SetUnderline("single")
          if (s.underline === false && target.SetUnderline) target.SetUnderline("none")
          if (s.strike === true && target.SetStrikeout) target.SetStrikeout(true)
          if (s.strike === false && target.SetStrikeout) target.SetStrikeout(false)
          if (s.font && target.SetFontFamily) target.SetFontFamily(s.font)
          if (s.size && target.SetFontSize) target.SetFontSize(Number(s.size) * 2)
          if (s.color && target.SetColor) target.SetColor(s.color.r, s.color.g, s.color.b)
          if (s.highlight && target.SetHighlight) target.SetHighlight(s.highlight)
          if (s.align && target.SetJc) target.SetJc(s.align)
          if (s.spacing_after != null && target.SetSpacingAfter) target.SetSpacingAfter(Math.round(Number(s.spacing_after) * 20))
          if (s.spacing_before != null && target.SetSpacingBefore) target.SetSpacingBefore(Math.round(Number(s.spacing_before) * 20))
          if (s.style) {
            var style = doc.GetStyle(s.style)
            if (!style && s.style === "Heading 1") style = doc.GetStyle("Título 1")
            if (!style && s.style === "Heading 2") style = doc.GetStyle("Título 2")
            if (!style && s.style === "Title") style = doc.GetStyle("Título")
            if (style && target.SetStyle) target.SetStyle(style)
          }
          var para = typeof doc.GetCurrentParagraph === "function" ? doc.GetCurrentParagraph() : target
          if (s.list && para) {
            if (s.list === "none" && para.RemoveNumbering) para.RemoveNumbering()
            if ((s.list === "bullet" || s.list === "number") && para.SetNumbering && Api.CreateNumbering) {
              var numbering = Api.CreateNumbering(s.list === "number" ? "numbered" : "bullet")
              if (numbering && numbering.GetLevel) para.SetNumbering(numbering.GetLevel(0))
            }
          }
          return true
        }, function (ok, error) {
          if (error) reply(requestId, null, error)
          else reply(requestId, { formatted: ok !== false })
        })
        return
      }

      if (type === "table") {
        setScope(payload)
        callDoc(function () {
          var s = Asc.scope
          var data = s.data || []
          var rows = Math.max(1, Math.min(40, Number(s.rows) || data.length || 2))
          var cols = Math.max(1, Math.min(12, Number(s.cols) || (data[0] && data[0].length) || 2))
          var table = Api.CreateTable(cols, rows)
          if (table.SetWidth) table.SetWidth("percent", Number(s.width_percent) || 100)
          var r
          var c
          for (r = 0; r < rows; r++) {
            var row = table.GetRow(r)
            if (!row) continue
            for (c = 0; c < cols; c++) {
              var cell = row.GetCell(c)
              if (!cell || !cell.GetContent) continue
              var content = cell.GetContent()
              var para = content && content.GetElement ? content.GetElement(0) : null
              if (!para && Api.CreateParagraph && content && content.Push) {
                para = Api.CreateParagraph()
                content.Push(para)
              }
              var value = data[r] && data[r][c] != null ? String(data[r][c]) : ""
              if (para && para.AddText) para.AddText(value)
              if (s.header && r === 0 && para && para.SetBold) para.SetBold(true)
            }
          }
          Api.GetDocument().InsertContent([table])
          return { rows: rows, cols: cols }
        }, function (info, error) {
          if (error) reply(requestId, null, error)
          else reply(requestId, info || { ok: true })
        })
        return
      }

      if (type === "layout") {
        setScope(payload)
        callDoc(function () {
          var s = Asc.scope
          var doc = Api.GetDocument()
          var section = doc.GetFinalSection ? doc.GetFinalSection() : null
          if (!section) return false
          if (s.orientation && section.SetPageOrient) section.SetPageOrient(s.orientation)
          var sizes = {
            a4: [11906, 16838],
            letter: [12240, 15840],
            legal: [12240, 20160]
          }
          if (s.page && sizes[s.page] && section.SetPageSize) {
            var size = sizes[s.page]
            if (s.orientation === "landscape") section.SetPageSize(size[1], size[0])
            else section.SetPageSize(size[0], size[1])
          }
          var mm = function (value) {
            return Math.round(Number(value) * 56.7)
          }
          if (s.margin_mm != null && section.SetPageMargins) {
            var all = mm(s.margin_mm)
            section.SetPageMargins(all, all, all, all)
          }
          if (s.margins && section.SetPageMargins) {
            section.SetPageMargins(
              mm(s.margins.left != null ? s.margins.left : 25),
              mm(s.margins.top != null ? s.margins.top : 25),
              mm(s.margins.right != null ? s.margins.right : 25),
              mm(s.margins.bottom != null ? s.margins.bottom : 25)
            )
          }
          if (s.columns && section.SetTypeCols) section.SetTypeCols(Number(s.columns))
          if (s.break === "page") {
            var para = Api.CreateParagraph()
            if (para.AddPageBreak) para.AddPageBreak()
            doc.InsertContent([para])
          }
          return true
        }, function (ok, error) {
          if (error) reply(requestId, null, error)
          else reply(requestId, { layout: ok !== false })
        })
        return
      }

      if (type === "goto" || type === "select") {
        var headingText = ""
        if (payload.heading != null && String(payload.heading).trim()) {
          headingText = String(payload.heading).trim()
        } else if (payload.text != null && String(payload.text).trim()) {
          headingText = String(payload.text).trim()
        }
        var rawIndex = payload.index
        var placement = normalizePlacement(payload.placement || payload.at)
        // Prefer heading when the model sends index 0 as a default alongside text.
        var useHeading =
          !!headingText &&
          (rawIndex == null ||
            rawIndex === "" ||
            (type === "goto" && Number(rawIndex) === 0))
        setScope({
          kind: type,
          heading: useHeading ? headingText : headingText || null,
          index: useHeading ? null : rawIndex,
          table_index: payload.table_index,
          flat_table_index: payload.flat_table_index,
          row: payload.row != null ? payload.row : payload.row_index,
          col:
            payload.col != null
              ? payload.col
              : payload.column != null
                ? payload.column
                : payload.col_index,
          page: payload.page,
          target: payload.target,
          placement: placement
        })
        if (payload.target === "all" && type === "select") {
          runMethod("SelectAll", [], function (_value, error) {
            if (error) reply(requestId, null, error)
            else reply(requestId, { selected: "all" })
          })
          return
        }
        callDoc(
          function () {
            var s = Asc.scope
            var doc = Api.GetDocument()
            var cellTarget = parseRowCol(s)
            if (cellTarget.row != null && cellTarget.col != null) {
              return selectTableCell(doc, {
                table_index: s.table_index,
                flat_table_index: s.flat_table_index,
                row: cellTarget.row,
                col: cellTarget.col
              })
            }
            var resolved = resolveDocTarget(doc, {
              heading: s.heading,
              index: s.index,
              table_index: s.table_index,
              target: s.target,
              placement: s.placement
            })
            if (!resolved.ok) return resolved
            var target = doc.GetElement(resolved.index)
            if (!target) return { ok: false, error: "Element missing at index " + resolved.index }
            // before/before_table: park the caret on the previous element so a following
            // PasteText/InsertContent does not land inside the table. Prefer atomic
            // doc_insert_text with placement=before_table (AddElement) for markers.
            if (resolved.placement === "before" || resolved.placement === "before_table") {
              if (resolved.index > 0) {
                var prev = doc.GetElement(resolved.index - 1)
                if (prev) {
                  if (typeof prev.GetRange === "function") {
                    var prevText =
                      typeof prev.GetText === "function" ? String(prev.GetText() || "") : ""
                    try {
                      var endRange = prev.GetRange(prevText.length, prevText.length)
                      if (endRange && typeof endRange.Select === "function") endRange.Select()
                      else selectAtElementStart(prev)
                    } catch (err) {
                      selectAtElementStart(prev)
                    }
                  } else if (prev.Select) {
                    prev.Select()
                  }
                } else {
                  selectAtElementStart(target)
                }
              } else {
                selectAtElementStart(target)
              }
            } else if (resolved.placement === "after" || resolved.placement === "after_table") {
              var next = doc.GetElement(resolved.insert_at)
              if (next) selectAtElementStart(next)
              else if (target.Select) target.Select()
            } else if (s.kind === "select" && target.Select) {
              target.Select()
            } else {
              selectAtElementStart(target)
            }
            var tableMeta = null
            if (resolved.type === "table" && isTableElement(target)) {
              var dims = tableDims(target)
              tableMeta = { rows: dims.rows, cols: dims.cols }
              if (typeof doc.GetActiveTable === "function") {
                try {
                  tableMeta.verified_in_table = !!doc.GetActiveTable()
                } catch (errActive) {
                  tableMeta.verified_in_table = null
                }
              }
            }
            return {
              ok: true,
              index: resolved.index,
              insert_at: resolved.insert_at,
              placement: resolved.placement,
              heading: resolved.heading,
              type: resolved.type,
              table: tableMeta
            }
          },
          function (info, error) {
            replyCommand(requestId, info, error)
          },
          { timeoutMs: 20000, retries: 1 }
        )
        return
      }

      if (type === "delete") {
        runMethod("InputText", ["", ""], function (_value, error) {
          if (error) reply(requestId, null, error)
          else reply(requestId, { deleted: true })
        })
        return
      }

      if (type === "find_replace") {
        var find = String(payload.find || "")
        if (!find) {
          reply(requestId, null, "find is required.")
          return
        }
        if (payload.all === true) {
          runMethod(
            "SearchAndReplace",
            [{
              searchString: find,
              replaceString: String(payload.replace || ""),
              matchCase: payload.match_case === true
            }],
            function (_value, error) {
              if (error) reply(requestId, null, error)
              else reply(requestId, { replaced: true, find: find, all: true })
            }
          )
          return
        }
        setScope({
          find: find,
          replace: String(payload.replace || ""),
          match_case: payload.match_case === true
        })
        callDoc(function () {
          var s = Asc.scope
          var doc = Api.GetDocument()
          if (!doc.Search) return false
          var hits = doc.Search(s.find)
          if (!hits || !hits.length) return { replaced: false, find: s.find }
          var first = hits[0]
          if (first.Select) first.Select()
          if (first.Delete) first.Delete()
          var para = doc.GetCurrentParagraph ? doc.GetCurrentParagraph() : null
          if (para && para.AddText) para.AddText(s.replace)
          return { replaced: true, find: s.find, all: false }
        }, function (info, error) {
          if (error) reply(requestId, null, error)
          else reply(requestId, info || { replaced: false, find: find })
        })
        return
      }

      // Line-oriented Word tools (like files_grep / files_read / files_replace).
      // Indexing: 1-based line = top-level element ordinal (same as outline.elements).
      if (type === "grep" || type === "read" || type === "replace") {
        setScope({
          cmd: type,
          pattern: payload.pattern != null ? String(payload.pattern) : "",
          old_string: payload.old_string != null ? String(payload.old_string) : "",
          new_string: payload.new_string != null ? String(payload.new_string) : "",
          replace_all: payload.replace_all === true,
          start_line: payload.start_line,
          end_line: payload.end_line,
          context: payload.context,
          max_matches: payload.max_matches
        })
        callDoc(
          function () {
            var s = Asc.scope
            if (typeof Api.GetDocument !== "function") {
              return { ok: false, error: "Not a Word document." }
            }
            var doc = Api.GetDocument()
            var all = listLogicalLines(doc)
            var lineCount = all.length
            var startLine =
              s.start_line != null && s.start_line !== "" ? Number(s.start_line) : null
            var endLine = s.end_line != null && s.end_line !== "" ? Number(s.end_line) : null
            var MAX_READ = 80
            var MAX_MATCHES = 40
            var MAX_CONTEXT = 3
            var CLIP = 500

            function inScope(line) {
              if (startLine != null && !isNaN(startLine) && line.line < startLine) return false
              if (endLine != null && !isNaN(endLine) && line.line > endLine) return false
              return true
            }

            function scopedRows() {
              return all.filter(inScope)
            }

            if (s.cmd === "grep") {
              var pattern = String(s.pattern || "")
              if (!pattern) return { ok: false, error: "pattern is required." }
              var re
              try {
                re = new RegExp(pattern)
              } catch (errRe) {
                return { ok: false, error: "invalid pattern: " + (errRe && errRe.message ? errRe.message : errRe) }
              }
              var ctx = Number(s.context)
              if (isNaN(ctx) || ctx < 0) ctx = 2
              if (ctx > MAX_CONTEXT) ctx = MAX_CONTEXT
              var cap = Number(s.max_matches)
              if (isNaN(cap) || cap < 1) cap = 20
              if (cap > MAX_MATCHES) cap = MAX_MATCHES
              var rows = scopedRows()
              var matches = []
              var truncated = false
              var mi
              for (mi = 0; mi < rows.length; mi++) {
                var row = rows[mi]
                if (!re.test(row.text || "")) continue
                if (matches.length >= cap) {
                  truncated = true
                  break
                }
                var from = Math.max(0, row.element_index - ctx)
                var to = Math.min(all.length - 1, row.element_index + ctx)
                var previewParts = []
                var pi
                for (pi = from; pi <= to; pi++) {
                  var pref = all[pi]
                  var sign = pref.line === row.line ? ">" : " "
                  previewParts.push(
                    sign + " " + pref.line + "|" + clipText(pref.text || "", CLIP)
                  )
                }
                matches.push({
                  line: row.line,
                  element_index: row.element_index,
                  type: row.type,
                  text: clipText(row.text || "", CLIP),
                  preview: previewParts.join("\n"),
                  table_index: row.table_index
                })
              }
              return {
                ok: true,
                indexing: "paragraph",
                line_count: lineCount,
                match_count: matches.length,
                truncated: truncated,
                matches: matches
              }
            }

            if (s.cmd === "read") {
              var startN = startLine != null && !isNaN(startLine) ? startLine : 1
              if (startN < 1) return { ok: false, error: "start_line must be >= 1" }
              if (lineCount === 0 && startN === 1) {
                return {
                  ok: true,
                  indexing: "paragraph",
                  line_count: 0,
                  start_line: 1,
                  end_line: 0,
                  clipped: false,
                  lines: []
                }
              }
              if (lineCount === 0) return { ok: false, error: "start_line is past the end" }
              if (startN > lineCount) {
                return {
                  ok: false,
                  error: "start_line " + startN + " is past line " + lineCount
                }
              }
              var requestedEnd =
                endLine != null && !isNaN(endLine) ? endLine : startN + MAX_READ - 1
              if (requestedEnd < startN) {
                return { ok: false, error: "end_line must be >= start_line" }
              }
              var finish = Math.min(requestedEnd, startN + MAX_READ - 1, lineCount)
              var outLines = []
              var li
              for (li = startN; li <= finish; li++) {
                var item = all[li - 1]
                outLines.push({
                  line: item.line,
                  element_index: item.element_index,
                  type: item.type,
                  text: item.text || "",
                  table_index: item.table_index
                })
              }
              return {
                ok: true,
                indexing: "paragraph",
                line_count: lineCount,
                start_line: startN,
                end_line: finish,
                clipped: requestedEnd > finish,
                lines: outLines
              }
            }

            // replace
            var oldString = String(s.old_string || "")
            var patternSrc = String(s.pattern || "")
            var newString = String(s.new_string != null ? s.new_string : "")
            var replaceAll = s.replace_all === true
            if (!oldString && !patternSrc) {
              return { ok: false, error: "Pass old_string or pattern" }
            }
            if (oldString && patternSrc) {
              return { ok: false, error: "Pass old_string or pattern, not both" }
            }
            if (oldString && !oldString.length) {
              return { ok: false, error: "old_string is empty" }
            }
            var replaceRe = null
            if (patternSrc) {
              try {
                replaceRe = new RegExp(patternSrc, replaceAll ? "g" : "")
              } catch (errPat) {
                return {
                  ok: false,
                  error: "invalid pattern: " + (errPat && errPat.message ? errPat.message : errPat)
                }
              }
            }

            var candidates = scopedRows()
            var hitLines = []
            var totalMatches = 0
            var ci
            for (ci = 0; ci < candidates.length; ci++) {
              var cand = candidates[ci]
              if (cand.type === "table") {
                var tableHit = false
                if (oldString) tableHit = (cand.text || "").indexOf(oldString) >= 0
                else tableHit = replaceRe.test(cand.text || "")
                if (tableHit) {
                  return {
                    ok: false,
                    error:
                      "Match is inside a table line (line " +
                      cand.line +
                      ", table_index=" +
                      cand.table_index +
                      "). Use doc_insert_text / cell tools with table_index+row+col, not doc_replace.",
                    line: cand.line,
                    table_index: cand.table_index
                  }
                }
                continue
              }
              var hay = cand.text || ""
              var nHits = 0
              if (oldString) nHits = countLiteralMatches(hay, oldString)
              else {
                var flags = replaceAll ? "g" : ""
                var scanRe
                try {
                  scanRe = new RegExp(patternSrc, flags)
                } catch (errScan) {
                  return { ok: false, error: "invalid pattern" }
                }
                var m
                if (flags.indexOf("g") >= 0) {
                  while ((m = scanRe.exec(hay)) !== null) {
                    if (!m[0].length) break
                    nHits += 1
                    if (nHits > 100) break
                  }
                } else if (scanRe.test(hay)) {
                  nHits = 1
                }
              }
              if (nHits > 0) {
                totalMatches += nHits
                hitLines.push({ line: cand.line, element_index: cand.element_index, hits: nHits, before: hay })
              }
            }

            if (totalMatches === 0) return { ok: false, error: "No match" }
            if (!replaceAll && totalMatches !== 1) {
              return {
                ok: false,
                error:
                  "Matched " +
                  totalMatches +
                  " times. Pass replace_all or a unique old_string/pattern.",
                match_count: totalMatches,
                match_lines: hitLines.map(function (h) {
                  return h.line
                }),
                snippets: hitLines.slice(0, 8).map(function (h) {
                  return { line: h.line, text: clipText(h.before, 160) }
                })
              }
            }

            var replacements = 0
            var hunkParts = []
            var hi
            for (hi = 0; hi < hitLines.length; hi++) {
              var hit = hitLines[hi]
              var para = doc.GetElement(hit.element_index)
              if (!para || isTableElement(para)) continue
              var before = hit.before
              var after
              if (oldString) {
                after = replaceAll
                  ? before.split(oldString).join(newString)
                  : before.replace(oldString, newString)
              } else {
                var applyRe = new RegExp(patternSrc, replaceAll ? "g" : "")
                after = before.replace(applyRe, newString)
              }
              if (after === before) continue
              if (!rewriteParagraphText(para, after)) {
                return {
                  ok: false,
                  error: "Could not rewrite paragraph at line " + hit.line + "."
                }
              }
              replacements += replaceAll ? hit.hits : 1
              hunkParts.push("- " + hit.line + "|" + clipText(before, CLIP))
              hunkParts.push("+ " + hit.line + "|" + clipText(after, CLIP))
            }

            return {
              ok: true,
              indexing: "paragraph",
              replacements: replacements,
              line_count: lineCount,
              hunk: hunkParts.join("\n")
            }
          },
          function (info, error) {
            replyCommand(requestId, info, error)
          },
          { timeoutMs: 20000, retries: 1 }
        )
        return
      }

      if (type === "get_range" || type === "set_range" || type === "goto_cell") {
        // callCommand stringifies this function and evals it in the editor.
        // The outer `type` is not visible there, so the command travels on Asc.scope.
        setScope({
          cmd: type,
          range: payload.range || payload.cell || "A1",
          cell: payload.cell,
          sheet: payload.sheet,
          value: payload.value,
          values: payload.values,
          rename: payload.rename
        })
        callDoc(function () {
          var s = Asc.scope
          function originOf(addr) {
            var head = String(addr || "A1").split(":")[0].replace(/\$/g, "")
            var match = head.match(/^([A-Za-z]+)(\d+)$/)
            if (!match) return null
            var letters = match[1].toUpperCase()
            var col = 0
            var i
            for (i = 0; i < letters.length; i++) col = col * 26 + (letters.charCodeAt(i) - 64)
            return { col: col - 1, row: parseInt(match[2], 10) - 1 }
          }
          function topLeft(value) {
            var cur = value
            var guard = 0
            while (Array.isArray(cur) && guard < 4) {
              if (!cur.length) return ""
              cur = cur[0]
              guard++
            }
            return cur
          }
          if (typeof Api.GetActiveSheet !== "function") return JSON.stringify({ error: "Not a spreadsheet." })
          var sheet = Api.GetActiveSheet()
          if (s.sheet && Api.GetSheet) {
            var named = Api.GetSheet(s.sheet)
            if (named) sheet = named
          }
          if (s.sheet && sheet.SetName && s.cmd === "goto_cell" && s.rename) sheet.SetName(s.sheet)
          var addr = s.range || s.cell || "A1"
          var range = sheet.GetRange(addr)
          if (!range) return JSON.stringify({ error: "Range not found." })
          if (s.cmd === "get_range") {
            return JSON.stringify({
              sheet: sheet.GetName ? sheet.GetName() : s.sheet,
              range: addr,
              values: range.GetValue ? range.GetValue() : null
            })
          }
          if (s.cmd === "set_range") {
            var raw = s.values != null ? s.values : s.value
            if (raw == null) return JSON.stringify({ error: "value is required." })
            var matrix = null
            if (Array.isArray(raw)) matrix = raw.length && Array.isArray(raw[0]) ? raw : [raw]
            if (!matrix) {
              if (!range.SetValue) return JSON.stringify({ error: "SetValue is unavailable." })
              range.SetValue(raw)
            } else if (typeof sheet.GetRangeByNumber === "function") {
              var origin = originOf(addr)
              if (!origin) return JSON.stringify({ error: "Range not found." })
              var r
              var c
              for (r = 0; r < matrix.length; r++) {
                var row = matrix[r] || []
                if (!Array.isArray(row)) row = [row]
                for (c = 0; c < row.length; c++) {
                  var cell = sheet.GetRangeByNumber(origin.row + r, origin.col + c)
                  if (!cell || !cell.SetValue) return JSON.stringify({ error: "Cell not found." })
                  var piece = row[c]
                  if (piece == null) piece = ""
                  cell.SetValue(piece)
                }
              }
            } else if (range.SetValue) {
              range.SetValue(matrix)
            } else {
              return JSON.stringify({ error: "SetValue is unavailable." })
            }
            var got = range.GetValue ? range.GetValue() : null
            var expect = matrix ? (matrix[0] && matrix[0].length ? matrix[0][0] : "") : raw
            var formula = typeof expect === "string" && expect.charAt(0) === "="
            if (!formula && String(topLeft(got)) !== String(expect == null ? "" : expect)) {
              return JSON.stringify({
                error: "Cell was not updated.",
                range: addr,
                expected: expect,
                actual: topLeft(got)
              })
            }
            return JSON.stringify({
              sheet: sheet.GetName ? sheet.GetName() : s.sheet,
              range: addr,
              wrote: true,
              values: got
            })
          }
          if (s.cmd === "goto_cell") {
            if (range.Select) range.Select()
            return JSON.stringify({ sheet: sheet.GetName ? sheet.GetName() : s.sheet, range: addr })
          }
          return JSON.stringify({ error: "Unsupported sheet command." })
        }, function (info, error) {
          if (error) {
            reply(requestId, null, error)
            return
          }
          var parsed = info
          if (typeof info === "string") {
            try {
              parsed = JSON.parse(info)
            } catch (err) {
              parsed = null
            }
          }
          if (parsed && parsed.onlyoffice_id_error_message) {
            reply(requestId, null, String(parsed.onlyoffice_id_error_message))
            return
          }
          if (!parsed || parsed.error) {
            reply(requestId, null, (parsed && parsed.error) || "Spreadsheet command did not run.")
            return
          }
          reply(requestId, parsed)
        })
        return
      }

      if (type === "context" || type === "suggest" || type === "suggest_update" || type === "suggest_accept" || type === "suggest_dismiss") {
        handleGhost(type, payload, requestId)
        return
      }

      if (type === "undo" || type === "redo") {
        runMethod(type === "undo" ? "Undo" : "Redo", [], function (_value, error) {
          if (error) reply(requestId, null, error)
          else reply(requestId, { ok: true })
        })
        return
      }

      reply(requestId, null, "Unsupported document command.")
    } catch (err) {
      reply(requestId, null, err && err.message ? err.message : String(err))
    }
  }

  function emit(event, result) {
    var payload = {
      type: "agent-doc",
      event: event,
      requestId: event,
      result: result || null,
      error: null
    }
    var raw = JSON.stringify(payload)
    var seen = []
    var win = window
    while (win) {
      try {
        if (seen.indexOf(win) >= 0) break
        seen.push(win)
        win.postMessage(raw, "*")
      } catch (err) {}
      var next = null
      try {
        next = win.parent
      } catch (err) {
        break
      }
      if (!next || next === win) break
      win = next
    }
  }

  function ghostState() {
    if (!window.__jiGhost) window.__jiGhost = { text: "", active: false }
    return window.__jiGhost
  }

  function handleGhost(type, payload, requestId) {
    if (type === "context") {
      callDoc(function () {
        if (typeof Api.GetDocument !== "function") return { prefix: "", suffix: "" }
        var para = Api.GetDocument().GetCurrentParagraph && Api.GetDocument().GetCurrentParagraph()
        var text = para && para.GetText ? String(para.GetText()) : ""
        var ghost = window.__jiGhost && window.__jiGhost.active ? String(window.__jiGhost.text || "") : ""
        if (ghost && text.slice(-ghost.length) === ghost) text = text.slice(0, -ghost.length)
        return { prefix: text, suffix: "" }
      }, function (info) {
        var prefix = info && info.prefix ? String(info.prefix) : ""
        var ghost = ghostState().active ? String(ghostState().text || "") : ""
        if (ghost && prefix.slice(-ghost.length) === ghost) prefix = prefix.slice(0, -ghost.length)
        reply(requestId, { prefix: prefix, suffix: (info && info.suffix) || "" })
      })
      return
    }

    if (type === "suggest_dismiss") {
      dismissGhost(function () {
        reply(requestId, { dismissed: true })
        emit("ghost", { shown: false, text: "" })
      })
      return
    }

    if (type === "suggest_accept") {
      var current = ghostState()
      if (!current.active || !current.text) {
        reply(requestId, { accepted: false })
        return
      }
      setScope({ ghost: current.text })
      callDoc(function () {
        var s = Asc.scope
        var doc = Api.GetDocument()
        if (!doc || !doc.Search) return false
        var hits = doc.Search(s.ghost)
        if (!hits || !hits.length) return false
        var last = hits[hits.length - 1]
        if (last.SetItalic) last.SetItalic(false)
        if (last.SetColor) last.SetColor(0, 0, 0)
        if (last.SetHighlight) last.SetHighlight("none")
        return true
      }, function (ok) {
        ghostState().active = false
        ghostState().text = ""
        reply(requestId, { accepted: ok !== false })
        emit("ghost", { shown: false, text: "", accepted: true })
      })
      return
    }

    var next = String(payload.text || "")
    if (!next) {
      reply(requestId, { shown: false })
      return
    }
    dismissGhost(function () {
      setScope({ text: next })
      callDoc(function () {
        var s = Asc.scope
        if (typeof Api.GetDocument !== "function") return false
        var para = Api.GetDocument().GetCurrentParagraph && Api.GetDocument().GetCurrentParagraph()
        if (!para) para = Api.CreateParagraph()
        var run = Api.CreateRun ? Api.CreateRun() : null
        if (run) {
          run.AddText(s.text || "")
          if (run.SetColor) run.SetColor(160, 160, 160)
          if (run.SetItalic) run.SetItalic(true)
          if (para.AddElement) para.AddElement(run)
        } else if (para.AddText) {
          para.AddText(s.text || "")
        }
        return true
      }, function () {
        var state = ghostState()
        state.active = true
        state.text = next
        reply(requestId, { shown: true, text: next })
        emit("ghost", { shown: true, text: next })
      })
    })
  }

  function dismissGhost(done) {
    var state = ghostState()
    if (!state.active || !state.text) {
      state.active = false
      state.text = ""
      if (done) done()
      return
    }
    setScope({ ghost: state.text })
    callDoc(function () {
      var s = Asc.scope
      var doc = Api.GetDocument()
      if (!doc || !doc.Search) return false
      var hits = doc.Search(s.ghost)
      if (!hits || !hits.length) return false
      var last = hits[hits.length - 1]
      if (last.Delete) last.Delete()
      return true
    }, function () {
      state.active = false
      state.text = ""
      if (done) done()
    })
  }

  function onEditorKey(event) {
    var key = event && (event.KeyCode || event.keyCode || event.keycode)
    if (key === 9 && ghostState().active) {
      handleGhost("suggest_accept", {}, "suggest_accept")
      return false
    }
    if (key === 27 && ghostState().active) {
      handleGhost("suggest_dismiss", {}, "suggest_dismiss")
      return false
    }
    if (ghostState().active && key !== 9 && key !== 27) {
      handleGhost("suggest_dismiss", {}, "suggest_dismiss")
    }
    return true
  }

  function isHostCommand(data) {
    return !!(
      data &&
      typeof data === "object" &&
      typeof data.requestId === "string" &&
      typeof data.type === "string" &&
      data.type !== "agent-doc"
    )
  }

  function onHostMessage(event) {
    var data = event && event.data
    if (typeof data === "string") {
      try {
        data = JSON.parse(data)
      } catch (err) {
        return
      }
    }
    if (!isHostCommand(data)) return
    handle(data)
  }

  var AGENT_GUID = "asc.{7E4A1C90-2B6D-4F11-9A33-8C0E5D71B2A4}"

  function ensurePluginApi() {
    try {
      if (!window.Asc || !window.Asc.plugin) return
      if (typeof window.Asc.plugin.callCommand === "function") return
      if (typeof window.startPluginApi === "function") {
        try {
          window.Asc.plugin.isStarted = false
          window.startPluginApi()
        } catch (err) {}
        window.Asc.plugin.isStarted = true
        return
      }
      // plugins.js loads ./config.json then posts {type:"initialize"}. Under some
      // third-party /embed embeds that XHR never completes (guid stays empty,
      // _initInternal stays false) so plugin_init / callCommand never arrive.
      if (!window.Asc.plugin.guid) {
        window.Asc.plugin.guid = AGENT_GUID
      }
      var guid = window.Asc.plugin.guid
      if (!guid) return
      var payload = { type: "initialize", guid: guid }
      try {
        if (window.Asc.plugin.windowID) payload.windowID = window.Asc.plugin.windowID
      } catch (err) {}
      // Mark ready to receive plugin_init even if config.json onload never fired.
      window.Asc.plugin._initInternal = true
      var raw = JSON.stringify(payload)
      var win = window
      var hops = 0
      while (win && hops < 4) {
        try {
          win.parent && win.parent !== win && win.parent.postMessage(raw, "*")
        } catch (err) {}
        try {
          if (!win.parent || win.parent === win) break
          win = win.parent
        } catch (err) {
          break
        }
        hops++
      }
    } catch (err) {}
  }

  function attachHandlersAndReady() {
    if (install.readySent) return
    if (typeof window.Asc.plugin.callCommand !== "function") {
      // Still missing — do not lie to the host that the Agent can run commands.
      return
    }
    install.readySent = true
    window.Asc.plugin.attachEvent("onExternalPluginMessage", handle)
    window.Asc.plugin.onExternalPluginMessage = handle
    window.Asc.plugin.event_onKeyDown = onEditorKey
    try {
      window.Asc.plugin.attachEvent("onKeyDown", onEditorKey)
      // Do not attach onTargetPositionChanged → callCommand (WASM dies after minutes).
    } catch (err) {}
    reply("ready", {
      ok: true,
      hasCallCommand: true,
      hasExecuteMethod: typeof window.Asc.plugin.executeMethod === "function",
      hasStartApi: typeof window.startPluginApi === "function"
    })
  }

  function install() {
    if (!install.listening) {
      install.listening = true
      window.addEventListener("message", onHostMessage)
    }
    if (!window.Asc || !window.Asc.plugin) {
      install.tries = (install.tries || 0) + 1
      if (install.tries > 200) return
      window.setTimeout(install, 100)
      return
    }

    if (!install.hookedInit) {
      install.hookedInit = true
      // OnlyOffice adds callCommand inside window.startPluginApi(), which runs on
      // the editor's "init" message. Hook init so we never advertise ready early,
      // and force-start the API if the editor skipped it (third-party /embed).
      var prevInit = window.Asc.plugin.init
      window.Asc.plugin.init = function () {
        ensurePluginApi()
        if (typeof prevInit === "function" && prevInit !== window.Asc.plugin.init) {
          try {
            prevInit.apply(this, arguments)
          } catch (err) {}
        }
        attachHandlersAndReady()
      }
      window.Asc.plugin.button = function () {}
    }

    if (typeof window.Asc.plugin.callCommand === "function") {
      attachHandlersAndReady()
      return
    }

    // Asc.plugin exists but API not started yet — keep polling.
    install.apiTries = (install.apiTries || 0) + 1
    ensurePluginApi()
    if (typeof window.Asc.plugin.callCommand === "function") {
      attachHandlersAndReady()
      return
    }
    if (install.apiTries <= 150) {
      window.setTimeout(install, 200)
    }
  }

  install()
})()
