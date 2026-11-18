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
    if (!value || value === "cursor") return "cursor"
    if (value === "before" || value === "before_element") return "before"
    if (value === "after" || value === "after_element") return "after"
    if (
      value === "before_table" ||
      value === "before_next_table" ||
      value === "before_following_table"
    ) {
      return "before_table"
    }
    if (value === "after_table" || value === "after_next_table") return "after_table"
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
      for (hi = 0; hi < n; hi++) {
        var text = elementText(doc.GetElement(hi)).toLowerCase()
        if (text.indexOf(needle) >= 0) {
          index = hi
          break
        }
      }
      if (index == null || isNaN(index)) {
        return { ok: false, error: "Heading/text not found: " + heading }
      }
    }

    if (targetKind === "table" && (index == null || isNaN(index))) {
      return { ok: false, error: "Pass table_index or heading to target a table." }
    }

    if (
      (placement === "before_table" || placement === "after_table") &&
      index != null &&
      !isNaN(index)
    ) {
      var start = isTableElement(doc.GetElement(index)) ? index : index + 1
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

  function insertParasAt(doc, paras, insertAt, placement) {
    if (!paras || !paras.length) return false
    var usePos =
      placement === "before" ||
      placement === "after" ||
      placement === "before_table" ||
      placement === "after_table"
    if (usePos && typeof doc.AddElement === "function" && insertAt != null) {
      var i
      // AddElement(nPos) inserts before the current element at nPos; add last→first.
      for (i = paras.length - 1; i >= 0; i--) {
        doc.AddElement(insertAt, paras[i])
      }
      return true
    }
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
        window.Asc.plugin.callCommand(fn, false, true, function (value) {
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
                tables.push({
                  index: i,
                  table_index: tableOrdinal,
                  text: row.text
                })
                tableOrdinal += 1
              }
              elements.push(row)
              var match = styleName.match(/heading\s*(\d)|t[ií]tulo\s*(\d)/i)
              if (!match) continue
              items.push({
                level: parseInt(match[1] || match[2], 10),
                text: typeof el.GetText === "function" ? el.GetText() : preview,
                index: i
              })
            }
            return { outline: items, elements: elements, tables: tables }
          },
          function (info, error) {
            if (error) reply(requestId, null, error)
            else {
              var body = info && typeof info === "object" ? info : { outline: info || [] }
              if (Array.isArray(body)) body = { outline: body, elements: [], tables: [] }
              reply(requestId, {
                outline: body.outline || [],
                elements: body.elements || [],
                tables: body.tables || []
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
            return {
              ok: true,
              index: resolved.index,
              insert_at: resolved.insert_at,
              placement: resolved.placement,
              heading: resolved.heading,
              type: resolved.type
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
