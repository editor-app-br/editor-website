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

  function callDoc(fn, done) {
    var finished = false
    var timer = window.setTimeout(function () {
      if (finished) return
      finished = true
      done(undefined, "OnlyOffice callCommand timeout.")
    }, 10000)
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
        callDoc(function () {
          if (typeof Api.GetDocument !== "function") return []
          var doc = Api.GetDocument()
          var items = []
          var n = doc.GetElementsCount()
          var i
          for (i = 0; i < n; i++) {
            var el = doc.GetElement(i)
            if (!el || typeof el.GetStyle !== "function") continue
            var style = el.GetStyle()
            var name = style && typeof style.GetName === "function" ? String(style.GetName()) : ""
            var match = name.match(/heading\s*(\d)|t[ií]tulo\s*(\d)/i)
            if (!match) continue
            items.push({
              level: parseInt(match[1] || match[2], 10),
              text: typeof el.GetText === "function" ? el.GetText() : "",
              index: i
            })
          }
          return items
        }, function (items, error) {
          if (error) reply(requestId, null, error)
          else reply(requestId, { outline: items || [] })
        })
        return
      }

      if (type === "get_selection") {
        runMethod("GetSelectedText", [{ Numbering: false }], function (text, error) {
          if (error) reply(requestId, null, error)
          else reply(requestId, { text: text || "" })
        })
        return
      }

      if (type === "insert_text" || type === "type") {
        var insert = normalizeNewlines(payload.text || "")
        var needsFormat = !!(
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
          payload.break
        )
        // Plain text: PasteText (string arg). Wrong InputText([{text}]) never completes.
        if (!needsFormat) {
          runMethod("PasteText", [insert], function (_value, error) {
            if (error) reply(requestId, null, error)
            else reply(requestId, { inserted: insert.length })
          })
          return
        }
        setScope({
          text: insert,
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
          break: payload.break
        })
        callDoc(function () {
          var s = Asc.scope
          if (typeof Api.GetPresentation === "function" && typeof Api.GetDocument !== "function") {
            var pres = Api.GetPresentation()
            var slide = pres.GetCurrentSlide ? pres.GetCurrentSlide() : null
            if (!slide) return false
            var shape = typeof Api.CreateShape === "function" ? Api.CreateShape("rect", 100 * 36000, 40 * 36000) : null
            if (!shape) return false
            var box = shape.GetDocContent && shape.GetDocContent()
            var para = box && box.GetElement ? box.GetElement(0) : null
            if (para && para.AddText) para.AddText(s.text || "")
            if (slide.AddObject) slide.AddObject(shape)
            return true
          }
          if (typeof Api.GetDocument !== "function") return false
          var doc = Api.GetDocument()
          var lines = String(s.text || "").split("\n")
          if (!lines.length) lines = [""]
          var paras = []
          var i
          for (i = 0; i < lines.length; i++) {
            var run = typeof Api.CreateRun === "function" ? Api.CreateRun() : null
            if (run) {
              run.AddText(lines[i])
              if (s.bold === true) run.SetBold(true)
              if (s.italic === true) run.SetItalic(true)
              if (s.underline === true && run.SetUnderline) run.SetUnderline("single")
              if (s.strike === true && run.SetStrikeout) run.SetStrikeout(true)
              if (s.font && run.SetFontFamily) run.SetFontFamily(s.font)
              if (s.size && run.SetFontSize) run.SetFontSize(Number(s.size) * 2)
              if (s.color && run.SetColor) run.SetColor(s.color.r, s.color.g, s.color.b)
              if (s.highlight && run.SetHighlight) run.SetHighlight(s.highlight)
            }
            var para = Api.CreateParagraph()
            if (run && para.AddElement) para.AddElement(run)
            else if (para.AddText) para.AddText(lines[i])
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
          if (doc.InsertContent) doc.InsertContent(paras)
          return true
        }, function (ok, error) {
          if (error) reply(requestId, null, error)
          else reply(requestId, { inserted: insert.length, ok: ok !== false })
        })
        return
      }

      if (type === "replace_selection") {
        runMethod("PasteText", [normalizeNewlines(payload.text || "")], function (_value, error) {
          if (error) reply(requestId, null, error)
          else reply(requestId, { replaced: true })
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
        setScope(Object.assign({ kind: type }, payload))
        if (payload.target === "all" && type === "select") {
          runMethod("SelectAll", [], function (_value, error) {
            if (error) reply(requestId, null, error)
            else reply(requestId, { selected: "all" })
          })
          return
        }
        callDoc(function () {
          var s = Asc.scope
          var doc = Api.GetDocument()
          var index = s.index
          if (s.heading != null && index == null) {
            var needle = String(s.heading).toLowerCase()
            var n = doc.GetElementsCount()
            var i
            for (i = 0; i < n; i++) {
              var el = doc.GetElement(i)
              var text = el && el.GetText ? String(el.GetText()).toLowerCase() : ""
              if (text.indexOf(needle) >= 0) {
                index = i
                break
              }
            }
          }
          if (index == null) index = 0
          var target = doc.GetElement(Number(index))
          if (!target) return false
          if (target.Select) target.Select()
          return { index: Number(index) }
        }, function (info, error) {
          if (error) reply(requestId, null, error)
          else reply(requestId, info || { ok: true })
        })
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
        setScope(payload)
        callDoc(function () {
          var s = Asc.scope
          if (typeof Api.GetActiveSheet !== "function") return { error: "Not a spreadsheet." }
          var sheet = Api.GetActiveSheet()
          if (s.sheet && Api.GetSheet) {
            var named = Api.GetSheet(s.sheet)
            if (named) sheet = named
          }
          if (s.sheet && sheet.SetName && type === "goto_cell" && s.rename) sheet.SetName(s.sheet)
          var addr = s.range || s.cell || "A1"
          var range = sheet.GetRange(addr)
          if (!range) return { error: "Range not found." }
          if (type === "get_range") {
            return {
              sheet: sheet.GetName ? sheet.GetName() : s.sheet,
              range: addr,
              values: range.GetValue ? range.GetValue() : null
            }
          }
          if (type === "set_range") {
            var values = s.values != null ? s.values : s.value
            if (range.SetValue) range.SetValue(values)
            return { sheet: sheet.GetName ? sheet.GetName() : s.sheet, range: addr, wrote: true }
          }
          if (range.Select) range.Select()
          return { sheet: sheet.GetName ? sheet.GetName() : s.sheet, range: addr }
        }, function (info) {
          if (info && info.error) reply(requestId, null, info.error)
          else reply(requestId, info || { ok: true })
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
      // plugin_init never arrived — re-request the editor handshake.
      var guid = window.Asc.plugin.guid
      if (!guid) return
      var payload = { type: "initialize", guid: guid }
      try {
        if (window.Asc.plugin.windowID) payload.windowID = window.Asc.plugin.windowID
      } catch (err) {}
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
