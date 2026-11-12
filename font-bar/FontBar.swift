// 🔤 Font bar, native side: a small panel stuck in PowerPoint's title bar, right of
// the deck's name, with the selection's font name and size — editable.
// Child of server.js (font-bar.js), which only keeps it alive; it exits when its
// stdin closes (the server is gone).
// It asks PowerPoint itself, over AppleScript, what the selection's font is and
// sets it the same way — any deck, Victor Tools open or not. That takes the
// Automation permission (System Settings › Privacy › Automation), asked once.
// Window frames come from CGWindowList, which gives them to everyone.
import AppKit

let powerPointBundle = "com.microsoft.Powerpoint"
let barWidth: CGFloat = 228
let barHeight: CGFloat = 26
// Measured on PowerPoint 16.10x: the title row's middle is 20 pt below the
// window's top; the comments icon starts ~165 pt left of its right edge.
let titleRowMiddle: CGFloat = 20
let rightInset: CGFloat = 180
let fieldFont = NSFont.systemFont(ofSize: 14)
let sizes: [Double] = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96]

// The font of the selection: nil name/size when mixed.
struct Selection: Equatable {
  let name: String?
  let size: Double?
  let sample: String?   // the selected text, or the word at the caret
}

// PowerPoint over AppleScript, on a queue of its own: an Apple Event waits for
// PowerPoint, which can be busy (a dialog, a save) — never on the main thread.
// The selection: the selected text if there is some; otherwise (a caret, or
// shapes selected as objects) all the text of the selected shapes.
final class PowerPointScript {
  private let queue = DispatchQueue(label: "font-bar.applescript")
  private var reader: NSAppleScript?
  private var lastError: String?
  private static let US = "\u{1F}", RS = "\u{1E}"

  // Names clash with PowerPoint's terms (full, out, offset…): every variable is v-something.
  private static let readSource = """
    on fontOf(vRange)
      tell application "Microsoft PowerPoint"
        set vFont to font of vRange
        set vName to font name of vFont
        set vSize to font size of vFont
      end tell
      if vName is missing value then set vName to ""
      if vSize is missing value then set vSize to 0
      return vName & (character id 31) & (vSize as text)
    end fontOf
    set vUS to character id 31
    set vRS to character id 30
    with timeout of 2 seconds
      tell application "Microsoft PowerPoint"
        if (count of document windows) = 0 then return "none"
        set vSel to selection of active window
        set vType to selection type of vSel
        if vType is selection type text then
          set vRange to text range of vSel
          set vProps to properties of vRange
          set vWhole to ""
          try
            set vWhole to content of text range of text frame of shape 1 of shape range of vSel
          end try
          return "text" & vRS & my fontOf(vRange) & vRS & ((offset of vProps) as text) & vUS & ((text length of vProps) as text) & vRS & vWhole
        else if vType is selection type shapes then
          set vOut to "shapes"
          set vShapes to shape range of vSel
          repeat with i from 1 to count of shapes of vShapes
            try
              set vShape to shape i of vShapes
              if has text of text frame of vShape then set vOut to vOut & vRS & my fontOf(text range of text frame of vShape)
            end try
          end repeat
          return vOut
        else
          return "other"
        end if
      end tell
    end timeout
    """

  // The selection's font, or nil when there's no text to show (no deck, a
  // picture, the slide sorter, PowerPoint busy).
  func read(_ done: @escaping (Selection?) -> Void) {
    queue.async {
      if self.reader == nil {
        self.reader = NSAppleScript(source: Self.readSource)
        var err: NSDictionary?
        self.reader?.compileAndReturnError(&err)
        if let err { log("read script does not compile: \(err)") }
      }
      var err: NSDictionary?
      let out = self.reader?.executeAndReturnError(&err).stringValue
      // Once per new error, not every 0.8 s; -1712 is PowerPoint busy, -1728 nothing selected.
      let code = err?[NSAppleScript.errorNumber] as? Int
      let message = err.map { "read: \($0[NSAppleScript.errorMessage] ?? $0)" }
      if message != self.lastError, code != -1712, code != -1728 { message.map(log) }
      self.lastError = message
      let sel = out.flatMap(Self.parse)
      DispatchQueue.main.async { done(sel) }
    }
  }

  static func parse(_ out: String) -> Selection? {
    let records = out.components(separatedBy: RS)
    func font(_ r: String) -> (String?, Double?) {
      let f = r.components(separatedBy: US)
      let name = f.first.flatMap { $0.isEmpty ? nil : $0 }
      let size = f.count > 1 ? Double(f[1].replacingOccurrences(of: ",", with: ".")).flatMap { $0 > 0 ? $0 : nil } : nil
      return (name, size)
    }
    switch records.first {
    case "text" where records.count >= 4:
      let (name, size) = font(records[1])
      let pos = records[2].components(separatedBy: US).compactMap { Int($0) }
      let whole = records[3...].joined(separator: RS)
      return Selection(name: name, size: size, sample: pos.count == 2 ? sample(whole, offset: pos[0], length: pos[1]) : nil)
    case "shapes" where records.count >= 2:
      let fonts = records.dropFirst().map(font)
      return Selection(name: common(fonts.map { $0.0 }), size: common(fonts.map { $0.1 }), sample: nil)
    default:
      return nil
    }
  }

  // One value when they all agree, nil when mixed or missing.
  static func common<T: Hashable>(_ values: [T?]) -> T? {
    let set = Set(values)
    return set.count == 1 ? set.first! : nil
  }

  // What the font list draws in each font: the selected text (first line), or
  // the word the caret is in or next to. PowerPoint counts from 1, in UTF-16 units.
  static func sample(_ whole: String, offset: Int, length: Int) -> String? {
    let text = whole as NSString
    let start = max(0, min(text.length, offset - 1))
    var picked: String
    if length > 0 {
      picked = text.substring(with: NSRange(location: start, length: min(length, text.length - start)))
    } else {
      let word = try! NSRegularExpression(pattern: "[\\p{L}\\p{N}_'’-]")
      func isWord(_ i: Int) -> Bool {
        word.firstMatch(in: whole, range: NSRange(location: i, length: 1)) != nil
      }
      var a = start, b = start
      while a > 0 && isWord(a - 1) { a -= 1 }
      while b < text.length && isWord(b) { b += 1 }
      picked = text.substring(with: NSRange(location: a, length: b - a))
    }
    let line = picked.components(separatedBy: CharacterSet(charactersIn: "\r\n\u{0B}"))
      .map { $0.trimmingCharacters(in: .whitespaces) }.first { !$0.isEmpty } ?? ""
    if line.isEmpty { return nil }
    return line.count > 24 ? String(line.prefix(23)) + "…" : line
  }

  // Sets the font name or size on the selection, as `read` sees it.
  func apply(name: String? = nil, size: Double? = nil, then done: @escaping () -> Void) {
    let property: String
    if let name {
      let quoted = name.replacingOccurrences(of: "\\", with: "\\\\").replacingOccurrences(of: "\"", with: "\\\"")
      property = "font name of font of vRange to \"\(quoted)\""
    } else if let size {
      property = "font size of font of vRange to \(size)"
    } else { return }
    let source = """
      with timeout of 5 seconds
        tell application "Microsoft PowerPoint"
          set vSel to selection of active window
          if selection type of vSel is selection type text then
            set vRange to text range of vSel
            set \(property)
          else if selection type of vSel is selection type shapes then
            set vShapes to shape range of vSel
            repeat with i from 1 to count of shapes of vShapes
              try
                set vShape to shape i of vShapes
                if has text of text frame of vShape then
                  set vRange to text range of text frame of vShape
                  set \(property)
                end if
              end try
            end repeat
          end if
        end tell
      end timeout
      """
    queue.async {
      var err: NSDictionary?
      NSAppleScript(source: source)?.executeAndReturnError(&err)
      if let err { log("apply: \(err[NSAppleScript.errorMessage] ?? err)") }
      DispatchQueue.main.async(execute: done)
    }
  }
}

func log(_ message: String) {
  FileHandle.standardError.write("\(Date()) \(message)\n".data(using: .utf8)!)
}

final class BarPanel: NSPanel {
  override var canBecomeKey: Bool { true }
  override var canBecomeMain: Bool { false }
}

// The drop-down list never takes the keyboard: typing stays in the field.
final class ListPanel: NSPanel {
  override var canBecomeKey: Bool { false }
  override var canBecomeMain: Bool { false }
}

// Rows take the very first click, though their panel is never key.
final class ClickTable: NSTableView {
  override func acceptsFirstMouse(for event: NSEvent?) -> Bool { true }
}

// The field's look, drawn by hand: the system bezel's focus ring spills out of
// the 26 pt panel and gets cut into stray pixels. A soft rounded fill like
// PowerPoint's own boxes; an accent border only while editing.
final class FieldBox: NSView {
  var editing = false { didSet { needsDisplay = true } }
  override var wantsUpdateLayer: Bool { true }
  override init(frame: NSRect) {
    super.init(frame: frame)
    wantsLayer = true
    layer?.cornerRadius = 6
    layer?.cornerCurve = .continuous
  }
  required init?(coder: NSCoder) { fatalError() }
  override func updateLayer() {
    effectiveAppearance.performAsCurrentDrawingAppearance {
      let dark = effectiveAppearance.bestMatch(from: [.darkAqua, .aqua]) == .darkAqua
      layer?.backgroundColor = (dark ? NSColor.white.withAlphaComponent(0.09) : NSColor.black.withAlphaComponent(0.06)).cgColor
      layer?.borderColor = NSColor.controlAccentColor.cgColor
      layer?.borderWidth = editing ? 1.5 : 0
    }
  }
}

// A field that opens its list on click and selects all, ready to type over.
final class PickerField: NSTextField {
  var onClick: (() -> Void)?
  override func mouseDown(with event: NSEvent) {
    super.mouseDown(with: event)   // starts editing; returns at mouse up
    onClick?()
  }
}

// A list row: the item on the left; on the right, the selection's own text
// drawn in that row's font — to see how it would look before picking.
final class RowCell: NSTableCellView {
  let sampleField = NSTextField(labelWithString: "")
  override init(frame: NSRect) {
    super.init(frame: frame)
    let label = NSTextField(labelWithString: "")
    label.lineBreakMode = .byTruncatingTail
    label.setContentCompressionResistancePriority(.defaultLow, for: .horizontal)
    sampleField.lineBreakMode = .byTruncatingTail
    sampleField.textColor = .secondaryLabelColor
    sampleField.alignment = .right
    sampleField.setContentCompressionResistancePriority(.defaultHigh, for: .horizontal)
    for v in [label, sampleField] { v.translatesAutoresizingMaskIntoConstraints = false; addSubview(v) }
    textField = label
    NSLayoutConstraint.activate([
      label.leadingAnchor.constraint(equalTo: leadingAnchor, constant: 10),
      label.centerYAnchor.constraint(equalTo: centerYAnchor),
      sampleField.trailingAnchor.constraint(equalTo: trailingAnchor, constant: -10),
      sampleField.centerYAnchor.constraint(equalTo: centerYAnchor),
      sampleField.leadingAnchor.constraint(greaterThanOrEqualTo: label.trailingAnchor, constant: 12),
      sampleField.widthAnchor.constraint(lessThanOrEqualTo: widthAnchor, multiplier: 0.5),
    ])
  }
  required init?(coder: NSCoder) { fatalError() }
}

// Decoration only: clicks go through to the field under it.
final class Chevron: NSImageView {
  override func hitTest(_ point: NSPoint) -> NSView? { nil }
}

func formatSize(_ s: Double) -> String {
  s == s.rounded() ? String(Int(s)) : String(s)
}

// A text field with a searchable drop-down: click → the whole list, current value
// highlighted; type → only the items containing it (those starting with it first);
// ↑/↓ move, Return or a click picks, Escape gives up.
final class Picker: NSObject, NSTextFieldDelegate, NSTableViewDataSource, NSTableViewDelegate {
  let box = FieldBox()
  let field = PickerField()
  let items: [String]
  let preview: Bool               // each font name drawn in its own font
  let baseListWidth: CGFloat
  var listWidth: CGFloat { sample.isEmpty ? baseListWidth : baseListWidth + 120 }
  // Text drawn on the right of each row, in that row's font (font list only).
  var sample = "" {
    didSet { if sample != oldValue && isOpen { table.reloadData(); place() } }
  }
  // Typed text that wins over the highlighted row (a size that isn't in the list).
  let typedValue: (String) -> String?
  var onPick: (String) -> Void = { _ in }
  var onCancel: () -> Void = {}
  private var shownItems: [String] = []
  let list = ListPanel(contentRect: NSRect(x: 0, y: 0, width: 200, height: 200), styleMask: [.nonactivatingPanel, .borderless],
                               backing: .buffered, defer: true)
  private let table = ClickTable()
  private let rowHeight: CGFloat
  private var fonts: [String: NSFont] = [:]

  init(items: [String], preview: Bool, listWidth: CGFloat, typedValue: @escaping (String) -> String? = { _ in nil }) {
    self.items = items
    self.preview = preview
    self.baseListWidth = listWidth
    self.typedValue = typedValue
    rowHeight = preview ? 26 : 22
    super.init()
    field.font = fieldFont
    field.isBezeled = false
    field.drawsBackground = false
    field.focusRingType = .none
    field.usesSingleLineMode = true
    field.lineBreakMode = .byTruncatingTail
    field.delegate = self
    field.onClick = { [weak self] in self?.open() }
    let chevron = Chevron(image: NSImage(systemSymbolName: "chevron.down", accessibilityDescription: nil)!)
    chevron.contentTintColor = .secondaryLabelColor
    chevron.symbolConfiguration = .init(pointSize: 10, weight: .semibold)
    chevron.translatesAutoresizingMaskIntoConstraints = false
    field.translatesAutoresizingMaskIntoConstraints = false
    box.addSubview(field)
    box.addSubview(chevron)
    NSLayoutConstraint.activate([
      chevron.trailingAnchor.constraint(equalTo: box.trailingAnchor, constant: -8),
      chevron.centerYAnchor.constraint(equalTo: box.centerYAnchor),
      field.leadingAnchor.constraint(equalTo: box.leadingAnchor, constant: 6),
      field.trailingAnchor.constraint(equalTo: chevron.leadingAnchor, constant: -4),
      field.centerYAnchor.constraint(equalTo: box.centerYAnchor),
    ])

    let column = NSTableColumn(identifier: .init("item"))
    table.addTableColumn(column)
    table.headerView = nil
    table.rowHeight = rowHeight
    table.intercellSpacing = .zero
    table.backgroundColor = .clear
    table.style = .inset
    table.dataSource = self
    table.delegate = self
    table.target = self
    table.action = #selector(rowClicked)
    let scroll = NSScrollView()
    scroll.documentView = table
    scroll.hasVerticalScroller = true
    scroll.drawsBackground = false
    scroll.autoresizingMask = [.width, .height]
    let background = NSVisualEffectView(frame: NSRect(x: 0, y: 0, width: 200, height: 200))
    background.material = .menu
    background.state = .active
    background.wantsLayer = true
    background.layer?.cornerRadius = 6
    background.layer?.masksToBounds = true
    background.autoresizingMask = [.width, .height]
    scroll.frame = background.bounds
    background.addSubview(scroll)
    list.contentView = background
    list.isOpaque = false
    list.backgroundColor = .clear
    list.hasShadow = true
    list.level = .popUpMenu
    list.hidesOnDeactivate = false
  }

  var isOpen: Bool { list.isVisible }

  func open() {
    guard field.isEnabled, !isOpen else { return }
    box.editing = true
    shownItems = items
    table.reloadData()
    field.currentEditor()?.selectAll(nil)
    place()
    if let i = items.firstIndex(where: { $0.caseInsensitiveCompare(field.stringValue) == .orderedSame }) {
      highlight(i, center: true)
    } else {
      table.deselectAll(nil)
      table.scrollRowToVisible(0)
    }
  }

  func close() {
    list.orderOut(nil)
    box.editing = false
  }

  private func place() {
    guard let window = field.window else { return }
    let anchor = window.convertToScreen(box.convert(box.bounds, to: nil))
    // The inset style pads the rows top and bottom: measure, don't multiply.
    table.layoutSubtreeIfNeeded()
    let rows = min(16, shownItems.count)
    let pad = shownItems.isEmpty ? 0 : table.rect(ofRow: 0).minY
    let height = rows == 0 ? rowHeight : table.rect(ofRow: rows - 1).maxY + pad
    var x = anchor.minX
    if let screen = window.screen, x + listWidth > screen.visibleFrame.maxX - 4 {
      x = screen.visibleFrame.maxX - 4 - listWidth
    }
    list.setFrame(NSRect(x: x, y: anchor.minY - height - 3, width: listWidth, height: height), display: true)
    if shownItems.isEmpty { list.orderOut(nil) } else { list.orderFront(nil) }
  }

  private func highlight(_ row: Int, center: Bool = false) {
    guard row >= 0, row < shownItems.count else { return }
    table.selectRowIndexes([row], byExtendingSelection: false)
    if center {
      let visible = Int(table.enclosingScrollView!.contentSize.height / rowHeight)
      table.scrollRowToVisible(min(shownItems.count - 1, row + visible / 2))
      table.scrollRowToVisible(max(0, row - visible / 2))
    } else {
      table.scrollRowToVisible(row)
    }
  }

  func controlTextDidChange(_ obj: Notification) {
    let q = field.stringValue.trimmingCharacters(in: .whitespaces).lowercased()
    if q.isEmpty {
      shownItems = items
    } else {
      let starts = items.filter { $0.lowercased().hasPrefix(q) }
      let contains = items.filter { !$0.lowercased().hasPrefix(q) && $0.lowercased().contains(q) }
      shownItems = starts + contains
    }
    table.reloadData()
    place()
    if shownItems.isEmpty { table.deselectAll(nil) } else { highlight(0) }
  }

  func control(_ control: NSControl, textView: NSTextView, doCommandBy selector: Selector) -> Bool {
    switch selector {
    case #selector(NSResponder.moveDown(_:)):
      if !isOpen { open() } else { highlight(min(shownItems.count - 1, table.selectedRow + 1)) }
    case #selector(NSResponder.moveUp(_:)):
      highlight(max(0, table.selectedRow - 1))
    case #selector(NSResponder.insertNewline(_:)):
      let typed = typedValue(field.stringValue)
      let row = table.selectedRow
      let value = typed ?? (isOpen && row >= 0 && row < shownItems.count ? shownItems[row] : field.stringValue)
      close()
      onPick(value)
    case #selector(NSResponder.cancelOperation(_:)):
      close()
      onCancel()
    default:
      return false
    }
    return true
  }

  @objc private func rowClicked() {
    let row = table.clickedRow
    guard row >= 0, row < shownItems.count else { return }
    close()
    onPick(shownItems[row])
  }

  func numberOfRows(in tableView: NSTableView) -> Int { shownItems.count }

  func tableView(_ tableView: NSTableView, viewFor tableColumn: NSTableColumn?, row: Int) -> NSView? {
    let id = NSUserInterfaceItemIdentifier("cell")
    let cell = tableView.makeView(withIdentifier: id, owner: nil) as? RowCell ?? {
      let cell = RowCell()
      cell.identifier = id
      return cell
    }()
    let item = shownItems[row]
    cell.textField?.stringValue = item
    cell.textField?.font = preview ? previewFont(item) : fieldFont
    let showSample = preview && !sample.isEmpty
    cell.sampleField.isHidden = !showSample
    if showSample {
      cell.sampleField.stringValue = sample
      // The real face even for symbol fonts: Wingdings shows what Wingdings would do.
      cell.sampleField.font = NSFontManager.shared.font(withFamily: item, traits: [], weight: 5, size: 15) ?? fieldFont
    }
    return cell
  }

  // The family in its own face; symbol fonts (Wingdings, Webdings…) can't spell
  // their own name, so those stay in the system font.
  private func previewFont(_ family: String) -> NSFont {
    if let font = fonts[family] { return font }
    var font = fieldFont
    if let f = NSFontManager.shared.font(withFamily: family, traits: [], weight: 5, size: 15),
       f.coveredCharacterSet.isSuperset(of: CharacterSet(charactersIn: family)) {
      font = f
    }
    fonts[family] = font
    return font
  }
}

final class FontBar: NSObject {
  let panel = BarPanel(contentRect: NSRect(x: 0, y: 0, width: barWidth, height: barHeight),
                       styleMask: [.nonactivatingPanel, .borderless], backing: .buffered, defer: false)
  let name = Picker(items: NSFontManager.shared.availableFontFamilies
                      .filter { !$0.hasPrefix(".") }
                      .sorted { $0.localizedCaseInsensitiveCompare($1) == .orderedAscending },
                    preview: true, listWidth: 280)
  let size = Picker(items: sizes.map(formatSize), preview: false, listWidth: 100) { text in
    // A typed number goes as typed, in or out of the list.
    Double(text.trimmingCharacters(in: .whitespaces).replacingOccurrences(of: ",", with: ".")).map(formatSize)
  }
  let powerPointScript = PowerPointScript()
  var shown: Selection?
  var timer: Timer?
  var reading = false
  var lastRead = Date.distantPast

  override init() {
    super.init()
    panel.isOpaque = false
    panel.backgroundColor = .clear
    panel.hasShadow = false
    panel.level = .floating
    panel.hidesOnDeactivate = false
    panel.collectionBehavior = [.fullScreenNone, .moveToActiveSpace, .ignoresCycle]

    name.field.toolTip = "Font of the selection (Victor Tools). Click, type to search, Return or click to apply."
    size.field.toolTip = "Font size of the selection (Victor Tools). Click, type a size, Return to apply."
    name.onPick = { [weak self] family in self?.send(name: family) }
    size.onPick = { [weak self] text in
      guard let value = Double(text.replacingOccurrences(of: ",", with: ".")), value >= 1, value <= 4000 else {
        self?.giveBackFocus(); return
      }
      self?.send(size: (value * 2).rounded() / 2)   // PowerPoint keeps half points
    }
    name.onCancel = { [weak self] in self?.giveBackFocus() }
    size.onCancel = { [weak self] in self?.giveBackFocus() }

    let stack = NSStackView(views: [name.box, size.box])
    stack.spacing = 4
    stack.alignment = .centerY
    stack.frame = panel.contentView!.bounds
    stack.autoresizingMask = [.width, .height]
    NSLayoutConstraint.activate([
      name.box.widthAnchor.constraint(equalToConstant: 152),
      size.box.widthAnchor.constraint(equalToConstant: 72),
      name.box.heightAnchor.constraint(equalToConstant: barHeight),
      size.box.heightAnchor.constraint(equalToConstant: barHeight),
    ])
    panel.contentView!.addSubview(stack)
    // Nothing focused until clicked: otherwise the font name shows up selected.
    panel.initialFirstResponder = panel.contentView
    panel.makeFirstResponder(nil)

    // Clicking back on the slide: drop the lists and what was half typed.
    NotificationCenter.default.addObserver(forName: NSWindow.didResignKeyNotification, object: panel, queue: .main) {
      [weak self] _ in
      guard let self else { return }
      self.name.close()
      self.size.close()
      self.panel.makeFirstResponder(nil)
      self.show(self.shown)
    }
    NSWorkspace.shared.notificationCenter.addObserver(
      self, selector: #selector(appChanged), name: NSWorkspace.didActivateApplicationNotification, object: nil)
    // A click is what usually changes the selection: read it right after, not at
    // the next poll. Mouse events need no permission to watch (keys would).
    NSEvent.addGlobalMonitorForEvents(matching: [.leftMouseUp]) { [weak self] _ in
      DispatchQueue.main.asyncAfter(deadline: .now() + 0.12) { self?.readSelection() }
    }
    if !CommandLine.arguments.contains("--preview") {
      appChanged()
      readStdin()
    }
  }

  // MARK: Following PowerPoint's front window

  var powerPoint: NSRunningApplication? {
    NSWorkspace.shared.frontmostApplication.flatMap { $0.bundleIdentifier == powerPointBundle ? $0 : nil }
  }

  @objc func appChanged() {
    timer?.invalidate()
    timer = nil
    if powerPoint != nil {
      timer = Timer.scheduledTimer(withTimeInterval: 0.12, repeats: true) { [weak self] _ in self?.follow() }
      follow()
    } else if !panel.isKeyWindow {
      hide()
    }
  }

  func hide() {
    name.close()
    size.close()
    panel.orderOut(nil)
  }

  // The front PowerPoint window: its frame (Cocoa coordinates) and title, if readable.
  func frontWindow(pid: pid_t) -> (frame: NSRect, title: String?)? {
    guard let list = CGWindowListCopyWindowInfo([.optionOnScreenOnly, .excludeDesktopElements], kCGNullWindowID)
            as? [[String: Any]] else { return nil }
    for w in list {
      guard (w[kCGWindowOwnerPID as String] as? pid_t) == pid,
            (w[kCGWindowLayer as String] as? Int) == 0,
            let b = w[kCGWindowBounds as String] as? [String: CGFloat],
            let x = b["X"], let y = b["Y"], let width = b["Width"], let height = b["Height"] else { continue }
      if width < 500 || height < 300 { continue }   // Find box, dialogs, popovers
      let primaryHeight = NSScreen.screens.first?.frame.height ?? 0
      let frame = NSRect(x: x, y: primaryHeight - y - height, width: width, height: height)
      return (frame, w[kCGWindowName as String] as? String)
    }
    return nil
  }

  func follow() {
    guard let app = powerPoint, let (frame, _) = frontWindow(pid: app.processIdentifier) else {
      if !panel.isKeyWindow { hide() }
      return
    }
    // Slide show or full screen: no title bar to sit in.
    let fullScreen = NSScreen.screens.contains { $0.frame == frame }
    // Too narrow and the bar would sit on the centered title.
    let left = frame.maxX - rightInset - barWidth
    let tooNarrow = left < frame.midX + 190
    if fullScreen || tooNarrow {
      hide()
      return
    }
    let target = NSRect(x: left, y: frame.maxY - titleRowMiddle - barHeight / 2, width: barWidth, height: barHeight)
    if panel.frame != target { panel.setFrame(target, display: false) }
    if !panel.isVisible { panel.orderFrontRegardless(); readSelection() }
    // Typing or the ribbon's own boxes change the font without a click here: poll.
    if Date().timeIntervalSince(lastRead) > readEvery { readSelection() }
  }

  // Each read is ~10 Apple Events (~120 ms of PowerPoint's time): not faster than this.
  let readEvery: TimeInterval = 0.8

  func readSelection() {
    guard !reading, powerPoint != nil, panel.isVisible, !panel.isKeyWindow else { return }
    reading = true
    lastRead = Date()
    powerPointScript.read { [weak self] sel in
      guard let self else { return }
      self.reading = false
      self.lastRead = Date()
      if sel != self.shown || !self.panel.isKeyWindow { self.show(sel) }
    }
  }

  // MARK: Showing and editing

  // nil: nothing with text selected — the boxes stay empty and grey.
  func show(_ sel: Selection?) {
    shown = sel
    name.field.isEnabled = sel != nil
    size.field.isEnabled = sel != nil
    // Hints only when there's text selected (mixed fonts); nothing selected: blank boxes.
    name.field.placeholderString = sel == nil ? nil : "Font"
    size.field.placeholderString = sel == nil ? nil : "–"
    name.sample = sel?.sample ?? ""
    if panel.isKeyWindow { return }   // don't overwrite what's being typed
    let nameText = sel?.name ?? ""
    let sizeText = sel?.size.map(formatSize) ?? ""
    if name.field.stringValue != nameText { name.field.stringValue = nameText }
    if size.field.stringValue != sizeText { size.field.stringValue = sizeText }
  }

  func send(name family: String? = nil, size points: Double? = nil) {
    if shown != nil {
      powerPointScript.apply(name: family, size: points) { [weak self] in self?.readSelection() }
    }
    giveBackFocus()
  }

  // After Return/Escape/a pick, the keyboard goes back to the slide.
  func giveBackFocus() {
    panel.makeFirstResponder(nil)
    panel.resignKey()
    NSRunningApplication.runningApplications(withBundleIdentifier: powerPointBundle).first?.activate()
    show(shown)
  }

  // MARK: The server

  // Nothing comes in; the end of stdin means the server is gone.
  func readStdin() {
    FileHandle.standardInput.readabilityHandler = { handle in
      if handle.availableData.isEmpty { exit(0) }
    }
  }
}

// `font-bar --preview DIR dark|light [typed text]`: draws the bar off screen,
// list open, and saves bar.png + list.png in DIR — to iterate on the look
// without touching the mouse or PowerPoint.
func preview(_ bar: FontBar, dir: String, typed: String?) {
  bar.panel.setFrame(NSRect(x: -3000, y: 200, width: barWidth, height: barHeight), display: true)
  // `typed` = "empty": nothing with text selected.
  bar.show(typed == "empty" ? nil : Selection(name: "Calibri", size: 28, sample: "Token ROI"))
  bar.panel.orderFrontRegardless()
  // `typed` = "size" opens the size list instead.
  let picker = typed == "size" ? bar.size : bar.name
  picker.open()
  if let typed, typed != "size", typed != "empty" {
    picker.field.stringValue = typed
    picker.controlTextDidChange(Notification(name: NSControl.textDidChangeNotification))
  }
  DispatchQueue.main.asyncAfter(deadline: .now() + 0.6) {
    for (window, file) in [(bar.panel as NSWindow, "bar.png"), (picker.list as NSWindow, "list.png")] {
      guard let view = window.contentView,
            let rep = view.bitmapImageRepForCachingDisplay(in: view.bounds) else {
        FileHandle.standardError.write("preview: \(file) not drawn (visible: \(window.isVisible))\n".data(using: .utf8)!)
        continue
      }
      view.cacheDisplay(in: view.bounds, to: rep)
      let png = rep.representation(using: .png, properties: [:])!
      do { try png.write(to: URL(fileURLWithPath: dir).appendingPathComponent(file)) }
      catch { FileHandle.standardError.write("preview: \(error)\n".data(using: .utf8)!) }
    }
    exit(0)
  }
}

// Office keeps its own fonts (Calibri, Aptos, Consolas…) inside its bundle, out
// of the system's sight: register them for this process, so they are listed —
// and drawn in their own face — like in PowerPoint's Font box.
func registerOfficeFonts() {
  let dir = URL(fileURLWithPath: "/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts")
  let urls = (try? FileManager.default.contentsOfDirectory(at: dir, includingPropertiesForKeys: nil)) ?? []
  let fonts = urls.filter { ["ttf", "otf", "ttc"].contains($0.pathExtension.lowercased()) }
  CTFontManagerRegisterFontURLs(fonts as CFArray, .process, true, nil)
}

// Started by node (server.js), macOS would ask *node* for the Automation
// permission — and a background node never gets the prompt, only a silent "not
// authorized". So the bar first replaces itself (same pid, same stdin) with a
// copy that answers for itself, as `font-bar` (Info.plist inside the binary).
@_silgen_name("responsibility_spawnattrs_setdisclaim")
func responsibility_spawnattrs_setdisclaim(_ attr: UnsafeMutablePointer<posix_spawnattr_t?>, _ disclaim: Int32) -> Int32

func disclaimResponsibility() {
  guard ProcessInfo.processInfo.environment["FONT_BAR_DISCLAIMED"] == nil else { return }
  setenv("FONT_BAR_DISCLAIMED", "1", 1)
  var attr: posix_spawnattr_t?
  posix_spawnattr_init(&attr)
  posix_spawnattr_setflags(&attr, Int16(POSIX_SPAWN_SETEXEC))
  let disclaimed = responsibility_spawnattrs_setdisclaim(&attr, 1)
  if disclaimed != 0 { log("responsibility_spawnattrs_setdisclaim: \(disclaimed)") }
  let argv = CommandLine.unsafeArgv
  var pid: pid_t = 0
  let rc = posix_spawn(&pid, Bundle.main.executablePath ?? CommandLine.arguments[0], nil, &attr, argv, environ)
  log("could not disclaim responsibility (\(rc)), going on as node's")   // SETEXEC returns only on failure
}

disclaimResponsibility()
registerOfficeFonts()
let app = NSApplication.shared
app.setActivationPolicy(.accessory)
let args = CommandLine.arguments
if let i = args.firstIndex(of: "--preview") {
  app.appearance = NSAppearance(named: args.count > i + 2 && args[i + 2] == "light" ? .aqua : .darkAqua)
}
// `font-bar --probe`: prints what it reads from PowerPoint once, and exits.
if args.contains("--probe") {
  PowerPointScript().read { sel in print(sel.map { "\($0)" } ?? "nothing with text selected"); exit(0) }
  app.run()
}
let bar = FontBar()
if let i = args.firstIndex(of: "--preview") {
  DispatchQueue.main.async { preview(bar, dir: args[i + 1], typed: args.count > i + 3 ? args[i + 3] : nil) }
}
app.run()
