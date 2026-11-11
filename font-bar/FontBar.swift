// 🔤 Font bar, native side: a small panel stuck in PowerPoint's title bar, right of
// the deck's name, with the selection's font name and size — editable.
// Child of server.js (font-bar.js), which feeds it JSON lines on stdin
//   {"panes":[{"pane","doc","name","size","at"}]}
// and reads the edits from stdout
//   {"pane":…, "name":…} | {"pane":…, "size":…}
// No permissions needed: window frames come from CGWindowList, which gives
// bounds (not titles) to everyone. With Screen Recording it also reads the
// window titles and matches the deck exactly; without, it shows the pane that
// reported last, which is the deck you're clicking around in.
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

struct PaneState: Decodable {
  let pane: String
  let doc: String
  let name: String?
  let size: Double?
  let at: Double
}
struct Update: Decodable { let panes: [PaneState] }

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
  let listWidth: CGFloat
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
    self.listWidth = listWidth
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
    let cell = tableView.makeView(withIdentifier: id, owner: nil) as? NSTableCellView ?? {
      let cell = NSTableCellView()
      cell.identifier = id
      let label = NSTextField(labelWithString: "")
      label.lineBreakMode = .byTruncatingTail
      label.translatesAutoresizingMaskIntoConstraints = false
      cell.addSubview(label)
      cell.textField = label
      NSLayoutConstraint.activate([
        label.leadingAnchor.constraint(equalTo: cell.leadingAnchor, constant: 10),
        label.trailingAnchor.constraint(equalTo: cell.trailingAnchor, constant: -10),
        label.centerYAnchor.constraint(equalTo: cell.centerYAnchor),
      ])
      return cell
    }()
    let item = shownItems[row]
    cell.textField?.stringValue = item
    cell.textField?.font = preview ? previewFont(item) : fieldFont
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
  var panes: [PaneState] = []
  var shown: PaneState?
  var timer: Timer?

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
    size.field.placeholderString = "–"
    name.onPick = { [weak self] family in self?.send(["name": family]) }
    size.onPick = { [weak self] text in
      guard let value = Double(text.replacingOccurrences(of: ",", with: ".")), value >= 1, value <= 4000 else {
        self?.giveBackFocus(); return
      }
      self?.send(["size": (value * 2).rounded() / 2])   // PowerPoint keeps half points
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
    guard let app = powerPoint, let (frame, title) = frontWindow(pid: app.processIdentifier) else {
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
    show(pickPane(title: title))
    if !panel.isVisible { panel.orderFrontRegardless() }
  }

  func pickPane(title: String?) -> PaneState? {
    if let title, let match = panes.filter({ !$0.doc.isEmpty && title.hasPrefix($0.doc) }).max(by: { $0.at < $1.at }) {
      return match
    }
    return panes.max(by: { $0.at < $1.at })
  }

  // MARK: Showing and editing

  func show(_ pane: PaneState?) {
    shown = pane
    name.field.isEnabled = pane != nil
    size.field.isEnabled = pane != nil
    name.field.placeholderString = pane == nil ? "Victor Tools closed" : "Font"
    if panel.isKeyWindow { return }   // don't overwrite what's being typed
    let nameText = pane?.name ?? ""
    let sizeText = pane?.size.map(formatSize) ?? ""
    if name.field.stringValue != nameText { name.field.stringValue = nameText }
    if size.field.stringValue != sizeText { size.field.stringValue = sizeText }
  }

  func send(_ change: [String: Any]) {
    if let pane = shown {
      var cmd = change
      cmd["pane"] = pane.pane
      if let data = try? JSONSerialization.data(withJSONObject: cmd), let line = String(data: data, encoding: .utf8) {
        FileHandle.standardOutput.write((line + "\n").data(using: .utf8)!)
      }
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

  // MARK: Input from the server

  func readStdin() {
    var buffer = Data()
    FileHandle.standardInput.readabilityHandler = { handle in
      let chunk = handle.availableData
      if chunk.isEmpty { exit(0) }   // the server is gone
      buffer.append(chunk)
      while let nl = buffer.firstIndex(of: 0x0A) {
        let line = buffer.subdata(in: buffer.startIndex..<nl)
        buffer.removeSubrange(buffer.startIndex...nl)
        if let update = try? JSONDecoder().decode(Update.self, from: line) {
          DispatchQueue.main.async { self.panes = update.panes; if self.powerPoint != nil { self.follow() } }
        }
      }
    }
  }
}

// `font-bar --preview DIR dark|light [typed text]`: draws the bar off screen,
// list open, and saves bar.png + list.png in DIR — to iterate on the look
// without touching the mouse or PowerPoint.
func preview(_ bar: FontBar, dir: String, typed: String?) {
  bar.panes = [PaneState(pane: "p", doc: "Deck", name: "Calibri", size: 28, at: 1)]
  bar.panel.setFrame(NSRect(x: -3000, y: 200, width: barWidth, height: barHeight), display: true)
  bar.show(bar.panes.first)
  bar.panel.orderFrontRegardless()
  // `typed` = "size" opens the size list instead.
  let picker = typed == "size" ? bar.size : bar.name
  picker.open()
  if let typed, typed != "size" {
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

registerOfficeFonts()
let app = NSApplication.shared
app.setActivationPolicy(.accessory)
let args = CommandLine.arguments
if let i = args.firstIndex(of: "--preview") {
  app.appearance = NSAppearance(named: args.count > i + 2 && args[i + 2] == "light" ? .aqua : .darkAqua)
}
let bar = FontBar()
if let i = args.firstIndex(of: "--preview") {
  DispatchQueue.main.async { preview(bar, dir: args[i + 1], typed: args.count > i + 3 ? args[i + 3] : nil) }
}
app.run()
