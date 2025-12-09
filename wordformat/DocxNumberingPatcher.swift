//
//  DocxNumberingPatcher.swift
//  wordformat
//
//  Created by Assistant on 08.12.2025.
//

import Foundation

/// Post-processing patcher that injects Word-native numbering into a saved DOCX by editing the XML.
struct DocxNumberingPatcher {
    enum PatcherError: Error {
        case unzipFailed
        case documentXMLNotFound
        case rezippingFailed
    }

    /// Represents a paragraph that should be numbered at a given level.
    struct NumberingTarget {
        let paragraphIndex: Int
        let level: Int // 0-based: 0 -> main (1.), 1 -> sub (a)), 2 -> subsub (i))
        let textPrefix: String // First N chars of paragraph for content matching

        init(paragraphIndex: Int, level: Int, textPrefix: String = "") {
            self.paragraphIndex = paragraphIndex
            self.level = level
            self.textPrefix = textPrefix
        }
    }

    /// Apply numbering in place on the saved DOCX file.
    func applyNumbering(to docxURL: URL, targets: [NumberingTarget]) throws {
        Logger.shared.log("Patcher: start for \(docxURL.lastPathComponent) with \(targets.count) targets", category: "PATCH")

        let fm = FileManager.default
        let tempRoot = URL(fileURLWithPath: NSTemporaryDirectory(), isDirectory: true)
        let workDir = tempRoot.appendingPathComponent("wordformat-patch-\(UUID().uuidString)", isDirectory: true)
        try fm.createDirectory(at: workDir, withIntermediateDirectories: true)

        // 1) Unzip
        let unzipOK = runShell("unzip -qq \"\(docxURL.path)\" -d \"\(workDir.path)\"")
        guard unzipOK else { throw PatcherError.unzipFailed }
        Logger.shared.log("Patcher: unzip ok to \(workDir.path)", category: "PATCH")

        let documentXML = workDir.appendingPathComponent("word/document.xml")
        guard fm.fileExists(atPath: documentXML.path) else { throw PatcherError.documentXMLNotFound }

        // 2) Find a safe numId that doesn't conflict with existing numbering
        let numberingXML = workDir.appendingPathComponent("word/numbering.xml")
        var safeNumId = 100  // Start high to avoid conflicts

        if fm.fileExists(atPath: numberingXML.path) {
            Logger.shared.log("Patcher: numbering.xml exists, finding safe numId", category: "PATCH")
            let existingXML = try String(contentsOf: numberingXML, encoding: .utf8)
            safeNumId = findSafeNumId(in: existingXML)

            // Append our numbering definition to existing file
            try appendNumberingDefinition(to: numberingXML, numId: safeNumId, abstractNumId: safeNumId)
        } else {
            Logger.shared.log("Patcher: numbering.xml missing, creating with numId=\(safeNumId)", category: "PATCH")
            let xml = numberingTemplate(numId: safeNumId, abstractNumId: safeNumId)
            try xml.data(using: .utf8)!.write(to: numberingXML)
        }

        Logger.shared.log("Patcher: using numId=\(safeNumId) for our numbering", category: "PATCH")

        // 3) Ensure rels and content type
        try ensureNumberingRelationship(in: workDir)

        // 4) Index-based numbering injection
        try injectNumberingByContent(into: documentXML, targets: targets, numId: safeNumId)

        // 5) Re-zip to same docx
        let parent = docxURL.deletingLastPathComponent()
        let tempOut = parent.appendingPathComponent(docxURL.lastPathComponent + ".tmp")
        let zipCmd = "cd \"\(workDir.path)\" && zip -qq -r \"\(tempOut.path)\" ."
        guard runShell(zipCmd) else { throw PatcherError.rezippingFailed }
        Logger.shared.log("Patcher: rezipped to \(tempOut.path)", category: "PATCH")

        // Replace original
        if fm.fileExists(atPath: docxURL.path) {
            try fm.removeItem(at: docxURL)
        }
        try fm.moveItem(at: tempOut, to: docxURL)
        Logger.shared.log("Patcher: replaced original DOCX with patched version", category: "PATCH")

        // Cleanup
        try? fm.removeItem(at: workDir)
    }

    // MARK: - Helpers

    private func runShell(_ command: String) -> Bool {
        let task = Process()
        task.launchPath = "/bin/zsh"
        task.arguments = ["-lc", command]
        let pipe = Pipe()
        task.standardOutput = pipe
        task.standardError = pipe
        do {
            try task.run()
            task.waitUntilExit()
            if task.terminationStatus != 0 {
                let data = pipe.fileHandleForReading.readDataToEndOfFile()
                let out = String(data: data, encoding: .utf8) ?? ""
                Logger.shared.log("Shell fail (\(task.terminationStatus)): \(out)", category: "PATCH")
                return false
            }
            return true
        } catch {
            Logger.shared.log("Shell exception: \(error.localizedDescription)", category: "PATCH")
            return false
        }
    }

    private func numberingTemplate(numId: Int, abstractNumId: Int) -> String {
        return """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:abstractNum w:abstractNumId="\(abstractNumId)">
            <w:lvl w:ilvl="0">
              <w:start w:val="1"/>
              <w:numFmt w:val="decimal"/>
              <w:lvlText w:val="%1."/>
              <w:lvlJc w:val="left"/>
              <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
            </w:lvl>
            <w:lvl w:ilvl="1">
              <w:start w:val="1"/>
              <w:numFmt w:val="lowerLetter"/>
              <w:lvlText w:val="(%2)"/>
              <w:lvlRestart w:val="1"/>
              <w:lvlJc w:val="left"/>
              <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>
            </w:lvl>
            <w:lvl w:ilvl="2">
              <w:start w:val="1"/>
              <w:numFmt w:val="lowerRoman"/>
              <w:lvlText w:val="(%3)"/>
              <w:lvlRestart w:val="2"/>
              <w:lvlJc w:val="left"/>
              <w:pPr><w:ind w:left="2160" w:hanging="360"/></w:pPr>
            </w:lvl>
          </w:abstractNum>
          <w:num w:numId="\(numId)">
            <w:abstractNumId w:val="\(abstractNumId)"/>
          </w:num>
        </w:numbering>
        """
    }

    /// Find the highest existing numId and abstractNumId in numbering.xml and return a safe value
    private func findSafeNumId(in xml: String) -> Int {
        var maxId = 0

        // Find all numId values
        let numIdPattern = try? NSRegularExpression(pattern: "w:numId=\"(\\d+)\"", options: [])
        numIdPattern?.enumerateMatches(in: xml, range: NSRange(xml.startIndex..., in: xml)) { match, _, _ in
            if let range = match?.range(at: 1),
               let swiftRange = Range(range, in: xml),
               let id = Int(xml[swiftRange]) {
                maxId = max(maxId, id)
            }
        }

        // Find all abstractNumId values
        let abstractPattern = try? NSRegularExpression(pattern: "w:abstractNumId=\"(\\d+)\"", options: [])
        abstractPattern?.enumerateMatches(in: xml, range: NSRange(xml.startIndex..., in: xml)) { match, _, _ in
            if let range = match?.range(at: 1),
               let swiftRange = Range(range, in: xml),
               let id = Int(xml[swiftRange]) {
                maxId = max(maxId, id)
            }
        }

        Logger.shared.log("Patcher: max existing numId/abstractNumId = \(maxId)", category: "PATCH")
        return maxId + 100  // Use a high offset to be safe
    }

    /// Append our numbering definition to an existing numbering.xml
    private func appendNumberingDefinition(to numberingXML: URL, numId: Int, abstractNumId: Int) throws {
        var xml = try String(contentsOf: numberingXML, encoding: .utf8)

        // Our abstract numbering definition
        let abstractDef = """
          <w:abstractNum w:abstractNumId="\(abstractNumId)">
            <w:lvl w:ilvl="0">
              <w:start w:val="1"/>
              <w:numFmt w:val="decimal"/>
              <w:lvlText w:val="%1."/>
              <w:lvlJc w:val="left"/>
              <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
            </w:lvl>
            <w:lvl w:ilvl="1">
              <w:start w:val="1"/>
              <w:numFmt w:val="lowerLetter"/>
              <w:lvlText w:val="(%2)"/>
              <w:lvlRestart w:val="1"/>
              <w:lvlJc w:val="left"/>
              <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>
            </w:lvl>
            <w:lvl w:ilvl="2">
              <w:start w:val="1"/>
              <w:numFmt w:val="lowerRoman"/>
              <w:lvlText w:val="(%3)"/>
              <w:lvlRestart w:val="2"/>
              <w:lvlJc w:val="left"/>
              <w:pPr><w:ind w:left="2160" w:hanging="360"/></w:pPr>
            </w:lvl>
          </w:abstractNum>
        """

        // Our num definition linking to abstract
        let numDef = """
          <w:num w:numId="\(numId)">
            <w:abstractNumId w:val="\(abstractNumId)"/>
          </w:num>
        """

        // Insert abstractNum before the first <w:num> or before </w:numbering>
        if let numRange = xml.range(of: "<w:num ") {
            xml.insert(contentsOf: abstractDef + "\n", at: numRange.lowerBound)
        } else if let endRange = xml.range(of: "</w:numbering>") {
            xml.insert(contentsOf: abstractDef + "\n", at: endRange.lowerBound)
        }

        // Insert num before </w:numbering>
        if let endRange = xml.range(of: "</w:numbering>") {
            xml.insert(contentsOf: numDef + "\n", at: endRange.lowerBound)
        }

        try xml.data(using: .utf8)!.write(to: numberingXML)
        Logger.shared.log("Patcher: appended numbering definition with numId=\(numId)", category: "PATCH")
    }

    private func ensureNumberingRelationship(in workDir: URL) throws {
        let fm = FileManager.default
        let relsURL = workDir.appendingPathComponent("word/_rels/document.xml.rels")
        if !fm.fileExists(atPath: relsURL.path) {
            try fm.createDirectory(at: relsURL.deletingLastPathComponent(), withIntermediateDirectories: true)
            let base = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>"
            try base.data(using: .utf8)!.write(to: relsURL)
        }
        var rels = try String(contentsOf: relsURL, encoding: .utf8)
        if !rels.contains("numbering.xml") {
            let relLine = "<Relationship Id=\"rIdNumbering\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering\" Target=\"numbering.xml\"/>"
            rels = rels.replacingOccurrences(of: "</Relationships>", with: "  \(relLine)\n</Relationships>")
            try rels.data(using: .utf8)!.write(to: relsURL)
            Logger.shared.log("Patcher: added rel for numbering.xml", category: "PATCH")
        }

        let ctypesURL = workDir.appendingPathComponent("[Content_Types].xml")
        guard fm.fileExists(atPath: ctypesURL.path) else { return }
        var ctypes = try String(contentsOf: ctypesURL, encoding: .utf8)
        if !ctypes.contains("word/numbering.xml") {
            let override = "<Override PartName=\"/word/numbering.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml\"/>"
            ctypes = ctypes.replacingOccurrences(of: "</Types>", with: "  \(override)\n</Types>")
            try ctypes.data(using: .utf8)!.write(to: ctypesURL)
            Logger.shared.log("Patcher: added content-type for numbering.xml", category: "PATCH")
        }
    }

    /// Content-based numbering - matches paragraphs by normalized text content (not index)
    private func injectNumberingByContent(into documentXML: URL, targets: [NumberingTarget], numId: Int) throws {
        Logger.shared.log("Patcher: content-based injection into document.xml with numId=\(numId)", category: "PATCH")
        let raw = try String(contentsOf: documentXML, encoding: .utf8)

        // Build a map from normalized text prefix to level
        // Use longer prefixes for more accurate matching
        var prefixToLevel: [String: Int] = [:]
        var usedPrefixes: Set<String> = []  // Track used prefixes to avoid double-numbering

        var level0Count = 0
        var level1Count = 0
        var level2Count = 0

        for target in targets {
            let normalized = normalizeForMatching(target.textPrefix)
            if !normalized.isEmpty && normalized.count >= 10 {  // Require meaningful prefix
                prefixToLevel[normalized] = target.level
                switch target.level {
                case 0: level0Count += 1
                case 1: level1Count += 1
                case 2: level2Count += 1
                default: break
                }
            }
        }
        Logger.shared.log("Patcher: \(prefixToLevel.count) unique prefixes (L0:\(level0Count) L1:\(level1Count) L2:\(level2Count))", category: "PATCH")

        // Find table boundaries to skip table content
        let tableStartPattern = try NSRegularExpression(pattern: "<w:tbl[^>]*>", options: [])
        let tableEndPattern = try NSRegularExpression(pattern: "</w:tbl>", options: [])

        var tableRanges: [NSRange] = []
        var tableStarts: [Int] = []

        tableStartPattern.enumerateMatches(in: raw, range: NSRange(raw.startIndex..., in: raw)) { match, _, _ in
            if let range = match?.range { tableStarts.append(range.location) }
        }
        tableEndPattern.enumerateMatches(in: raw, range: NSRange(raw.startIndex..., in: raw)) { match, _, _ in
            if let range = match?.range, !tableStarts.isEmpty {
                let start = tableStarts.removeFirst()
                tableRanges.append(NSRange(location: start, length: range.location + range.length - start))
            }
        }

        func isInsideTable(_ location: Int) -> Bool {
            tableRanges.contains { location >= $0.location && location < $0.location + $0.length }
        }

        // Find all paragraphs outside tables
        let paragraphPattern = try NSRegularExpression(pattern: "<w:p(?:\\s[^>]*)?>(?:(?!</w:p>).)*</w:p>", options: [.dotMatchesLineSeparators])

        var allParagraphs: [(range: NSRange, xml: String, text: String)] = []
        paragraphPattern.enumerateMatches(in: raw, range: NSRange(raw.startIndex..., in: raw)) { match, _, _ in
            guard let matchRange = match?.range else { return }
            if isInsideTable(matchRange.location) { return }  // Skip table paragraphs

            let startIdx = raw.index(raw.startIndex, offsetBy: matchRange.location)
            let endIdx = raw.index(raw.startIndex, offsetBy: matchRange.location + matchRange.length)
            let paraXML = String(raw[startIdx..<endIdx])
            let text = extractTextFromParagraphXML(paraXML)
            allParagraphs.append((matchRange, paraXML, text))
        }

        Logger.shared.log("Patcher: found \(allParagraphs.count) paragraphs outside tables", category: "PATCH")

        // Build replacements by matching content
        var replacements: [(range: NSRange, newContent: String)] = []
        var numberedCount = 0

        for (index, para) in allParagraphs.enumerated() {
            // Skip empty paragraphs
            let normalized = normalizeForMatching(String(para.text.prefix(80)))
            if normalized.isEmpty { continue }

            // Find matching prefix from targets and get the level from buildNumberingTargetsFromAnalysis
            // That function uses INDENTATION and context to detect subparagraphs, not just text markers
            var matchedLevel: Int? = nil
            var matchedPrefix: String? = nil

            for (prefix, level) in prefixToLevel {
                // Check if this prefix was already used
                if usedPrefixes.contains(prefix) { continue }

                // Match: normalized text starts with prefix OR prefix starts with normalized text
                if normalized.hasPrefix(prefix) || (prefix.hasPrefix(normalized) && normalized.count >= 15) {
                    matchedLevel = level
                    matchedPrefix = prefix
                    break
                }
            }

            // Process if we found a matching prefix
            if var level = matchedLevel, let prefix = matchedPrefix {
                usedPrefixes.insert(prefix)  // Mark as used

                // Check if the XML text has a letter/roman marker that wasn't detected
                // This handles cases where the text is "a. Something" but was detected as level 0
                let textMarkerLevel = detectLevelFromText(para.text)
                if textMarkerLevel > level {
                    Logger.shared.log("Patcher: upgrading para \(index) from level \(level) to \(textMarkerLevel) based on text marker", category: "PATCH")
                    level = textMarkerLevel
                }

                var newParaXML = para.xml

                // Remove existing Word numbering if present - we'll apply our own
                if newParaXML.contains("<w:numPr>") {
                    newParaXML = removeExistingNumbering(from: newParaXML)
                }

                // Strip letter/roman markers if detected in text (regardless of final level)
                // This handles "a. Text" -> "Text" when we apply "(a)" numbering
                if textMarkerLevel > 0 {
                    newParaXML = stripMarkerFromParagraphXML(newParaXML, level: textMarkerLevel)
                }

                let newPara = injectNumPrIntoParagraph(newParaXML, level: level, numId: numId)
                replacements.append((para.range, newPara))
                numberedCount += 1

                if numberedCount <= 20 || numberedCount % 25 == 0 {
                    Logger.shared.log("Patcher: #\(numberedCount) para \(index) level \(level): \(para.text.prefix(50))", category: "PATCH")
                }
            }
        }

        // Apply replacements in reverse order to maintain valid indices
        var result = raw
        for replacement in replacements.reversed() {
            let startIdx = result.index(result.startIndex, offsetBy: replacement.range.location)
            let endIdx = result.index(result.startIndex, offsetBy: replacement.range.location + replacement.range.length)
            result.replaceSubrange(startIdx..<endIdx, with: replacement.newContent)
        }

        Logger.shared.log("Patcher: numbered \(numberedCount) of \(targets.count) target paragraphs", category: "PATCH")
        try result.data(using: .utf8)!.write(to: documentXML)
        Logger.shared.log("Patcher: inject complete", category: "PATCH")
    }

    /// Extract text content from XML paragraph
    private func extractTextFromParagraphXML(_ xml: String) -> String {
        // Find all <w:t>...</w:t> elements and concatenate their content
        var text = ""
        let pattern = try? NSRegularExpression(pattern: "<w:t[^>]*>([^<]*)</w:t>", options: [])
        pattern?.enumerateMatches(in: xml, range: NSRange(xml.startIndex..., in: xml)) { match, _, _ in
            if let range = match?.range(at: 1) {
                let startIdx = xml.index(xml.startIndex, offsetBy: range.location)
                let endIdx = xml.index(xml.startIndex, offsetBy: range.location + range.length)
                text += String(xml[startIdx..<endIdx])
            }
        }
        return text
    }

    /// Remove existing <w:numPr> element from paragraph XML
    /// This allows us to replace incorrect existing numbering with our own
    private func removeExistingNumbering(from xml: String) -> String {
        // Pattern to match <w:numPr>...</w:numPr> including nested content
        guard let pattern = try? NSRegularExpression(
            pattern: "<w:numPr>.*?</w:numPr>",
            options: [.dotMatchesLineSeparators]
        ) else {
            return xml
        }

        return pattern.stringByReplacingMatches(
            in: xml,
            range: NSRange(location: 0, length: xml.utf16.count),
            withTemplate: ""
        )
    }

    /// Remove leading markers from the first text run in a paragraph XML.
    /// For subparagraphs (level 1, 2), also strips any incorrect number prefix like "145."
    /// The level parameter indicates what type of marker to strip.
    private func stripMarkerFromParagraphXML(_ xml: String, level: Int) -> String {
        // Only strip markers for subparagraphs (level 1 and 2)
        // Level 0 (main paragraphs) should not have their text stripped
        if level == 0 { return xml }

        guard let pattern = try? NSRegularExpression(pattern: "(<w:t[^>]*>)([^<]*)(</w:t>)", options: [.dotMatchesLineSeparators]) else {
            return xml
        }
        guard let match = pattern.firstMatch(in: xml, range: NSRange(xml.startIndex..., in: xml)) else {
            return xml
        }
        let textRange = match.range(at: 2)
        guard let swiftRange = Range(textRange, in: xml) else { return xml }
        var textContent = String(xml[swiftRange])
        let originalContent = textContent

        // First, strip any leading number prefix (like "145. " or "145) ") that shouldn't be there
        // This handles cases like "145. a. Gas..." where 145 is an incorrect paragraph number
        let numberPrefixPattern = "^\\s*\\d+[.):]?\\s*"
        if let numRegex = try? NSRegularExpression(pattern: numberPrefixPattern, options: []) {
            textContent = numRegex.stringByReplacingMatches(in: textContent, range: NSRange(location: 0, length: textContent.utf16.count), withTemplate: "")
        }

        // Strip subparagraph markers based on level
        if level == 1 {
            // Level 1: Strip letter markers like "(a)", "a)", "a."
            let letterPatterns = [
                "^\\s*\\([a-zA-Z]\\)\\s*",    // "(a) "
                "^\\s*[a-zA-Z]\\)\\s*",        // "a) "
                "^\\s*[a-zA-Z]\\.\\s*"         // "a. " or "a."
            ]
            for markerPattern in letterPatterns {
                guard let regex = try? NSRegularExpression(pattern: markerPattern, options: []) else { continue }
                if regex.firstMatch(in: textContent, range: NSRange(location: 0, length: textContent.utf16.count)) != nil {
                    textContent = regex.stringByReplacingMatches(in: textContent, range: NSRange(location: 0, length: textContent.utf16.count), withTemplate: "")
                    break
                }
            }
        } else if level == 2 {
            // Level 2: Strip roman numeral markers like "(i)", "i)", "i."
            let romanPatterns = [
                "^\\s*\\([ivxIVX]+\\)\\s*",   // "(i) ", "(ii) "
                "^\\s*[ivxIVX]+\\)\\s*",       // "i) ", "ii) "
                "^\\s*[ivxIVX]+\\.\\s*"        // "i. ", "ii. "
            ]
            for markerPattern in romanPatterns {
                guard let regex = try? NSRegularExpression(pattern: markerPattern, options: [.caseInsensitive]) else { continue }
                if regex.firstMatch(in: textContent, range: NSRange(location: 0, length: textContent.utf16.count)) != nil {
                    textContent = regex.stringByReplacingMatches(in: textContent, range: NSRange(location: 0, length: textContent.utf16.count), withTemplate: "")
                    break
                }
            }
        }

        // Only modify XML if we actually stripped something
        if textContent != originalContent {
            var newXML = xml
            newXML.replaceSubrange(swiftRange, with: textContent)
            return newXML
        }

        return xml
    }

    /// Normalize text for content matching - removes markers, lowercases, collapses whitespace
    private func normalizeForMatching(_ text: String) -> String {
        var result = text.lowercased()

        // Remove common leading markers that might differ between source formats
        let markerPatterns = [
            "^\\s*\\d+[.)]\\s*",           // "1.", "1)"
            "^\\s*\\([a-zA-Z]\\)\\s*",     // "(a)"
            "^\\s*[a-zA-Z][).]\\s*",       // "a)", "a."
            "^\\s*\\([ivxIVX]+\\)\\s*",    // "(i)"
            "^\\s*[ivxIVX]+[).]\\s*"       // "i)", "i."
        ]

        for pattern in markerPatterns {
            if let regex = try? NSRegularExpression(pattern: pattern, options: .caseInsensitive) {
                result = regex.stringByReplacingMatches(in: result, range: NSRange(location: 0, length: result.utf16.count), withTemplate: "")
            }
        }

        // Collapse whitespace and trim
        return result
            .components(separatedBy: .whitespacesAndNewlines)
            .filter { !$0.isEmpty }
            .joined(separator: " ")
            .trimmingCharacters(in: .whitespaces)
    }

    /// Detect numbering level from text content
    /// Returns: 0 = main paragraph, 1 = subparagraph (a), 2 = sub-subparagraph (i)
    private func detectLevelFromText(_ text: String) -> Int {
        let trimmed = text.trimmingCharacters(in: .whitespacesAndNewlines)
        if trimmed.isEmpty { return 0 }

        // Optional number prefix that may exist from original Word numbering
        // e.g., "145. a. Text" or "145) a. Text" or "145 a. Text"
        let optionalNumPrefix = "(?:\\d+[.):]?\\s*)?"

        // Level 2 patterns (roman numerals ii, iii, iv, etc.) - check first as they're more specific
        // Single 'i' is ambiguous, so we handle it separately
        let level2Patterns = [
            "^" + optionalNumPrefix + "\\((?:ii|iii|iv|vi|vii|viii|ix|xi|xii)[ivxIVX]*\\)",  // "(ii)", "(iii)", "(iv)"...
            "^" + optionalNumPrefix + "(?:ii|iii|iv|vi|vii|viii|ix|xi|xii)[ivxIVX]*\\)",     // "ii)", "iii)"...
            "^" + optionalNumPrefix + "(?:ii|iii|iv|vi|vii|viii|ix|xi|xii)[ivxIVX]*\\.",     // "ii.", "iii."...
            "^\\((?:ii|iii|iv|vi|vii|viii|ix|xi|xii)[ivxIVX]*\\)"                            // Just "(ii)" at start
        ]
        for pattern in level2Patterns {
            if trimmed.range(of: pattern, options: [.regularExpression, .caseInsensitive]) != nil {
                return 2
            }
        }

        // Level 1 patterns (single letters a-z)
        // These patterns detect subparagraph markers like "(a)", "a)", "a."
        let level1Patterns = [
            "^" + optionalNumPrefix + "\\([a-zA-Z]\\)\\s*",        // "(a) " or "(a)"
            "^" + optionalNumPrefix + "[a-zA-Z]\\)\\s*",            // "a) " or "a)"
            "^" + optionalNumPrefix + "[a-zA-Z]\\.\\s+",            // "a. " (requires space after dot)
            "^" + optionalNumPrefix + "[a-zA-Z]\\.[A-Z]",           // "a.Text" (no space, capital follows)
            "^\\([a-zA-Z]\\)\\s*",                                  // Just "(a)" at start
            "^[a-zA-Z]\\)\\s*"                                      // Just "a)" at start
        ]
        for pattern in level1Patterns {
            if trimmed.range(of: pattern, options: .regularExpression) != nil {
                return 1
            }
        }

        // Check for single "(i)" which could be level 1 or level 2 - treat as level 1
        let singleRomanPattern = "^" + optionalNumPrefix + "\\([ivxIVX]\\)\\s*"
        if trimmed.range(of: singleRomanPattern, options: [.regularExpression, .caseInsensitive]) != nil {
            return 1  // Single roman numeral in parens - treat as subparagraph
        }

        return 0
    }

    /// Inject numPr into a paragraph XML
    private func injectNumPrIntoParagraph(_ paraXML: String, level: Int, numId: Int) -> String {
        let numPrXML = "<w:numPr><w:ilvl w:val=\"\(level)\"/><w:numId w:val=\"\(numId)\"/></w:numPr>"

        // Check if paragraph already has pPr
        if let pPrRange = paraXML.range(of: "<w:pPr>") {
            // Insert numPr right after <w:pPr>
            var modified = paraXML
            modified.insert(contentsOf: numPrXML, at: pPrRange.upperBound)
            return modified
        } else if let pPrRange = paraXML.range(of: "<w:pPr ") {
            // Has pPr with attributes, find closing > and insert after
            if let closeRange = paraXML.range(of: ">", range: pPrRange.lowerBound..<paraXML.endIndex) {
                var modified = paraXML
                modified.insert(contentsOf: numPrXML, at: closeRange.upperBound)
                return modified
            }
        }

        // No pPr exists, need to add one after <w:p> or <w:p ...>
        if let pRange = paraXML.range(of: "<w:p>") {
            var modified = paraXML
            modified.insert(contentsOf: "<w:pPr>\(numPrXML)</w:pPr>", at: pRange.upperBound)
            return modified
        } else if let pRange = paraXML.range(of: "<w:p ") {
            // Find the closing > of the opening tag
            if let closeRange = paraXML.range(of: ">", range: pRange.lowerBound..<paraXML.endIndex) {
                var modified = paraXML
                modified.insert(contentsOf: "<w:pPr>\(numPrXML)</w:pPr>", at: closeRange.upperBound)
                return modified
            }
        }

        return paraXML
    }
}
