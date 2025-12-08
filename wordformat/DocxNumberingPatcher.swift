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

        // 2) Ensure numbering.xml exists
        let numberingXML = workDir.appendingPathComponent("word/numbering.xml")
        if !fm.fileExists(atPath: numberingXML.path) {
            Logger.shared.log("Patcher: numbering.xml missing, creating…", category: "PATCH")
            let xml = numberingTemplate()
            try xml.data(using: .utf8)!.write(to: numberingXML)
        } else {
            Logger.shared.log("Patcher: numbering.xml already present", category: "PATCH")
        }

        // 3) Ensure rels and content type
        try ensureNumberingRelationship(in: workDir)

        // 4) Content-based numbering injection
        try injectNumberingByContent(into: documentXML, targets: targets)

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

    private func numberingTemplate() -> String {
        return """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:abstractNum w:abstractNumId="0">
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
              <w:lvlText w:val="%2)"/>
              <w:lvlJc w:val="left"/>
              <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>
            </w:lvl>
            <w:lvl w:ilvl="2">
              <w:start w:val="1"/>
              <w:numFmt w:val="lowerRoman"/>
              <w:lvlText w:val="%3)"/>
              <w:lvlJc w:val="left"/>
              <w:pPr><w:ind w:left="2160" w:hanging="360"/></w:pPr>
            </w:lvl>
          </w:abstractNum>
          <w:num w:numId="1">
            <w:abstractNumId w:val="0"/>
          </w:num>
        </w:numbering>
        """
    }

    private func ensureNumberingRelationship(in workDir: URL) throws {
        let fm = FileManager.default
        let relsURL = workDir.appendingPathComponent("word/_rels/document.xml.rels")
        if !fm.fileExists(atPath: relsURL.path) {
            try fm.createDirectory(at: relsURL.deletingLastPathComponent(), withIntermediateDirectories: true)
            let base = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>"
            try base.data(using: .utf8)!.write(to: relsURL)
        }
        var rels = try String(contentsOf: relsURL)
        if !rels.contains("numbering.xml") {
            let relLine = "<Relationship Id=\"rIdNumbering\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering\" Target=\"numbering.xml\"/>"
            rels = rels.replacingOccurrences(of: "</Relationships>", with: "  \(relLine)\n</Relationships>")
            try rels.data(using: .utf8)!.write(to: relsURL)
            Logger.shared.log("Patcher: added rel for numbering.xml", category: "PATCH")
        }

        let ctypesURL = workDir.appendingPathComponent("[Content_Types].xml")
        guard fm.fileExists(atPath: ctypesURL.path) else { return }
        var ctypes = try String(contentsOf: ctypesURL)
        if !ctypes.contains("word/numbering.xml") {
            let override = "<Override PartName=\"/word/numbering.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml\"/>"
            ctypes = ctypes.replacingOccurrences(of: "</Types>", with: "  \(override)\n</Types>")
            try ctypes.data(using: .utf8)!.write(to: ctypesURL)
            Logger.shared.log("Patcher: added content-type for numbering.xml", category: "PATCH")
        }
    }

    /// Content-based numbering - matches paragraphs by text content
    private func injectNumberingByContent(into documentXML: URL, targets: [NumberingTarget]) throws {
        Logger.shared.log("Patcher: content-based injection into document.xml", category: "PATCH")
        let raw = try String(contentsOf: documentXML)

        // Build a set of text prefixes to match (normalized)
        var prefixToLevel: [String: Int] = [:]
        for target in targets {
            let normalized = normalizeText(target.textPrefix)
            if !normalized.isEmpty {
                prefixToLevel[normalized] = target.level
            }
        }
        Logger.shared.log("Patcher: \(prefixToLevel.count) unique text prefixes to match", category: "PATCH")

        // Find all paragraphs and their text content
        var result = raw
        var numberedCount = 0
        var tableDepth = 0

        // Process using regex to find paragraphs
        // Pattern matches <w:p...>...</w:p> including nested content
        let paragraphPattern = try NSRegularExpression(pattern: "<w:p[^>]*>.*?</w:p>", options: [.dotMatchesLineSeparators])
        let tableStartPattern = try NSRegularExpression(pattern: "<w:tbl[^>]*>", options: [])
        let tableEndPattern = try NSRegularExpression(pattern: "</w:tbl>", options: [])

        // Find table boundaries first
        var tableRanges: [NSRange] = []
        var tableStarts: [Int] = []

        tableStartPattern.enumerateMatches(in: raw, range: NSRange(raw.startIndex..., in: raw)) { match, _, _ in
            if let range = match?.range {
                tableStarts.append(range.location)
            }
        }

        tableEndPattern.enumerateMatches(in: raw, range: NSRange(raw.startIndex..., in: raw)) { match, _, _ in
            if let range = match?.range, !tableStarts.isEmpty {
                let start = tableStarts.removeFirst()
                tableRanges.append(NSRange(location: start, length: range.location + range.length - start))
            }
        }

        // Check if a location is inside a table
        func isInsideTable(_ location: Int) -> Bool {
            for range in tableRanges {
                if location >= range.location && location < range.location + range.length {
                    return true
                }
            }
            return false
        }

        // Find paragraphs to number
        var replacements: [(range: NSRange, newContent: String)] = []

        paragraphPattern.enumerateMatches(in: raw, range: NSRange(raw.startIndex..., in: raw)) { match, _, _ in
            guard let matchRange = match?.range else { return }

            // Skip paragraphs inside tables
            if isInsideTable(matchRange.location) { return }

            let startIdx = raw.index(raw.startIndex, offsetBy: matchRange.location)
            let endIdx = raw.index(raw.startIndex, offsetBy: matchRange.location + matchRange.length)
            let paraXML = String(raw[startIdx..<endIdx])

            // Extract text content from paragraph
            let text = extractTextFromParagraphXML(paraXML)
            let normalized = normalizeText(String(text.prefix(100)))

            // Check if this paragraph should be numbered
            var matchedLevel: Int? = nil
            for (prefix, level) in prefixToLevel {
                if normalized.hasPrefix(prefix) || prefix.hasPrefix(normalized) {
                    matchedLevel = level
                    break
                }
            }

            if let level = matchedLevel {
                let newPara = injectNumPrIntoParagraph(paraXML, level: level)
                replacements.append((matchRange, newPara))
                numberedCount += 1
            }
        }

        // Apply replacements in reverse order to maintain valid indices
        for replacement in replacements.reversed() {
            let startIdx = result.index(result.startIndex, offsetBy: replacement.range.location)
            let endIdx = result.index(result.startIndex, offsetBy: replacement.range.location + replacement.range.length)
            result.replaceSubrange(startIdx..<endIdx, with: replacement.newContent)
        }

        Logger.shared.log("Patcher: numbered \(numberedCount) paragraphs by content match", category: "PATCH")
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

    /// Normalize text for matching (lowercase, remove extra whitespace)
    private func normalizeText(_ text: String) -> String {
        return text.lowercased()
            .components(separatedBy: .whitespacesAndNewlines)
            .filter { !$0.isEmpty }
            .joined(separator: " ")
            .trimmingCharacters(in: .whitespaces)
    }

    /// Inject numPr into a paragraph XML
    private func injectNumPrIntoParagraph(_ paraXML: String, level: Int) -> String {
        let numPrXML = "<w:numPr><w:ilvl w:val=\"\(level)\"/><w:numId w:val=\"1\"/></w:numPr>"

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
