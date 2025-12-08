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

    /// Index-based numbering - matches paragraphs by their sequential index in the document
    private func injectNumberingByContent(into documentXML: URL, targets: [NumberingTarget], numId: Int) throws {
        Logger.shared.log("Patcher: index-based injection into document.xml with numId=\(numId)", category: "PATCH")
        let raw = try String(contentsOf: documentXML, encoding: .utf8)

        // Build maps
        var indexToLevel: [Int: Int] = [:]
        var prefixQueue = targets.filter { !$0.textPrefix.isEmpty }
        prefixQueue.sort { $0.paragraphIndex < $1.paragraphIndex }
        for target in targets {
            indexToLevel[target.paragraphIndex] = target.level
        }
        Logger.shared.log("Patcher: \(indexToLevel.count) paragraphs to number (index map), \(prefixQueue.count) with prefixes", category: "PATCH")

        // Find all paragraphs in document order (including tables to keep indices aligned)
        let paragraphPattern = try NSRegularExpression(pattern: "<w:p(?:\\s[^>]*)?>(?:(?!</w:p>).)*</w:p>", options: [.dotMatchesLineSeparators])

        var allParagraphs: [(range: NSRange, xml: String)] = []
        paragraphPattern.enumerateMatches(in: raw, range: NSRange(raw.startIndex..., in: raw)) { match, _, _ in
            guard let matchRange = match?.range else { return }
            let startIdx = raw.index(raw.startIndex, offsetBy: matchRange.location)
            let endIdx = raw.index(raw.startIndex, offsetBy: matchRange.location + matchRange.length)
            let paraXML = String(raw[startIdx..<endIdx])
            allParagraphs.append((matchRange, paraXML))
        }

        Logger.shared.log("Patcher: found \(allParagraphs.count) paragraphs (all)", category: "PATCH")

        // Build replacements based on paragraph index
        var replacements: [(range: NSRange, newContent: String)] = []
        var numberedCount = 0

        for (index, para) in allParagraphs.enumerated() {
            var level: Int? = indexToLevel[index]

            // If no direct index match, attempt prefix match in order
            if level == nil, !prefixQueue.isEmpty {
                let text = extractTextFromParagraphXML(para.xml)
                if let pos = prefixQueue.firstIndex(where: { !text.isEmpty && text.hasPrefix($0.textPrefix) }) {
                    level = prefixQueue[pos].level
                    prefixQueue.remove(at: pos)
                }
            }

            if let level = level {
                // Skip if paragraph already has numbering
                if para.xml.contains("<w:numPr>") {
                    Logger.shared.log("Patcher: para \(index) already has numPr, skipping", category: "PATCH")
                    continue
                }

                var newParaXML = para.xml
                // Strip existing textual marker for sublevels to avoid "(a) (a)" output
                if level > 0 {
                    newParaXML = stripMarkerFromParagraphXML(newParaXML)
                }
                let newPara = injectNumPrIntoParagraph(newParaXML, level: level, numId: numId)
                replacements.append((para.range, newPara))
                numberedCount += 1

                if numberedCount <= 5 || numberedCount % 20 == 0 {
                    let text = extractTextFromParagraphXML(para.xml)
                    Logger.shared.log("Patcher: numbering para \(index) level \(level): \(text.prefix(50))", category: "PATCH")
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

        Logger.shared.log("Patcher: numbered \(numberedCount) paragraphs by index", category: "PATCH")
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

    /// Remove leading manual marker like "1. " or "(a)" from the first text run in a paragraph XML.
    private func stripMarkerFromParagraphXML(_ xml: String) -> String {
        guard let pattern = try? NSRegularExpression(pattern: "(<w:t[^>]*>)([^<]*)(</w:t>)", options: [.dotMatchesLineSeparators]) else {
            return xml
        }
        guard let match = pattern.firstMatch(in: xml, range: NSRange(xml.startIndex..., in: xml)) else {
            return xml
        }
        let textRange = match.range(at: 2)
        guard let swiftRange = Range(textRange, in: xml) else { return xml }
        let textContent = String(xml[swiftRange])

        // Detect leading markers (1. ) (a) a) (i) etc.
        let markerPattern = try? NSRegularExpression(pattern: "^\\s*(\\d+[.)]|\\([a-zA-Z]\\)|[a-zA-Z]\\)|\\([ivxIVX]+\\))\\s+", options: [])
        if let marker = markerPattern,
           marker.firstMatch(in: textContent, range: NSRange(location: 0, length: textContent.utf16.count)) != nil {
            let cleaned = marker.stringByReplacingMatches(in: textContent, options: [], range: NSRange(location: 0, length: textContent.utf16.count), withTemplate: "")
            var newXML = xml
            newXML.replaceSubrange(swiftRange, with: cleaned)
            return newXML
        }
        return xml
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
