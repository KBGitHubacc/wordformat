//
//  DocxNumberingPatcher.swift
//  wordformat
//
//  Created by Assistant on 08.12.2025.
//

import Foundation

/// Post-processing patcher that injects Word-native numbering into a saved DOCX by editing the XML.
/// Steps:
/// 1) Unzip to a temp folder.
/// 2) Add numbering.xml (3-level: decimal, lower-alpha, lower-roman) if missing.
/// 3) Add relationships + content-type override for numbering.
/// 4) For target paragraphs, inject <w:numPr> with numId and ilvl.
/// 5) Re-zip to the destination DOCX.
struct DocxNumberingPatcher {
    enum PatcherError: Error {
        case unzipFailed
        case documentXMLNotFound
        case rezippingFailed
    }
    
    /// Represents a paragraph that should be numbered at a given level.
    struct NumberingTarget {
        let paragraphIndex: Int
        let level: Int // 0-based: 0 -> main, 1 -> (a), 2 -> (i)
    }
    
    /// Apply numbering in place on the saved DOCX file.
    /// - Parameters:
    ///   - docxURL: path to the saved DOCX
    ///   - targets: paragraphs to number with levels
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
        
        // 4) Inject numPr into document.xml
        try injectNumbering(into: documentXML, targets: targets)
        
        // 5) Re-zip to same docx
        let parent = docxURL.deletingLastPathComponent()
        let tempOut = parent.appendingPathComponent(docxURL.lastPathComponent + ".tmp")
        let cwd = fm.currentDirectoryPath
        defer { _ = runShell("cd \"\(cwd)\"") } // no-op reset
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
    
    // MARK: helpers
    
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
        // 1) _rels/.rels not needed; we need word/_rels/document.xml.rels
        let relsURL = workDir.appendingPathComponent("word/_rels/document.xml.rels")
        if !fm.fileExists(atPath: relsURL.path) {
            Logger.shared.log("Patcher: document.xml.rels missing, creating empty", category: "PATCH")
            try fm.createDirectory(at: relsURL.deletingLastPathComponent(), withIntermediateDirectories: true)
            let base = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>"
            try base.data(using: .utf8)!.write(to: relsURL)
        }
        var rels = try String(contentsOf: relsURL)
        if !rels.contains("numbering.xml") {
            let id = "rIdNumbering"
            let relLine = "<Relationship Id=\"\(id)\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering\" Target=\"numbering.xml\"/>"
            rels = rels.replacingOccurrences(of: "</Relationships>", with: "  \(relLine)\n</Relationships>")
            try rels.data(using: .utf8)!.write(to: relsURL)
            Logger.shared.log("Patcher: added rel for numbering.xml", category: "PATCH")
        }
        
        // 2) [Content_Types].xml override
        let ctypesURL = workDir.appendingPathComponent("[Content_Types].xml")
        guard fm.fileExists(atPath: ctypesURL.path) else { return }
        var ctypes = try String(contentsOf: ctypesURL)
        if !ctypes.contains("word/numbering.xml") {
            let override = "<Override PartName=\"/word/numbering.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml\"/>"
            ctypes = ctypes.replacingOccurrences(of: "</Types>", with: "  \(override)\n</Types>")
            try ctypes.data(using: .utf8)!.write(to: ctypesURL)
            Logger.shared.log("Patcher: added content-type override for numbering.xml", category: "PATCH")
        }
    }
    
    private func injectNumbering(into documentXML: URL, targets: [NumberingTarget]) throws {
        Logger.shared.log("Patcher: injecting numbering into document.xml", category: "PATCH")
        let raw = try String(contentsOf: documentXML)
        var paragraphs = raw.components(separatedBy: "<w:p")
        // The first split element is the header before first <w:p, keep as-is
        if paragraphs.count <= 1 {
            Logger.shared.log("Patcher: no paragraphs found", category: "PATCH")
            return
        }
        
        var numberedSet = Set(targets.map { $0.paragraphIndex })
        let levelForIndex = Dictionary(uniqueKeysWithValues: targets.map { ($0.paragraphIndex, $0.level) })
        
        var rebuilt = paragraphs[0]
        for idx in 1..<paragraphs.count {
            var chunk = paragraphs[idx]
            // Re-add the removed "<w:p"
            chunk = "<w:p" + chunk
            
            // Only body paragraphs, skip if not targeted
            if numberedSet.contains(idx - 1) {
                let level = levelForIndex[idx - 1] ?? 0
                let numPr = """
                <w:pPr><w:numPr><w:ilvl w:val="\(level)"/><w:numId w:val="1"/></w:numPr>
                """
                // If <w:pPr> exists, inject numPr inside it; else create new pPr
                if let range = chunk.range(of: "<w:pPr>") {
                    chunk.insert(contentsOf: "<w:numPr><w:ilvl w:val=\"\(level)\"/><w:numId w:val=\"1\"/></w:numPr>", at: range.upperBound)
                } else if let range = chunk.range(of: "<w:pPr ") { // attributes version
                    if let close = chunk.range(of: ">", range: range.lowerBound..<chunk.endIndex) {
                        chunk.insert(contentsOf: "<w:numPr><w:ilvl w:val=\"\(level)\"/><w:numId w:val=\"1\"/></w:numPr>", at: close.upperBound)
                    }
                } else {
                    // No pPr; insert one right after <w:p
                    if let close = chunk.range(of: ">") {
                        chunk.insert(contentsOf: numPr + "</w:pPr>", at: close.upperBound)
                    }
                }
            }
            
            rebuilt += chunk
        }
        
        try rebuilt.data(using: .utf8)!.write(to: documentXML)
        Logger.shared.log("Patcher: inject complete", category: "PATCH")
    }
}
