//
//  Formatting.swift
//  wordformat
//
//  Created by Krisztian Bito on 08.12.2025.
//

import Foundation
import AppKit

// MARK: - Defaults

enum LegalFormattingDefaults {
    static let fontFamily = "Times New Roman"
    static let fontSize: CGFloat = 12.0
}

// MARK: - Main Entry Point

func applyUKLegalFormatting(
    to document: NSMutableAttributedString,
    header: LegalHeaderMetadata,
    analysis: AnalysisResult?
) {
    guard document.length > 0 else { return }
    
    // 1. Analyse Structure
    // Use AI analysis if provided, otherwise fallback to local heuristic
    var structure = analysis?.classifiedRanges ?? []
    if structure.isEmpty {
        structure = detectDocumentStructure(in: document)
    }
    
    // 2. Reconstruct using Markdown Engine
    // This generates a completely new NSAttributedString with perfect internal list structures.
    if let newDocument = rebuildDocumentUsingMarkdown(from: document, structure: structure, metadata: header) {
        document.setAttributedString(newDocument)
    }
}

// MARK: - Markdown Reconstruction Engine

private func rebuildDocumentUsingMarkdown(
    from original: NSAttributedString,
    structure: [AnalysisResult.FormattedRange],
    metadata: LegalHeaderMetadata
) -> NSAttributedString? {
    
    var markdown = ""
    
    // -- Step A: Header --
    // We treat the header as a bold block.
    // If the original doc DOES NOT have a header detected, we generate one.
    // If it DOES, we use the original text but format it.
    
    let hasDetectedHeader = structure.contains { $0.type == .headerMetadata }
    if !hasDetectedHeader && !metadata.caseReference.isEmpty {
        markdown += generateHeaderMarkdown(metadata)
    }
    
    // -- Step B: Body Processing --
    // We iterate through the blocks and convert them to Markdown syntax.
    
    for item in structure {
        let rawText = (original.string as NSString).substring(with: item.range)
        let cleanText = rawText.trimmingCharacters(in: .whitespacesAndNewlines)
        
        if cleanText.isEmpty { continue }
        
        switch item.type {
        case .headerMetadata:
            // Centered Bold Text (Markdown doesn't support center natively, we fix in post-process)
            // We just make it bold here.
            markdown += "\n\n**\(cleanText.replacingOccurrences(of: "\n", with: "  \n"))**\n\n"
            
        case .documentTitle:
            // Bold Uppercase
            markdown += "\n\n**\(cleanText.uppercased())**\n\n"
            
        case .heading:
            // Bold
            markdown += "\n\n**\(cleanText)**\n\n"
            
        case .body:
            // DYNAMIC NUMBERING
            // In Markdown, "1. Text" creates an ordered list.
            // "   a. Text" creates a sub-list.
            
            let stripped = stripManualNumbering(cleanText)
            
            if isSubPoint(cleanText) {
                // Indent with 4 spaces for sub-level
                markdown += "    1. \(stripped)\n"
            } else {
                // Main level
                markdown += "1. \(stripped)\n"
            }
            
        case .quote:
            // Blockquote
            markdown += "> \(cleanText)\n\n"
            
        case .intro, .statementOfTruth, .signature:
            // Standard text
            markdown += "\(cleanText)\n\n"
            
        case .unknown:
            markdown += "\(cleanText)\n\n"
        }
    }
    
    // -- Step C: Convert Markdown to NSAttributedString --
    do {
        var options = AttributedString.MarkdownParsingOptions()
        options.interpretedSyntax = .inlineOnlyPreservingWhitespace // Adjust if needed
        
        // We use the modern NSAttributedString(markdown:) initializer (macOS 12+)
        let attrStr = try NSMutableAttributedString(
            markdown: markdown,
            options: .init(interpretedSyntax: .full)
        )
        
        // -- Step D: Post-Processing Styles (The Polish) --
        // Markdown handles the lists/bolding. We must handle Fonts and Alignment.
        
        let fullRange = NSRange(location: 0, length: attrStr.length)
        
        // 1. Force Times New Roman 12pt globally
        applyBaseFont(to: attrStr, range: fullRange)
        
        // 2. Fix Alignment based on content analysis (Heuristic Post-Pass)
        // Since Markdown swallowed the structure types, we re-scan briefly to center the header/title.
        // Or we rely on the formatting defaults.
        // List items are automatically justified/left by the system.
        
        // Let's center the top part if it looks like a header (uppercase, "IN THE", etc)
        centerHeaders(in: attrStr)
        
        return attrStr
        
    } catch {
        print("Markdown conversion failed: \(error)")
        return nil
    }
}

// MARK: - Post-Processing Helpers

private func applyBaseFont(to doc: NSMutableAttributedString, range: NSRange) {
    let targetFont = NSFont(name: LegalFormattingDefaults.fontFamily, size: LegalFormattingDefaults.fontSize)
        ?? NSFont.systemFont(ofSize: LegalFormattingDefaults.fontSize)
    
    doc.enumerateAttribute(.font, in: range, options: []) { value, subRange, _ in
        if let existing = value as? NSFont {
            // Keep Bold/Italic traits
            let traits = existing.fontDescriptor.symbolicTraits
            // Build a descriptor from the target font descriptor, preserving traits
            let mergedDescriptor = targetFont.fontDescriptor.withSymbolicTraits(traits)
            let newFont = NSFont(descriptor: mergedDescriptor, size: LegalFormattingDefaults.fontSize) ?? targetFont
            doc.addAttribute(.font, value: newFont, range: subRange)
        } else {
            doc.addAttribute(.font, value: targetFont, range: subRange)
        }
        
        // Fix Paragraph Spacing
        let para = NSMutableParagraphStyle()
        para.paragraphSpacing = 12
        doc.addAttribute(.paragraphStyle, value: para, range: subRange)
    }
}

private func centerHeaders(in doc: NSMutableAttributedString) {
    let string = doc.string as NSString
    string.enumerateSubstrings(in: NSRange(location: 0, length: string.length), options: .byParagraphs) { substring, range, _, _ in
        guard let text = substring?.trimmingCharacters(in: .whitespacesAndNewlines).lowercased() else { return }
        
        // Heuristic for Centering: Header keywords
        if text.contains("in the") || text.contains("case reference") || text.contains("between:") || text.contains("witness statement") || text.contains("-and-") || text.contains("applicant") {
            
            let style = NSMutableParagraphStyle()
            style.alignment = .center
            style.paragraphSpacing = 6
            doc.addAttribute(.paragraphStyle, value: style, range: range)
        }
    }
}

private func generateHeaderMarkdown(_ metadata: LegalHeaderMetadata) -> String {
    return """
    **IN THE \(metadata.tribunalName.uppercased())** Case Reference: \(metadata.caseReference)  
    
    **BETWEEN:** **\(metadata.applicantName.uppercased())** Applicant  
    
    **-and-** **\(metadata.respondentName.uppercased())** Respondent
    
    
    """
}

// MARK: - Utilities

private func stripManualNumbering(_ text: String) -> String {
    // Remove "1.", "304.", "(a)" etc
    let pattern = "^\\s*(?:\\d+|\\([a-zA-Z]\\)|[a-zA-Z]\\.)[.)]?\\s+"
    if let regex = try? NSRegularExpression(pattern: pattern) {
        let range = NSRange(location: 0, length: text.utf16.count)
        return regex.stringByReplacingMatches(in: text, options: [], range: range, withTemplate: "")
    }
    return text
}

private func isSubPoint(_ text: String) -> Bool {
    return text.range(of: "^\\s*\\(?[a-zA-Z]\\)[.)]", options: .regularExpression) != nil
}

// MARK: - Heuristic Fallback (Same as before, used if AI fails)

private func detectDocumentStructure(in document: NSAttributedString) -> [AnalysisResult.FormattedRange] {
    var ranges: [AnalysisResult.FormattedRange] = []
    let fullString = document.string as NSString
    let fullRange = NSRange(location: 0, length: fullString.length)
    
    // Simple Keyword State Machine
    let titleKeywords = ["witness statement"]
    let introKeywords = ["will say as follows", "states as follows"]
    
    // We treat everything before "Witness Statement" as Header
    var foundTitle = false
    var foundIntro = false
    
    fullString.enumerateSubstrings(in: fullRange, options: .byParagraphs) { substring, range, _, _ in
        guard let text = substring?.trimmingCharacters(in: .whitespacesAndNewlines) else { return }
        let lower = text.lowercased()
        
        var type: LegalParagraphType = .body
        
        if !foundTitle {
            if lower.contains("witness statement") {
                type = .documentTitle
                foundTitle = true
            } else {
                type = .headerMetadata
            }
        } else if !foundIntro {
            if lower.contains("will say as follows") || lower.contains("states as follows") {
                type = .intro
                foundIntro = true
            } else {
                type = .intro
            }
        } else {
            // Body / Heading
            if text.range(of: "^[A-Z0-9\\.]+\\s+[A-Z]", options: .regularExpression) != nil && text.count < 60 {
                type = .heading
            } else if lower.contains("statement of truth") {
                type = .statementOfTruth
            } else {
                type = .body
            }
        }
        
        ranges.append(AnalysisResult.FormattedRange(range: range, type: type))
    }
    
    return ranges
}

