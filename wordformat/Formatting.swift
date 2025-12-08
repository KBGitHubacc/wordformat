//
//  Formatting.swift
//  wordformat
//
//  Created by Krisztian Bito on 08.12.2025.
//

import Foundation
import AppKit

// MARK: - Configuration

enum LegalFormattingDefaults {
    static let fontFamily = "Times New Roman"
    static let fontSize: CGFloat = 12.0
    
    // Spacing
    static let paragraphSpacing: CGFloat = 12.0
    static let lineHeight: CGFloat = 1.0
    
    // List Indentation (Points)
    // 36pts = 0.5 inches (Standard Word Indent)
    static let level1TextIndent: CGFloat = 36.0
    static let level1NumberIndent: CGFloat = 0.0
    
    static let level2TextIndent: CGFloat = 72.0
    static let level2NumberIndent: CGFloat = 36.0
}

// MARK: - Main Entry Point

func applyUKLegalFormatting(
    to document: NSMutableAttributedString,
    header: LegalHeaderMetadata,
    analysis: AnalysisResult?
) {
    guard document.length > 0 else { return }
    
    // 1. Analyse Structure
    // We scan the document to classify headers, body, quotes, etc.
    let structure = detectDocumentStructure(in: document)
    
    // 2. Re-Build Document
    // We create a fresh NSMutableAttributedString to ensure clean attributes.
    // We stitch it together paragraph by paragraph.
    let newDocument = rebuildDocument(from: document, structure: structure, metadata: header)
    
    // 3. Global Font Polish
    // Ensure the font is consistent (Times New Roman 12)
    let fullRange = NSRange(location: 0, length: newDocument.length)
    applyBaseFont(to: newDocument, range: fullRange)
    
    // 4. Replace content
    document.setAttributedString(newDocument)
}

// MARK: - Document Reconstruction Engine

private func rebuildDocument(
    from original: NSAttributedString,
    structure: [AnalysisResult.FormattedRange],
    metadata: LegalHeaderMetadata
) -> NSMutableAttributedString {
    
    let output = NSMutableAttributedString()
    
    // -- Step A: Header --
    let hasHeader = structure.first?.type == .headerMetadata
    if !hasHeader && !metadata.caseReference.isEmpty {
        output.append(generateLegalHeader(metadata))
    }
    
    // -- Step B: Create Persistent List Objects --
    // Crucial: Reusing these objects ensures the numbering is continuous (1, 2, 3...)
    // If we created a new NSTextList for every paragraph, they would all be "1."
    let rootList = NSTextList(markerFormat: .decimal, options: 0) // 1, 2, 3
    rootList.startingItemNumber = 1
    
    // We maintain a map of sublists if needed, but a single generic sublist usually suffices for simple legal docs
    let subList = NSTextList(markerFormat: .lowercaseAlpha, options: 0) // a, b, c
    
    // -- Step C: Process Paragraphs --
    for item in structure {
        let originalRange = item.range
        let originalText = (original.string as NSString).substring(with: originalRange)
        
        // Skip empty paragraphs to keep list tight, unless they are spacers
        let cleanText = originalText.trimmingCharacters(in: .whitespacesAndNewlines)
        if cleanText.isEmpty {
            // Optional: Insert a spacer if it's not a list item
            if item.type != .body {
                output.append(NSAttributedString(string: "\n"))
            }
            continue
        }
        
        let paragraphContent = NSMutableAttributedString(string: cleanText)
        
        switch item.type {
        case .headerMetadata:
            applyHeaderStyle(to: paragraphContent)
            output.append(paragraphContent)
            output.append(NSAttributedString(string: "\n"))
            
        case .documentTitle:
            applyTitleStyle(to: paragraphContent)
            output.append(paragraphContent)
            output.append(NSAttributedString(string: "\n"))
            
        case .intro:
            applyBodyStyle(to: paragraphContent)
            output.append(paragraphContent)
            output.append(NSAttributedString(string: "\n"))
            
        case .heading:
            applyHeadingStyle(to: paragraphContent)
            output.append(paragraphContent)
            output.append(NSAttributedString(string: "\n"))
            
        case .body:
            // LIST LOGIC
            // 1. Detect Level
            let isSub = isSubPoint(cleanText)
            
            // 2. Strip existing manual numbers (e.g. "1.", "(a)")
            let strippedText = stripManualNumbering(cleanText)
            
            // 3. Create the list item text
            // Ideally, we prepend a TAB. This helps text engines align the text to the tab stop.
            let itemString = NSMutableAttributedString(string: strippedText) // No manual number here!
            
            // 4. Create Paragraph Style for List
            let style = NSMutableParagraphStyle()
            style.paragraphSpacing = LegalFormattingDefaults.paragraphSpacing
            style.alignment = .justified
            
            if isSub {
                // LEVEL 2
                style.headIndent = LegalFormattingDefaults.level2TextIndent
                style.firstLineHeadIndent = LegalFormattingDefaults.level2NumberIndent
                
                // Important: Apply BOTH lists to indicate nesting
                style.textLists = [rootList, subList]
                
                // Add tab stops to match indent
                style.tabStops = [
                    NSTextTab(textAlignment: .left, location: LegalFormattingDefaults.level2TextIndent, options: [:])
                ]
            } else {
                // LEVEL 1
                style.headIndent = LegalFormattingDefaults.level1TextIndent
                style.firstLineHeadIndent = LegalFormattingDefaults.level1NumberIndent
                
                // Apply Root List
                style.textLists = [rootList]
                
                // Add tab stops
                style.tabStops = [
                    NSTextTab(textAlignment: .left, location: LegalFormattingDefaults.level1TextIndent, options: [:])
                ]
            }
            
            // 5. Apply Attributes
            let range = NSRange(location: 0, length: itemString.length)
            itemString.addAttribute(.paragraphStyle, value: style, range: range)
            
            // 6. Append to Document
            output.append(itemString)
            output.append(NSAttributedString(string: "\n"))
            
        case .quote:
            applyQuoteStyle(to: paragraphContent)
            output.append(paragraphContent)
            output.append(NSAttributedString(string: "\n"))
            
        case .statementOfTruth:
            applyBodyStyle(to: paragraphContent)
            boldEverything(in: paragraphContent)
            output.append(paragraphContent)
            output.append(NSAttributedString(string: "\n"))
            
        case .signature:
            applyBodyStyle(to: paragraphContent)
            output.append(paragraphContent)
            output.append(NSAttributedString(string: "\n"))
            
        case .unknown:
            break
        }
    }
    
    return output
}

// MARK: - Style Applicators

private func applyHeaderStyle(to str: NSMutableAttributedString) {
    let style = NSMutableParagraphStyle()
    style.alignment = .center
    style.paragraphSpacing = 0
    str.addAttribute(.paragraphStyle, value: style, range: NSRange(location: 0, length: str.length))
    
    // Bold specific words
    let text = str.string.lowercased()
    if text.contains("in the") || text.contains("between") || text.contains("witness") {
        boldEverything(in: str)
    }
}

private func applyTitleStyle(to str: NSMutableAttributedString) {
    let style = NSMutableParagraphStyle()
    style.alignment = .center
    style.paragraphSpacing = 24
    str.addAttribute(.paragraphStyle, value: style, range: NSRange(location: 0, length: str.length))
    boldEverything(in: str)
    
    let upper = str.string.uppercased()
    str.replaceCharacters(in: NSRange(location: 0, length: str.length), with: upper)
}

private func applyBodyStyle(to str: NSMutableAttributedString) {
    let style = NSMutableParagraphStyle()
    style.alignment = .left
    style.paragraphSpacing = 12
    str.addAttribute(.paragraphStyle, value: style, range: NSRange(location: 0, length: str.length))
}

private func applyHeadingStyle(to str: NSMutableAttributedString) {
    let style = NSMutableParagraphStyle()
    style.alignment = .left
    style.paragraphSpacingBefore = 18
    style.paragraphSpacing = 6
    str.addAttribute(.paragraphStyle, value: style, range: NSRange(location: 0, length: str.length))
    boldEverything(in: str)
}

private func applyQuoteStyle(to str: NSMutableAttributedString) {
    let style = NSMutableParagraphStyle()
    style.alignment = .justified
    style.headIndent = 36
    style.firstLineHeadIndent = 36
    style.paragraphSpacing = 12
    str.addAttribute(.paragraphStyle, value: style, range: NSRange(location: 0, length: str.length))
}

// MARK: - Text Processing Helpers

private func stripManualNumbering(_ text: String) -> String {
    // Matches "1.", "304.", "1 ", "(a)", "a)"
    // We strip this so the auto-numbering doesn't duplicate it.
    let pattern = "^\\s*(\\d+|\\([a-zA-Z0-9]+\\)|[a-zA-Z0-9]+\\))[.)]?\\s+"
    
    if let regex = try? NSRegularExpression(pattern: pattern, options: []) {
        let range = NSRange(location: 0, length: text.utf16.count)
        // Only remove if it's at the very start
        let modified = regex.stringByReplacingMatches(in: text, options: [], range: range, withTemplate: "")
        
        // Only return modified if we actually stripped something substantial (avoid stripping single words by accident)
        // Heuristic: If we removed < 6 chars, it's likely a number.
        if text.count - modified.count > 0 {
            return modified
        }
    }
    return text
}

private func isSubPoint(_ text: String) -> Bool {
    // Detects "(a)", "a.", "a)"
    let pattern = "^\\s*\\(?([a-zA-Z])\\)[.)]?\\s+"
    return text.range(of: pattern, options: .regularExpression) != nil
}

// MARK: - Structure Analysis

private func detectDocumentStructure(in document: NSAttributedString) -> [AnalysisResult.FormattedRange] {
    var ranges: [AnalysisResult.FormattedRange] = []
    let fullString = document.string as NSString
    let fullRange = NSRange(location: 0, length: fullString.length)
    
    let headerKeywords = ["case no", "case ref", "claim no", "in the", "tribunal", "between:", "applicant", "respondent", "-v-", "-and-"]
    let titleKeywords = ["witness statement"]
    let introKeywords = ["will say as follows", "states as follows"]
    let truthKeywords = ["statement of truth", "believe that the facts"]
    
    enum ScanState { case header, preBody, body, backMatter }
    var currentState: ScanState = .header
    
    fullString.enumerateSubstrings(in: fullRange, options: .byParagraphs) { substring, substringRange, _, _ in
        guard let text = substring?.trimmingCharacters(in: .whitespacesAndNewlines).lowercased(), !text.isEmpty else { return }
        
        var type: LegalParagraphType = .body
        
        switch currentState {
        case .header:
            if containsAny(text, keywords: titleKeywords) {
                type = .documentTitle
                currentState = .preBody
            } else if containsAny(text, keywords: headerKeywords) || substringRange.location < 800 {
                type = .headerMetadata
            } else {
                type = .intro
                currentState = .preBody
            }
        case .preBody:
            if containsAny(text, keywords: introKeywords) {
                type = .intro
                currentState = .body
            } else if containsAny(text, keywords: titleKeywords) {
                type = .documentTitle
            } else {
                type = .intro
            }
        case .body:
            if containsAny(text, keywords: truthKeywords) {
                type = .statementOfTruth
                currentState = .backMatter
            } else if isHeading(substring ?? "") {
                type = .heading
            } else {
                type = .body
            }
        case .backMatter:
            type = .signature
        }
        
        ranges.append(AnalysisResult.FormattedRange(range: substringRange, type: type))
    }
    return ranges
}

// MARK: - Utilities

private func generateLegalHeader(_ metadata: LegalHeaderMetadata) -> NSAttributedString {
    let content = """
    IN THE \(metadata.tribunalName.uppercased())
    Case Reference: \(metadata.caseReference)
    
    BETWEEN:
    
    \(metadata.applicantName.uppercased())
        Applicant
    
    -and-
    
    \(metadata.respondentName.uppercased())
        Respondent
    
    
    """
    let attr = NSMutableAttributedString(string: content)
    applyHeaderStyle(to: attr)
    return attr
}

private func applyBaseFont(to doc: NSMutableAttributedString, range: NSRange) {
    let baseFont = NSFont(name: LegalFormattingDefaults.fontFamily, size: LegalFormattingDefaults.fontSize)
        ?? NSFont.systemFont(ofSize: LegalFormattingDefaults.fontSize)
    
    doc.enumerateAttribute(.font, in: range, options: []) { value, subRange, _ in
        if let existing = value as? NSFont {
            // Preserve existing traits (e.g., bold/italic) while forcing family/size
            let existingTraits = existing.fontDescriptor.symbolicTraits
            let targetDesc = baseFont.fontDescriptor.withSymbolicTraits(existingTraits)
            let newFont = NSFont(descriptor: targetDesc, size: LegalFormattingDefaults.fontSize) ?? baseFont
            doc.addAttribute(.font, value: newFont, range: subRange)
        } else {
            doc.addAttribute(.font, value: baseFont, range: subRange)
        }
    }
}

private func boldEverything(in str: NSMutableAttributedString) {
    let range = NSRange(location: 0, length: str.length)
    applyBold(to: str, range: range)
}

private func applyBold(to str: NSMutableAttributedString, range: NSRange) {
    let font = NSFont(name: LegalFormattingDefaults.fontFamily, size: LegalFormattingDefaults.fontSize)
        ?? NSFont.systemFont(ofSize: LegalFormattingDefaults.fontSize)
    let boldDesc = font.fontDescriptor.withSymbolicTraits(.bold)
    let boldFont = NSFont(descriptor: boldDesc, size: LegalFormattingDefaults.fontSize) ?? font
    str.addAttribute(.font, value: boldFont, range: range)
}

private func containsAny(_ text: String, keywords: [String]) -> Bool {
    for k in keywords { if text.contains(k) { return true } }
    return false
}

private func isHeading(_ text: String) -> Bool {
    let clean = text.trimmingCharacters(in: .whitespaces)
    guard clean.count < 100 else { return false }
    return clean.range(of: "^[A-Z0-9]+\\.\\s", options: .regularExpression) != nil
}
