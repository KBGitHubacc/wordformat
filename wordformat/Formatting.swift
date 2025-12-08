//
//  Formatting.swift
//  wordformat
//
//  Created by Krisztian Bito on 08.12.2025.
//

import Foundation
import AppKit

// MARK: - Defaults / Policy

enum LegalFormattingDefaults {
    static let fontFamilyName = "Times New Roman"
    static let fontSize: CGFloat = 12.0
    
    // Layout
    static let bodyLineHeightMultiple: CGFloat = 1.0
    static let bodyParagraphSpacing: CGFloat = 12.0
    
    // Indentation
    // Standard indent for numbered lists (approx 1.27cm / 0.5 inch)
    static let listHeadIndent: CGFloat = 36.0
    static let listFirstLineIndent: CGFloat = 0.0 // Marker sits here
}

// MARK: - Main Entry Point

func applyUKLegalFormatting(
    to document: NSMutableAttributedString,
    header: LegalHeaderMetadata,
    analysis: AnalysisResult?
) {
    guard document.length > 0 else { return }
    
    // 1. Analyse Structure
    // We rely on the state machine to identify the zones (Header vs Body vs Back Matter)
    let structure = detectDocumentStructure(in: document)
    
    // 2. Header Handling
    // Check if header exists; if not, and we have metadata, insert it.
    let hasHeader = structure.first?.type == .headerMetadata
    if !hasHeader && !header.caseReference.isEmpty {
        insertGeneratedHeader(to: document, metadata: header)
        // Re-scan required after insertion as ranges shifted
        let newStructure = detectDocumentStructure(in: document)
        applyStructureStyles(to: document, structure: newStructure)
    } else {
        applyStructureStyles(to: document, structure: structure)
    }
    
    // 3. Global Font Normalisation
    // Convert everything to Times New Roman 12pt, preserving traits (Bold/Italic)
    let fullRange = NSRange(location: 0, length: document.length)
    applyBaseFontFamily(
        to: document,
        in: fullRange,
        familyName: LegalFormattingDefaults.fontFamilyName,
        pointSize: LegalFormattingDefaults.fontSize
    )
}

// MARK: - Structure Detection (State Machine)

private func detectDocumentStructure(in document: NSAttributedString) -> [AnalysisResult.FormattedRange] {
    var ranges: [AnalysisResult.FormattedRange] = []
    let fullString = document.string as NSString
    let fullRange = NSRange(location: 0, length: fullString.length)
    
    // Keywords
    let headerKeywords = ["case no", "case ref", "claim no", "in the", "tribunal", "between:", "applicant", "respondent", "-v-", "-and-"]
    let titleKeywords = ["witness statement"]
    let introKeywords = ["will say as follows", "states as follows", "say as follows"]
    let truthKeywords = ["statement of truth", "believe that the facts"]
    
    enum ScanState {
        case header
        case preBody
        case body
        case backMatter
    }
    
    var currentState: ScanState = .header
    
    fullString.enumerateSubstrings(in: fullRange, options: .byParagraphs) { substring, substringRange, _, _ in
        guard let rawText = substring else { return }
        let text = rawText.trimmingCharacters(in: .whitespacesAndNewlines)
        let lowerText = text.lowercased()
        
        // Keep empty lines as 'unknown' to preserve spacing, but don't format them
        if text.isEmpty {
            ranges.append(AnalysisResult.FormattedRange(range: substringRange, type: .unknown))
            return
        }
        
        var type: LegalParagraphType = .body
        
        switch currentState {
        case .header:
            if containsAny(lowerText, keywords: titleKeywords) {
                type = .documentTitle
                currentState = .preBody
            } else if containsAny(lowerText, keywords: headerKeywords) || substringRange.location < 800 {
                type = .headerMetadata
            } else {
                type = .intro
                currentState = .preBody
            }
            
        case .preBody:
            if containsAny(lowerText, keywords: introKeywords) {
                type = .intro
                currentState = .body
            } else if containsAny(lowerText, keywords: titleKeywords) {
                type = .documentTitle
            } else {
                type = .intro
            }
            
        case .body:
            if containsAny(lowerText, keywords: truthKeywords) {
                type = .statementOfTruth
                currentState = .backMatter
            } else if isHeading(text) {
                type = .heading
            } else {
                type = .body
            }
            
        case .backMatter:
            if containsAny(lowerText, keywords: truthKeywords) {
                type = .statementOfTruth
            } else {
                type = .signature
            }
        }
        
        ranges.append(AnalysisResult.FormattedRange(range: substringRange, type: type))
    }
    
    return ranges
}

// MARK: - Styling & Native Numbering

private func applyStructureStyles(
    to document: NSMutableAttributedString,
    structure: [AnalysisResult.FormattedRange]
) {
    // 1. Create a Single List Instance
    // NSTextList objects must be reused across paragraphs to maintain sequential numbering (1, 2, 3...)
    // If we create a new NSTextList for every paragraph, they will all be "1.".
    let masterNumberingList = NSTextList(markerFormat: .decimal, options: 0)
    masterNumberingList.startingItemNumber = 1
    
    // We will assume sub-lists might restart, so we create them as needed or keep a running one.
    // For simplicity in legal docs, (a) usually restarts under each number.
    var currentSubList: NSTextList? = nil
    
    // We process in REVERSE to handle text deletion (stripping old numbers) without invalidating future ranges.
    // However, since we are stripping text, we must be careful with the 'structure' ranges.
    // It is safer to calculate the range dynamically or update offsets.
    // Simplest approach: Process Reverse.
    
    for item in structure.reversed() {
        let style = NSMutableParagraphStyle()
        style.paragraphSpacing = LegalFormattingDefaults.bodyParagraphSpacing
        style.lineHeightMultiple = LegalFormattingDefaults.bodyLineHeightMultiple
        
        // We need the current text to check for patterns
        let currentRange = item.range
        // Note: Because we are iterating backwards, the 'location' of earlier items remains valid.
        // The 'length' might technically change if we modify *this* paragraph, but we handle that locally.
        let paragraphText = (document.string as NSString).substring(with: currentRange)
        
        switch item.type {
        case .headerMetadata:
            style.alignment = .center
            style.paragraphSpacing = 0
            boldImportantHeaderLines(in: document, range: currentRange)
            document.addAttribute(.paragraphStyle, value: style, range: currentRange)
            
        case .documentTitle:
            style.alignment = .center
            style.paragraphSpacing = 24
            applyTrait(.boldFontMask, to: document, range: currentRange)
            uppercaseRange(document, range: currentRange)
            document.addAttribute(.paragraphStyle, value: style, range: currentRange)
            
        case .intro:
            style.alignment = .left
            document.addAttribute(.paragraphStyle, value: style, range: currentRange)
            
        case .heading:
            style.alignment = .left
            style.paragraphSpacingBefore = 18
            applyTrait(.boldFontMask, to: document, range: currentRange)
            document.addAttribute(.paragraphStyle, value: style, range: currentRange)
            
        case .body:
            // Detect if this paragraph HAS a number to replace, or is a continuation.
            // Regex for "1." or "2)"
            let level1Pattern = "^\\s*(\\d+)[.)]\\s+"
            // Regex for "(a)" or "a." or "a)"
            let level2Pattern = "^\\s*\\(?([a-zA-Z])[.)]\\s+"
            
            if let match1 = rangeOfPattern(level1Pattern, in: paragraphText) {
                // LEVEL 1: Numbered Paragraph
                // 1. Strip the old text number ("1. ")
                let globalMatchRange = NSRange(location: currentRange.location + match1.location, length: match1.length)
                document.replaceCharacters(in: globalMatchRange, with: "")
                
                // 2. Apply Native List Style
                style.headIndent = LegalFormattingDefaults.listHeadIndent
                style.firstLineHeadIndent = LegalFormattingDefaults.listFirstLineIndent
                style.textLists = [masterNumberingList]
                
                // Reset sublist context because we hit a new main number
                currentSubList = nil
                
                // Re-calculate range since we shortened the text
                let newLength = currentRange.length - match1.length
                let fixRange = NSRange(location: currentRange.location, length: newLength)
                document.addAttribute(.paragraphStyle, value: style, range: fixRange)
                
            } else if let match2 = rangeOfPattern(level2Pattern, in: paragraphText) {
                // LEVEL 2: Sub-paragraph
                // 1. Strip the old text marker ("(a) ")
                let globalMatchRange = NSRange(location: currentRange.location + match2.location, length: match2.length)
                document.replaceCharacters(in: globalMatchRange, with: "")
                
                // 2. Create or Reuse Sublist
                // In legal docs, sublists usually reset per paragraph.
                // However, if we are going backwards, we can't easily know the "parent".
                // Strategy: Use a generic alpha list.
                if currentSubList == nil {
                    if #available(macOS 14.0, *) {
                        currentSubList = NSTextList(markerFormat: .lowercaseLatin, options: 0)
                    } else {
                        // Fallback to decimal or roman if lowercaseLatin is unavailable
                        currentSubList = NSTextList(markerFormat: .decimal, options: 0)
                    }
                    // Note: NSTextList does not render parentheses; this will typically appear as "a.".
                }
                
                style.headIndent = LegalFormattingDefaults.listHeadIndent + 36.0 // Indent further
                style.firstLineHeadIndent = LegalFormattingDefaults.listHeadIndent
                // Nesting: Root -> Sub
                style.textLists = [masterNumberingList, currentSubList!]
                
                let newLength = currentRange.length - match2.length
                let fixRange = NSRange(location: currentRange.location, length: newLength)
                document.addAttribute(.paragraphStyle, value: style, range: fixRange)
                
            } else {
                // CONTINUATION PARAGRAPH (No number detected)
                // Just align it with the text of the numbered items
                style.alignment = .justified
                style.headIndent = LegalFormattingDefaults.listHeadIndent
                style.firstLineHeadIndent = LegalFormattingDefaults.listHeadIndent
                // DO NOT add textLists, so it has no number but looks aligned
                document.addAttribute(.paragraphStyle, value: style, range: currentRange)
            }
            
        case .quote:
            style.alignment = .left
            style.headIndent = LegalFormattingDefaults.listHeadIndent * 2
            style.firstLineHeadIndent = LegalFormattingDefaults.listHeadIndent * 2
            document.addAttribute(.paragraphStyle, value: style, range: currentRange)
            
        case .statementOfTruth:
            style.alignment = .left
            style.paragraphSpacingBefore = 24
            applyTrait(.boldFontMask, to: document, range: currentRange)
            document.addAttribute(.paragraphStyle, value: style, range: currentRange)
            
        case .signature:
            style.alignment = .left
            document.addAttribute(.paragraphStyle, value: style, range: currentRange)
            
        case .unknown:
            break
        }
    }
}

// MARK: - Utilities & Helpers

private func rangeOfPattern(_ pattern: String, in text: String) -> NSRange? {
    guard let regex = try? NSRegularExpression(pattern: pattern, options: []) else { return nil }
    let range = NSRange(location: 0, length: text.utf16.count)
    if let match = regex.firstMatch(in: text, options: [], range: range) {
        return match.range
    }
    return nil
}

private func containsAny(_ text: String, keywords: [String]) -> Bool {
    for k in keywords {
        if text.contains(k) { return true }
    }
    return false
}

private func isHeading(_ text: String) -> Bool {
    let clean = text.trimmingCharacters(in: .whitespaces)
    guard clean.count < 100 else { return false }
    // Detects "A. INTRODUCTION"
    let pattern = "^[A-Za-z0-9IVX]+\\.\\s"
    return rangeOfPattern(pattern, in: clean) != nil
}

private func boldImportantHeaderLines(in document: NSMutableAttributedString, range: NSRange) {
    let text = (document.string as NSString).substring(with: range).lowercased()
    if text.contains("in the") || text.contains("witness statement") || text.contains("between") {
        applyTrait(.boldFontMask, to: document, range: range)
    }
}

private func uppercaseRange(_ document: NSMutableAttributedString, range: NSRange) {
    let text = (document.string as NSString).substring(with: range)
    document.replaceCharacters(in: range, with: text.uppercased())
}

private func applyTrait(_ trait: NSFontTraitMask, to document: NSMutableAttributedString, range: NSRange) {
    let manager = NSFontManager.shared
    document.enumerateAttribute(.font, in: range, options: []) { value, subRange, _ in
        let font = (value as? NSFont) ?? NSFont.systemFont(ofSize: 12)
        let newFont = manager.convert(font, toHaveTrait: trait)
        document.addAttribute(.font, value: newFont, range: subRange)
    }
}

private func applyBaseFontFamily(to document: NSMutableAttributedString, in range: NSRange, familyName: String, pointSize: CGFloat) {
    let manager = NSFontManager.shared
    document.enumerateAttribute(.font, in: range, options: []) { value, subrange, _ in
        let currentFont = (value as? NSFont) ?? NSFont.systemFont(ofSize: pointSize)
        var newFont = manager.convert(currentFont, toFamily: familyName)
        let desc = newFont.fontDescriptor.withSize(pointSize)
        if let sized = NSFont(descriptor: desc, size: pointSize) {
            newFont = sized
        }
        document.removeAttribute(.font, range: subrange)
        document.addAttribute(.font, value: newFont, range: subrange)
    }
}

private func insertGeneratedHeader(to document: NSMutableAttributedString, metadata: LegalHeaderMetadata) {
    let headerText = """
    IN THE \(metadata.tribunalName.uppercased())
    Case Reference: \(metadata.caseReference)
    
    BETWEEN:
    
    \(metadata.applicantName.uppercased())
        Applicant
    
    -and-
    
    \(metadata.respondentName.uppercased())
        Respondent
    
    
    """
    let attr = NSMutableAttributedString(string: headerText)
    document.insert(attr, at: 0)
}

