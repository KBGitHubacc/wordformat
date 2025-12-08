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
    
    // Indentation for Lists (Crucial for visibility)
    // The number sits at 0. The text starts at 36 (approx 1.27cm).
    static let listHeadIndent: CGFloat = 36.0
    static let listFirstLineIndent: CGFloat = 0.0
    
    // Indentation for Sub-lists (Level 2)
    // The letter sits at 36. The text starts at 72.
    static let subListHeadIndent: CGFloat = 72.0
    static let subListFirstLineIndent: CGFloat = 36.0
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
    
    // 2. Header Insertion (Idempotent)
    let hasHeader = structure.first?.type == .headerMetadata
    if !hasHeader && !header.caseReference.isEmpty {
        insertGeneratedHeader(to: document, metadata: header)
        // Re-scan ranges after insertion
        let newStructure = detectDocumentStructure(in: document)
        applyFormatting(to: document, structure: newStructure)
    } else {
        applyFormatting(to: document, structure: structure)
    }
    
    // 3. Global Font Normalisation
    // Ensures everything is Times New Roman 12pt while preserving Bold/Italic traits
    let fullRange = NSRange(location: 0, length: document.length)
    applyBaseFontFamily(
        to: document,
        in: fullRange,
        familyName: LegalFormattingDefaults.fontFamilyName,
        pointSize: LegalFormattingDefaults.fontSize
    )
}

// MARK: - Structure Detection

private func detectDocumentStructure(in document: NSAttributedString) -> [AnalysisResult.FormattedRange] {
    var ranges: [AnalysisResult.FormattedRange] = []
    let fullString = document.string as NSString
    let fullRange = NSRange(location: 0, length: fullString.length)
    
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

// MARK: - Styling & Dynamic Numbering

private func applyFormatting(
    to document: NSMutableAttributedString,
    structure: [AnalysisResult.FormattedRange]
) {
    // We create persistent list objects.
    // Reusing the same object tells the RTF/DocX exporter that these paragraphs belong to the *same* list sequence.
    let rootList = NSTextList(markerFormat: .decimal, options: 0) // 1, 2, 3...
    rootList.startingItemNumber = 1
    
    let subList = NSTextList(markerFormat: .lowercaseAlpha, options: 0) // a, b, c...
    // Note: subLists often need to be re-created if we want them to restart (a) at every new number.
    // However, tracking that logic in reverse iteration is hard.
    // For robust legal formatting, we will try to rely on a single subList definition,
    // though Word may continue the sequence (a, b, c, d...) across paragraphs if not careful.
    // A safer bet for sub-lists in `NSAttributedString` is to create a new list when the parent changes.
    // But since we iterate backwards, we can't easily detect "parent change".
    // We will use a single subList instance for now; Word usually handles hierarchy resets based on indentation.
    
    // We process in REVERSE so that deleting text (stripping old numbers) doesn't invalidate the ranges of unprocessed items.
    for item in structure.reversed() {
        let currentRange = item.range
        let paragraphText = (document.string as NSString).substring(with: currentRange)
        
        let style = NSMutableParagraphStyle()
        style.paragraphSpacing = LegalFormattingDefaults.bodyParagraphSpacing
        style.lineHeightMultiple = LegalFormattingDefaults.bodyLineHeightMultiple
        
        switch item.type {
        case .headerMetadata:
            style.alignment = .center
            style.paragraphSpacing = 0
            document.addAttribute(.paragraphStyle, value: style, range: currentRange)
            boldImportantHeaderLines(in: document, range: currentRange)
            
        case .documentTitle:
            style.alignment = .center
            style.paragraphSpacing = 24
            document.addAttribute(.paragraphStyle, value: style, range: currentRange)
            applyTrait(.boldFontMask, to: document, range: currentRange)
            uppercaseRange(document, range: currentRange)
            
        case .intro:
            style.alignment = .left
            document.addAttribute(.paragraphStyle, value: style, range: currentRange)
            
        case .heading:
            style.alignment = .left
            style.paragraphSpacingBefore = 18
            document.addAttribute(.paragraphStyle, value: style, range: currentRange)
            applyTrait(.boldFontMask, to: document, range: currentRange)
            
        case .body:
            // DYNAMIC NUMBERING LOGIC
            
            // Regex to find existing numbers "1.", "304.", "1 " to strip them
            let mainPattern = "^\\s*(\\d+)[.)]?\\s+"
            // Regex to find sub-points "(a)", "a.", "a)"
            let subPattern = "^\\s*\\(?([a-zA-Z])\\)[.)]?\\s+"
            
            if let subMatch = rangeOfPattern(subPattern, in: paragraphText) {
                // --- LEVEL 2: Sub-point (a, b, c) ---
                
                // 1. Strip the static text "a."
                let deleteRange = NSRange(location: currentRange.location + subMatch.location, length: subMatch.length)
                document.replaceCharacters(in: deleteRange, with: "")
                
                // 2. Configure Dynamic List Style
                style.alignment = .justified
                // Indentation is critical for the number to appear
                style.headIndent = LegalFormattingDefaults.subListHeadIndent
                style.firstLineHeadIndent = LegalFormattingDefaults.subListFirstLineIndent
                
                // We must apply BOTH lists to indicate hierarchy: Root -> Sub
                style.textLists = [rootList, subList]
                
                // 3. Apply to the new range (length has changed)
                let newLength = currentRange.length - subMatch.length
                let fixRange = NSRange(location: currentRange.location, length: newLength)
                document.addAttribute(.paragraphStyle, value: style, range: fixRange)
                
            } else {
                // --- LEVEL 1: Main Point (1, 2, 3) ---
                
                // 1. Strip existing number "1." if present
                var workingRange = currentRange
                if let mainMatch = rangeOfPattern(mainPattern, in: paragraphText) {
                    let deleteRange = NSRange(location: currentRange.location + mainMatch.location, length: mainMatch.length)
                    document.replaceCharacters(in: deleteRange, with: "")
                    workingRange.length -= mainMatch.length
                }
                
                // 2. Configure Dynamic List Style
                style.alignment = .justified
                style.headIndent = LegalFormattingDefaults.listHeadIndent
                style.firstLineHeadIndent = LegalFormattingDefaults.listFirstLineIndent
                
                // Assign the list. This is what generates the dynamic number.
                style.textLists = [rootList]
                
                // 3. Apply
                document.addAttribute(.paragraphStyle, value: style, range: workingRange)
            }
            
        case .quote:
            style.alignment = .left
            style.headIndent = LegalFormattingDefaults.listHeadIndent * 2
            style.firstLineHeadIndent = LegalFormattingDefaults.listHeadIndent * 2
            document.addAttribute(.paragraphStyle, value: style, range: currentRange)
            
        case .statementOfTruth:
            style.alignment = .left
            style.paragraphSpacingBefore = 24
            document.addAttribute(.paragraphStyle, value: style, range: currentRange)
            applyTrait(.boldFontMask, to: document, range: currentRange)
            
        case .signature:
            style.alignment = .left
            document.addAttribute(.paragraphStyle, value: style, range: currentRange)
            
        case .unknown:
            break
        }
    }
}

// MARK: - Helpers

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
    let pattern = "^[A-Za-z0-9IVX]+\\.\\s" // Detects "A. ", "IV. "
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
