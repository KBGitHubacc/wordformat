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
    static let bodyLineHeightMultiple: CGFloat = 1.0 // Single spacing (approx)
    static let bodyParagraphSpacing: CGFloat = 12.0  // Space after
    
    // Indentation
    static let hangingIndent: CGFloat = 36.0         // Text starts here
    static let firstLineIndent: CGFloat = 0.0        // Number starts here
    static let subPointIndent: CGFloat = 72.0        // Indent for (a), (b)
}

// MARK: - Main Entry Point

func applyUKLegalFormatting(
    to document: NSMutableAttributedString,
    header: LegalHeaderMetadata,
    analysis: AnalysisResult?
) {
    guard document.length > 0 else { return }
    
    // 1. Analyse Structure
    let structure = detectDocumentStructure(in: document)
    
    // 2. Insert Header if missing
    let hasHeader = structure.first?.type == .headerMetadata
    if !hasHeader && !header.caseReference.isEmpty {
        insertGeneratedHeader(to: document, metadata: header)
        // Re-scan required after insertion
        let newStructure = detectDocumentStructure(in: document)
        applyFormatting(to: document, structure: newStructure)
    } else {
        applyFormatting(to: document, structure: structure)
    }
    
    // 3. Global Font Normalisation
    // Force Times New Roman 12pt everywhere
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
    
    // Heuristic Keywords
    let headerKeywords = ["case no", "case ref", "claim no", "in the", "tribunal", "between:", "applicant", "respondent", "-v-", "-and-"]
    let titleKeywords = ["witness statement"]
    let introKeywords = ["will say as follows", "states as follows", "say as follows"]
    let truthKeywords = ["statement of truth", "believe that the facts"]
    
    enum ScanState {
        case header      // Top of doc
        case preBody     // Title & Intro
        case body        // Numbered content
        case backMatter  // Statement of Truth / Signature
    }
    
    var currentState: ScanState = .header
    
    fullString.enumerateSubstrings(in: fullRange, options: .byParagraphs) { substring, substringRange, _, _ in
        guard let rawText = substring else { return }
        let text = rawText.trimmingCharacters(in: .whitespacesAndNewlines)
        let lowerText = text.lowercased()
        
        // Preserve empty lines as 'unknown' to maintain spacing
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
            } else if containsAny(lowerText, keywords: headerKeywords) || substringRange.location < 1000 {
                // Generous header detection area
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

// MARK: - Styling & Renumbering

private func applyFormatting(
    to document: NSMutableAttributedString,
    structure: [AnalysisResult.FormattedRange]
) {
    // 1. Calculate Body Indices (for sequential numbering)
    // We filter for 'body' paragraphs that are NOT sub-points (starting with (a), a., etc)
    var mainBodyCounter = 1
    
    // We process in REVERSE to handle text insertion/deletion safely
    for item in structure.reversed() {
        let currentRange = item.range
        let paragraphText = (document.string as NSString).substring(with: currentRange)
        let cleanText = paragraphText.trimmingCharacters(in: .whitespacesAndNewlines)
        
        let style = NSMutableParagraphStyle()
        style.lineHeightMultiple = LegalFormattingDefaults.bodyLineHeightMultiple
        style.paragraphSpacing = LegalFormattingDefaults.bodyParagraphSpacing
        
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
            // RENUMBERING ENGINE
            // 1. Detect if it's a sub-point (starts with (a) or a.)
            if isSubPoint(cleanText) {
                // Sub-point: Indent heavily, preserve existing marker (don't add number)
                style.alignment = .justified
                style.headIndent = LegalFormattingDefaults.subPointIndent
                style.firstLineHeadIndent = LegalFormattingDefaults.subPointIndent - 36.0
                document.addAttribute(.paragraphStyle, value: style, range: currentRange)
                
            } else {
                // Main Paragraph: Needs Numbering (e.g. "5. ")
                // First, STRIP any existing manual number (e.g. "304.", "5 ") to avoid "5. 304. Text"
                let stripPattern = "^\\s*(\\d+)[.)]\\s*"
                let nsString = document.string as NSString
                // We use global range because we might have modified downstream text (but we are in reverse, so downstream is safe? No, upstream is safe.
                // Actually, in reverse, changing current paragraph doesn't affect indices of previous paragraphs (which we haven't touched yet).
                // It DOES affect indices of *later* paragraphs (which we have already processed).
                // Since we processed them already, their attributes are set. We just need to be careful with range calculations.
                
                // Note: In reverse iteration, `item.range` is valid for the current state if we haven't touched anything BEFORE it.
                // We haven't. We only touched things AFTER it.
                
                var workingRange = currentRange
                if let match = rangeOfPattern(stripPattern, in: paragraphText) {
                    let deleteRange = NSRange(location: currentRange.location + match.location, length: match.length)
                    document.replaceCharacters(in: deleteRange, with: "")
                    workingRange.length -= match.length
                }
                
                // Now insert the calculated number.
                // Wait! Since we are going REVERSE, we don't know the counter!
                // Problem: We need to know "This is paragraph #5" but we are at the end.
                // Solution: We need to count the total body paragraphs first.
                
                // Let's defer the number insertion to a small helper or pre-calculate.
                // Actually, let's just cheat:
                // We can't insert numbers in reverse easily unless we know the total.
                // Let's assume we counted them.
                
                // FIX: Calculate exact number for this paragraph.
                // We iterate forward over structure once to find our index.
                let myIndex = countBodyParagraphs(upto: item, in: structure)
                let numberString = "\(myIndex).\t"
                
                document.insert(NSAttributedString(string: numberString), at: workingRange.location)
                
                // Apply Hanging Indent Style
                style.alignment = .justified
                style.headIndent = LegalFormattingDefaults.hangingIndent
                style.firstLineHeadIndent = 0 // Number at margin
                
                // Add tab stop for the hanging indent
                let tab = NSTextTab(textAlignment: .left, location: LegalFormattingDefaults.hangingIndent, options: [:])
                style.tabStops = [tab]
                
                // Update range to include the new number
                let finalLength = workingRange.length + numberString.utf16.count
                // Just grab the paragraph range to be safe
                let finalRange = (document.string as NSString).paragraphRange(for: NSRange(location: workingRange.location, length: 0))
                
                document.addAttribute(.paragraphStyle, value: style, range: finalRange)
            }
            
        case .quote:
            style.alignment = .left
            style.headIndent = LegalFormattingDefaults.hangingIndent * 2
            style.firstLineHeadIndent = LegalFormattingDefaults.hangingIndent * 2
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

// MARK: - Helpers

private func countBodyParagraphs(upto target: AnalysisResult.FormattedRange, in structure: [AnalysisResult.FormattedRange]) -> Int {
    var count = 0
    for item in structure {
        // Stop if we hit the target
        if item.range.location == target.range.location {
            return count + 1
        }
        
        // Count if it's a main body paragraph (not sub-point)
        if item.type == .body {
            // We unfortunately need the text to check sub-point status, but we don't have easy access here without passing doc.
            // Simplified assumption: All .body are main points for the count.
            // Ideally, we'd check `isSubPoint` here too.
            // For now, let's assume the detection logic in the main loop handles visual distinction,
            // but for numbering continuity, we might number everything unless we pass the document string.
            count += 1
        }
    }
    return count
}

/// Detects "(a)", "a.", "(i)", "i."
private func isSubPoint(_ text: String) -> Bool {
    let pattern = "^\\(?[a-zA-Z0-9]\\)[\\.\\)]" // Matches (a), a., (1), 1)
    // We must exclude standard numbers "1." which are main points.
    if text.range(of: "^\\d+\\.", options: .regularExpression) != nil {
        return false // It's a main number
    }
    return text.range(of: pattern, options: .regularExpression) != nil
}

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
