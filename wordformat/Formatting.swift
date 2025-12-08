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
    static let bodyLineHeightMultiple: CGFloat = 1.0 // Standard spacing
    static let bodyParagraphSpacing: CGFloat = 12.0  // Space after paragraph
    
    // Indentation
    static let hangingIndent: CGFloat = 36.0         // 1.27cm approx
}

// MARK: - Main Entry Point

func applyUKLegalFormatting(
    to document: NSMutableAttributedString,
    header: LegalHeaderMetadata,
    analysis: AnalysisResult?
) {
    guard document.length > 0 else { return }
    
    // 1. Analyse Structure
    // We scan the document to map out what is Header, Body, Heading, etc.
    let structure = detectDocumentStructure(in: document)
    
    // 2. Insert Header ONLY if totally missing (idempotent check)
    // If the doc starts with "Case No" or "IN THE", we assume it has a header.
    let hasHeader = structure.first?.type == .headerMetadata
    if !hasHeader && !header.caseReference.isEmpty {
        insertGeneratedHeader(to: document, metadata: header)
        // Re-scan required after insertion
        let newStructure = detectDocumentStructure(in: document)
        applyFormatting(to: document, structure: newStructure)
    } else {
        applyFormatting(to: document, structure: structure)
    }
    
    // 3. Global Font Normalisation (Last step to ensure uniformity)
    // Enforce Times New Roman 12pt everywhere, preserving Bold/Italic traits.
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
    
    // Keywords for detecting zones
    let headerKeywords = ["case no", "case ref", "claim no", "in the", "tribunal", "between:", "applicant", "respondent", "-v-", "-and-"]
    let titleKeywords = ["witness statement"]
    let introKeywords = ["will say as follows", "states as follows", "say as follows"]
    let truthKeywords = ["statement of truth", "believe that the facts"]
    
    enum ScanState {
        case header      // Top of doc
        case preBody     // Between title and "will say as follows"
        case body        // The numbered paragraphs
        case backMatter  // Statement of Truth / Signature
    }
    
    var currentState: ScanState = .header
    
    // We iterate by paragraphs
    fullString.enumerateSubstrings(in: fullRange, options: .byParagraphs) { substring, substringRange, _, _ in
        guard let rawText = substring else { return }
        let text = rawText.trimmingCharacters(in: .whitespacesAndNewlines)
        let lowerText = text.lowercased()
        
        // Skip empty lines (but keep them in the range list to preserve spacing)
        if text.isEmpty {
            ranges.append(AnalysisResult.FormattedRange(range: substringRange, type: .unknown))
            return
        }
        
        var type: LegalParagraphType = .body // Default
        
        switch currentState {
        case .header:
            // logic: If we hit "Witness Statement", switch to Title.
            // If we are still at the top and see "Case No" or "Applicant", it's Header.
            if containsAny(lowerText, keywords: titleKeywords) {
                type = .documentTitle
                currentState = .preBody
            } else if containsAny(lowerText, keywords: headerKeywords) || substringRange.location < 800 {
                // Generous limit for header detection (800 chars)
                type = .headerMetadata
            } else {
                // If we drifted too far, assume Intro
                type = .intro
                currentState = .preBody
            }
            
        case .preBody:
            if containsAny(lowerText, keywords: introKeywords) {
                type = .intro
                currentState = .body // Next paragraph is definitely body
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

// MARK: - Styling & Numbering Application

private func applyFormatting(
    to document: NSMutableAttributedString,
    structure: [AnalysisResult.FormattedRange]
) {
    // We must process in REVERSE order because we are inserting text (numbers).
    // If we process forwards, the ranges for later paragraphs will shift and become invalid.
    var bodyCounter = 1
    
    // First pass: Calculate which items are body paragraphs to assign numbers correctly (forward pass)
    // We need to know the number *before* we iterate backwards.
    var bodyIndices: [Int] = []
    for (index, item) in structure.enumerated() {
        if item.type == .body {
            bodyIndices.append(index)
        }
    }
    
    let totalBodyParagraphs = bodyIndices.count
    
    // Second pass: Apply changes in REVERSE order
    for (index, item) in structure.enumerated().reversed() {
        let style = NSMutableParagraphStyle()
        style.lineHeightMultiple = LegalFormattingDefaults.bodyLineHeightMultiple
        style.paragraphSpacing = LegalFormattingDefaults.bodyParagraphSpacing
        
        switch item.type {
        case .headerMetadata:
            style.alignment = .center
            style.paragraphSpacing = 0
            // Bold specific lines (Court name, Parties)
            boldImportantHeaderLines(in: document, range: item.range)
            
        case .documentTitle:
            style.alignment = .center
            style.paragraphSpacing = 24
            applyTrait(.boldFontMask, to: document, range: item.range)
            uppercaseRange(document, range: item.range)
            
        case .intro:
            style.alignment = .left
            style.firstLineHeadIndent = 0
            
        case .heading:
            style.alignment = .left
            style.paragraphSpacingBefore = 18
            applyTrait(.boldFontMask, to: document, range: item.range)
            // Ensure Heading doesn't get a number
            
        case .body:
            // HARD NUMBERING INSERTION
            // We calculate the number for this specific paragraph based on its position in the body list
            if let order = bodyIndices.firstIndex(of: index) {
                let paragraphNumber = order + 1
                let numberString = "\(paragraphNumber).\t"
                
                // Insert the number string at the start of the range
                document.insert(NSAttributedString(string: numberString), at: item.range.location)
                
                // Adjust paragraph style for hanging indent
                style.alignment = .justified
                style.headIndent = LegalFormattingDefaults.hangingIndent
                style.firstLineHeadIndent = 0
                style.tabStops = [NSTextTab(textAlignment: .left, location: LegalFormattingDefaults.hangingIndent, options: [:])]
            }
            
        case .quote:
            style.alignment = .left
            style.headIndent = LegalFormattingDefaults.hangingIndent
            style.firstLineHeadIndent = LegalFormattingDefaults.hangingIndent
            
        case .statementOfTruth:
            style.alignment = .left
            style.paragraphSpacingBefore = 24
            applyTrait(.boldFontMask, to: document, range: item.range)
            
        case .signature:
            style.alignment = .left
            
        case .unknown:
            break
        }
        
        // The range length might have changed if we inserted text, but since we work backwards,
        // the `item.range.location` is stable for the *current* item.
        // However, we must account for the inserted length for the *style* application.
        // For simplicity in this reverse loop, we calculate the range length dynamically if needed,
        // or just apply to the range starting at location.
        // Better safety: Just apply paragraph style to the paragraph covering that location.
        
        let safeLength = (document.string as NSString).paragraphRange(for: NSRange(location: item.range.location, length: 0)).length
        let applyRange = NSRange(location: item.range.location, length: safeLength)
        
        document.addAttribute(.paragraphStyle, value: style, range: applyRange)
    }
}

// MARK: - Helper Functions

private func containsAny(_ text: String, keywords: [String]) -> Bool {
    for k in keywords {
        if text.contains(k) { return true }
    }
    return false
}

private func isHeading(_ text: String) -> Bool {
    // Detects "A. INTRODUCTION" or "1. BACKGROUND" patterns
    let clean = text.trimmingCharacters(in: .whitespaces)
    guard clean.count < 100 else { return false }
    
    // Regex for Heading: Start of line, 1-2 chars (letter/digit), dot, space
    // e.g. "A. ", "1. ", "IV. "
    let pattern = "^[A-Za-z0-9IVX]+\\.\\s"
    if clean.range(of: pattern, options: .regularExpression) != nil {
        // It looks like a heading. Ensure it's not just a numbered paragraph we haven't processed yet.
        // Usually headings are Uppercase or Title Case.
        let upperCount = clean.filter { $0.isUppercase }.count
        if Double(upperCount) / Double(clean.count) > 0.4 { return true }
    }
    return false
}

private func boldImportantHeaderLines(in document: NSMutableAttributedString, range: NSRange) {
    let text = (document.string as NSString).substring(with: range)
    let lower = text.lowercased()
    
    // Bold specific keywords in the header
    if lower.contains("in the") || lower.contains("witness statement") || lower.contains("between") {
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
        
        // Ensure size
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
