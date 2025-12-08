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
    
    // Line spacing
    static let bodyLineHeightMultiple: CGFloat = 1.5   // Standard 1.5 spacing
    static let singleLineHeightMultiple: CGFloat = 1.0 // For signatures/quotes
    
    // Spacing
    static let bodyParagraphSpacing: CGFloat = 12.0
    static let headingParagraphSpacing: CGFloat = 12.0
    
    // Indentation
    static let numberedHeadIndent: CGFloat = 36.0      // Indent for text after number
    static let numberedFirstLineIndent: CGFloat = 0.0  // Number sits at margin
    static let quoteHeadIndent: CGFloat = 36.0         // Block quote indent
}

// MARK: - Main Entry Point

func applyUKLegalFormatting(
    to document: NSMutableAttributedString,
    header: LegalHeaderMetadata,
    analysis: AnalysisResult?
) {
    guard document.length > 0 else { return }
    
    // 1. Insert Header (Idempotent check inside)
    applyLegalHeader(to: document, metadata: header)
    
    // 2. Classify the document structure
    // If we have AI analysis, use it. Otherwise, use heuristics.
    let structure = analysis?.classifiedRanges.isEmpty == false
        ? analysis!.classifiedRanges
        : detectDocumentStructure(in: document)
    
    // 3. Apply Formatting based on classification
    // We process in reverse order (bottom up) if we were changing text lengths,
    // but since we are only changing attributes, forward is fine.
    for item in structure {
        applyStyle(to: document, range: item.range, type: item.type)
    }
    
    // 4. Global Font Normalisation (Times New Roman 12pt)
    // We do this last to ensure consistency, preserving Bold/Italic traits.
    let fullRange = NSRange(location: 0, length: document.length)
    applyBaseFontFamily(
        to: document,
        in: fullRange,
        familyName: LegalFormattingDefaults.fontFamilyName,
        pointSize: LegalFormattingDefaults.fontSize
    )
}

// MARK: - Structure Detection (Heuristic Fallback)

/// Scans the document to guess where the Body starts and ends.
private func detectDocumentStructure(in document: NSAttributedString) -> [AnalysisResult.FormattedRange] {
    var ranges: [AnalysisResult.FormattedRange] = []
    let fullString = document.string as NSString
    let fullRange = NSRange(location: 0, length: fullString.length)
    
    // Markers for heuristic detection
    // Note: These strings should be robust enough to catch standard phrasing
    let startMarker = "will say as follows"
    let endMarker = "statement of truth"
    let beliefMarker = "i believe that the facts"
    
    var bodyStartIndex = 0
    var bodyEndIndex = fullString.length
    
    // 1. Find Start of Body
    let rangeOfStart = fullString.range(of: startMarker, options: .caseInsensitive)
    if rangeOfStart.location != NSNotFound {
        // The body usually starts the paragraph *after* "will say as follows"
        let lineRange = fullString.lineRange(for: rangeOfStart)
        bodyStartIndex = NSMaxRange(lineRange)
    }
    
    // 2. Find End of Body (Statement of Truth)
    let rangeOfTruth = fullString.range(of: endMarker, options: .caseInsensitive)
    let rangeOfBelief = fullString.range(of: beliefMarker, options: .caseInsensitive)
    
    // Pick the earliest occurrence of an end marker
    let foundEndLocation: Int?
    if rangeOfTruth.location != NSNotFound && rangeOfBelief.location != NSNotFound {
        foundEndLocation = min(rangeOfTruth.location, rangeOfBelief.location)
    } else if rangeOfTruth.location != NSNotFound {
        foundEndLocation = rangeOfTruth.location
    } else if rangeOfBelief.location != NSNotFound {
        foundEndLocation = rangeOfBelief.location
    } else {
        foundEndLocation = nil
    }
    
    if let endLoc = foundEndLocation {
        // The body ends at the start of the paragraph containing the marker
        let lineRange = fullString.paragraphRange(for: NSRange(location: endLoc, length: 0))
        bodyEndIndex = lineRange.location
    }
    
    // 3. Classify Paragraphs
    fullString.enumerateSubstrings(in: fullRange, options: .byParagraphs) { substring, substringRange, _, _ in
        guard let text = substring?.trimmingCharacters(in: .whitespacesAndNewlines), !text.isEmpty else { return }
        
        var type: LegalParagraphType = .unknown
        
        if substringRange.location < bodyStartIndex {
            // Front Matter
            if text.lowercased().contains("witness statement") {
                type = .documentTitle
            } else {
                type = .intro
            }
        } else if substringRange.location >= bodyEndIndex {
            // Back Matter
            if text.lowercased().contains("believe") || text.lowercased().contains("truth") {
                type = .statementOfTruth
            } else {
                type = .signature
            }
        } else {
            // Main Body Area
            if isHeading(text) {
                type = .heading
            } else if isQuote(text) {
                type = .quote
            } else {
                type = .body // Numbered paragraph
            }
        }
        
        ranges.append(AnalysisResult.FormattedRange(range: substringRange, type: type))
    }
    
    return ranges
}

/// Simple helper to guess if a line is a Heading (e.g., "A. INTRODUCTION")
private func isHeading(_ text: String) -> Bool {
    // Logic: Uppercase, short, starts with A., B., 1., etc.
    let clean = text.trimmingCharacters(in: .whitespaces)
    guard clean.count < 60 else { return false } // Headings typically aren't huge
    
    // Heuristic: Is largely uppercase?
    let letters = clean.filter { $0.isLetter }
    let upper = letters.filter { $0.isUppercase }
    if !letters.isEmpty && Double(upper.count) / Double(letters.count) > 0.8 {
        return true
    }
    return false
}

/// Simple helper to guess if a line is a quote
private func isQuote(_ text: String) -> Bool {
    // This is hard to guess without context/AI.
    // For now, assume false unless specifically flagged by AI.
    return false
}

// MARK: - Styling Logic

private func applyStyle(to document: NSMutableAttributedString, range: NSRange, type: LegalParagraphType) {
    let style = NSMutableParagraphStyle()
    
    // -- Defaults --
    style.lineHeightMultiple = LegalFormattingDefaults.bodyLineHeightMultiple
    style.paragraphSpacing = LegalFormattingDefaults.bodyParagraphSpacing
    style.alignment = .justified // UK legal usually justified for body
    
    // -- Specifics --
    switch type {
    case .headerMetadata:
        // Already handled by applyLegalHeader, but just in case:
        style.alignment = .center
        style.lineHeightMultiple = 1.1
        
    case .documentTitle:
        style.alignment = .center
        style.paragraphSpacing = 24
        // Make Bold
        applyTrait(.boldFontMask, to: document, range: range)
        
    case .intro:
        style.alignment = .left
        style.firstLineHeadIndent = 0
        style.paragraphSpacing = 12
        
    case .body:
        // Numbered List
        style.alignment = .justified
        style.headIndent = LegalFormattingDefaults.numberedHeadIndent
        style.firstLineHeadIndent = LegalFormattingDefaults.numberedFirstLineIndent
        
        // Create the list numbering
        let list = NSTextList(markerFormat: .decimal, options: 0)
        // Note: prepending "\t" is sometimes needed for NSAttributedString to render the gap
        // But NSTextList usually handles the indent via headIndent.
        style.textLists = [list]
        
    case .heading:
        style.alignment = .left
        style.paragraphSpacing = 6
        style.paragraphSpacingBefore = 18
        // Make Bold
        applyTrait(.boldFontMask, to: document, range: range)
        
    case .quote:
        style.alignment = .left
        style.headIndent = LegalFormattingDefaults.quoteHeadIndent
        style.firstLineHeadIndent = LegalFormattingDefaults.quoteHeadIndent
        style.lineHeightMultiple = LegalFormattingDefaults.singleLineHeightMultiple
        
    case .statementOfTruth:
        style.alignment = .justified
        style.paragraphSpacingBefore = 24
        
    case .signature:
        style.alignment = .left
        style.lineHeightMultiple = 1.0
        style.paragraphSpacing = 4
        
    case .unknown:
        break
    }
    
    document.addAttribute(.paragraphStyle, value: style, range: range)
}

private func applyTrait(_ trait: NSFontTraitMask, to document: NSMutableAttributedString, range: NSRange) {
    document.enumerateAttribute(.font, in: range, options: []) { value, subRange, _ in
        if let font = value as? NSFont {
            let newFont = NSFontManager.shared.convert(font, toHaveTrait: trait)
            document.addAttribute(.font, value: newFont, range: subRange)
        } else {
            // If no font set, apply trait to a base font using defaults
            let base = NSFont(name: LegalFormattingDefaults.fontFamilyName, size: LegalFormattingDefaults.fontSize) ?? NSFont.systemFont(ofSize: LegalFormattingDefaults.fontSize)
            let newFont = NSFontManager.shared.convert(base, toHaveTrait: trait)
            document.addAttribute(.font, value: newFont, range: subRange)
        }
    }
}


// MARK: - Font Normalisation

private func applyBaseFontFamily(
    to document: NSMutableAttributedString,
    in range: NSRange,
    familyName: String,
    pointSize: CGFloat
) {
    let manager = NSFontManager.shared
    
    document.enumerateAttribute(.font, in: range, options: []) { value, subrange, _ in
        let existingFont = (value as? NSFont) ?? NSFont.systemFont(ofSize: pointSize)
        
        // 1. Convert to Family
        var converted = manager.convert(existingFont, toFamily: familyName)
        
        // 2. Enforce Size (convert doesn't always guarantee size)
        let descriptor = converted.fontDescriptor.withSize(pointSize)
        if let sizedFont = NSFont(descriptor: descriptor, size: pointSize) {
            converted = sizedFont
        }
        
        document.removeAttribute(.font, range: subrange)
        document.addAttribute(.font, value: converted, range: subrange)
    }
}

// MARK: - Header Construction

private func applyLegalHeader(to document: NSMutableAttributedString, metadata: LegalHeaderMetadata) {
    let prefixLength = min(200, document.length)
    let prefix = (document.string as NSString).substring(to: prefixLength)
    let headerMarker = "IN THE \(metadata.tribunalName.uppercased())"
    
    if prefix.contains(headerMarker) { return } // Already has header
    
    let baseFont = NSFont(name: LegalFormattingDefaults.fontFamilyName, size: LegalFormattingDefaults.fontSize)
        ?? NSFont.systemFont(ofSize: LegalFormattingDefaults.fontSize)
    
    let headerString = makeLegalHeaderString(metadata: metadata, baseFont: baseFont)
    document.insert(headerString, at: 0)
}

private func makeLegalHeaderString(
    metadata: LegalHeaderMetadata,
    baseFont: NSFont
) -> NSMutableAttributedString {
    
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
    
    let paragraphStyle = NSMutableParagraphStyle()
    paragraphStyle.alignment = .center
    paragraphStyle.lineHeightMultiple = 1.1
    paragraphStyle.paragraphSpacing = 6
    
    attr.addAttributes(
        [.font: baseFont,
         .paragraphStyle: paragraphStyle],
        range: NSRange(location: 0, length: attr.length)
    )
    
    let boldFont = NSFontManager.shared.convert(baseFont, toHaveTrait: .boldFontMask)
    
    func boldFirstOccurrence(of text: String) {
        if let range = (attr.string as NSString).range(of: text, options: .caseInsensitive).toOptional() {
            attr.addAttribute(.font, value: boldFont, range: range)
        }
    }
    
    boldFirstOccurrence(of: metadata.tribunalName.uppercased())
    boldFirstOccurrence(of: metadata.applicantName.uppercased())
    boldFirstOccurrence(of: metadata.respondentName.uppercased())
    
    return attr
}

private extension NSRange {
    func toOptional() -> NSRange? {
        if location == NSNotFound { return nil }
        return self
    }
}
