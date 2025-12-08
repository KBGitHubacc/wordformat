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
    static let fontPostScriptName = "TimesNewRomanPSMT"
    static let fontSize: CGFloat = 12.0
    
    // Layout
    static let bodyLineHeightMultiple: CGFloat = 1.2 // Standard legal line spacing (approx 1.5 equivalent in Word)
    static let bodyParagraphSpacing: CGFloat = 12.0
    
    // Indentation for Numbered Lists
    // "1." sits at 0. Text starts at 36.
    static let numberedHeadIndent: CGFloat = 36.0
    static let numberedFirstLineHeadIndent: CGFloat = 0.0
}

// MARK: - Main Entry Point

func applyUKLegalFormatting(
    to document: NSMutableAttributedString,
    header: LegalHeaderMetadata,
    analysis: AnalysisResult?
) {
    guard document.length > 0 else { return }
    
    // 1. Analyse Structure (Heuristic State Machine)
    // We do this first so we know what everything is (Header, Title, Body, etc.)
    let structure = detectDocumentStructure(in: document)
    
    // 2. Check if we found an existing header
    let hasExistingHeader = structure.contains { $0.type == .headerMetadata }
    
    // 3. Insert Header ONLY if completely missing (and user provided metadata)
    if !hasExistingHeader && !header.caseReference.isEmpty {
        insertGeneratedHeader(to: document, metadata: header)
        // Re-run detection since indices shifted
        let newStructure = detectDocumentStructure(in: document)
        applyStructureStyles(to: document, structure: newStructure)
    } else {
        // Just format what is there
        applyStructureStyles(to: document, structure: structure)
    }
    
    // 4. Global Font Normalisation
    // Convert everything to Times New Roman 12pt, preserving Bold/Italic traits
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
    
    // -- Keywords for Detection --
    let headerKeywords = ["case no", "case ref", "claim no", "in the", "tribunal", "sitting at", "between:", "applicant", "respondent", "-v-", "-and-", "date:"]
    let titleKeywords = ["witness statement"]
    let introKeywords = ["will say as follows", "states as follows", "say as follows"]
    let truthKeywords = ["statement of truth", "believe that the facts", "believes that the facts"]
    
    // -- State Machine --
    enum ScanState {
        case header      // Top of doc, looking for court/parties
        case preBody     // Found title, looking for start of paragraphs
        case body        // The main numbered content
        case backMatter  // Statement of truth, signature
    }
    
    var currentState: ScanState = .header
    
    fullString.enumerateSubstrings(in: fullRange, options: .byParagraphs) { substring, substringRange, _, _ in
        guard let rawText = substring else { return }
        let text = rawText.trimmingCharacters(in: .whitespacesAndNewlines).lowercased()
        
        // Skip purely empty lines (but keep them in ranges to ensure continuity if needed,
        // though usually we style the text content)
        if text.isEmpty { return }
        
        var type: LegalParagraphType = .body // Default assumption
        
        switch currentState {
        case .header:
            // Heuristic: If it looks like header metadata OR it's very early in doc
            if textContainsAny(text, keywords: headerKeywords) || substringRange.location < 500 {
                type = .headerMetadata
                
                // Transition: If we hit "Witness Statement", we leave header mode
                if textContainsAny(text, keywords: titleKeywords) {
                    type = .documentTitle
                    currentState = .preBody
                }
            } else {
                // If we drifted too far without hitting title, assume we are in body/intro
                // But let's check if it's the Title first
                if textContainsAny(text, keywords: titleKeywords) {
                    type = .documentTitle
                    currentState = .preBody
                } else {
                    // Fallback to intro
                    type = .intro
                    currentState = .preBody
                }
            }
            
        case .preBody:
            if textContainsAny(text, keywords: introKeywords) {
                type = .intro
                currentState = .body // Next paragraph is definitely body
            } else if textContainsAny(text, keywords: titleKeywords) {
                type = .documentTitle
            } else {
                // It's likely intro text (e.g. "1. I am the claimant...")
                // Note: Sometimes the intro IS the first numbered paragraph.
                // Let's treat it as body if it doesn't match specific intro keywords?
                // For safety, let's call it Intro until we hit "will say" OR a clear Heading.
                type = .intro
                
                // Failsafe: if this paragraph is long (>200 chars), it's probably body already
                if text.count > 200 {
                    type = .body
                    currentState = .body
                }
            }
            
        case .body:
            // Check for end of body
            if textContainsAny(text, keywords: truthKeywords) {
                type = .statementOfTruth
                currentState = .backMatter
            } else if isHeading(text) {
                type = .heading
            } else {
                type = .body
            }
            
        case .backMatter:
            if textContainsAny(text, keywords: truthKeywords) {
                type = .statementOfTruth
            } else {
                type = .signature
            }
        }
        
        ranges.append(AnalysisResult.FormattedRange(range: substringRange, type: type))
    }
    
    return ranges
}

private func textContainsAny(_ text: String, keywords: [String]) -> Bool {
    for k in keywords {
        if text.contains(k) { return true }
    }
    return false
}

// MARK: - Styling Application

private func applyStructureStyles(
    to document: NSMutableAttributedString,
    structure: [AnalysisResult.FormattedRange]
) {
    for item in structure {
        let style = NSMutableParagraphStyle()
        
        // -- Global Defaults --
        style.paragraphSpacing = LegalFormattingDefaults.bodyParagraphSpacing
        style.lineHeightMultiple = LegalFormattingDefaults.bodyLineHeightMultiple
        
        switch item.type {
        case .headerMetadata:
            style.alignment = .center
            style.paragraphSpacing = 6
            // Bold specific lines if they look like names or court
            boldIfImportant(in: document, range: item.range)
            
        case .documentTitle:
            style.alignment = .center
            style.paragraphSpacing = 24
            applyTrait(.boldFontMask, to: document, range: item.range)
            // Ensure uppercase?
            uppercaseText(in: document, range: item.range)
            
        case .intro:
            style.alignment = .left
            style.firstLineHeadIndent = 0
            
        case .heading:
            style.alignment = .left
            style.paragraphSpacingBefore = 18
            applyTrait(.boldFontMask, to: document, range: item.range)
            
        case .body:
            // THE KEY FIX: Numbering
            style.alignment = .justified
            
            // Indentation for the text wrapping
            style.headIndent = LegalFormattingDefaults.numberedHeadIndent
            
            // The number sits at 0
            style.firstLineHeadIndent = LegalFormattingDefaults.numberedFirstLineHeadIndent
            
            // Tab stop to align text after the number
            let tab = NSTextTab(textAlignment: .left, location: LegalFormattingDefaults.numberedHeadIndent, options: [:])
            style.tabStops = [tab]
            
            // Apply the list marker (1, 2, 3...)
            let list = NSTextList(markerFormat: .decimal, options: 0)
            style.textLists = [list]
            
            // Note: In some TextKit implementations, you must prepend "\t" to the string
            // for the number to appear in the tab stop. However, NSTextList usually handles this.
            // If numbering doesn't appear, it's often because the exporter ignores NSTextList.
            // But this is the correct NSAttributedString way.
            
        case .quote:
            style.alignment = .left
            style.headIndent = LegalFormattingDefaults.numberedHeadIndent
            style.firstLineHeadIndent = LegalFormattingDefaults.numberedHeadIndent
            
        case .statementOfTruth:
            style.alignment = .justified
            style.paragraphSpacingBefore = 24
            // Bold the title "Statement of Truth" if inside
            boldSubstring("Statement of Truth", in: document, range: item.range)
            
        case .signature:
            style.alignment = .left
            style.lineHeightMultiple = 1.0
            style.paragraphSpacing = 4
            
        case .unknown:
            break
        }
        
        document.addAttribute(.paragraphStyle, value: style, range: item.range)
    }
}

// MARK: - Helper Formatting Functions

private func boldIfImportant(in document: NSMutableAttributedString, range: NSRange) {
    let text = (document.string as NSString).substring(with: range).lowercased()
    // Bold if it looks like the court name or a party name (heuristic: mostly uppercase or contains specific words)
    // For safety, just bold the "IN THE..." line
    if text.contains("in the") || text.contains("between") || text.contains("witness statement") {
        applyTrait(.boldFontMask, to: document, range: range)
    }
}

private func boldSubstring(_ substring: String, in document: NSMutableAttributedString, range: NSRange) {
    let fullText = document.string as NSString
    let subRange = fullText.range(of: substring, options: .caseInsensitive, range: range)
    if subRange.location != NSNotFound {
        applyTrait(.boldFontMask, to: document, range: subRange)
    }
}

private func uppercaseText(in document: NSMutableAttributedString, range: NSRange) {
    let text = (document.string as NSString).substring(with: range)
    document.replaceCharacters(in: range, with: text.uppercased())
}

private func applyTrait(_ trait: NSFontTraitMask, to document: NSMutableAttributedString, range: NSRange) {
    let manager = NSFontManager.shared
    document.enumerateAttribute(.font, in: range, options: []) { value, subRange, _ in
        if let font = value as? NSFont {
            let newFont = manager.convert(font, toHaveTrait: trait)
            document.addAttribute(.font, value: newFont, range: subRange)
        } else {
            // Default font
            let base = NSFont(name: LegalFormattingDefaults.fontFamilyName, size: LegalFormattingDefaults.fontSize)
                ?? NSFont.systemFont(ofSize: LegalFormattingDefaults.fontSize)
            let newFont = manager.convert(base, toHaveTrait: trait)
            document.addAttribute(.font, value: newFont, range: subRange)
        }
    }
}

private func applyBaseFontFamily(
    to document: NSMutableAttributedString,
    in range: NSRange,
    familyName: String,
    pointSize: CGFloat
) {
    let manager = NSFontManager.shared
    document.enumerateAttribute(.font, in: range, options: []) { value, subrange, _ in
        let existingFont = (value as? NSFont) ?? NSFont.systemFont(ofSize: pointSize)
        
        // Convert to family
        var converted = manager.convert(existingFont, toFamily: familyName)
        
        // Ensure size
        let descriptor = converted.fontDescriptor.withSize(pointSize)
        if let sizedFont = NSFont(descriptor: descriptor, size: pointSize) {
            converted = sizedFont
        }
        
        document.removeAttribute(.font, range: subrange)
        document.addAttribute(.font, value: converted, range: subrange)
    }
}

// MARK: - Header Insertion (Legacy Fallback)

private func insertGeneratedHeader(to document: NSMutableAttributedString, metadata: LegalHeaderMetadata) {
    let baseFont = NSFont(name: LegalFormattingDefaults.fontFamilyName, size: LegalFormattingDefaults.fontSize)
        ?? NSFont.systemFont(ofSize: LegalFormattingDefaults.fontSize)
    
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
    
    attr.addAttributes([.font: baseFont, .paragraphStyle: paragraphStyle], range: NSRange(location: 0, length: attr.length))
    
    // Bold specific parts
    let boldFont = NSFontManager.shared.convert(baseFont, toHaveTrait: .boldFontMask)
    let nsString = attr.string as NSString
    
    ["IN THE", metadata.tribunalName, metadata.applicantName, metadata.respondentName].forEach { pattern in
        let r = nsString.range(of: pattern, options: .caseInsensitive)
        if r.location != NSNotFound {
            attr.addAttribute(.font, value: boldFont, range: r)
        }
    }
    
    document.insert(attr, at: 0)
}

// MARK: - Utilities

private func isHeading(_ text: String) -> Bool {
    // Basic heuristic: Starts with "A.", "1.", and is short
    let clean = text.trimmingCharacters(in: .whitespaces)
    guard clean.count < 80 else { return false }
    
    // Check for "A. " or "1. " pattern at start
    let range = clean.range(of: "^[A-Z0-9]+\\.\\s", options: .regularExpression)
    return range != nil
}
