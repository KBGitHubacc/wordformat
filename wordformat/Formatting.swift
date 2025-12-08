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
    static let fontFamilyName = "Times New Roman"          // Family name used for conversions
    static let fontPostScriptName = "TimesNewRomanPSMT"    // Canonical PostScript name on macOS
    static let fontSize: CGFloat = 12.0
    
    static let bodyLineHeightMultiple: CGFloat = 1.5
    static let bodyParagraphSpacing: CGFloat = 8.0
    
    static let headerLineHeightMultiple: CGFloat = 1.2
    static let headerParagraphSpacing: CGFloat = 6.0
    
    // Simple hanging indent for numbered paragraphs
    static let numberedHeadIndent: CGFloat = 24.0
    static let numberedFirstLineHeadIndent: CGFloat = 0.0
}

// MARK: - Main entry point

func applyUKLegalFormatting(
    to document: NSMutableAttributedString,
    header: LegalHeaderMetadata,
    analysis: AnalysisResult?
) {
    // 0. Sanity check – if there is no text, nothing to format.
    guard document.length > 0 else { return }
    
    // 1. Idempotence guard – if document already appears to have this header, bail out.
    let prefixLength = min(200, document.length)
    let prefix = (document.string as NSString).substring(to: prefixLength)
    let headerMarker = "IN THE \(header.tribunalName.uppercased())"
    
    if prefix.contains(headerMarker) {
        // Assume formatting (including header) has already been applied.
        return
    }
    
    // 2. Prepare base font (used mainly for header construction; we convert everything
    //    to Times New Roman later in a trait-preserving pass).
    let baseFont: NSFont = {
        if let psFont = NSFont(name: LegalFormattingDefaults.fontPostScriptName,
                               size: LegalFormattingDefaults.fontSize) {
            return psFont
        }
        if let familyFont = NSFontManager.shared.font(
            withFamily: LegalFormattingDefaults.fontFamilyName,
            traits: [],
            weight: 5,
            size: LegalFormattingDefaults.fontSize
        ) {
            return familyFont
        }
        return NSFont.systemFont(ofSize: LegalFormattingDefaults.fontSize)
    }()
    
    // 3. Build the UK legal header block and insert at the top.
    let headerString = makeLegalHeaderString(
        metadata: header,
        baseFont: baseFont
    )
    
    let headerLength = headerString.length
    document.insert(headerString, at: 0)
    
    // 4. Apply numbered paragraph formatting to the **body only**.
    let fullRangeAfterHeader = NSRange(
        location: headerLength,
        length: document.length - headerLength
    )
    applyNumberedParagraphs(to: document, in: fullRangeAfterHeader)
    
    // 5. Use analysis (from OpenAI) to style headings, tables, etc. if available.
    if let analysis = analysis {
        applyHeadingStyles(using: analysis, in: document)
        applyTableStyles(using: analysis, in: document)
    } else {
        // Optional: simple heuristic for headings present in the text.
        applySimpleHeuristicHeadings(
            in: document,
            bodyRange: fullRangeAfterHeader,
            baseFont: baseFont
        )
    }
    
    // 6. Apply body line spacing (1.5) and paragraph spacing ONLY to the body text,
    //    preserving other paragraph style properties (alignment, lists, etc.).
    applyBodyLineSpacing(
        to: document,
        in: fullRangeAfterHeader
    )
    
    // 7. Convert all fonts to Times New Roman 12pt while preserving traits
    //    (bold, italic, etc.), including in the header.
    let fullRange = NSRange(location: 0, length: document.length)
    applyBaseFontFamily(
        to: document,
        in: fullRange,
        familyName: LegalFormattingDefaults.fontFamilyName,
        pointSize: LegalFormattingDefaults.fontSize
    )
}

// MARK: - Header construction

/// Build the UK-style legal header.
///
/// Example:
/// IN THE EMPLOYMENT TRIBUNAL LONDON
///
/// Case Reference: 2401234/2025
///
/// BETWEEN:
///
/// JOHN SMITH
///     Applicant
///
/// -and-
///
/// ACME LTD
///     Respondent
///
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
    paragraphStyle.lineHeightMultiple = LegalFormattingDefaults.headerLineHeightMultiple
    paragraphStyle.paragraphSpacing = LegalFormattingDefaults.headerParagraphSpacing
    
    attr.addAttributes(
        [.font: baseFont,
         .paragraphStyle: paragraphStyle],
        range: NSRange(location: 0, length: attr.length)
    )
    
    // Emphasise key lines (tribunal and party names) in bold.
    let boldFont = NSFontManager.shared.convert(baseFont, toHaveTrait: .boldFontMask)
    
    func boldFirstOccurrence(of text: String) {
        if let range = (attr.string as NSString)
            .range(of: text, options: .caseInsensitive)
            .toOptional() {
            attr.addAttribute(.font, value: boldFont, range: range)
        }
    }
    
    boldFirstOccurrence(of: metadata.tribunalName.uppercased())
    boldFirstOccurrence(of: metadata.applicantName.uppercased())
    boldFirstOccurrence(of: metadata.respondentName.uppercased())
    
    return attr
}

// MARK: - Paragraph numbering

/// Apply numbered paragraphs to the body range, skipping blank paragraphs.
/// This uses NSTextList so Word sees it as a proper numbered list.
private func applyNumberedParagraphs(
    to document: NSMutableAttributedString,
    in range: NSRange
) {
    let textList = NSTextList(markerFormat: .decimal, options: 0)
    
    let fullNSString = document.string as NSString
    let bodyString = fullNSString.substring(with: range) as NSString
    
    bodyString.enumerateSubstrings(
        in: NSRange(location: 0, length: bodyString.length),
        options: .byParagraphs
    ) { _, paragraphRange, _, _ in
        
        // Map local range back to global range in the document.
        let globalRange = NSRange(
            location: range.location + paragraphRange.location,
            length: paragraphRange.length
        )
        
        let paragraphText = fullNSString.substring(with: globalRange)
        let trimmed = paragraphText.trimmingCharacters(in: .whitespacesAndNewlines)
        
        // Skip empty paragraphs.
        guard !trimmed.isEmpty else { return }
        
        let currentStyle: NSMutableParagraphStyle
        if let style = document.attribute(.paragraphStyle,
                                          at: globalRange.location,
                                          effectiveRange: nil) as? NSParagraphStyle {
            currentStyle = style.mutableCopy() as! NSMutableParagraphStyle
        } else {
            currentStyle = NSMutableParagraphStyle()
        }
        
        // Attach the text list (numbered list).
        currentStyle.textLists = [textList]
        
        // Optional: hanging indent.
        currentStyle.headIndent = LegalFormattingDefaults.numberedHeadIndent
        currentStyle.firstLineHeadIndent = LegalFormattingDefaults.numberedFirstLineHeadIndent
        
        document.addAttribute(.paragraphStyle, value: currentStyle, range: globalRange)
    }
}

// MARK: - Line spacing for body text

/// Apply 1.5 line spacing and paragraph spacing to body paragraphs,
/// preserving other paragraph style attributes (alignment, lists, etc.).
private func applyBodyLineSpacing(
    to document: NSMutableAttributedString,
    in range: NSRange
) {
    let fullNSString = document.string as NSString
    
    fullNSString.enumerateSubstrings(
        in: range,
        options: .byParagraphs
    ) { _, paragraphRange, _, _ in
        
        // paragraphRange is already in the coordinate space of the full string.
        let globalRange = paragraphRange
        
        let currentStyle: NSMutableParagraphStyle
        if let style = document.attribute(.paragraphStyle,
                                          at: globalRange.location,
                                          effectiveRange: nil) as? NSParagraphStyle {
            currentStyle = style.mutableCopy() as! NSMutableParagraphStyle
        } else {
            currentStyle = NSMutableParagraphStyle()
        }
        
        currentStyle.lineHeightMultiple = LegalFormattingDefaults.bodyLineHeightMultiple
        currentStyle.paragraphSpacing = LegalFormattingDefaults.bodyParagraphSpacing
        
        document.addAttribute(.paragraphStyle, value: currentStyle, range: globalRange)
    }
}

// MARK: - Font family normalisation

/// Convert all fonts in the given range to the specified family and point size,
/// preserving traits such as bold/italic.
private func applyBaseFontFamily(
    to document: NSMutableAttributedString,
    in range: NSRange,
    familyName: String,
    pointSize: CGFloat
) {
    let manager = NSFontManager.shared

    document.enumerateAttribute(.font, in: range, options: []) { value, subrange, _ in
        let existingFont = (value as? NSFont) ?? NSFont.systemFont(ofSize: pointSize)

        // Convert to the desired family while preserving traits (bold, italic, etc.).
        let converted = manager.convert(existingFont, toFamily: familyName)

        // Normalise to the standard point size.
        let finalFont = NSFont(
            descriptor: converted.fontDescriptor,
            size: pointSize
        ) ?? converted

        document.removeAttribute(.font, range: subrange)
        document.addAttribute(.font, value: finalFont, range: subrange)
    }
}

// MARK: - Heading / table styling hooks

/// Use OpenAI analysis output to apply heading styles.
/// Extend this once AnalysisResult contains headingRanges.
private func applyHeadingStyles(
    using analysis: AnalysisResult,
    in document: NSMutableAttributedString
) {
    // TODO: when AnalysisResult has headingRanges, apply bold font and extra spacing.
}

/// Use OpenAI analysis output to convert text tables into real tables.
/// Extend this once AnalysisResult contains tableRanges.
private func applyTableStyles(
    using analysis: AnalysisResult,
    in document: NSMutableAttributedString
) {
    // TODO: when AnalysisResult has tableRanges, build NSTextTable and apply via paragraph styles.
}

// MARK: - Simple heuristic heading styling

/// Very simple heuristic for headings of the form "A. INTRODUCTION", "B. FACTS", etc.
/// Makes them bold with extra spacing.
private func applySimpleHeuristicHeadings(
    in document: NSMutableAttributedString,
    bodyRange: NSRange,
    baseFont: NSFont
) {
    let fullNSString = document.string as NSString
    let bodyString = fullNSString.substring(with: bodyRange) as NSString
    
    let boldFont = NSFontManager.shared.convert(baseFont, toHaveTrait: .boldFontMask)
    
    bodyString.enumerateSubstrings(
        in: NSRange(location: 0, length: bodyString.length),
        options: .byParagraphs
    ) { substring, subrange, _, _ in
        
        guard let line = substring?
            .trimmingCharacters(in: .whitespacesAndNewlines),
              !line.isEmpty else {
            return
        }
        
        // Simple pattern: "A. SOMETHING", "B. FACTS" etc.
        if let firstChar = line.first,
           firstChar.isUppercase {
            let remainder = line.dropFirst()
                .trimmingCharacters(in: .whitespaces)
            
            if remainder.first == "." {
                // This looks like "A. ..." – treat as heading.
                let globalRange = NSRange(
                    location: bodyRange.location + subrange.location,
                    length: subrange.length
                )
                
                let paraStyle = NSMutableParagraphStyle()
                paraStyle.paragraphSpacingBefore = 12
                paraStyle.paragraphSpacing = 6
                
                document.addAttributes(
                    [.font: boldFont,
                     .paragraphStyle: paraStyle],
                    range: globalRange
                )
            }
        }
    }
}

// MARK: - Utilities

private extension NSRange {
    /// Convert NSRange(location: NSNotFound, length: 0) into nil for convenience.
    func toOptional() -> NSRange? {
        if location == NSNotFound { return nil }
        return self
    }
}
