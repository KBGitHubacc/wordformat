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
    
    // 1. Identify the "Split Point" between Header/Intro and the Numbered Body
    // Prefer AI-derived split if available; otherwise fallback to heuristic.
    let splitIndex = findBodyStartIndex(in: document.string, analysis: analysis)
    
    // 2. Separate the Document into Two Parts
    let headerRange = NSRange(location: 0, length: max(0, splitIndex))
    let bodyRange = NSRange(location: max(0, splitIndex), length: document.length - max(0, splitIndex))
    
    // 3. Process Header (Native Styling - Safe)
    // We strictly preserve the text but apply UK Legal styling (Center/Bold)
    let headerPart = document.attributedSubstring(from: headerRange).mutableCopy() as! NSMutableAttributedString
    styleHeaderPart(headerPart)
    
    // 4. Process Body (NSTextList - True Word List Objects)
    // We extract the text, clean it, and rebuild it using NSTextList so DOCX gets real numbering.
    let bodyText = document.attributedSubstring(from: bodyRange).string
    let formattedBody = generateDynamicListBody(from: bodyText, analysis: analysis)
    
    // 5. Stitch Together
    let finalDoc = NSMutableAttributedString()
    finalDoc.append(headerPart)
    finalDoc.append(NSAttributedString(string: "\n\n")) // Buffer
    finalDoc.append(formattedBody)
    
    // 6. Final Polish (Global Font)
    let fullRange = NSRange(location: 0, length: finalDoc.length)
    applyBaseFont(to: finalDoc, range: fullRange)
    
    // 7. Update Document
    document.setAttributedString(finalDoc)
}

// MARK: - Step 1: Safe Header Styling

private func styleHeaderPart(_ attrString: NSMutableAttributedString) {
    let fullRange = NSRange(location: 0, length: attrString.length)
    
    // 1. Base Style (Left Aligned initially)
    let baseStyle = NSMutableParagraphStyle()
    baseStyle.alignment = .left
    baseStyle.paragraphSpacing = 12
    attrString.addAttribute(.paragraphStyle, value: baseStyle, range: fullRange)
    
    // 2. Center specific blocks (The "Header" proper)
    // We scan for keywords like "IN THE", "BETWEEN", "WITNESS STATEMENT"
    let string = attrString.string as NSString
    
    string.enumerateSubstrings(in: fullRange, options: .byParagraphs) { substring, substringRange, _, _ in
        guard let text = substring?.trimmingCharacters(in: .whitespacesAndNewlines).lowercased() else { return }
        
        let isCenterable = text.contains("in the") ||
                           text.contains("between") ||
                           text.contains("witness statement") ||
                           text.contains("case ref") ||
                           text.contains("claim no") ||
                           text.contains("-and-") ||
                           text.contains("applicant") ||
                           text.contains("respondent")
        
        if isCenterable {
            let centerStyle = NSMutableParagraphStyle()
            centerStyle.alignment = .center
            centerStyle.paragraphSpacing = 6
            attrString.addAttribute(.paragraphStyle, value: centerStyle, range: substringRange)
            
            // Apply Bold
            applyBold(to: attrString, range: substringRange)
            
            // Uppercase "WITNESS STATEMENT" titles
            if text.contains("witness statement") {
                let upper = (attrString.string as NSString).substring(with: substringRange).uppercased()
                attrString.replaceCharacters(in: substringRange, with: upper)
            }
        }
    }
}

// MARK: - Step 2: Dynamic List Generation (NSTextList -> Word-native numbering)

private func generateDynamicListBody(from text: String, analysis: AnalysisResult?) -> NSAttributedString {
    let result = NSMutableAttributedString()
    let baseFont = NSFont(name: LegalFormattingDefaults.fontFamily, size: LegalFormattingDefaults.fontSize)
        ?? NSFont.systemFont(ofSize: LegalFormattingDefaults.fontSize)
    
    // Split into logical paragraphs
    let paragraphs = text.components(separatedBy: .newlines)
    
    // Tracking nested lists; reuse the same NSTextList instances across contiguous runs
    var listStack: [NSTextList] = []
    var currentLevel = 0
    
    // Regex for detecting list items
    // Level 1: "1.", "304."
    let pattern1 = "^\\s*\\d+[.)]\\s+"
    // Level 2: "(a)", "a."
    let pattern2 = "^\\s*\\(?[a-zA-Z]\\)[.)]\\s+"
    // Level 3: "(i)", "i."
    let pattern3 = "^\\s*\\(?[ivx]+\\)[.)]\\s+"
    
    for line in paragraphs {
        let cleanLine = line.trimmingCharacters(in: .whitespacesAndNewlines)
        if cleanLine.isEmpty {
            // Blank lines break list sequences in most legal docs.
            if currentLevel > 0 {
                listStack.removeAll()
                currentLevel = 0
            }
            continue
        }
        
        // Statement of truth breaks lists and is bolded
        if cleanLine.lowercased().contains("statement of truth") {
            // Close lists
            if currentLevel > 0 {
                listStack.removeAll()
                currentLevel = 0
            }
            let p = paragraph(text: cleanLine, font: baseFont, alignment: .justified, bold: true)
            result.append(p)
            result.append(NSAttributedString(string: "\n"))
            continue
        }
        
        // Determine Level and content
        var level = 0
        var content = cleanLine
        
        if rangeOfPattern(pattern3, in: cleanLine) != nil {
            level = 3
            content = stripPattern(pattern3, from: cleanLine).trimmingCharacters(in: .whitespaces)
        } else if rangeOfPattern(pattern2, in: cleanLine) != nil {
            level = 2
            content = stripPattern(pattern2, from: cleanLine).trimmingCharacters(in: .whitespaces)
        } else if rangeOfPattern(pattern1, in: cleanLine) != nil {
            level = 1
            content = stripPattern(pattern1, from: cleanLine).trimmingCharacters(in: .whitespaces)
        } else {
            level = 0
        }
        
        // Adjust list depth
        if level > currentLevel {
            // Open new lists up to the target level
            while currentLevel < level {
                let format = markerFormat(for: currentLevel + 1)
                let newList = NSTextList(markerFormat: format, options: 0)
                listStack.append(newList)
                currentLevel += 1
            }
        } else if level < currentLevel && level > 0 {
            // Close lists down to target level
            while currentLevel > level {
                _ = listStack.popLast()
                currentLevel -= 1
            }
        } else if level == 0 && currentLevel > 0 {
            // Non-numbered paragraph breaks list
            listStack.removeAll()
            currentLevel = 0
        }
        
        // Output content
        if level > 0 {
            // Paragraph styled as a list item via NSTextList (no manual marker in text)
            let style = NSMutableParagraphStyle()
            style.alignment = .justified
            style.paragraphSpacing = 12
            style.textLists = listStack
            // Indentation per level (approx 18-24 pt per level)
            let indentPerLevel: CGFloat = 24.0
            style.headIndent = indentPerLevel * CGFloat(level)
            style.firstLineHeadIndent = indentPerLevel * CGFloat(level)
            
            let attrs: [NSAttributedString.Key: Any] = [
                .font: baseFont,
                .paragraphStyle: style
            ]
            let para = NSAttributedString(string: content, attributes: attrs)
            result.append(para)
            result.append(NSAttributedString(string: "\n"))
        } else {
            // Heading or plain text
            if isHeading(cleanLine) {
                let p = paragraph(text: content, font: baseFont, alignment: .justified, bold: true)
                result.append(p)
                result.append(NSAttributedString(string: "\n"))
            } else {
                let p = paragraph(text: content, font: baseFont, alignment: .justified, bold: false)
                result.append(p)
                result.append(NSAttributedString(string: "\n"))
            }
        }
    }
    
    return result
}

// MARK: - Utilities

private func paragraph(text: String, font: NSFont, alignment: NSTextAlignment, bold: Bool) -> NSAttributedString {
    let style = NSMutableParagraphStyle()
    style.alignment = alignment
    style.paragraphSpacing = 12
    
    let usedFont: NSFont
    if bold {
        let boldDesc = font.fontDescriptor.withSymbolicTraits(.bold)
        // NSFont(descriptor:size:) is optional on macOS; fall back to preserving family via NSFontManager.
        usedFont = NSFont(descriptor: boldDesc, size: font.pointSize)
            ?? NSFontManager.shared.convert(font, toHaveTrait: .boldFontMask)
    } else {
        usedFont = font
    }
    
    return NSAttributedString(string: text, attributes: [
        .font: usedFont,
        .paragraphStyle: style
    ])
}

private func markerFormat(for level: Int) -> NSTextList.MarkerFormat {
    switch level {
    case 1: return .decimal          // 1, 2, 3
    case 2: return .lowercaseAlpha   // a, b, c
    case 3: return .lowercaseRoman   // i, ii, iii
    default: return .decimal
    }
}

private func findBodyStartIndex(in text: String, analysis: AnalysisResult?) -> Int {
    // Prefer AI: find first block classified as body or heading and use its start
    if let analysis {
        // Find the earliest occurrence of types that usually mark the body start
        let candidates = analysis.classifiedRanges.filter { $0.type == .body || $0.type == .heading || $0.type == .intro }
        if let earliest = candidates.min(by: { $0.range.location < $1.range.location }) {
            return earliest.range.location
        }
    }
    // Fallback heuristic
    return findBodyStartIndex(in: text)
}

private func findBodyStartIndex(in text: String) -> Int {
    // We look for the intro line "will say as follows"
    let target = "will say as follows"
    if let range = text.range(of: target, options: .caseInsensitive) {
        // We split AFTER this line (end of paragraph)
        let nsRange = NSRange(range, in: text)
        let paragraphRange = (text as NSString).paragraphRange(for: nsRange)
        return paragraphRange.location + paragraphRange.length
    }
    // Fallback: If not found, split after "WITNESS STATEMENT... Name"
    if let range = text.range(of: "WITNESS STATEMENT", options: .caseInsensitive) {
        return NSRange(range, in: text).upperBound
    }
    return 0 // No header detected
}

private func rangeOfPattern(_ pattern: String, in text: String) -> NSRange? {
    guard let regex = try? NSRegularExpression(pattern: pattern) else { return nil }
    return regex.firstMatch(in: text, range: NSRange(location: 0, length: text.utf16.count))?.range
}

private func stripPattern(_ pattern: String, from text: String) -> String {
    guard let regex = try? NSRegularExpression(pattern: pattern) else { return text }
    let range = NSRange(location: 0, length: text.utf16.count)
    return regex.stringByReplacingMatches(in: text, options: [], range: range, withTemplate: "")
}

private func isHeading(_ text: String) -> Bool {
    // Detects "A. INTRODUCTION"
    let clean = text.trimmingCharacters(in: .whitespaces)
    return clean.range(of: "^[A-Z0-9]+\\.\\s+[A-Z]", options: .regularExpression) != nil
}

private func applyBaseFont(to doc: NSMutableAttributedString, range: NSRange) {
    let baseFont = NSFont(name: LegalFormattingDefaults.fontFamily, size: LegalFormattingDefaults.fontSize)
        ?? NSFont.systemFont(ofSize: LegalFormattingDefaults.fontSize)
    
    doc.enumerateAttribute(.font, in: range, options: []) { value, subRange, _ in
        if let existing = value as? NSFont {
            let traits = existing.fontDescriptor.symbolicTraits
            let desc = baseFont.fontDescriptor.withSymbolicTraits(traits)
            // NSFont(descriptor:size:) is optional; if it fails, fall back to baseFont.
            let newFont = NSFont(descriptor: desc, size: LegalFormattingDefaults.fontSize) ?? baseFont
            doc.addAttribute(.font, value: newFont, range: subRange)
        } else {
            doc.addAttribute(.font, value: baseFont, range: subRange)
        }
    }
}

private func applyBold(to str: NSMutableAttributedString, range: NSRange) {
    let baseFont = NSFont(name: LegalFormattingDefaults.fontFamily, size: LegalFormattingDefaults.fontSize)
        ?? NSFont.systemFont(ofSize: LegalFormattingDefaults.fontSize)
    let boldDesc = baseFont.fontDescriptor.withSymbolicTraits(.bold)
    // Prefer descriptor-based bold; fall back to preserving family via NSFontManager.
    let boldFont = NSFont(descriptor: boldDesc, size: LegalFormattingDefaults.fontSize)
        ?? NSFontManager.shared.convert(baseFont, toHaveTrait: .boldFontMask)
    str.addAttribute(.font, value: boldFont, range: range)
}
