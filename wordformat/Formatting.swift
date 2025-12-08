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
    Logger.shared.log("Formatting start. Total length: \(document.length)", category: "FORMAT")
    
    // 1. Identify the "Split Point" between Header/Intro and the Numbered Body
    // Prefer AI-derived split if available; otherwise fallback to heuristic.
    let splitIndex = findBodyStartIndex(in: document.string, analysis: analysis)
    Logger.shared.log("Split index at \(splitIndex)", category: "FORMAT")
    
    // 2. Separate the Document into Two Parts
    let headerRange = NSRange(location: 0, length: max(0, splitIndex))
    let bodyRange = NSRange(location: max(0, splitIndex), length: document.length - max(0, splitIndex))
    
    // 3. Process Header (Native Styling - Safe)
    // We strictly preserve the text but apply UK Legal styling (Center/Bold)
    let headerPart = document.attributedSubstring(from: headerRange).mutableCopy() as! NSMutableAttributedString
    styleHeaderPart(headerPart)
    
    // 4. Process Body (Preserve and rebuild numbering based on list attributes)
    // IMPORTANT: When we read a DOCX into an attributed string, Word list numbers are NOT in the plain text.
    // They live in paragraph attributes (`textLists`). Converting to `string` would lose them, so we work
    // with the attributed body directly.
    let bodyPart = document.attributedSubstring(from: bodyRange)
    Logger.shared.log("Header length: \(headerRange.length), body length: \(bodyRange.length)", category: "FORMAT")
    let formattedBody = generateDynamicListBody(from: bodyPart, analysis: analysis)
    
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

// MARK: - Step 2: Dynamic List Generation (Preserve Word-native numbering)

private func generateDynamicListBody(from originalBody: NSAttributedString, analysis: AnalysisResult?) -> NSAttributedString {
    let result = NSMutableAttributedString()
    let baseFont = NSFont(name: LegalFormattingDefaults.fontFamily, size: LegalFormattingDefaults.fontSize)
        ?? NSFont.systemFont(ofSize: LegalFormattingDefaults.fontSize)
    
    let nsText = originalBody.string as NSString
    var cursor = 0
    
    // Fresh list objects to ensure deterministic numbering (auto-renumber on insert/delete).
    let level1List = NSTextList(markerFormat: .decimal, options: 0)
    let level2List = NSTextList(markerFormat: .lowercaseAlpha, options: 0)
    let level3List = NSTextList(markerFormat: .lowercaseRoman, options: 0)

    // Regex helpers for manual markers (if present, we strip them and rely on lists).
    let patternLevel1 = "^\\s*\\d+[.)]\\s+"
    let patternLevel2 = "^\\s*\\(?[a-zA-Z]\\)[.)]\\s+"
    let patternLevel3 = "^\\s*\\(?[ivx]+\\)[.)]\\s+"
    
    // Optional AI guidance lookup
    let typeLookup = makeTypeLookup(analysis: analysis)
    let ns = originalBody.string as NSString
    let paraCount = ns.components(separatedBy: .newlines).count
    Logger.shared.log("Body paragraphs (approx): \(paraCount)", category: "FORMAT")
    
    var levelCounts = [0: 0, 1: 0, 2: 0, 3: 0]
    var sampleLogged = 0
    
    while cursor < originalBody.length {
        let paraRange = nsText.paragraphRange(for: NSRange(location: cursor, length: 0))
        let rawParagraphAttr = originalBody.attributedSubstring(from: paraRange)
        let rawParagraph = nsText.substring(with: paraRange)
        let cleanParagraph = rawParagraph.trimmingCharacters(in: .whitespacesAndNewlines)
        
        // Classification (AI guided if available)
        let classifiedType = typeLookup(paraRange) ?? .unknown
        let isHeadingPara = isHeading(cleanParagraph) || classifiedType == .heading
        let isStatementOfTruth = cleanParagraph.lowercased().contains("statement of truth") || classifiedType == .statementOfTruth
        
        // Blank lines: keep as spacer, but do not attach list
        if cleanParagraph.isEmpty {
            cursor = paraRange.upperBound
            continue
        }
        
        // Prepare mutable paragraph with preserved inline styling
        let mutablePara = rawParagraphAttr.mutableCopy() as! NSMutableAttributedString
        let paraStyle = (mutablePara.attribute(.paragraphStyle, at: 0, effectiveRange: nil) as? NSParagraphStyle)?.mutableCopy() as? NSMutableParagraphStyle ?? NSMutableParagraphStyle()
        paraStyle.alignment = .justified
        paraStyle.paragraphSpacing = 12
        
        // Decide list level
        var level = 0
        if !isHeadingPara && !isStatementOfTruth {
            if rangeOfPattern(patternLevel3, in: cleanParagraph) != nil {
                level = 3
            } else if rangeOfPattern(patternLevel2, in: cleanParagraph) != nil {
                level = 2
            } else {
                // Default: treat as main numbered paragraph
                level = 1
            }
        }
        
        // Strip manual markers if present (content is preserved otherwise).
        var paragraphContent = cleanParagraph
        if level == 3 {
            paragraphContent = stripPattern(patternLevel3, from: paragraphContent).trimmingCharacters(in: .whitespaces)
        } else if level == 2 {
            paragraphContent = stripPattern(patternLevel2, from: paragraphContent).trimmingCharacters(in: .whitespaces)
        } else if level == 1 && rangeOfPattern(patternLevel1, in: paragraphContent) != nil {
            paragraphContent = stripPattern(patternLevel1, from: paragraphContent).trimmingCharacters(in: .whitespaces)
        }
        
        // Apply list styles
        if level == 0 {
            paraStyle.textLists = []
            paraStyle.firstLineHeadIndent = 0
            paraStyle.headIndent = 0
        } else if level == 1 {
            paraStyle.textLists = [level1List]
            paraStyle.headIndent = 24
            paraStyle.firstLineHeadIndent = 24
        } else if level == 2 {
            paraStyle.textLists = [level1List, level2List]
            paraStyle.headIndent = 48
            paraStyle.firstLineHeadIndent = 48
        } else {
            paraStyle.textLists = [level1List, level2List, level3List]
            paraStyle.headIndent = 72
            paraStyle.firstLineHeadIndent = 72
        }
        
        levelCounts[level, default: 0] += 1
        if sampleLogged < 25 {
            Logger.shared.log("Para level \(level) heading:\(isHeadingPara) truth:\(isStatementOfTruth) text: \(paragraphContent.prefix(120))", category: "FORMAT")
            sampleLogged += 1
        }
        
        // Replace text with stripped content while keeping inline attributes
        mutablePara.mutableString.setString(paragraphContent)
        mutablePara.addAttribute(.paragraphStyle, value: paraStyle, range: NSRange(location: 0, length: mutablePara.length))
        
        // Font handling (preserve existing traits, bold statement of truth)
        let existingFont = mutablePara.attribute(.font, at: 0, effectiveRange: nil) as? NSFont
        let appliedFont = baseFontWithExistingTraits(baseFont: baseFont, existing: existingFont)
        let finalFont = isStatementOfTruth ? boldVariant(of: appliedFont) : appliedFont
        mutablePara.addAttribute(.font, value: finalFont, range: NSRange(location: 0, length: mutablePara.length))
        
        result.append(mutablePara)
        result.append(NSAttributedString(string: "\n"))
        
        cursor = paraRange.upperBound
    }
    
    Logger.shared.log("Level counts -> level0:\(levelCounts[0, default:0]) level1:\(levelCounts[1, default:0]) level2:\(levelCounts[2, default:0]) level3:\(levelCounts[3, default:0])", category: "FORMAT")
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
    // Prefer style-aware paragraph types
    if let types = analysis?.paragraphTypes, !types.isEmpty {
        // Find first paragraph marked as intro/body/heading
        let targetTypes: Set<LegalParagraphType> = [.intro, .body, .heading]
        if let firstTarget = types.sorted(by: { $0.key < $1.key }).first(where: { targetTypes.contains($0.value) }) {
            if let range = paragraphRange(forIndex: firstTarget.key, in: text) {
                return range.location
            }
        }
    }
    // Legacy classifiedRanges
    if let analysis {
        let candidates = analysis.classifiedRanges.filter { $0.type == .body || $0.type == .heading || $0.type == .intro }
        if let earliest = candidates.min(by: { $0.range.location < $1.range.location }) {
            return earliest.range.location
        }
    }
    // Fallback heuristic
    return findBodyStartIndex(in: text)
}

private func paragraphRange(forIndex target: Int, in text: String) -> NSRange? {
    let ns = text as NSString
    var idx = 0
    var result: NSRange?
    ns.enumerateSubstrings(in: NSRange(location: 0, length: ns.length), options: .byParagraphs) { _, range, _, stop in
        if idx == target {
            result = range
            stop.pointee = true
        }
        idx += 1
    }
    return result
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

// MARK: - Font helpers

/// Returns the base font while keeping any existing symbolic traits (bold/italic) from the source font.
private func baseFontWithExistingTraits(baseFont: NSFont, existing: NSFont?) -> NSFont {
    guard let existing else { return baseFont }
    let traits = existing.fontDescriptor.symbolicTraits
    let descriptor = baseFont.fontDescriptor.withSymbolicTraits(traits)
    return NSFont(descriptor: descriptor, size: baseFont.pointSize) ?? baseFont
}

// MARK: - Numbering targets extraction

/// Build numbering targets from the fully formatted document, using AI paragraph levels when available.
func buildNumberingTargets(from doc: NSAttributedString, analysis: AnalysisResult?) -> [DocxNumberingPatcher.NumberingTarget] {
    let ns = doc.string as NSString
    var targets: [DocxNumberingPatcher.NumberingTarget] = []
    let aiLevels = analysis?.paragraphLevels ?? [:]
    let aiTypes = analysis?.paragraphTypes ?? [:]
    let splitIndex = findBodyStartIndex(in: doc.string, analysis: analysis)
    Logger.shared.log("Target builder using splitIndex \(splitIndex)", category: "PATCH")
    
    var paraIndex = 0
    ns.enumerateSubstrings(in: NSRange(location: 0, length: ns.length), options: .byParagraphs) { substring, range, _, _ in
        defer { paraIndex += 1 }
        guard range.location >= splitIndex else { return } // skip header
        let text = (substring ?? "").trimmingCharacters(in: .whitespacesAndNewlines)
        if text.isEmpty { return }
        
        // Heading detection
        let aiType = aiTypes[paraIndex]
        let isHeadingPara = isHeading(text) || aiType == .heading || aiType == .documentTitle || aiType == .headerMetadata
        if isHeadingPara { return }
        
        let attrStyle = doc.attribute(.paragraphStyle, at: range.location, effectiveRange: nil) as? NSParagraphStyle
        let indent = attrStyle?.headIndent ?? 0
        
        // Determine level
        var level = aiLevels[paraIndex] ?? 0
        if level == 0 {
            // Fallback to marker/indent heuristics
            if rangeOfPattern("^\\s*\\(?[a-zA-Z]\\)", in: text) != nil || indent >= 44 {
                level = 1
            }
            if rangeOfPattern("^\\s*\\(?[ivx]+\\)", in: text.lowercased()) != nil || indent >= 80 {
                level = 2
            }
        }
        
        // Record all numbered levels (including main level 0)
        targets.append(.init(paragraphIndex: paraIndex, level: max(0, level)))
    }
    if !targets.isEmpty {
        let sample = targets.prefix(10).map { "[\($0.paragraphIndex):\($0.level)]" }.joined(separator: " ")
        Logger.shared.log("Built \(targets.count) numbering targets from document paragraphs. Sample: \(sample)", category: "PATCH")
    } else {
        Logger.shared.log("No numbering targets built", category: "PATCH")
    }
    return targets
}
/// Bold variant of a given font, preserving family where possible.
private func boldVariant(of font: NSFont) -> NSFont {
    let desc = font.fontDescriptor.withSymbolicTraits(.bold)
    return NSFont(descriptor: desc, size: font.pointSize) ?? NSFontManager.shared.convert(font, toHaveTrait: .boldFontMask)
}

// MARK: - AI type lookup helper

private func makeTypeLookup(analysis: AnalysisResult?) -> (NSRange) -> LegalParagraphType? {
    guard let analysis else {
        return { _ in nil }
    }
    let ranges = analysis.classifiedRanges
    return { paraRange in
        // Pick the first matching type that intersects this paragraph.
        for item in ranges {
            if NSIntersectionRange(item.range, paraRange).length > 0 {
                return item.type
            }
        }
        return nil
    }
}
