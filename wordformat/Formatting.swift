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
    // Returns both character index and paragraph index for proper AI alignment
    let (splitCharIndex, splitParaIndex) = findBodyStartIndexWithParaIndex(in: document.string, analysis: analysis)
    Logger.shared.log("Split index at char \(splitCharIndex), para \(splitParaIndex)", category: "FORMAT")

    // 2. Separate the Document into Two Parts
    let headerRange = NSRange(location: 0, length: max(0, splitCharIndex))
    let bodyRange = NSRange(location: max(0, splitCharIndex), length: document.length - max(0, splitCharIndex))

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
    let formattedBody = generateDynamicListBody(from: bodyPart, analysis: analysis, bodyStartParaIndex: splitParaIndex)

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

private func generateDynamicListBody(from originalBody: NSAttributedString, analysis: AnalysisResult?, bodyStartParaIndex: Int) -> NSAttributedString {
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
    // Level 1: Main paragraphs - "1.", "2)", "10." etc.
    let patternLevel1 = "^\\s*\\d+[.)]\\s*"
    // Level 2: Sub-paragraphs - "(a)", "a)", "(b)" etc. (single letter)
    let patternLevel2 = "^\\s*\\(?[a-zA-Z]\\)\\s*"
    // Level 3: Sub-sub-paragraphs - "(i)", "ii)", "(iii)", "(iv)", "(v)" etc. (roman numerals)
    let patternLevel3 = "^\\s*\\(?[ivxIVX]+\\)\\s*"

    // AI guidance lookups (using global paragraph indices)
    let aiTypes = analysis?.paragraphTypes ?? [:]
    let aiLevels = analysis?.paragraphLevels ?? [:]

    let ns = originalBody.string as NSString
    let paraCount = ns.components(separatedBy: .newlines).count
    Logger.shared.log("Body paragraphs (approx): \(paraCount), starting at global index \(bodyStartParaIndex)", category: "FORMAT")

    var levelCounts = [0: 0, 1: 0, 2: 0, 3: 0]
    var sampleLogged = 0
    var localParaIndex = 0  // Index within body portion

    while cursor < originalBody.length {
        let paraRange = nsText.paragraphRange(for: NSRange(location: cursor, length: 0))
        let rawParagraphAttr = originalBody.attributedSubstring(from: paraRange)
        let rawParagraph = nsText.substring(with: paraRange)
        let cleanParagraph = rawParagraph.trimmingCharacters(in: .whitespacesAndNewlines)

        // Global paragraph index for AI lookup
        let globalParaIndex = bodyStartParaIndex + localParaIndex
        localParaIndex += 1

        // Classification (AI guided if available, using global index)
        let classifiedType = aiTypes[globalParaIndex] ?? .unknown

        // Determine if this paragraph should NOT be numbered
        let isHeadingPara = isHeading(cleanParagraph) || classifiedType == .heading
        let isStatementOfTruth = cleanParagraph.lowercased().contains("statement of truth") || classifiedType == .statementOfTruth
        let isTitle = classifiedType == .documentTitle
        let isIntro = classifiedType == .intro || cleanParagraph.lowercased().contains("will say as follows")
        let isHeader = classifiedType == .headerMetadata
        let isSignature = classifiedType == .signature
        let isNonNumbered = isHeadingPara || isStatementOfTruth || isTitle || isIntro || isHeader || isSignature

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

        // Decide list level - first check AI guidance, then fall back to pattern detection
        var level = 0
        if !isNonNumbered {
            // First check AI-provided level
            if let aiLevel = aiLevels[globalParaIndex] {
                level = aiLevel + 1  // AI uses 0=main, 1=sub, 2=subsub; we use 1,2,3 for numbered
            } else {
                // Fall back to pattern detection
                if rangeOfPattern(patternLevel3, in: cleanParagraph) != nil {
                    level = 3
                } else if rangeOfPattern(patternLevel2, in: cleanParagraph) != nil {
                    level = 2
                } else {
                    // Default: treat as main numbered paragraph
                    level = 1
                }
            }
        }

        // Strip manual markers if present while PRESERVING inline attributes
        let strippedAttr = stripMarkersPreservingAttributes(from: mutablePara, level: level, patternLevel1: patternLevel1, patternLevel2: patternLevel2, patternLevel3: patternLevel3)
        let paragraphContent = strippedAttr.string.trimmingCharacters(in: .whitespacesAndNewlines)

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
            Logger.shared.log("Para level \(level) type:\(classifiedType.rawValue) heading:\(isHeadingPara) intro:\(isIntro) title:\(isTitle) text: \(paragraphContent.prefix(100))", category: "FORMAT")
            sampleLogged += 1
        }

        // Apply paragraph style to the stripped attributed string (preserves inline bold/italic)
        strippedAttr.addAttribute(.paragraphStyle, value: paraStyle, range: NSRange(location: 0, length: strippedAttr.length))

        // Apply base font while preserving traits (bold/italic) - enumerate to preserve inline variations
        applyBaseFontPreservingTraits(to: strippedAttr, baseFont: baseFont, makeBold: isStatementOfTruth)

        result.append(strippedAttr)
        result.append(NSAttributedString(string: "\n"))

        cursor = paraRange.upperBound
    }

    Logger.shared.log("Level counts -> level0:\(levelCounts[0, default:0]) level1:\(levelCounts[1, default:0]) level2:\(levelCounts[2, default:0]) level3:\(levelCounts[3, default:0])", category: "FORMAT")
    return result
}

/// Strip numbering markers from attributed string while preserving inline formatting (bold, italic, etc.)
private func stripMarkersPreservingAttributes(from attrStr: NSMutableAttributedString, level: Int, patternLevel1: String, patternLevel2: String, patternLevel3: String) -> NSMutableAttributedString {
    let text = attrStr.string
    var rangeToRemove: NSRange?

    if level == 3 {
        rangeToRemove = rangeOfPattern(patternLevel3, in: text)
    } else if level == 2 {
        rangeToRemove = rangeOfPattern(patternLevel2, in: text)
    } else if level >= 1 {
        rangeToRemove = rangeOfPattern(patternLevel1, in: text)
    }

    // Also strip leading whitespace after removing pattern
    if let range = rangeToRemove, range.location == 0 || text.prefix(range.location).trimmingCharacters(in: .whitespaces).isEmpty {
        // Remove the marker range from the attributed string (keeps other attributes intact)
        let startTrim = text.prefix(range.location).count
        let totalToRemove = startTrim + range.length
        if totalToRemove > 0 && totalToRemove <= attrStr.length {
            attrStr.deleteCharacters(in: NSRange(location: 0, length: totalToRemove))
        }
    }

    // Trim leading whitespace
    while attrStr.length > 0 {
        let firstChar = (attrStr.string as NSString).substring(with: NSRange(location: 0, length: 1))
        if firstChar.trimmingCharacters(in: .whitespaces).isEmpty {
            attrStr.deleteCharacters(in: NSRange(location: 0, length: 1))
        } else {
            break
        }
    }

    return attrStr
}

/// Apply base font to attributed string while preserving bold/italic traits at each position
private func applyBaseFontPreservingTraits(to attrStr: NSMutableAttributedString, baseFont: NSFont, makeBold: Bool) {
    guard attrStr.length > 0 else { return }
    let fullRange = NSRange(location: 0, length: attrStr.length)

    attrStr.enumerateAttribute(.font, in: fullRange, options: []) { value, subRange, _ in
        var traits: NSFontDescriptor.SymbolicTraits = []
        if let existing = value as? NSFont {
            traits = existing.fontDescriptor.symbolicTraits
        }
        if makeBold {
            traits.insert(.bold)
        }
        let desc = baseFont.fontDescriptor.withSymbolicTraits(traits)
        let newFont = NSFont(descriptor: desc, size: baseFont.pointSize) ?? baseFont
        attrStr.addAttribute(.font, value: newFont, range: subRange)
    }
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

/// Find body start index using AI analysis or heuristics
/// Returns: (characterIndex, paragraphIndex) tuple for proper AI alignment
private func findBodyStartIndexWithParaIndex(in text: String, analysis: AnalysisResult?) -> (Int, Int) {
    // Prefer style-aware paragraph types - find first BODY paragraph (not intro/title/header)
    if let types = analysis?.paragraphTypes, !types.isEmpty {
        // Find first paragraph that is actually body content (numbered paragraphs)
        // Skip intro, title, header - these should not be numbered
        let bodyTypes: Set<LegalParagraphType> = [.body, .heading]
        if let firstBody = types.sorted(by: { $0.key < $1.key }).first(where: { bodyTypes.contains($0.value) }) {
            if let range = paragraphRange(forIndex: firstBody.key, in: text) {
                Logger.shared.log("AI-guided split: first body/heading at para \(firstBody.key), type \(firstBody.value.rawValue)", category: "FORMAT")
                return (range.location, firstBody.key)
            }
        }
    }

    // Legacy classifiedRanges fallback
    if let analysis, !analysis.classifiedRanges.isEmpty {
        let candidates = analysis.classifiedRanges.filter { $0.type == .body || $0.type == .heading }
        if let earliest = candidates.min(by: { $0.range.location < $1.range.location }) {
            let paraIdx = paragraphIndexAt(charLocation: earliest.range.location, in: text)
            return (earliest.range.location, paraIdx)
        }
    }

    // Fallback heuristic - find first actual body paragraph after intro
    return findBodyStartIndexHeuristic(in: text)
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

/// Find which paragraph index corresponds to a character location
private func paragraphIndexAt(charLocation: Int, in text: String) -> Int {
    let ns = text as NSString
    var idx = 0
    var result = 0
    ns.enumerateSubstrings(in: NSRange(location: 0, length: ns.length), options: .byParagraphs) { _, range, _, stop in
        if range.location <= charLocation && charLocation < range.location + range.length {
            result = idx
            stop.pointee = true
        }
        idx += 1
    }
    return result
}

/// Heuristic-based body start detection
/// Returns: (characterIndex, paragraphIndex) tuple
private func findBodyStartIndexHeuristic(in text: String) -> (Int, Int) {
    let ns = text as NSString
    var paraIdx = 0
    var foundIntro = false
    var resultCharIdx = 0
    var resultParaIdx = 0

    ns.enumerateSubstrings(in: NSRange(location: 0, length: ns.length), options: .byParagraphs) { substring, range, _, stop in
        let paraText = (substring ?? "").trimmingCharacters(in: .whitespacesAndNewlines).lowercased()

        // Skip empty paragraphs
        if paraText.isEmpty {
            paraIdx += 1
            return
        }

        // If we found intro, the NEXT non-empty paragraph is the body start
        if foundIntro {
            // Make sure we're not hitting another header-like paragraph
            let isLikelyBody = !paraText.contains("witness statement") &&
                               !paraText.contains("in the ") &&
                               !paraText.contains("between") &&
                               !paraText.hasPrefix("case")
            if isLikelyBody {
                resultCharIdx = range.location
                resultParaIdx = paraIdx
                stop.pointee = true
                return
            }
        }

        // Look for intro paragraph
        if paraText.contains("will say as follows") {
            foundIntro = true
        }

        // Also check for section headings like "A. INTRODUCTION" as body start
        if !foundIntro {
            let upperText = (substring ?? "").trimmingCharacters(in: .whitespacesAndNewlines)
            if upperText.range(of: "^[A-Z]+\\.\\s+[A-Z]", options: .regularExpression) != nil {
                resultCharIdx = range.location
                resultParaIdx = paraIdx
                stop.pointee = true
                return
            }
        }

        paraIdx += 1
    }

    // If we found intro but no body after (edge case), return end of document
    if resultCharIdx == 0 && foundIntro {
        Logger.shared.log("Warning: Found intro but no body paragraphs detected", category: "FORMAT")
    }

    return (resultCharIdx, resultParaIdx)
}

/// Legacy function for backward compatibility
private func findBodyStartIndex(in text: String, analysis: AnalysisResult?) -> Int {
    return findBodyStartIndexWithParaIndex(in: text, analysis: analysis).0
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
/// IMPORTANT: The doc parameter here is the FORMATTED document, which may have different paragraph
/// structure than the original. We need to track paragraph indices correctly.
func buildNumberingTargets(from doc: NSAttributedString, analysis: AnalysisResult?) -> [DocxNumberingPatcher.NumberingTarget] {
    let ns = doc.string as NSString
    var targets: [DocxNumberingPatcher.NumberingTarget] = []
    let aiLevels = analysis?.paragraphLevels ?? [:]
    let aiTypes = analysis?.paragraphTypes ?? [:]

    // Find split point in the formatted document
    let (splitCharIndex, _) = findBodyStartIndexWithParaIndex(in: doc.string, analysis: analysis)
    Logger.shared.log("Target builder using splitCharIndex \(splitCharIndex)", category: "PATCH")

    // Non-numbered paragraph types
    let nonNumberedTypes: Set<LegalParagraphType> = [.headerMetadata, .documentTitle, .intro, .heading, .statementOfTruth, .signature, .quote]

    var xmlParaIndex = 0  // Index for the XML patcher (counts all paragraphs in output)
    ns.enumerateSubstrings(in: NSRange(location: 0, length: ns.length), options: .byParagraphs) { substring, range, _, _ in
        defer { xmlParaIndex += 1 }

        // Skip paragraphs before split point (header area)
        guard range.location >= splitCharIndex else { return }

        let text = (substring ?? "").trimmingCharacters(in: .whitespacesAndNewlines)

        // Skip empty paragraphs
        if text.isEmpty { return }

        // Check AI type for this paragraph (if available)
        // Note: We need to map back to original paragraph index if AI data is indexed differently
        // For now, check using pattern-based detection as primary
        let aiType = aiTypes[xmlParaIndex]

        // Determine if this is a non-numbered paragraph type
        let isHeadingPara = isHeading(text)
        let isIntro = text.lowercased().contains("will say as follows")
        let isTitle = text.uppercased().contains("WITNESS STATEMENT OF")
        let isStatementOfTruth = text.lowercased().contains("statement of truth")

        // Skip non-numbered types based on both AI and heuristics
        if let t = aiType, nonNumberedTypes.contains(t) {
            Logger.shared.log("Skipping para \(xmlParaIndex) - AI type: \(t.rawValue)", category: "PATCH")
            return
        }
        if isHeadingPara || isIntro || isTitle || isStatementOfTruth {
            Logger.shared.log("Skipping para \(xmlParaIndex) - heuristic: heading=\(isHeadingPara) intro=\(isIntro) title=\(isTitle) truth=\(isStatementOfTruth)", category: "PATCH")
            return
        }

        // Get paragraph style to check for existing textLists (indicates numbered paragraph)
        let attrStyle = doc.attribute(.paragraphStyle, at: range.location, effectiveRange: nil) as? NSParagraphStyle
        let textLists = attrStyle?.textLists ?? []
        let indent = attrStyle?.headIndent ?? 0

        // Only add to targets if paragraph has textLists (was marked for numbering by formatter)
        // OR if indent suggests it should be numbered
        guard !textLists.isEmpty || indent >= 20 else {
            Logger.shared.log("Skipping para \(xmlParaIndex) - no textLists and low indent (\(indent))", category: "PATCH")
            return
        }

        // Determine level from textLists count or AI
        var level = textLists.count - 1  // textLists array length indicates nesting depth
        if level < 0 {
            // Fallback to AI level or indent-based detection
            level = aiLevels[xmlParaIndex] ?? 0
            if level == 0 {
                if indent >= 72 { level = 2 }
                else if indent >= 48 { level = 1 }
                else { level = 0 }
            }
        }

        targets.append(.init(paragraphIndex: xmlParaIndex, level: max(0, level)))
    }

    if !targets.isEmpty {
        let sample = targets.prefix(10).map { "[\($0.paragraphIndex):\($0.level)]" }.joined(separator: " ")
        Logger.shared.log("Built \(targets.count) numbering targets. Sample: \(sample)", category: "PATCH")
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
