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

    // 1. Identify Split Point
    let (splitCharIndex, splitParaIndex) = findBodyStartIndexWithParaIndex(in: document.string, analysis: analysis)
    Logger.shared.log("Split index at char \(splitCharIndex), para \(splitParaIndex)", category: "FORMAT")

    // 2. Separate Document
    let headerRange = NSRange(location: 0, length: max(0, splitCharIndex))
    let bodyRange = NSRange(location: max(0, splitCharIndex), length: document.length - max(0, splitCharIndex))

    // 3. Process Header
    let headerPart = document.attributedSubstring(from: headerRange).mutableCopy() as! NSMutableAttributedString
    styleHeaderPart(headerPart)

    // 4. Process Body
    let bodyPart = document.attributedSubstring(from: bodyRange)
    Logger.shared.log("Header length: \(headerRange.length), body length: \(bodyRange.length)", category: "FORMAT")
    let formattedBody = generateDynamicListBody(from: bodyPart, analysis: analysis, bodyStartParaIndex: splitParaIndex)

    // 5. Stitch
    let finalDoc = NSMutableAttributedString()
    finalDoc.append(headerPart)
    finalDoc.append(NSAttributedString(string: "\n\n"))
    finalDoc.append(formattedBody)

    // 6. Global Polish
    let fullRange = NSRange(location: 0, length: finalDoc.length)
    applyBaseFont(to: finalDoc, range: fullRange)

    // 7. Apply
    document.setAttributedString(finalDoc)
}

private func styleHeaderPart(_ attrString: NSMutableAttributedString) {
    let fullRange = NSRange(location: 0, length: attrString.length)
    
    // Base Style
    let baseStyle = NSMutableParagraphStyle()
    baseStyle.alignment = .left
    baseStyle.paragraphSpacing = 12
    attrString.addAttribute(.paragraphStyle, value: baseStyle, range: fullRange)
    
    // Center specific blocks
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
            applyBold(to: attrString, range: substringRange)
            
            if text.contains("witness statement") {
                let upper = (attrString.string as NSString).substring(with: substringRange).uppercased()
                attrString.replaceCharacters(in: substringRange, with: upper)
            }
        }
    }
}

// MARK: - Step 2: Dynamic List Generation

private func generateDynamicListBody(from originalBody: NSAttributedString, analysis: AnalysisResult?, bodyStartParaIndex: Int) -> NSAttributedString {
    let result = NSMutableAttributedString()
    let baseFont = NSFont(name: LegalFormattingDefaults.fontFamily, size: LegalFormattingDefaults.fontSize)
        ?? NSFont.systemFont(ofSize: LegalFormattingDefaults.fontSize)

    let nsText = originalBody.string as NSString
    var cursor = 0

    // List definitions
    let level1List = NSTextList(markerFormat: .decimal, options: 0)
    let level2List = NSTextList(markerFormat: .lowercaseAlpha, options: 0)
    let level3List = NSTextList(markerFormat: .lowercaseRoman, options: 0)

    // --- IMPROVED REGEX ---
    // Level 1: "1.", "1)", "10."
    let patternLevel1 = "^\\s*\\d+[.)]\\s+"
    
    // Level 2: "(a)", "a)", "a." - Now strictly supports dot delimiter
    // Matches: whitespace -> optional ( -> letter -> required ) or . -> whitespace
    let patternLevel2 = "^\\s*(?:\\([a-zA-Z]\\)|[a-zA-Z][).])\\s+"
    
    // Level 3: "(i)", "i)", "i." - Roman numerals
    let patternLevel3 = "^\\s*(?:\\([ivxIVX]+\\)|[ivxIVX]+[).])\\s+"

    let aiTypes = analysis?.paragraphTypes ?? [:]
    let aiLevels = analysis?.paragraphLevels ?? [:]

    var localParaIndex = 0

    while cursor < originalBody.length {
        let paraRange = nsText.paragraphRange(for: NSRange(location: cursor, length: 0))
        let rawParagraphAttr = originalBody.attributedSubstring(from: paraRange)
        let cleanParagraph = rawParagraphAttr.string.trimmingCharacters(in: .whitespacesAndNewlines)

        let globalParaIndex = bodyStartParaIndex + localParaIndex
        localParaIndex += 1

        if cleanParagraph.isEmpty {
            cursor = paraRange.upperBound
            continue
        }

        // Classification
        let classifiedType = aiTypes[globalParaIndex] ?? .unknown
        let isHeadingPara = isHeading(cleanParagraph) || classifiedType == .heading
        let isStatement = cleanParagraph.lowercased().contains("statement of truth") || classifiedType == .statementOfTruth
        let isTitle = classifiedType == .documentTitle
        let isIntro = classifiedType == .intro || cleanParagraph.lowercased().contains("will say as follows")
        let isHeader = classifiedType == .headerMetadata
        let isSignature = classifiedType == .signature
        let isNonNumbered = isHeadingPara || isStatement || isTitle || isIntro || isHeader || isSignature

        // Prepare Paragraph
        let mutablePara = rawParagraphAttr.mutableCopy() as! NSMutableAttributedString
        let paraStyle = NSMutableParagraphStyle()
        paraStyle.alignment = .justified
        paraStyle.paragraphSpacing = 12

        // Decide list level
        var level = 0
        if !isNonNumbered {
            if let aiLevel = aiLevels[globalParaIndex] {
                level = aiLevel + 1
            } else {
                // Regex fallback
                if rangeOfPattern(patternLevel3, in: cleanParagraph) != nil {
                    level = 3
                } else if rangeOfPattern(patternLevel2, in: cleanParagraph) != nil {
                    level = 2
                } else if rangeOfPattern(patternLevel1, in: cleanParagraph) != nil {
                    level = 1
                }
            }
        }

        // Strip markers
        let strippedAttr = stripMarkersPreservingAttributes(
            from: mutablePara,
            level: level,
            patternLevel1: patternLevel1,
            patternLevel2: patternLevel2,
            patternLevel3: patternLevel3,
            allowStrip: !isNonNumbered && level > 0
        )

        // Apply List Attributes
        if level == 0 {
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

        strippedAttr.addAttribute(.paragraphStyle, value: paraStyle, range: NSRange(location: 0, length: strippedAttr.length))
        applyBaseFontPreservingTraits(to: strippedAttr, baseFont: baseFont, makeBold: isStatement)

        result.append(strippedAttr)
        result.append(NSAttributedString(string: "\n"))

        cursor = paraRange.upperBound
    }

    return result
}

// MARK: - Utilities

private func stripMarkersPreservingAttributes(from attrStr: NSMutableAttributedString, level: Int, patternLevel1: String, patternLevel2: String, patternLevel3: String, allowStrip: Bool = true) -> NSMutableAttributedString {
    let text = attrStr.string
    var rangeToRemove: NSRange?

    if level == 3 {
        rangeToRemove = rangeOfPattern(patternLevel3, in: text)
    } else if level == 2 {
        rangeToRemove = rangeOfPattern(patternLevel2, in: text)
    } else if level >= 1 {
        rangeToRemove = rangeOfPattern(patternLevel1, in: text)
    }

    if allowStrip, let range = rangeToRemove, range.location == 0 {
        if range.length <= attrStr.length {
            attrStr.deleteCharacters(in: range)
        }
    }
    return attrStr
}

private func applyBaseFontPreservingTraits(to attrStr: NSMutableAttributedString, baseFont: NSFont, makeBold: Bool) {
    guard attrStr.length > 0 else { return }
    let fullRange = NSRange(location: 0, length: attrStr.length)

    attrStr.enumerateAttribute(.font, in: fullRange, options: []) { value, subRange, _ in
        var traits: NSFontDescriptor.SymbolicTraits = []
        if let existing = value as? NSFont {
            traits = existing.fontDescriptor.symbolicTraits
        }
        if makeBold { traits.insert(.bold) }
        
        let desc = baseFont.fontDescriptor.withSymbolicTraits(traits)
        let newFont = NSFont(descriptor: desc, size: baseFont.pointSize) ?? baseFont
        attrStr.addAttribute(.font, value: newFont, range: subRange)
    }
}

private func findBodyStartIndexWithParaIndex(in text: String, analysis: AnalysisResult?) -> (Int, Int) {
    if let types = analysis?.paragraphTypes, !types.isEmpty {
        let bodyTypes: Set<LegalParagraphType> = [.body, .heading]
        if let firstBody = types.sorted(by: { $0.key < $1.key }).first(where: { bodyTypes.contains($0.value) }) {
            if let range = paragraphRange(forIndex: firstBody.key, in: text) {
                return (range.location, firstBody.key)
            }
        }
    }

    // Heuristic
    let ns = text as NSString
    var paraIdx = 0
    var foundIntro = false
    var resultCharIdx = 0
    var resultParaIdx = 0

    ns.enumerateSubstrings(in: NSRange(location: 0, length: ns.length), options: .byParagraphs) { substring, range, _, stop in
        let t = (substring ?? "").trimmingCharacters(in: .whitespacesAndNewlines).lowercased()
        if t.isEmpty { paraIdx += 1; return }

        if !foundIntro && t.range(of: "^[a-z]+\\.\\s+[a-z]+", options: .regularExpression) != nil {
            resultCharIdx = range.location
            resultParaIdx = paraIdx
            stop.pointee = true
            return
        }

        if t.contains("will say as follows") {
            foundIntro = true
        } else if foundIntro {
            resultCharIdx = range.location
            resultParaIdx = paraIdx
            stop.pointee = true
            return
        }
        paraIdx += 1
    }
    return (resultCharIdx, resultParaIdx)
}

private func paragraphRange(forIndex target: Int, in text: String) -> NSRange? {
    let ns = text as NSString
    var idx = 0
    var result: NSRange?
    ns.enumerateSubstrings(in: NSRange(location: 0, length: ns.length), options: .byParagraphs) { _, range, _, stop in
        if idx == target { result = range; stop.pointee = true }
        idx += 1
    }
    return result
}

private func rangeOfPattern(_ pattern: String, in text: String) -> NSRange? {
    guard let regex = try? NSRegularExpression(pattern: pattern) else { return nil }
    return regex.firstMatch(in: text, range: NSRange(location: 0, length: text.utf16.count))?.range
}

private func isHeading(_ text: String) -> Bool {
    let clean = text.trimmingCharacters(in: .whitespaces)
    return clean.range(of: "^[A-Z0-9]+\\.\\s+[A-Z]", options: .regularExpression) != nil
}

private func applyBaseFont(to doc: NSMutableAttributedString, range: NSRange) {
    let baseFont = NSFont(name: LegalFormattingDefaults.fontFamily, size: LegalFormattingDefaults.fontSize)
        ?? NSFont.systemFont(ofSize: LegalFormattingDefaults.fontSize)
    applyBaseFontPreservingTraits(to: doc, baseFont: baseFont, makeBold: false)
}

private func applyBold(to str: NSMutableAttributedString, range: NSRange) {
    let font = NSFont(name: LegalFormattingDefaults.fontFamily, size: LegalFormattingDefaults.fontSize) ?? NSFont.systemFont(ofSize: 12)
    let boldFont = NSFontManager.shared.convert(font, toHaveTrait: .boldFontMask)
    str.addAttribute(.font, value: boldFont, range: range)
}

// MARK: - Updated Target Builder for Patcher

func buildNumberingTargetsFromAnalysis(doc: NSAttributedString, analysis: AnalysisResult?) -> [DocxNumberingPatcher.NumberingTarget] {
    let aiTypes = analysis?.paragraphTypes ?? [:]
    let aiLevels = analysis?.paragraphLevels ?? [:]
    let nonNumberedTypes: Set<LegalParagraphType> = [.headerMetadata, .documentTitle, .intro, .heading, .statementOfTruth, .signature, .quote]

    var targets: [DocxNumberingPatcher.NumberingTarget] = []
    let ns = doc.string as NSString
    let (_, bodyStartPara) = findBodyStartIndexWithParaIndex(in: doc.string, analysis: analysis)

    // Improved patterns for level detection (more flexible)
    // Optional number prefix: "145. " or "145) " or "145 " (number followed by punctuation and optional space)
    let optionalNumPrefix = "(?:\\d+[.):]?\\s*)?"

    // Level 1 (subparagraph): "(a)", "a)", "a." - single letter
    // Made more flexible: space after marker is now optional with \\s* instead of \\s+
    let level1Patterns = [
        "^\\s*" + optionalNumPrefix + "\\([a-zA-Z]\\)\\s*",           // "(a)" or "145. (a)"
        "^\\s*" + optionalNumPrefix + "[a-zA-Z]\\)\\s*",               // "a)" or "145. a)"
        "^\\s*" + optionalNumPrefix + "[a-zA-Z]\\.\\s+",               // "a. " or "145. a. " (with space)
        "^\\s*" + optionalNumPrefix + "[a-zA-Z]\\.[A-Z]",              // "a.Text" (no space, capital letter follows)
        "^\\s*\\([a-zA-Z]\\)\\s*",                                      // Just "(a)" at start
        "^\\s*[a-zA-Z]\\)\\s*"                                          // Just "a)" at start
    ]

    // Level 2 (sub-subparagraph): "(i)", "i)", "i." - roman numerals (ii, iii, iv, etc.)
    // Note: Single 'i' is ambiguous so we check multi-char roman numerals
    let level2Patterns = [
        "^\\s*" + optionalNumPrefix + "\\((?:ii|iii|iv|vi|vii|viii|ix|xi|xii)[ivxIVX]*\\)\\s*",  // "(ii)", "(iii)"
        "^\\s*" + optionalNumPrefix + "(?:ii|iii|iv|vi|vii|viii|ix|xi|xii)[ivxIVX]*\\)\\s*",     // "ii)", "iii)"
        "^\\s*" + optionalNumPrefix + "(?:ii|iii|iv|vi|vii|viii|ix|xi|xii)[ivxIVX]*\\.\\s*",     // "ii. ", "iii. "
        "^\\s*\\((?:ii|iii|iv|vi|vii|viii|ix|xi|xii)[ivxIVX]*\\)\\s*"                            // Just "(ii)" at start
    ]

    // First pass: collect all paragraph texts for context-based detection
    var paragraphTexts: [(index: Int, text: String, range: NSRange)] = []
    var tempIndex = 0
    ns.enumerateSubstrings(in: NSRange(location: 0, length: ns.length), options: .byParagraphs) { substring, substringRange, _, _ in
        let text = (substring ?? "").trimmingCharacters(in: .whitespacesAndNewlines)
        if tempIndex >= bodyStartPara && !text.isEmpty {
            paragraphTexts.append((tempIndex, text, substringRange))
        }
        tempIndex += 1
    }

    var sampleLogged = 0
    var inListContext = false  // Track if we're inside a list (after ":")
    var detectionSamplesLogged = 0

    Logger.shared.log("Starting level detection for \(paragraphTexts.count) body paragraphs", category: "FORMAT")

    for (i, para) in paragraphTexts.enumerated() {
        let paraIndex = para.index
        let text = para.text
        let substringRange = para.range

        // Skip non-numbered types
        if let t = aiTypes[paraIndex], nonNumberedTypes.contains(t) { continue }

        // Skip Heuristics
        if text.lowercased().contains("statement of truth") || isHeading(text) { continue }

        // Determine Level
        var level = 0
        var detectionMethod = "none"

        // Pattern detection first
        let isLevel2 = level2Patterns.contains { pattern in
            text.range(of: pattern, options: .regularExpression) != nil
        }
        let isLevel1 = !isLevel2 && level1Patterns.contains { pattern in
            text.range(of: pattern, options: .regularExpression) != nil
        }
        if isLevel2 {
            level = 2; detectionMethod = "pattern-roman"
        } else if isLevel1 {
            level = 1; detectionMethod = "pattern-letter"
        }

        // AI override (use max to respect explicit markers)
        if let aiLevel = aiLevels[paraIndex] {
            if aiLevel > level { level = aiLevel; detectionMethod = "AI" }
        }

        // Context-based detection when still level 0
        if level == 0 && i > 0 {
            let prevText = paragraphTexts[i - 1].text
            let prevEndsWithColon = prevText.hasSuffix(":") || prevText.hasSuffix("that:") ||
                                    prevText.hasSuffix(": ") || prevText.hasSuffix(":\n")
            let thisEndsWithSemicolon = text.hasSuffix(";") ||
                                         text.hasSuffix("; and") ||
                                         text.hasSuffix("; or") ||
                                         text.hasSuffix(";and") ||
                                         text.hasSuffix(";or") ||
                                         text.hasSuffix("; ") ||
                                         text.hasSuffix(";\n")
            let thisEndsWithPeriodAfterList = inListContext && text.hasSuffix(".")

            if detectionSamplesLogged < 10 && (prevEndsWithColon || inListContext) {
                Logger.shared.log("Context check para \(paraIndex): prevColon=\(prevEndsWithColon) inList=\(inListContext) endsSemi=\(thisEndsWithSemicolon) text='\(text.suffix(30))'", category: "FORMAT")
                detectionSamplesLogged += 1
            }

            if prevEndsWithColon { inListContext = true }

            if inListContext && (thisEndsWithSemicolon || thisEndsWithPeriodAfterList) {
                level = 1
                detectionMethod = "context-semicolon"
                if text.hasSuffix(".") && !text.hasSuffix("etc.") { inListContext = false }
            } else if !thisEndsWithSemicolon && !prevEndsWithColon {
                inListContext = false
            }
        }

        // Fallback: indentation if still level 0
        if level == 0 && substringRange.location < doc.length {
            let safeLocation = min(substringRange.location, doc.length - 1)
            if let style = doc.attribute(.paragraphStyle, at: safeLocation, effectiveRange: nil) as? NSParagraphStyle {
                if style.headIndent >= 72 || style.firstLineHeadIndent >= 72 {
                    level = 2; detectionMethod = "indent-72"
                } else if style.headIndent >= 36 || style.firstLineHeadIndent >= 36 {
                    level = 1; detectionMethod = "indent-36"
                }
            }
        }

        // Log samples for debugging
        if sampleLogged < 30 && level > 0 {
            Logger.shared.log("Target para \(paraIndex) level \(level) (\(detectionMethod)): \(text.prefix(60))", category: "PATCH")
            sampleLogged += 1
        }

        // Also log some level 0 samples to see what's NOT being detected
        if sampleLogged < 5 && level == 0 && i < 20 {
            Logger.shared.log("Level0 para \(paraIndex): '\(text.prefix(50))...\(text.suffix(20))'", category: "FORMAT")
        }

        // We capture longer prefix for content matching
        targets.append(.init(paragraphIndex: paraIndex, level: level, textPrefix: String(text.prefix(80))))
    }

    // Log level distribution
    let levelCounts = Dictionary(grouping: targets, by: { $0.level }).mapValues { $0.count }
    Logger.shared.log("Target levels: \(levelCounts)", category: "PATCH")

    return targets
}
