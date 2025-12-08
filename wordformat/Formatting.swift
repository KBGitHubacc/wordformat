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
    static let fontSize: Int = 12
}

// MARK: - Main Entry Point

func applyUKLegalFormatting(
    to document: NSMutableAttributedString,
    header: LegalHeaderMetadata,
    analysis: AnalysisResult?
) {
    guard document.length > 0 else { return }
    
    // 1. Analyse Structure
    // Identify what is a Header, Body, Title, etc.
    let structure = detectDocumentStructure(in: document)
    
    // 2. Generate HTML Representation
    // We rebuild the document as HTML. This is the most reliable way to force
    // "Dynamic Lists" (Word <w:numPr>) to appear in the final DOCX.
    let htmlString = generateHTML(
        from: document,
        structure: structure,
        metadata: header
    )
    
    // 3. Convert HTML back to Attributed String
    // The system parser applies the correct hidden TextKit attributes for lists.
    if let newAttrString = convertHTMLToAttributedString(htmlString) {
        document.setAttributedString(newAttrString)
    }
}

// MARK: - HTML Generation Engine

private func generateHTML(
    from original: NSAttributedString,
    structure: [AnalysisResult.FormattedRange],
    metadata: LegalHeaderMetadata
) -> String {
    
    var html = """
    <!DOCTYPE html>
    <html>
    <head>
    <style>
        body {
            font-family: '\(LegalFormattingDefaults.fontFamily)';
            font-size: \(LegalFormattingDefaults.fontSize)pt;
            line-height: 1.2;
        }
        p {
            margin-bottom: 12pt;
            text-align: justify;
        }
        ol {
            margin-bottom: 12pt;
            padding-left: 36pt; /* Hanging Indent */
            list-style-type: decimal;
            list-style-position: outside;
        }
        ol.sublist {
            list-style-type: lower-alpha;
            list-style-position: outside;
        }
        li {
            text-align: justify;
            margin-bottom: 12pt;
            padding-left: 10pt;
        }
        .header { text-align: center; font-weight: bold; margin-bottom: 0px; }
        .title { text-align: center; font-weight: bold; margin-top: 24pt; margin-bottom: 24pt; text-transform: uppercase; }
        .heading { text-align: left; font-weight: bold; margin-top: 18pt; margin-bottom: 6pt; }
        .quote { margin-left: 36pt; margin-right: 36pt; }
    </style>
    </head>
    <body>
    """
    
    // -- Step A: Insert Header --
    // We construct the header block manually to ensure it's always perfect.
    // If the original doc had a header, we skip it in the loop below to avoid duplication.
    html += """
    <p class='header'>IN THE \(metadata.tribunalName.uppercased())</p>
    <p class='header' style='font-weight:normal'>Case Reference: \(metadata.caseReference)</p>
    <br>
    <p class='header'>BETWEEN:</p>
    <p class='header'>\(metadata.applicantName.uppercased())</p>
    <p class='header' style='font-weight:normal'>Applicant</p>
    <p class='header'>-and-</p>
    <p class='header'>\(metadata.respondentName.uppercased())</p>
    <p class='header' style='font-weight:normal'>Respondent</p>
    <br><br>
    """
    
    // -- Step B: Process Paragraphs --
    // Track numbering so the resulting DOCX relies on Word's native list
    // numbering instead of fixed text. This preserves numbering if paragraphs
    // are added or removed later.
    var isInMainList = false
    var hasOpenMainListItem = false
    var isInSubList = false
    var lastMainNumber: Int = 0
    var lastSubIndex: Int = 0
    
    for item in structure {
        // Skip header items from the original doc as we just generated a fresh one
        if item.type == .headerMetadata { continue }
        
        // Extract plain text
        let originalText = (original.string as NSString).substring(with: item.range)
        let cleanText = originalText.trimmingCharacters(in: .whitespacesAndNewlines)
        
        // Handle empty lines (paragraph breaks)
        if cleanText.isEmpty {
            continue
        }
        
        // Close list structures if we are switching away from body paragraphs
        if item.type != .body {
            if isInSubList {
                html += "</ol></li>"
                isInSubList = false
                hasOpenMainListItem = false
            } else if hasOpenMainListItem {
                html += "</li>"
                hasOpenMainListItem = false
            }
            if isInMainList {
                html += "</ol>"
                isInMainList = false
            }
        }
        
        switch item.type {
        case .documentTitle:
            html += "<p class='title'>\(cleanText)</p>"
            
        case .intro:
            html += "<p>\(cleanText)</p>"
            
        case .heading:
            html += "<p class='heading'>\(cleanText)</p>"
            
        case .body:
            let numbering = parseParagraphNumber(cleanText)
            let content = stripManualNumbering(cleanText)

            switch numbering {
            case .main(let number):
                // Close any open sublist tied to the previous main item
                if isInSubList {
                    html += "</ol></li>"
                    isInSubList = false
                    hasOpenMainListItem = false
                    lastSubIndex = 0
                } else if hasOpenMainListItem {
                    html += "</li>"
                    hasOpenMainListItem = false
                }

                // Start or restart the main list at the detected number
                if !isInMainList {
                    html += "<ol start='\(number)'>"
                    isInMainList = true
                } else if number != lastMainNumber + 1 {
                    // Restart list if manual numbering jumps (preserve original numbering)
                    html += "</ol><ol start='\(number)'>"
                }

                html += "<li><p>\(content)</p>"
                hasOpenMainListItem = true
                lastMainNumber = number

            case .sub(let letter):
                // Ensure we have a main list + open item to attach the sublist to
                if !isInMainList {
                    html += "<ol start='1'>"
                    isInMainList = true
                    lastMainNumber = 1
                }
                if !hasOpenMainListItem {
                    html += "<li>"
                    hasOpenMainListItem = true
                }

                let startValue = max(1, letterListIndex(letter))

                if !isInSubList {
                    html += "<ol class='sublist' type='a' start='\(startValue)'>"
                    isInSubList = true
                }

                if startValue != lastSubIndex + 1 {
                    html += "<li value='\(startValue)'><p>\(content)</p></li>"
                } else {
                    html += "<li><p>\(content)</p></li>"
                }

                lastSubIndex = startValue

            case .none:
                // Treat as a standard paragraph outside of numbered lists
                if isInSubList {
                    html += "</ol></li>"
                    isInSubList = false
                    hasOpenMainListItem = false
                    lastSubIndex = 0
                } else if hasOpenMainListItem {
                    html += "</li>"
                    hasOpenMainListItem = false
                }
                if isInMainList {
                    html += "</ol>"
                    isInMainList = false
                }
                html += "<p>\(cleanText)</p>"
            }
            
        case .quote:
            html += "<p class='quote'>\(cleanText)</p>"
            
        case .statementOfTruth:
            html += "<p style='margin-top:24pt; font-weight:bold'>\(cleanText)</p>"
            
        case .signature:
            html += "<p>\(cleanText)</p>"
            
        default:
            html += "<p>\(cleanText)</p>"
        }
    }

    if isInSubList {
        html += "</ol>"
        isInSubList = false
    }
    if hasOpenMainListItem {
        html += "</li>"
        hasOpenMainListItem = false
    }
    if isInMainList {
        html += "</ol>"
        isInMainList = false
    }
    
    html += "</body></html>"
    return html
}

// MARK: - HTML Conversion Helper

private func convertHTMLToAttributedString(_ html: String) -> NSAttributedString? {
    guard let data = html.data(using: .utf8) else { return nil }
    
    do {
        return try NSMutableAttributedString(
            data: data,
            options: [
                .documentType: NSAttributedString.DocumentType.html,
                .characterEncoding: String.Encoding.utf8.rawValue
            ],
            documentAttributes: nil
        )
    } catch {
        print("Error converting HTML: \(error)")
        return nil
    }
}

// MARK: - Structure Detection (State Machine)

private func detectDocumentStructure(in document: NSAttributedString) -> [AnalysisResult.FormattedRange] {
    var ranges: [AnalysisResult.FormattedRange] = []
    let fullString = document.string as NSString
    let fullRange = NSRange(location: 0, length: fullString.length)
    
    let headerKeywords = ["case no", "case ref", "claim no", "in the", "tribunal", "between:", "applicant", "respondent", "-v-", "-and-"]
    let titleKeywords = ["witness statement"]
    let introKeywords = ["will say as follows", "states as follows", "say as follows"]
    let truthKeywords = ["statement of truth", "believe that the facts"]
    
    enum ScanState { case header, preBody, body, backMatter }
    var currentState: ScanState = .header
    
    fullString.enumerateSubstrings(in: fullRange, options: .byParagraphs) { substring, substringRange, _, _ in
        guard let rawText = substring else { return }
        let text = rawText.trimmingCharacters(in: .whitespacesAndNewlines)
        let lower = text.lowercased()
        
        if text.isEmpty { return }
        
        var type: LegalParagraphType = .body
        
        switch currentState {
        case .header:
            if containsAny(lower, keywords: titleKeywords) {
                type = .documentTitle
                currentState = .preBody
            } else if containsAny(lower, keywords: headerKeywords) || substringRange.location < 800 {
                type = .headerMetadata
            } else {
                type = .intro
                currentState = .preBody
            }
        case .preBody:
            if containsAny(lower, keywords: introKeywords) {
                type = .intro
                currentState = .body
            } else if containsAny(lower, keywords: titleKeywords) {
                type = .documentTitle
            } else {
                type = .intro
            }
        case .body:
            if containsAny(lower, keywords: truthKeywords) {
                type = .statementOfTruth
                currentState = .backMatter
            } else if isHeading(text) {
                type = .heading
            } else {
                type = .body
            }
        case .backMatter:
            type = .signature
        }
        ranges.append(AnalysisResult.FormattedRange(range: substringRange, type: type))
    }
    return ranges
}

// MARK: - Utilities

private func stripManualNumbering(_ text: String) -> String {
    // Removes "1.", "1 ", "304.", "(a)" from start of string
    let pattern = "^\\s*(?:\\d+|\\([a-zA-Z]\\))[.)]?\\s+"
    if let regex = try? NSRegularExpression(pattern: pattern) {
        let range = NSRange(location: 0, length: text.utf16.count)
        return regex.stringByReplacingMatches(in: text, options: [], range: range, withTemplate: "")
    }
    return text
}

private enum ParagraphNumbering {
    case main(Int)
    case sub(Character)
    case none
}

/// Parses a paragraph's leading numbering so we can rebuild it using
/// native Word list semantics rather than fixed text.
private func parseParagraphNumber(_ text: String) -> ParagraphNumbering {
    let nsRange = NSRange(location: 0, length: text.utf16.count)

    if let mainMatch = try? NSRegularExpression(pattern: "^\\s*(\\d+)[.)]?") {
        if let result = mainMatch.firstMatch(in: text, options: [], range: nsRange),
           let range = Range(result.range(at: 1), in: text),
           let value = Int(text[range]) {
            return .main(value)
        }
    }

    if let subMatch = try? NSRegularExpression(pattern: "^\\s*\\(?([a-zA-Z])\\)[.)]?") {
        if let result = subMatch.firstMatch(in: text, options: [], range: nsRange),
           let range = Range(result.range(at: 1), in: text),
           let char = text[range].first {
            return .sub(char)
        }
    }

    return .none
}

/// Returns the alphabetical index for a subparagraph marker (a/A -> 1).
private func letterListIndex(_ letter: Character) -> Int {
    let alphabet = Array("abcdefghijklmnopqrstuvwxyz")
    if let idx = alphabet.firstIndex(of: Character(letter.lowercased())) {
        return idx + 1
    }
    return 1
}

private func containsAny(_ text: String, keywords: [String]) -> Bool {
    for k in keywords { if text.contains(k) { return true } }
    return false
}

private func isHeading(_ text: String) -> Bool {
    let clean = text.trimmingCharacters(in: .whitespaces)
    guard clean.count < 100 else { return false }
    return clean.range(of: "^[A-Z0-9]+\\.\\s", options: .regularExpression) != nil
}
