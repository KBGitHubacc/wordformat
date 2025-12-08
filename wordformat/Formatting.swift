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
    var listCounter = 1
    var isInList = false
    
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
        
        // Close list if we were in one but this item is NOT a body paragraph
        if isInList && item.type != .body {
            html += "</ol>"
            isInList = false
        }
        
        switch item.type {
        case .documentTitle:
            html += "<p class='title'>\(cleanText)</p>"
            
        case .intro:
            html += "<p>\(cleanText)</p>"
            
        case .heading:
            html += "<p class='heading'>\(cleanText)</p>"
            
        case .body:
            // List Handling
            // 1. Check for Sub-points "(a)"
            if isSubPoint(cleanText) {
                // To support nested lists in flat HTML generation, we close the main list,
                // start a type='a' list. This is complex to get right for continuity.
                // Simplified robust approach: Treat as an indented paragraph or a bullet.
                // Ideally: Word handles nested lists best if they are strictly hierarchical.
                
                // For now, we will render sub-points as part of the main list but manually styled,
                // OR simpler: Use a separate list type.
                
                // Let's stick to the MAIN numbered list for the "1, 2, 3" requirement.
                // If it's a sub-point, we strip the marker and just indent it?
                // No, user wants sub-paragraph numbering kept.
                
                // If we are in a main number list, pause it?
                // Let's assume for this version we format MAIN paragraphs as <ol>
                // and sub-paragraphs we print as text to avoid breaking the main sequence 1..2..3
                
                // STRATEGY: Strip the manual "1." but KEEP the manual "(a)".
                // Only apply <ol> to the main points.
                
                let stripped = stripManualNumbering(cleanText)
                if isInList {
                    html += "</ol>"
                    isInList = false
                }
                // Render sub-point as indented block
                html += "<p style='margin-left: 72pt;'>\(cleanText)</p>"
                
            } else {
                // Main Paragraph "1."
                if !isInList {
                    // Start list at current counter
                    html += "<ol start='\(listCounter)'>"
                    isInList = true
                }
                
                // Strip "1." or "304." so we don't get double numbers
                let content = stripManualNumbering(cleanText)
                html += "<li>\(content)</li>"
                listCounter += 1
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
    
    if isInList {
        html += "</ol>"
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
    // Removes "1.", "1 ", "304." from start of string
    let pattern = "^\\s*\\d+[.)]\\s+"
    if let regex = try? NSRegularExpression(pattern: pattern) {
        let range = NSRange(location: 0, length: text.utf16.count)
        return regex.stringByReplacingMatches(in: text, options: [], range: range, withTemplate: "")
    }
    return text
}

private func isSubPoint(_ text: String) -> Bool {
    // Detects (a), a., (i)
    return text.range(of: "^\\s*\\(?[a-zA-Z]\\)[.)]", options: .regularExpression) != nil
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
