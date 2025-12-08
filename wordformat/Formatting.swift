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
    static let bodySpacing: CGFloat = 12.0
}

// MARK: - Main Entry Point

func applyUKLegalFormatting(
    to document: NSMutableAttributedString,
    header: LegalHeaderMetadata,
    analysis: AnalysisResult?
) {
    guard document.length > 0 else { return }
    
    // 1. Analyse Structure (Identify Header, Body, Titles)
    let structure = detectDocumentStructure(in: document)
    
    // 2. Reconstruct the Document
    // Instead of patching the existing string, we build a new one.
    // This allows us to use HTML for lists (which guarantees dynamic numbering in Word)
    // and standard TextKit attributes for headers.
    let newDocument = reconstructDocument(from: document, structure: structure, metadata: header)
    
    // 3. Apply Global Font Normalisation (Safety pass)
    let fullRange = NSRange(location: 0, length: newDocument.length)
    applyBaseFont(to: newDocument, range: fullRange)
    
    // 4. Replace content
    document.setAttributedString(newDocument)
}

// MARK: - Reconstruction Engine

private func reconstructDocument(
    from original: NSAttributedString,
    structure: [AnalysisResult.FormattedRange],
    metadata: LegalHeaderMetadata
) -> NSMutableAttributedString {
    
    let output = NSMutableAttributedString()
    
    // -- Step A: Insert/Process Header --
    // Check if we detected an existing header
    let hasExistingHeader = structure.first?.type == .headerMetadata
    
    if !hasExistingHeader && !metadata.caseReference.isEmpty {
        // Generate new header if missing
        output.append(generateLegalHeader(metadata))
    }
    
    // -- Step B: Process Blocks --
    // We group consecutive 'body' paragraphs together to form a single HTML list.
    // Other types (headers, titles) are appended directly.
    
    var bodyBuffer: [String] = []
    
    for item in structure {
        let text = (original.string as NSString).substring(with: item.range).trimmingCharacters(in: .whitespacesAndNewlines)
        if text.isEmpty { continue }
        
        if item.type == .body {
            // Buffer body paragraphs to convert to HTML list later
            bodyBuffer.append(text)
        } else {
            // 1. Flush any buffered body paragraphs first
            if !bodyBuffer.isEmpty {
                if let listAttr = createDynamicListFromText(bodyBuffer) {
                    output.append(listAttr)
                }
                bodyBuffer.removeAll()
            }
            
            // 2. Process this non-body item (Title, Intro, Signature, etc.)
            let itemAttr = NSMutableAttributedString(string: text + "\n")
            applyStyle(to: itemAttr, type: item.type)
            output.append(itemAttr)
        }
    }
    
    // Flush any remaining body paragraphs at the end
    if !bodyBuffer.isEmpty {
        if let listAttr = createDynamicListFromText(bodyBuffer) {
            output.append(listAttr)
        }
    }
    
    return output
}

// MARK: - HTML List Generator (The "Brand New" Approach)

/// Converts an array of text strings into an NSAttributedString using HTML <ol> tags.
/// This forces the system to generate the correct NSTextList attributes that Word recognises.
private func createDynamicListFromText(_ paragraphs: [String]) -> NSAttributedString? {
    // 1. Clean the text (Strip existing manual numbers like "1.", "304.")
    let pattern = "^\\s*(\\d+|[a-zA-Z])+[.)]\\s+" // Matches "1.", "a)", "304."
    
    var listItemsHTML = ""
    
    for rawLine in paragraphs {
        // Remove old numbering so we don't get "1. 304. Text"
        let cleanLine = rawLine.replacingOccurrences(of: pattern, with: "", options: .regularExpression)
        
        // Detect nesting based on patterns (simple heuristic)
        // If the original line started with (a) or a., we ideally want a nested list.
        // For robustness in this version, we will stick to a single clean numeric list
        // as mixing levels in flat HTML string construction can be complex.
        // Word allows users to indent 'level 2' easily if the list object exists.
        
        listItemsHTML += "<li>\(cleanLine)</li>"
    }
    
    // 2. Build HTML Wrapper with CSS for Times New Roman
    let htmlString = """
    <html>
    <head>
    <style>
        body {
            font-family: 'Times New Roman';
            font-size: 12pt;
        }
        ol {
            margin-left: 0px;
            padding-left: 36px; /* Hanging indent simulation */
        }
        li {
            text-align: justify;
            margin-bottom: 12pt;
        }
    </style>
    </head>
    <body>
        <ol>
            \(listItemsHTML)
        </ol>
    </body>
    </html>
    """
    
    // 3. Convert HTML to NSAttributedString
    guard let data = htmlString.data(using: .utf8) else { return nil }
    
    do {
        let attrString = try NSMutableAttributedString(
            data: data,
            options: [
                .documentType: NSAttributedString.DocumentType.html,
                .characterEncoding: String.Encoding.utf8.rawValue
            ],
            documentAttributes: nil
        )
        // HTML import often adds a trailing newline, trim it if needed
        return attrString
    } catch {
        print("HTML conversion failed: \(error)")
        return nil
    }
}

// MARK: - Standard Styling (For Non-List Items)

private func applyStyle(to str: NSMutableAttributedString, type: LegalParagraphType) {
    let style = NSMutableParagraphStyle()
    style.lineHeightMultiple = 1.0
    
    let range = NSRange(location: 0, length: str.length)
    
    switch type {
    case .headerMetadata:
        style.alignment = .center
        style.paragraphSpacing = 0
        // Bold specific lines
        let text = str.string.lowercased()
        if text.contains("in the") || text.contains("between") || text.contains("witness") {
            applyBold(to: str, range: range)
        }
        
    case .documentTitle:
        style.alignment = .center
        style.paragraphSpacing = 24
        applyBold(to: str, range: range)
        // Uppercase
        let upper = str.string.uppercased()
        str.replaceCharacters(in: range, with: upper)
        
    case .intro:
        style.alignment = .left
        style.paragraphSpacing = 12
        
    case .heading:
        style.alignment = .left
        style.paragraphSpacingBefore = 18
        style.paragraphSpacing = 6
        applyBold(to: str, range: range)
        
    case .quote:
        style.alignment = .left
        style.headIndent = 36
        style.firstLineHeadIndent = 36
        style.paragraphSpacing = 12
        
    case .statementOfTruth:
        style.alignment = .left
        style.paragraphSpacingBefore = 24
        applyBold(to: str, range: range)
        
    case .signature:
        style.alignment = .left
        style.paragraphSpacing = 24
        
    default:
        style.alignment = .left
    }
    
    str.addAttribute(.paragraphStyle, value: style, range: range)
}

// MARK: - Document Analysis (State Machine)

private func detectDocumentStructure(in document: NSAttributedString) -> [AnalysisResult.FormattedRange] {
    var ranges: [AnalysisResult.FormattedRange] = []
    let fullString = document.string as NSString
    let fullRange = NSRange(location: 0, length: fullString.length)
    
    // Keywords
    let headerKeywords = ["case no", "case ref", "claim no", "in the", "tribunal", "between:", "-v-", "-and-"]
    let titleKeywords = ["witness statement"]
    let introKeywords = ["will say as follows", "states as follows"]
    let truthKeywords = ["statement of truth", "believe that the facts"]
    
    enum ScanState {
        case header, preBody, body, backMatter
    }
    
    var currentState: ScanState = .header
    
    fullString.enumerateSubstrings(in: fullRange, options: .byParagraphs) { substring, substringRange, _, _ in
        guard let text = substring?.trimmingCharacters(in: .whitespacesAndNewlines).lowercased(), !text.isEmpty else { return }
        
        var type: LegalParagraphType = .body
        
        switch currentState {
        case .header:
            if containsAny(text, keywords: titleKeywords) {
                type = .documentTitle
                currentState = .preBody
            } else if containsAny(text, keywords: headerKeywords) || substringRange.location < 800 {
                type = .headerMetadata
            } else {
                type = .intro
                currentState = .preBody
            }
            
        case .preBody:
            if containsAny(text, keywords: introKeywords) {
                type = .intro
                currentState = .body
            } else if containsAny(text, keywords: titleKeywords) {
                type = .documentTitle
            } else {
                type = .intro
            }
            
        case .body:
            if containsAny(text, keywords: truthKeywords) {
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

// MARK: - Helpers

private func generateLegalHeader(_ metadata: LegalHeaderMetadata) -> NSAttributedString {
    let content = """
    IN THE \(metadata.tribunalName.uppercased())
    Case Reference: \(metadata.caseReference)
    
    BETWEEN:
    
    \(metadata.applicantName.uppercased())
        Applicant
    
    -and-
    
    \(metadata.respondentName.uppercased())
        Respondent
    
    
    """
    let attr = NSMutableAttributedString(string: content)
    let style = NSMutableParagraphStyle()
    style.alignment = .center
    attr.addAttribute(.paragraphStyle, value: style, range: NSRange(location: 0, length: attr.length))
    return attr
}

private func applyBaseFont(to doc: NSMutableAttributedString, range: NSRange) {
    let baseFont = NSFont(name: LegalFormattingDefaults.fontFamily, size: LegalFormattingDefaults.fontSize)
        ?? NSFont.systemFont(ofSize: LegalFormattingDefaults.fontSize)
    
    doc.enumerateAttribute(.font, in: range, options: []) { value, subRange, _ in
        if let existing = value as? NSFont {
            let traits = existing.fontDescriptor.symbolicTraits
            let newDescriptor = baseFont.fontDescriptor.withSymbolicTraits(traits)
            if let newFont = NSFont(descriptor: newDescriptor, size: LegalFormattingDefaults.fontSize) {
                doc.addAttribute(.font, value: newFont, range: subRange)
            } else {
                // Fallback to baseFont if descriptor-based font creation fails
                doc.addAttribute(.font, value: baseFont, range: subRange)
            }
        } else {
            doc.addAttribute(.font, value: baseFont, range: subRange)
        }
    }
}

private func applyBold(to str: NSMutableAttributedString, range: NSRange) {
    let font = NSFont(name: LegalFormattingDefaults.fontFamily, size: LegalFormattingDefaults.fontSize)
        ?? NSFont.systemFont(ofSize: LegalFormattingDefaults.fontSize)
    let boldDesc = font.fontDescriptor.withSymbolicTraits(.bold)
    if let boldFont = NSFont(descriptor: boldDesc, size: LegalFormattingDefaults.fontSize) {
        str.addAttribute(.font, value: boldFont, range: range)
    } else {
        // Fallback to system bold if creation fails
        let fallbackBold = NSFont.boldSystemFont(ofSize: LegalFormattingDefaults.fontSize)
        str.addAttribute(.font, value: fallbackBold, range: range)
    }
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

