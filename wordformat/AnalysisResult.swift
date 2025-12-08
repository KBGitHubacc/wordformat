//
//  AnalysisResult.swift
//  wordformat
//
//  Created by Krisztian Bito on 08.12.2025.
//

import Foundation

/// Defines the types of content found in a UK Legal document.
enum LegalParagraphType: String, Codable {
    case headerMetadata // The case no, parties, etc.
    case documentTitle  // "WITNESS STATEMENT OF..."
    case intro          // "I, [Name], will say..."
    case heading        // "A. INTRODUCTION"
    case body           // Standard numbered paragraph
    case quote          // Block quote (indented, no number)
    case statementOfTruth // "I believe that the facts..."
    case signature      // Date and signature lines
    case unknown
}

/// A container for the AI analysis of the document structure.
struct AnalysisResult {
    /// A list of ranges and their determined type.
    /// The formatter will iterate through these and apply the correct style.
    var classifiedRanges: [FormattedRange] = []
    
    struct FormattedRange {
        var range: NSRange
        var type: LegalParagraphType
    }
}
