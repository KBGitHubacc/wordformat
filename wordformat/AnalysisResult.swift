//
//  AnalysisResult.swift
//  wordformat
//
//  Created by Krisztian Bito on 08.12.2025.
//

import Foundation

enum LegalParagraphType: String, Codable {
    case headerMetadata = "header"
    case documentTitle = "title"
    case intro = "intro"
    case heading = "heading"
    case body = "body"
    case quote = "quote"
    case statementOfTruth = "statementOfTruth"
    case signature = "signature"
    case unknown = "unknown"
}

struct AnalysisResult {
    var classifiedRanges: [FormattedRange] = []
    /// Optional per-paragraph level classification (0=main,1=sub,a;2=subsub,i)
    var paragraphLevels: [Int: Int] = [:]
    
    struct FormattedRange {
        var range: NSRange
        var type: LegalParagraphType
    }
}
