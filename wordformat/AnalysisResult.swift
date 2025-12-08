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
    /// Legacy range-based classifications (kept for backward compatibility).
    var classifiedRanges: [FormattedRange] = []
    /// Index-precise paragraph type map produced by style-aware AI.
    var paragraphTypes: [Int: LegalParagraphType] = [:]
    /// Optional per-paragraph level classification (0=main,1=sub,a;2=subsub,i)
    var paragraphLevels: [Int: Int] = [:]
    
    struct FormattedRange {
        var range: NSRange
        var type: LegalParagraphType
    }
}

struct ParagraphMetadata: Encodable {
    let id: Int
    let text: String
    let isBold: Bool
    let isUppercased: Bool
    let isCentered: Bool
    let wordCount: Int
}
