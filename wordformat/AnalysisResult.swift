//
//  AnalysisResult.swift
//  wordformat
//
//  Created by Krisztian Bito on 08.12.2025.
//

import Foundation

/// Placeholder for AI-driven document analysis.
/// Extend with additional fields as needed.
struct AnalysisResult {
    /// Ranges in the document that should be styled as headings.
    /// These are NSRange in the coordinate space of the full attributed string.
    var headingRanges: [NSRange] = []

    /// Ranges in the document that represent text tables to convert to real tables.
    var tableRanges: [NSRange] = []
}
