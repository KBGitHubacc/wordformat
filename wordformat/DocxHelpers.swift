//
//  DocxHelpers.swift
//  wordformat
//
//  Created by Krisztian Bito on 08.12.2025.
//

import Foundation
import AppKit

// MARK: - DOCX Loading & Saving

/// Loads a `.docx` file and returns it as a mutable attributed string.
/// - Parameter url: The file URL pointing to a DOCX document.
/// - Throws: Any error raised when reading or decoding the file.
/// - Returns: A `NSMutableAttributedString` representation of the Word document.
func loadDocx(from url: URL) throws -> NSMutableAttributedString {
    let options: [NSAttributedString.DocumentReadingOptionKey: Any] = [
        .documentType: NSAttributedString.DocumentType.officeOpenXML
    ]
    
    // Using NSMutableAttributedString(url: ...) automatically parses the DOCX.
    let attributedText = try NSMutableAttributedString(
        url: url,
        options: options,
        documentAttributes: nil
    )
    
    return attributedText
}

/// Saves a formatted attributed string as a `.docx` document.
/// - Parameters:
///   - attributedString: The formatted document.
///   - url: The destination URL for the new DOCX file.
/// - Throws: Errors from data generation or file writing.
func saveDocx(_ attributedString: NSAttributedString, to url: URL) throws {
    let fullRange = NSRange(location: 0, length: attributedString.length)
    
    let attributes: [NSAttributedString.DocumentAttributeKey: Any] = [
        .documentType: NSAttributedString.DocumentType.officeOpenXML
    ]
    
    let data = try attributedString.data(from: fullRange, documentAttributes: attributes)
    
    try data.write(to: url, options: .atomic)
}

/// Produces a new output URL with `_UKLegal` appended before the `.docx` extension.
/// - Parameter inputURL: The original input document URL.
/// - Returns: A new URL in the same folder with `_UKLegal.docx` appended.
func makeOutputURL(from inputURL: URL) -> URL {
    let baseName = inputURL.deletingPathExtension().lastPathComponent
    let folder = inputURL.deletingLastPathComponent()
    
    let newName = baseName + "_UKLegal.docx"
    
    return folder.appendingPathComponent(newName)
}
