//
//  OpenAIService.swift
//  wordformat
//
//  Created by Krisztian Bito on 08.12.2025.
//

import Foundation

struct OpenAIService {
    private let apiKey: String
    
    init(apiKey: String) {
        self.apiKey = apiKey
    }
    
    func analyseDocumentStructure(text: String) async throws -> [AnalysisResult.FormattedRange] {
        let url = URL(string: "https://api.openai.com/v1/chat/completions")!
        var request = URLRequest(url: url)
        request.httpMethod = "POST"
        request.addValue("Bearer \(apiKey)", forHTTPHeaderField: "Authorization")
        request.addValue("application/json", forHTTPHeaderField: "Content-Type")
        
        // We send the first 15k characters to avoid token limits, usually enough for structure
        let truncatedText = String(text.prefix(15000))
        
        let prompt = """
        Analyze the following legal document text. Return a JSON response with a list of blocks classifying the text.
        
        Categories:
        - "header": The court details, case reference, parties (BETWEEN...), up to the Witness Statement title.
        - "title": The main document title (e.g., "WITNESS STATEMENT OF...").
        - "intro": The introductory paragraph ("I, Name, will say...").
        - "body": The main numbered paragraphs of the witness statement.
        - "heading": Section headings (e.g., "A. INTRODUCTION").
        - "quote": Block quotes or indented text.
        - "statementOfTruth": The statement of truth paragraph.
        - "signature": The date and signature block.
        
        Return ONLY valid JSON in this format:
        {
            "blocks": [
                {"type": "header", "text": "start of text... end of text"},
                {"type": "body", "text": "..."}
            ]
        }
        
        Important:
        1. Do NOT summarize. Group consecutive paragraphs of the same type into one block where possible, OR list them sequentially.
        2. Ensure the "text" field matches the content so I can match it back to the document.
        
        Document Text:
        \(truncatedText)
        """
        
        let body: [String: Any] = [
            "model": "gpt-5-mini-2025-08-07",
            "messages": [
                ["role": "system", "content": "You are a legal document structure analyzer."],
                ["role": "user", "content": prompt]
            ],
            "response_format": ["type": "json_object"]
        ]
        
        request.httpBody = try JSONSerialization.data(withJSONObject: body)
        
        let (data, _) = try await URLSession.shared.data(for: request)
        
        // Parse Response
        guard let json = try JSONSerialization.jsonObject(with: data) as? [String: Any],
              let choices = json["choices"] as? [[String: Any]],
              let message = choices.first?["message"] as? [String: Any],
              let content = message["content"] as? String,
              let data = content.data(using: .utf8) else {
            throw NSError(domain: "OpenAI", code: -1, userInfo: [NSLocalizedDescriptionKey: "Invalid response"])
        }
        
        struct AIResponse: Decodable {
            struct Block: Decodable {
                let type: String
                let text: String
            }
            let blocks: [Block]
        }
        
        let responseObj = try JSONDecoder().decode(AIResponse.self, from: data)
        
        // Map back to ranges
        var ranges: [AnalysisResult.FormattedRange] = []
        let fullNSString = text as NSString
        
        for block in responseObj.blocks {
            let type: LegalParagraphType = LegalParagraphType(rawValue: block.type) ?? .body
            
            // Fuzzy search for the text in the document
            // We search within the full string. This is heuristic.
            let searchRange = fullNSString.range(of: block.text)
            if searchRange.location != NSNotFound {
                ranges.append(AnalysisResult.FormattedRange(range: searchRange, type: type))
            }
        }
        
        // If AI fails to cover everything, we might have gaps.
        // For this hybrid approach, if we get results, we use them.
        return ranges
    }
}
