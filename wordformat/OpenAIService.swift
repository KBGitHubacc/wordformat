//
//  OpenAIService.swift
//  wordformat
//
//  Created by Krisztian Bito on 08.12.2025.
//

import Foundation
import AppKit

struct OpenAIService {
    private let apiKey: String
    private static var cachedBestModel: String?
    
    init(apiKey: String) {
        self.apiKey = apiKey
    }
    
    // MARK: - Attribute Extraction
    
    /// Extract paragraph text plus key style hints to improve AI classification.
    func extractParagraphMetadata(from doc: NSAttributedString) -> [ParagraphMetadata] {
        var metas: [ParagraphMetadata] = []
        let ns = doc.string as NSString
        var index = 0
        ns.enumerateSubstrings(in: NSRange(location: 0, length: ns.length), options: .byParagraphs) { substring, range, _, _ in
            defer { index += 1 }
            let raw = substring ?? ""
            let text = raw.trimmingCharacters(in: .whitespacesAndNewlines)
            if text.isEmpty { return }
            
            let font = doc.attribute(.font, at: range.location, effectiveRange: nil) as? NSFont
            let style = doc.attribute(.paragraphStyle, at: range.location, effectiveRange: nil) as? NSParagraphStyle
            let isBold = font?.fontDescriptor.symbolicTraits.contains(.bold) ?? false
            let isCaps = text.count > 3 && text == text.uppercased()
            let isCentered = style?.alignment == .center
            let words = text.components(separatedBy: .whitespaces).filter { !$0.isEmpty }
            
            metas.append(ParagraphMetadata(
                id: index,
                text: text,
                isBold: isBold,
                isUppercased: isCaps,
                isCentered: isCentered,
                wordCount: words.count
            ))
        }
        return metas
    }

    /// Returns the preferred available model id from the account, with graceful fallback.
    /// Priority order (skips code/embedding/audio models): gpt-5.1-mini, gpt-5-mini, gpt-5.1, gpt-5,
    /// gpt-4.1, gpt-4.1-mini, gpt-4o-mini, gpt-4o.
    private func selectBestModel() async throws -> String {
        if let cached = Self.cachedBestModel { return cached }
        
        Logger.shared.log("Requesting model list…", category: "AI")
        struct ModelList: Decodable {
            struct Model: Decodable { let id: String }
            let data: [Model]
        }
        
        var request = URLRequest(url: URL(string: "https://api.openai.com/v1/models")!)
        request.httpMethod = "GET"
        request.addValue("Bearer \(apiKey)", forHTTPHeaderField: "Authorization")
        request.timeoutInterval = 30
        
        let (data, _) = try await URLSession.shared.data(for: request)
        Logger.shared.log("Received AI response (\(data.count) bytes)", category: "AI")
        let list = try JSONDecoder().decode(ModelList.self, from: data)
        let ids = list.data.map { $0.id }
        Logger.shared.log("Model list: \(ids)", category: "AI")
        
        let priority = [
            "gpt-5.1-mini",
            "gpt-5-mini",
            "gpt-5.1",
            "gpt-5",
            "gpt-4.1",
            "gpt-4.1-mini",
            "gpt-4o-mini",
            "gpt-4o"
        ]
        
        func isUsable(_ id: String) -> Bool {
            let lowered = id.lowercased()
            if lowered.contains("codex") { return false }
            if lowered.contains("whisper") { return false }
            if lowered.contains("audio") { return false }
            if lowered.contains("embed") { return false }
            if lowered.contains("tts") { return false }
            if lowered.contains("dall-e") { return false }
            return true
        }
        
        if let best = priority.first(where: { ids.contains($0) && isUsable($0) }) {
            Self.cachedBestModel = best
            Logger.shared.log("Selected model: \(best)", category: "AI")
            return best
        }
        
        // Fallback to a safe, widely available chat model.
        let fallback = ids.first(where: { $0.hasPrefix("gpt-4o") && isUsable($0) }) ?? "gpt-4o-mini"
        Self.cachedBestModel = fallback
        Logger.shared.log("Using fallback model: \(fallback)", category: "AI")
        return fallback
    }
    
    func analyseDocumentStructure(text: String) async throws -> [AnalysisResult.FormattedRange] {
        let url = URL(string: "https://api.openai.com/v1/chat/completions")!
        var request = URLRequest(url: url)
        request.httpMethod = "POST"
        request.addValue("Bearer \(apiKey)", forHTTPHeaderField: "Authorization")
        request.addValue("application/json", forHTTPHeaderField: "Content-Type")
        request.timeoutInterval = 120
        
        // We send the first 8k characters to reduce latency/timeouts while keeping structure context
        let truncatedText = String(text.prefix(8000))
        
        // Resolve best model available; fall back silently if listing fails.
        let model: String
        do {
            model = try await selectBestModel()
        } catch {
            model = "gpt-4o-mini"
            Logger.shared.log("Model listing failed: \(error.localizedDescription). Falling back to \(model)", category: "AI")
        }
        
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
        
        // Use a widely-available model to avoid "model not found" errors.
        let body: [String: Any] = [
            "model": model,
            "messages": [
                ["role": "system", "content": "You are a legal document structure analyzer."],
                ["role": "user", "content": prompt]
            ],
            "response_format": ["type": "json_object"]
        ]
        
        request.httpBody = try JSONSerialization.data(withJSONObject: body)
        Logger.shared.log("Sending chat request to \(url.host ?? "api.openai.com") with model \(model) and body \(truncatedText.count) chars", category: "AI")
        
        let (responseData, _) = try await URLSession.shared.data(for: request)
        
        // Parse Response
        guard let json = try JSONSerialization.jsonObject(with: responseData) as? [String: Any],
              let choices = json["choices"] as? [[String: Any]],
              let message = choices.first?["message"] as? [String: Any],
              let content = message["content"] as? String,
              let contentData = content.data(using: .utf8) else {
            throw NSError(domain: "OpenAI", code: -1, userInfo: [NSLocalizedDescriptionKey: "Invalid response"])
        }
        
        struct AIResponse: Decodable {
            struct Block: Decodable {
                let type: String
                let text: String
            }
            let blocks: [Block]
        }
        
        let responseObj = try JSONDecoder().decode(AIResponse.self, from: contentData)
        
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
    
    /// Style-aware analysis: chunk paragraphs with style metadata and classify by index.
    func analyseDocument(doc: NSAttributedString) async throws -> AnalysisResult {
        let metas = extractParagraphMetadata(from: doc)
        Logger.shared.log("AI style-aware: \(metas.count) non-empty paragraphs to classify", category: "AI")
        let allParagraphsAligned = extractAllParagraphs(from: doc)
        Logger.shared.log("Total paragraphs (including empty): \(allParagraphsAligned.count)", category: "AI")
        
        let chunkSize = 40
        var paragraphTypes: [Int: LegalParagraphType] = [:]
        
        var start = 0
        while start < metas.count {
            let end = min(start + chunkSize, metas.count)
            let chunk = Array(metas[start..<end])
            do {
                let chunkResult = try await sendMetaBatchToAI(chunk: chunk)
                paragraphTypes.merge(chunkResult) { _, new in new }
                Logger.shared.log("AI chunk \(start)-\(end) classified \(chunkResult.count)", category: "AI")
            } catch {
                Logger.shared.log("AI chunk \(start)-\(end) failed: \(error.localizedDescription)", category: "AI")
            }
            start = end
        }
        
        // Optionally also derive levels from text-only routine to reuse existing logic
        let levels = try? await analyseParagraphLevels(paragraphs: allParagraphsAligned)
        
        return AnalysisResult(
            classifiedRanges: [],
            paragraphTypes: paragraphTypes,
            paragraphLevels: levels ?? [:]
        )
    }
    
    private func sendMetaBatchToAI(chunk: [ParagraphMetadata]) async throws -> [Int: LegalParagraphType] {
        let model = try await selectBestModel()
        let jsonData = try JSONEncoder().encode(chunk)
        let jsonString = String(data: jsonData, encoding: .utf8) ?? "[]"
        
        let prompt = """
        You are a legal document formatter.
        Classify each paragraph in the JSON array below.
        
        Input element fields: id, text, isBold, isUppercased, isCentered, wordCount.
        
        Type rules:
        - "header": court/tribunal names, case refs, parties, BETWEEN, up to but not including the witness statement title.
        - "title": main document title (e.g., WITNESS STATEMENT OF ...), often uppercase & bold & centered.
        - "intro": the "I, Name, will say as follows" paragraph.
        - "heading": short (usually < 12 words), often bold/uppercase section headings like "A. INTRODUCTION".
        - "body": main numbered narrative paragraphs.
        - "statementOfTruth": contains statement of truth language.
        - "signature": signature/date lines.
        - "quote": indented or obviously quoted blocks.
        
        Return only JSON: {"classifications":[{"id":0,"type":"body"}, ...]}
        """
        
        var request = URLRequest(url: URL(string: "https://api.openai.com/v1/chat/completions")!)
        request.httpMethod = "POST"
        request.addValue("Bearer \(apiKey)", forHTTPHeaderField: "Authorization")
        request.addValue("application/json", forHTTPHeaderField: "Content-Type")
        request.timeoutInterval = 120
        
        let body: [String: Any] = [
            "model": model,
            "messages": [
                ["role": "system", "content": "You categorize legal paragraphs accurately with attention to style cues."],
                ["role": "user", "content": prompt + "\n\nInput:\n\(jsonString)"]
            ],
            "response_format": ["type": "json_object"]
        ]
        
        request.httpBody = try JSONSerialization.data(withJSONObject: body)
        Logger.shared.log("AI meta chunk send (\(chunk.count) items, model \(model))", category: "AI")
        let (data, _) = try await URLSession.shared.data(for: request)
        Logger.shared.log("AI meta chunk recv \(data.count) bytes", category: "AI")
        
        struct Resp: Decodable {
            struct Item: Decodable { let id: Int; let type: String }
            let classifications: [Item]
        }
        
        guard let json = try JSONSerialization.jsonObject(with: data) as? [String: Any],
              let choices = json["choices"] as? [[String: Any]],
              let msg = choices.first?["message"] as? [String: Any],
              let content = msg["content"] as? String,
              let cdata = content.data(using: .utf8) else {
            throw NSError(domain: "OpenAI", code: -1, userInfo: [NSLocalizedDescriptionKey: "Bad response structure"])
        }
        
        let parsed = try JSONDecoder().decode(Resp.self, from: cdata)
        var map: [Int: LegalParagraphType] = [:]
        for item in parsed.classifications {
            let t = LegalParagraphType(rawValue: item.type) ?? .body
            map[item.id] = t
        }
        return map
    }
    
    // Extract every paragraph string (including empty) to keep indices aligned.
    private func extractAllParagraphs(from doc: NSAttributedString) -> [String] {
        var arr: [String] = []
        let ns = doc.string as NSString
        ns.enumerateSubstrings(in: NSRange(location: 0, length: ns.length), options: .byParagraphs) { substring, _, _, _ in
            arr.append(substring ?? "")
        }
        return arr
    }
    
    /// Chunked per-paragraph analysis to classify level (0/1/2) and type.
    /// Returns a dictionary: paragraphIndex -> level (0-based main/sub/subsub).
    func analyseParagraphLevels(paragraphs: [String]) async throws -> [Int: Int] {
        var result: [Int: Int] = [:]
        let chunkSize = 40
        let overlap = 0
        var start = 0
        while start < paragraphs.count {
            let end = min(start + chunkSize, paragraphs.count)
            let slice = paragraphs[start..<end]
            let prompt = buildParagraphPrompt(startIndex: start, slice: Array(slice))
            let model = try await selectBestModel()
            
            var request = URLRequest(url: URL(string: "https://api.openai.com/v1/chat/completions")!)
            request.httpMethod = "POST"
            request.addValue("Bearer \(apiKey)", forHTTPHeaderField: "Authorization")
            request.addValue("application/json", forHTTPHeaderField: "Content-Type")
            request.timeoutInterval = 120
            
            let body: [String: Any] = [
                "model": model,
                "messages": [
                    ["role": "system", "content": "You classify legal paragraphs into numbered levels."],
                    ["role": "user", "content": prompt]
                ],
                "response_format": ["type": "json_object"]
            ]
            request.httpBody = try JSONSerialization.data(withJSONObject: body)
            
            Logger.shared.log("AI chunk \(start)-\(end) sending with model \(model)", category: "AI")
            let (data, _) = try await URLSession.shared.data(for: request)
            Logger.shared.log("AI chunk \(start)-\(end) response \(data.count) bytes", category: "AI")
            
            struct ChunkResp: Decodable {
                struct Item: Decodable { let i: Int; let level: Int }
                let items: [Item]
            }
            guard let choices = try JSONSerialization.jsonObject(with: data) as? [String: Any],
                  let ch = choices["choices"] as? [[String: Any]],
                  let msg = ch.first?["message"] as? [String: Any],
                  let content = msg["content"] as? String,
                  let cdata = content.data(using: .utf8) else { continue }
            let parsed = try JSONDecoder().decode(ChunkResp.self, from: cdata)
            for item in parsed.items {
                result[item.i] = item.level
            }
            
            start = end - overlap
        }
        Logger.shared.log("AI paragraph levels collected: \(result.count) of \(paragraphs.count)", category: "AI")
        return result
    }
    
    private func buildParagraphPrompt(startIndex: Int, slice: [String]) -> String {
        var lines: [String] = []
        for (offset, text) in slice.enumerated() {
            let idx = startIndex + offset
            let clean = text.replacingOccurrences(of: "\n", with: " ").trimmingCharacters(in: .whitespacesAndNewlines)
            lines.append("[\(idx)] \(clean)")
        }
        let joined = lines.joined(separator: "\n")
        return """
        For each paragraph below, return JSON: {"items":[{"i":index,"level":N}]}
        level 0 = main numbered paragraph
        level 1 = subparagraph (a,b,c)
        level 2 = sub-subparagraph (i,ii,iii)
        Headings or titles should be omitted from the list.
        Use best judgment based on numbering/indent/context.
        Paragraphs:
        \(joined)
        """
    }
}
