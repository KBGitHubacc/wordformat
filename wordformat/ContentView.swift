//
//  ContentView.swift
//  wordformat
//
//  Created by Krisztian Bito on 08.12.2025.
//

import SwiftUI
import AppKit
import UniformTypeIdentifiers

struct LegalHeaderMetadata {
    var tribunalName: String
    var caseReference: String
    var applicantName: String
    var respondentName: String
}

struct ContentView: View {
    @State private var caseReference: String = ""
    @State private var tribunalName: String = ""
    @State private var applicantName: String = ""
    @State private var respondentName: String = ""
    @State private var openAIKey: String = ""

    @State private var status: String = "Select a DOCX file to format."
    @State private var isProcessing: Bool = false

    var body: some View {
        VStack(alignment: .leading, spacing: 16) {
            Text("UK Legal DOCX Formatter")
                .font(.title)
                .bold()

            GroupBox("Header metadata") {
                VStack(alignment: .leading, spacing: 8) {
                    TextField("Tribunal name", text: $tribunalName)
                    TextField("Case reference", text: $caseReference)
                    TextField("Applicant", text: $applicantName)
                    TextField("Respondent", text: $respondentName)
                }
                .textFieldStyle(.roundedBorder)
            }

            GroupBox("AI Analysis (Optional)") {
                SecureField("OpenAI API Key (sk-...)", text: $openAIKey)
                    .textFieldStyle(.roundedBorder)
                Text("If provided, AI will be used to intelligently structure the document before formatting.")
                    .font(.caption)
                    .foregroundStyle(.secondary)
            }

            Button {
                openAndFormatDocx()
            } label: {
                if isProcessing {
                    ProgressView().controlSize(.small)
                    Text("Processing...")
                } else {
                    Text("Open DOCX and Format")
                }
            }
            .disabled(isProcessing)

            Divider()

            ScrollView {
                Text(status)
                    .frame(maxWidth: .infinity, alignment: .leading)
                    .textSelection(.enabled)
            }
        }
        .padding()
        .frame(minWidth: 600, minHeight: 500)
    }

    private func openAndFormatDocx() {
        let panel = NSOpenPanel()
        if let type = UTType(filenameExtension: "docx") { panel.allowedContentTypes = [type] }
        panel.allowsMultipleSelection = false

        if panel.runModal() == .OK, let url = panel.url {
            status = "Processing \(url.lastPathComponent)..."
            Logger.shared.log("Selected file: \(url.path)", category: "UI")
            isProcessing = true

            Task {
                await processDocxPreservingStructure(at: url)
            }
        }
    }

    /// New approach: Preserve original document structure (tables, images, etc.)
    /// Only apply numbering via XML patching
    @MainActor
    private func processDocxPreservingStructure(at url: URL) async {
        defer { isProcessing = false }
        do {
            // 1. Load DOCX for analysis only (we won't save this back)
            let docForAnalysis = try loadDocx(from: url)
            Logger.shared.log("Loaded DOCX for analysis. Length \(docForAnalysis.length) chars", category: "DOCX")
            Logger.shared.log("Log file path: \(Logger.shared.currentLogPath())", category: "LOG")

            // 2. AI Analysis Step (style-aware)
            var analysisResult: AnalysisResult? = nil
            if !openAIKey.isEmpty {
                status = "Analysing structure with AI..."
                let aiService = OpenAIService(apiKey: openAIKey)
                do {
                    analysisResult = try await aiService.analyseDocument(doc: docForAnalysis)
                    Logger.shared.log("AI paragraph types: \(analysisResult?.paragraphTypes.count ?? 0)", category: "AI")
                    Logger.shared.log("AI paragraph levels: \(analysisResult?.paragraphLevels.count ?? 0)", category: "AI")
                } catch {
                    status += "\nAI Error: \(error.localizedDescription).\nProceeding with offline heuristics only."
                    Logger.shared.log(error: error, category: "AI")
                }
            } else {
                Logger.shared.log("AI key not provided. Skipping AI analysis.", category: "AI")
            }

            // 3. Copy original DOCX to output location (preserves tables, images, all formatting)
            let outputURL = makeOutputURL(from: url)
            let fm = FileManager.default
            if fm.fileExists(atPath: outputURL.path) {
                try fm.removeItem(at: outputURL)
            }
            try fm.copyItem(at: url, to: outputURL)
            Logger.shared.log("Copied original DOCX to \(outputURL.path)", category: "DOCX")

            // 4. Build numbering targets from analysis
            status += "\nAnalysing paragraphs..."
            let targets = buildNumberingTargetsFromAnalysis(
                doc: docForAnalysis, // use unmodified doc with original markers
                analysis: analysisResult
            )

            // 5. Apply numbering via XML patching (preserves document structure)
            if !targets.isEmpty {
                status += "\nApplying numbering to \(targets.count) paragraphs..."
                Logger.shared.log("Patcher will apply numbering to \(targets.count) paragraphs", category: "PATCH")
                let patcher = DocxNumberingPatcher()
                do {
                    try patcher.applyNumbering(to: outputURL, targets: targets)
                    Logger.shared.log("Patcher completed successfully", category: "PATCH")
                } catch {
                    Logger.shared.log("Patcher failed: \(error)", category: "PATCH")
                    status += "\nWarning: Numbering patch failed - \(error.localizedDescription)"
                }
            } else {
                Logger.shared.log("Patcher skipped: no numbering targets detected", category: "PATCH")
            }

            status = "Success!\nSaved to: \(outputURL.path)"
        } catch {
            status = "Error: \(error.localizedDescription)"
            Logger.shared.log(error: error, category: "APP")
        }
    }
}
