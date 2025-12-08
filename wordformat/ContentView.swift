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
    @State private var openAIKey: String = "" // NEW
    
    @State private var status: String = "Select a DOCX file to format."
    @State private var isProcessing: Bool = false

    var body: some View {
        VStack(alignment: .leading, spacing: 16) {
            Text("UK Legal DOCX Formatter (Markdown Engine)")
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
            
            // NEW: AI Section
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
            isProcessing = true
            
            Task {
                await processDocx(at: url)
            }
        }
    }

    @MainActor
    private func processDocx(at url: URL) async {
        defer { isProcessing = false }
        do {
            let header = LegalHeaderMetadata(
                tribunalName: tribunalName,
                caseReference: caseReference,
                applicantName: applicantName,
                respondentName: respondentName
            )

            let mutableDoc = try loadDocx(from: url)
            
            // AI Analysis Step
            var analysisResult: AnalysisResult? = nil
            if !openAIKey.isEmpty {
                status = "Analysing structure with AI..."
                let aiService = OpenAIService(apiKey: openAIKey)
                // Extract plain text for AI
                let plainText = mutableDoc.string
                do {
                    let ranges = try await aiService.analyseDocumentStructure(text: plainText)
                    analysisResult = AnalysisResult(classifiedRanges: ranges)
                } catch {
                    status += "\nAI Error: \(error.localizedDescription). Falling back to standard logic."
                }
            }

            // Reconstruction
            status += "\nRebuilding document..."
            applyUKLegalFormatting(to: mutableDoc, header: header, analysis: analysisResult)

            let outputURL = makeOutputURL(from: url)
            try saveDocx(mutableDoc, to: outputURL)

            status = "Success!\nSaved to: \(outputURL.path)"
        } catch {
            status = "Error: \(error.localizedDescription)"
        }
    }
}
