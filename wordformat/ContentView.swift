//
//  ContentView.swift
//  wordformat
//
//  Created by Krisztian Bito on 08.12.2025.
//

import SwiftUI
import AppKit
import UniformTypeIdentifiers

/// Metadata required to build the UK legal header for the formatted document.
struct LegalHeaderMetadata {
    var tribunalName: String
    var caseReference: String
    var applicantName: String
    var respondentName: String
}

struct ContentView: View {
    // MARK: - Header input state
    @State private var caseReference: String = ""
    @State private var tribunalName: String = ""
    @State private var applicantName: String = ""
    @State private var respondentName: String = ""

    // MARK: - UI state
    @State private var status: String = "Select a DOCX file to format."
    @State private var isProcessing: Bool = false

    // MARK: - Body
    var body: some View {
        VStack(alignment: .leading, spacing: 16) {
            Text("UK Legal DOCX Formatter")
                .font(.title)
                .bold()

            GroupBox("Header metadata") {
                VStack(alignment: .leading, spacing: 8) {
                    TextField("Tribunal name (e.g. Employment Tribunal London)", text: $tribunalName)
                    TextField("Case reference (e.g. 2401234/2025)", text: $caseReference)
                    TextField("Applicant name (e.g. John Smith)", text: $applicantName)
                    TextField("Respondent name (e.g. ACME Ltd)", text: $respondentName)
                }
                .textFieldStyle(.roundedBorder)
            }

            Button {
                openAndFormatDocx()
            } label: {
                if isProcessing {
                    ProgressView()
                        .controlSize(.small)
                    Text("Processing…")
                } else {
                    Text("Open DOCX and apply UK legal format")
                }
            }
            .keyboardShortcut(.defaultAction)
            .disabled(isProcessing)

            Divider()

            Text("Status:")
                .font(.headline)
            ScrollView {
                Text(status)
                    .frame(maxWidth: .infinity, alignment: .leading)
                    .padding(.top, 4)
                    .textSelection(.enabled)
            }

            Spacer()
        }
        .padding()
        .frame(minWidth: 600, minHeight: 400)
    }

    // MARK: - Actions

    /// Presents an open panel to choose a `.docx` file and triggers formatting.
    private func openAndFormatDocx() {
        let panel = NSOpenPanel()
        // Prefer a specific UTType if available; otherwise fall back gracefully.
        if let wordType = UTType(filenameExtension: "docx") {
            panel.allowedContentTypes = [wordType]
        } else {
            // If UTType resolution fails for some reason, allow selecting any file
            // and rely on loadDocx throwing a clear error for non-DOCX content.
            panel.allowedContentTypes = []
        }
        panel.allowsMultipleSelection = false
        panel.canChooseDirectories = false

        if panel.runModal() == .OK, let url = panel.url {
            status = "Selected file: \(url.lastPathComponent)\nProcessing…"
            isProcessing = true

            Task {
                await processDocx(at: url)
            }
        } else {
            status = "File selection cancelled."
        }
    }

    /// Loads, formats, and saves the selected DOCX file.
    /// This method runs on the main actor because it updates view state.
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

            // 1. Load DOCX into attributed string from the ORIGINAL file (read-only)
            let mutableDoc = try loadDocx(from: url)

            // 2. (Optional) Call OpenAI here to analyse structure and return
            //    JSON describing headings, table ranges, etc.
            //
            //    let analysis = try await analyseWithOpenAI(text: mutableDoc.string)
            //
            //    For now, we just proceed without it.

            // 3. Apply formatting WITHOUT changing body text content
            applyUKLegalFormatting(to: mutableDoc, header: header, analysis: nil)

            // 4. Save as NEW DOCX next to original so the original is always preserved
            let outputURL = makeOutputURL(from: url)
            try saveDocx(mutableDoc, to: outputURL)

            status = """
            Formatting completed successfully.

            Original file (unchanged):
            \(url.path)

            New formatted copy:
            \(outputURL.path)
            """
        } catch {
            status = "Error: \(error.localizedDescription)"
        }
    }
}
