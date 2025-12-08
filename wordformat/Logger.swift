//
//  Logger.swift
//  wordformat
//
//  Created by Assistant on 08.12.2025.
//

import Foundation

/// Lightweight file logger for verbose diagnostics.
/// Tries multiple writable locations so logs are easy to retrieve:
/// 1) ~/Downloads/Developer/wordformat/logs/wordformat.log  (project workspace, easy for helper to read)
/// 2) ~/Documents/wordformat-logs/wordformat.log
/// 3) ~/Library/Application Support/wordformat-logs/wordformat.log
/// 4) NSTemporaryDirectory()/wordformat-logs/wordformat.log
/// Also echoes to console for quick inspection.
final class Logger {
    static let shared = Logger()
    
    private let queue = DispatchQueue(label: "wordformat.logger.queue", qos: .utility)
    private let logURL: URL
    private let chosenPathDescription: String
    
    private init() {
        var selectedURL: URL?
        var selectedDesc: String = "unknown"
        let fm = FileManager.default
        let home = fm.homeDirectoryForCurrentUser
        
        let candidates: [(URL, String)] = [
            (home.appendingPathComponent("Downloads/Developer/wordformat/logs", isDirectory: true), "Workspace"),
            (home.appendingPathComponent("Documents/wordformat-logs", isDirectory: true), "Documents"),
            (fm.urls(for: .applicationSupportDirectory, in: .userDomainMask).first!.appendingPathComponent("wordformat-logs", isDirectory: true), "Application Support"),
            (URL(fileURLWithPath: NSTemporaryDirectory(), isDirectory: true).appendingPathComponent("wordformat-logs", isDirectory: true), "Temporary")
        ]
        
        for (folder, desc) in candidates {
            do {
                try fm.createDirectory(at: folder, withIntermediateDirectories: true)
                selectedURL = folder.appendingPathComponent("wordformat.log")
                selectedDesc = desc
                break
            } catch {
                // Try next
                continue
            }
        }
        
        // Fallback: in-memory temp log if everything fails (unlikely)
        selectedURL = selectedURL ?? URL(fileURLWithPath: NSTemporaryDirectory()).appendingPathComponent("wordformat.log")
        self.logURL = selectedURL!
        self.chosenPathDescription = selectedDesc
        
        queue.async { [logURL] in
            let banner = "\n=== wordformat run \(Date()) ===\n"
            try? banner.appendLine(to: logURL)
        }
        
        print("[LOGGER] Logging to \(self.logURL.path) (\(self.chosenPathDescription))")
    }
    
    func log(_ message: String, category: String = "GENERAL") {
        queue.async { [logURL] in
            let line = "[\(self.timestamp())][\(category)] \(message)"
            print(line)
            try? line.appendLine(to: logURL)
        }
    }
    
    func log(error: Error, category: String = "ERROR") {
        log("Error: \(error.localizedDescription)", category: category)
    }
    
    private func timestamp() -> String {
        let formatter = DateFormatter()
        formatter.dateFormat = "yyyy-MM-dd HH:mm:ss.SSS"
        return formatter.string(from: Date())
    }
    
    /// Exposes the chosen log path for UI/tooling if needed.
    func currentLogPath() -> String {
        return logURL.path
    }
}

private extension String {
    func appendLine(to url: URL) throws {
        let data = (self + "\n").data(using: .utf8)!
        if FileManager.default.fileExists(atPath: url.path) {
            let handle = try FileHandle(forWritingTo: url)
            defer { try? handle.close() }
            try handle.seekToEnd()
            try handle.write(contentsOf: data)
        } else {
            try data.write(to: url, options: .atomic)
        }
    }
}
