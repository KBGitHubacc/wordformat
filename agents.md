# agents.md - Codebase Guide for AI Assistants

## Project Overview

**wordformat** (also known as **LegalFormatterApp**) is a macOS application built with SwiftUI that automatically formats DOCX documents according to UK legal standards. The application transforms witness statements and legal documents into properly formatted versions with hierarchical numbering, appropriate styling, and correct header metadata.

### Key Features

- Opens and processes .docx files using native macOS APIs
- Applies UK legal formatting standards (Times New Roman, justified text, specific spacing)
- Implements hierarchical numbering: main paragraphs (1, 2, 3), subparagraphs (a, b, c), sub-subparagraphs (i, ii, iii)
- Optional OpenAI integration for intelligent document structure analysis
- Post-processes DOCX XML to inject native Word numbering
- Comprehensive file logging for debugging

### Platform & Requirements

- **Platform**: macOS 26.1+
- **Language**: Swift 6.2
- **Framework**: SwiftUI
- **IDE**: Xcode
- **Development Team**: B7449FV8K6

## Repository Structure

```
wordformat/
├── wordformat/                    # Main application source
│   ├── LegalFormatterApp.swift   # App entry point (@main)
│   ├── ContentView.swift         # Main UI and orchestration
│   ├── DocxHelpers.swift         # DOCX I/O utilities
│   ├── Formatting.swift          # Core formatting engine
│   ├── OpenAIService.swift       # AI analysis integration
│   ├── DocxNumberingPatcher.swift # XML post-processing
│   ├── AnalysisResult.swift      # Data models for AI results
│   ├── Logger.swift              # File logging system
│   └── Assets.xcassets/          # App icons and resources
├── wordformatTests/              # Unit tests
├── wordformatUITests/            # UI tests
├── wordformat.xcodeproj/         # Xcode project configuration
└── logs/                         # Runtime logs and sample files
```

## Core Architecture

### 1. Application Flow

**Entry Point**: `LegalFormatterApp.swift`
- Simple SwiftUI `@main` app structure
- Displays `ContentView` as the main window

**Main UI**: `ContentView.swift` (lines 19-156)
- Accepts user metadata: tribunal name, case reference, applicant, respondent
- Optional OpenAI API key for AI-powered analysis
- File picker for selecting .docx files
- Async processing pipeline with status updates

**Processing Pipeline**:
1. Load DOCX → `loadDocx()` in `DocxHelpers.swift:17`
2. Optional AI Analysis → `OpenAIService.analyseDocument()` in `OpenAIService.swift:213`
3. Apply Formatting → `applyUKLegalFormatting()` in `Formatting.swift:20`
4. Save DOCX → `saveDocx()` in `DocxHelpers.swift:37`
5. Patch XML Numbering → `DocxNumberingPatcher.applyNumbering()` in `DocxNumberingPatcher.swift:34`

### 2. Key Components

#### DocxHelpers.swift
- **loadDocx()**: Reads .docx using `NSAttributedString` with `officeOpenXML` document type
- **saveDocx()**: Writes attributed string back to .docx format
- **makeOutputURL()**: Generates output filename with `_UKLegal` suffix
- **extractParagraphTexts()**: Extracts plain text paragraphs

#### Formatting.swift
Core formatting engine with two main phases:

**Phase 1 - Header Processing** (`styleHeaderPart()`, line 64):
- Centers and bolds specific header elements (tribunal names, case refs, parties)
- Uppercases "WITNESS STATEMENT" titles
- Preserves original text content

**Phase 2 - Body Processing** (`generateDynamicListBody()`, line 109):
- Detects paragraph levels using regex patterns or AI guidance
- Strips manual numbering markers (e.g., "1.", "a)", "i)")
- Applies native `NSTextList` with proper indentation
- Three-level numbering: decimal → lower-alpha → lower-roman
- Preserves inline formatting (bold, italic) from original

**Heuristics**:
- Split point detection: looks for "will say as follows" or "WITNESS STATEMENT"
- Heading detection: regex pattern `^[A-Z0-9]+\\.\\s+[A-Z]` (e.g., "A. INTRODUCTION")
- Level detection via regex patterns for "1.", "(a)", "(i)"

#### OpenAIService.swift
Provides AI-powered document analysis with two approaches:

**Style-Aware Analysis** (`analyseDocument()`, line 213):
- Extracts paragraph metadata: text, isBold, isUppercased, isCentered, wordCount
- Sends chunks of 40 paragraphs to OpenAI for classification
- Returns paragraph types (header, title, intro, heading, body, etc.) and levels (0/1/2)

**Model Selection** (`selectBestModel()`, line 54):
- Queries OpenAI API for available models
- Priority order: gpt-5.1-mini → gpt-5-mini → gpt-5.1 → gpt-5 → gpt-4.1 → gpt-4.1-mini → gpt-4o-mini → gpt-4o
- Filters out non-chat models (codex, whisper, embeddings, tts, dall-e)
- Caches selected model for session

**Classification Categories**:
- `header`: Court/tribunal metadata, case refs, parties
- `title`: Main document title (e.g., "WITNESS STATEMENT OF...")
- `intro`: Introductory paragraph ("I, Name, will say...")
- `heading`: Section headings (e.g., "A. INTRODUCTION")
- `body`: Main numbered narrative paragraphs
- `statementOfTruth`: Statement of truth paragraph
- `signature`: Signature/date lines
- `quote`: Indented or quoted blocks

#### DocxNumberingPatcher.swift
Post-processes saved DOCX files by manipulating the underlying XML:

**Process** (`applyNumbering()`, line 34):
1. Unzip DOCX to temporary directory
2. Create/ensure `word/numbering.xml` exists with 3-level definition
3. Add relationships in `word/_rels/document.xml.rels`
4. Add content-type override in `[Content_Types].xml`
5. Inject `<w:numPr>` elements into `word/document.xml` for target paragraphs
6. Re-zip to original location

**Numbering Template** (line 111):
- Abstract numbering definition with 3 levels
- Level 0 (ilvl=0): decimal format, "1."
- Level 1 (ilvl=1): lower-alpha format, "a)"
- Level 2 (ilvl=2): lower-roman format, "i)"

#### Logger.swift
Comprehensive file logging system:

**Log Locations** (tried in order, line 30):
1. `~/Downloads/Developer/wordformat/logs/wordformat.log` (workspace)
2. `~/Documents/wordformat-logs/wordformat.log`
3. `~/Library/Application Support/wordformat-logs/wordformat.log`
4. `NSTemporaryDirectory()/wordformat-logs/wordformat.log`

**Categories**: GENERAL, UI, DOCX, FORMAT, AI, PATCH, ERROR, LOG

**Usage**: `Logger.shared.log("message", category: "CATEGORY")`

## Development Conventions

### Code Style

1. **File Headers**: Every Swift file includes creation date and author
2. **Documentation**: Use `///` for documentation comments, `//` for inline comments
3. **MARK Comments**: Organize code sections with `// MARK: - Section Name`
4. **Naming**:
   - Functions: camelCase with descriptive verbs (`applyUKLegalFormatting`, `loadDocx`)
   - Private helpers: prefix with `private` modifier
   - Constants: `static let` in enums (e.g., `LegalFormattingDefaults`)

### Swift Patterns

1. **Error Handling**: Throws for file I/O, async for network calls
2. **Optionals**: Guard statements for early returns
3. **Attributes**: Extensive use of `NSAttributedString` for rich text
4. **Concurrency**: `async/await` for AI calls, `@MainActor` for UI updates
5. **Default Parameters**: Use sparingly (e.g., `category: String = "GENERAL"`)

### Key Constants

- **Font Family**: Times New Roman (`LegalFormattingDefaults.fontFamily`)
- **Font Size**: 12pt (`LegalFormattingDefaults.fontSize`)
- **Output Suffix**: `_UKLegal.docx`
- **AI Chunk Size**: 40 paragraphs
- **API Timeout**: 120 seconds for OpenAI calls

## Building & Running

### Xcode Build

1. Open `wordformat.xcodeproj` in Xcode
2. Select target: "wordformat"
3. Build Configuration: Debug or Release
4. Deployment Target: macOS 26.1
5. Build: Cmd+B, Run: Cmd+R

### Sandbox Permissions

The app is sandboxed with the following entitlements (configured in project.pbxproj):
- File access to Downloads folder: read/write
- User selected files: read/write
- Network connections: incoming and outgoing (for OpenAI API)
- Hardened runtime enabled

### Build Settings

- **Swift Version**: 5.0
- **Optimization Level**: Debug: `-Onone`, Release: `-wholemodule`
- **Code Signing**: Automatic, Team B7449FV8K6
- **Product Bundle ID**: `com.kbkh.wordformat`

## Testing

### Unit Tests

Location: `wordformatTests/wordformatTests.swift`
- Run: Cmd+U in Xcode
- Target: wordformatTests.xctest

### UI Tests

Location: `wordformatUITests/`
- Launch tests: `wordformatUITestsLaunchTests.swift`
- Interaction tests: `wordformatUITests.swift`

### Manual Testing

1. Launch app
2. Enter metadata: Tribunal, Case Ref, Applicant, Respondent
3. (Optional) Enter OpenAI API key for AI analysis
4. Click "Open DOCX and Format"
5. Select test file from `logs/` directory
6. Verify output file created with `_UKLegal` suffix
7. Check log file for detailed processing info

## Git Workflow

### Current Branch

Development occurs on feature branches with naming pattern:
`claude/claude-md-mixeabh850n8yy00-01MWtAbJ5mtfJztarcYQkAUu`

### Commit Message Style

Review recent commits for style (from git log):
- `just a quik update`
- `fixes more issues`
- `all changes`
- `fxing more problems`
- `fixis`

**Recommendation**: Use descriptive commit messages in imperative mood:
- Good: "Add AI-powered paragraph classification"
- Good: "Fix numbering patcher XML injection"
- Avoid: Generic messages like "fixes" or "updates"

### Push Protocol

- Always use: `git push -u origin <branch-name>`
- Branch must start with `claude/` and end with session ID
- Retry on network errors with exponential backoff (2s, 4s, 8s, 16s)

## Common Tasks for AI Assistants

### 1. Adding a New Formatting Rule

**Files to modify**:
- `Formatting.swift`: Add logic to `generateDynamicListBody()` or `styleHeaderPart()`
- Consider adding AI classification support in `OpenAIService.swift`

**Example**: To detect and format "Schedule A" sections differently:
1. Add regex pattern detection in `Formatting.swift`
2. Create new paragraph style with custom indentation
3. Add logging for verification
4. Update AI prompt in `OpenAIService.swift` if needed

### 2. Modifying Numbering Styles

**File**: `DocxNumberingPatcher.swift`

**Template Location**: `numberingTemplate()` function (line 111)

**Customization**:
- Change `<w:numFmt w:val="..."/>` for format (decimal, lowerLetter, upperRoman, etc.)
- Modify `<w:lvlText w:val="..."/>` for marker format
- Adjust `<w:ind w:left="..." w:hanging="..."/>` for indentation

### 3. Enhancing AI Analysis

**File**: `OpenAIService.swift`

**Key Functions**:
- `analyseDocument()`: Main entry point
- `sendMetaBatchToAI()`: Batch classification logic
- Modify prompt in line 251-268 for different classifications

**Adding Style Metadata**:
1. Extend `ParagraphMetadata` struct in `AnalysisResult.swift`
2. Update `extractParagraphMetadata()` in `OpenAIService.swift:22`
3. Update AI prompt to consider new metadata

### 4. Debugging Formatting Issues

**Strategy**:
1. Check log file (location printed at app launch)
2. Search for category: "FORMAT" for formatting decisions
3. Review first 25 paragraph classifications (logged with sample text)
4. Check level counts at end of formatting
5. Verify numbering targets built by patcher ("PATCH" category)

**Key Log Messages**:
- `"Split index at N"`: Where header ends and body begins
- `"Para level N heading:X truth:X text: ..."`: Per-paragraph classification
- `"Level counts -> level0:X level1:X..."`: Distribution of paragraph types
- `"Built N numbering targets"`: How many paragraphs will be numbered

### 5. Adding New Document Types

**Current**: Witness statements only

**To add support for other UK legal documents**:
1. Update `LegalParagraphType` enum in `AnalysisResult.swift`
2. Modify split point detection in `Formatting.swift:findBodyStartIndex()`
3. Add new heading/marker patterns to `generateDynamicListBody()`
4. Update AI prompts in `OpenAIService.swift` with document-specific rules
5. Consider adding document type selector in `ContentView.swift`

### 6. Handling Edge Cases

**Empty Paragraphs**: Skipped in formatting loop (line 148-150 in Formatting.swift)

**No AI Key**: Gracefully falls back to heuristic-only mode (line 122-124 in ContentView.swift)

**AI Errors**: Logged but don't halt processing; offline heuristics continue

**Missing Numbering**: Patcher checks if targets exist before applying (line 136-148 in ContentView.swift)

## Troubleshooting

### Issue: Formatting Not Applied

**Check**:
1. Log file shows "FORMAT" entries
2. Split index is correct (should be after header, before main body)
3. Paragraph level detection working (check "Para level" logs)
4. Verify input DOCX has proper structure

**Common Cause**: Split point detection failed, entire document treated as header

**Fix**: Ensure "will say as follows" or "WITNESS STATEMENT" appears in document

### Issue: Numbering Not Appearing

**Check**:
1. "PATCH" category logs show targets were built
2. No patcher errors in log
3. XML injection completed successfully
4. Word can open the output file

**Common Cause**: Sandbox permissions prevent zip/unzip operations

**Fix**: Verify file access permissions in Xcode project settings

### Issue: AI Analysis Fails

**Check**:
1. API key is valid and has credits
2. Network connection works
3. Model selection succeeded (check logs)
4. Timeout is sufficient (currently 120s)

**Common Cause**: Model not available or API key expired

**Fix**: Check OpenAI account status, try fallback model (gpt-4o-mini)

### Issue: App Crashes on Large Documents

**Check**:
1. Memory usage (attributed strings can be large)
2. AI chunking is working (40 paragraphs per batch)
3. Timeout values are appropriate

**Fix**: Increase chunk size or process in smaller sections

## Important Notes for AI Assistants

### What to Preserve

1. **Inline Formatting**: Never strip bold/italic from original document
2. **Text Content**: Only strip automatic numbering markers, keep all other text
3. **NSTextList Objects**: These provide native Word numbering, don't replace with text
4. **Log Statements**: Comprehensive logging is essential for debugging

### What to Avoid

1. **Don't hardcode numbering** in text (e.g., "1. Paragraph text")
2. **Don't modify header content** beyond styling
3. **Don't skip error handling** for file I/O operations
4. **Don't remove sandbox entitlements** without understanding security implications

### Testing Checklist

When making changes:
- [ ] Build succeeds without warnings
- [ ] App runs and opens file picker
- [ ] Processing completes with status message
- [ ] Output file created with correct name
- [ ] Log file contains expected entries
- [ ] Output file opens correctly in Microsoft Word
- [ ] Numbering appears correctly in Word
- [ ] Formatting matches UK legal standards

## API Reference Quick Guide

### NSAttributedString Key Attributes

- `.font`: NSFont (preserve traits with `fontDescriptor.symbolicTraits`)
- `.paragraphStyle`: NSParagraphStyle (alignment, spacing, indents, textLists)
- `.textLists`: Array of NSTextList objects for native numbering

### NSTextList Marker Formats

- `.decimal`: 1, 2, 3
- `.lowercaseAlpha`: a, b, c
- `.uppercaseAlpha`: A, B, C
- `.lowercaseRoman`: i, ii, iii
- `.uppercaseRoman`: I, II, III

### Paragraph Enumeration Pattern

```swift
let ns = attributedString.string as NSString
ns.enumerateSubstrings(in: range, options: .byParagraphs) { substring, range, _, _ in
    // Process paragraph
}
```

### Logger Usage

```swift
Logger.shared.log("Message", category: "CATEGORY")
Logger.shared.log(error: error, category: "ERROR")
Logger.shared.currentLogPath() // Get log file path
```

## Future Enhancement Ideas

1. Support for multiple document types (pleadings, skeleton arguments, etc.)
2. Configurable formatting profiles (different courts/tribunals)
3. Batch processing of multiple files
4. Preview before saving
5. Undo/redo for formatting operations
6. Custom numbering schemes per document section
7. Integration with document templates
8. Export to PDF with proper formatting
9. Offline mode improvements (better heuristics without AI)
10. Style preservation from original document

## Resources

- **OpenXML Spec**: https://learn.microsoft.com/en-us/office/open-xml/
- **NSAttributedString Docs**: Apple Developer Documentation
- **SwiftUI Documentation**: Apple Developer Documentation
- **UK Legal Formatting Standards**: Practice Direction 32 (Witness Statements)

---

**Last Updated**: 2025-12-08
**Maintainer**: Development Team B7449FV8K6
**For Questions**: Check logs/ directory for runtime diagnostics
